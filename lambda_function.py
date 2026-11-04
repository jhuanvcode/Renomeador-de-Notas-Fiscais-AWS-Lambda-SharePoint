# -*- coding: utf-8 -*-
"""
AWS LAMBDA - RENOMEADOR DE NOTAS FISCAIS NO SHAREPOINT
Gatilho: Webhook Microsoft Graph (Lambda Function URL)

Como funciona:
  1. O Power Automate (plano Basic) continua salvando as notas que chegam no Outlook
     direto nas pastas do SharePoint.
  2. Uma subscription do Microsoft Graph avisa esta Lambda (Function URL) quando um
     arquivo novo ou alterado aparece na pasta monitorada.
  3. O Airflow renova a subscription automaticamente antes de expirar (~3 dias).
  4. O Lambda processa apenas o arquivo notificado (PDF, XML ou CSV).
     Se o webhook vier sem id do arquivo (/drive/root), varre dinamicamente
     {ano}/{mês}/{PDF|XML|CSV} sem precisar atualizar env var todo mês.
  5. Identifica quem emitiu a nota (com trava de segurança bloqueando notas da própria empresa).
  6. Para PDFs escaneados (imagens), utiliza OCR via Amazon Textract.
  7. Renomeia o arquivo diretamente no SharePoint via Microsoft Graph API.
  8. DynamoDB evita reprocessar o mesmo arquivo (idempotência).

Ambiente de execução: AWS Lambda (Python 3.12+)
"""

import os
import re
import csv
import json
import zlib
import base64
import hmac
import logging
import urllib.parse
import urllib.request
import urllib.error
import unicodedata
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime, timezone, timedelta
from typing import Optional, List, Dict, Any, Tuple


# ============================================================
# LOG
# ============================================================

logger = logging.getLogger()
logger.setLevel(logging.INFO)


# ============================================================
# CONFIGURAÇÕES E CREDENCIAIS (MICROSOFT GRAPH & AWS)
# ============================================================

# Configure TODOS os valores na aba "Environment variables" do AWS Lambda.
# NUNCA coloque valores reais (credenciais, CNPJ, URLs internas) diretamente no código.

# 1. Credenciais do Azure Entra ID (App Registration):
AZURE_TENANT_ID = os.environ.get("AZURE_TENANT_ID", "")
AZURE_CLIENT_ID = os.environ.get("AZURE_CLIENT_ID", "")
AZURE_CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "")

# 2. Informações da Pasta no SharePoint:
# Exemplo de URL do site: https://SUA-EMPRESA.sharepoint.com/sites/SEU-SITE
SHAREPOINT_SITE_URL = os.environ.get("SHAREPOINT_SITE_URL", "")
# Caminho da pasta das notas dentro da biblioteca de documentos:
# Exemplo: "Notas Fiscais" ou "General/Notas Fiscais"
SHAREPOINT_FOLDER_PATH = os.environ.get("SHAREPOINT_FOLDER_PATH", "")

# Se você já tiver os IDs exatos (Site ID e Drive ID), pode preencher diretamente:
SHAREPOINT_SITE_ID = os.environ.get("SHAREPOINT_SITE_ID", "")
SHAREPOINT_DRIVE_ID = os.environ.get("SHAREPOINT_DRIVE_ID", "")

# 3. Regras de Segurança e Negócio:
# CNPJ da PRÓPRIA empresa (tomadora). Notas emitidas por ela/filiais nunca são renomeadas.
EMPRESA_CNPJ = os.environ.get("EMPRESA_CNPJ", "")
BLOQUEAR_TODAS_AS_FILIAIS = True      # Bloqueia qualquer CNPJ com a mesma raiz (8 primeiros dígitos) da empresa

# Termos do nome da própria empresa para bloqueio quando não há CNPJ legível.
# Formato: grupos separados por "|"; dentro do grupo, termos separados por "+" (todos devem aparecer).
# Exemplo: "minhaempresa+matriz|minhaempresa+filial"  (sem acentos, minúsculo)
EMPRESA_TERMOS_BLOQUEIO = os.environ.get("EMPRESA_TERMOS_BLOQUEIO", "")

NOME_MAX_CARACTERES = 20
SEPARADOR_PALAVRAS = ""                # "" = ABCDANEURO | "_" = ABC_DA_NEURO

# False = Renomeia de verdade no SharePoint | True = Apenas simula nos logs
DRY_RUN = os.environ.get("DRY_RUN", "false").lower() in ("true", "1")

# Documentos internos conhecidos que não são notas fiscais
# Nomes de arquivo (separados por vírgula, em minúsculo) a ignorar sem aviso.
ARQUIVOS_SEM_AVISO = {
    x.strip().lower()
    for x in os.environ.get("ARQUIVOS_IGNORADOS", "").split(",")
    if x.strip()
}

EXTENSOES_VALIDAS = (".pdf", ".xml", ".csv")

# OCR via AWS Textract para notas escaneadas
USAR_AWS_TEXTRACT = os.environ.get("USAR_AWS_TEXTRACT", "true").lower() in ("true", "1")

# Webhook Graph + DynamoDB
GRAPH_CLIENT_STATE = os.environ.get("GRAPH_CLIENT_STATE", "")  # obrigatório: segredo compartilhado com a subscription
DYNAMODB_TABLE = os.environ.get("DYNAMODB_TABLE", "notas-processadas")

# Fallback quando Graph notifica /drive/root (sem id do arquivo)
# Aumentado para 60 minutos para evitar que atrasos no Webhook descartem notas recém-chegadas
WEBHOOK_FALLBACK_MINUTES = int(os.environ.get("WEBHOOK_FALLBACK_MINUTES", "60"))
# Override opcional: pastas fixas (separadas por vírgula). Se vazio, varre ano/mês/tipo dinamicamente.
SHAREPOINT_WATCH_PATHS = os.environ.get("SHAREPOINT_WATCH_PATHS", "")
# Ano(s) para varredura dinâmica (ex: "2026" ou "2025,2026"). Vazio = ano atual.
SHAREPOINT_WATCH_YEARS = os.environ.get("SHAREPOINT_WATCH_YEARS", "")
# Subpastas de tipo dentro de cada mês (padrão: PDF, XML, CSV)
SHAREPOINT_TIPOS_PASTA = os.environ.get("SHAREPOINT_TIPOS_PASTA", "PDF,XML,CSV")

# Imports de pacotes opcionais
try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import boto3
except ImportError:
    boto3 = None

_dynamodb_table = None


def _tabela_dynamodb():
    global _dynamodb_table
    if _dynamodb_table is None and boto3 is not None and DYNAMODB_TABLE:
        _dynamodb_table = boto3.resource("dynamodb").Table(DYNAMODB_TABLE)
    return _dynamodb_table


# ============================================================
# NORMALIZAÇÃO DE TEXTO
# ============================================================

def remover_acentos(texto) -> str:
    if not texto:
        return ""
    return "".join(
        c for c in unicodedata.normalize("NFD", str(texto))
        if unicodedata.category(c) != "Mn"
    )


def normalizar_texto(texto) -> str:
    texto = remover_acentos(texto).lower()
    return re.sub(r"\s+", " ", texto).strip()


# ============================================================
# CNPJ/CPF (Validação e Bloqueio da própria empresa)
# ============================================================

_PESOS_1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
_PESOS_2 = [6] + _PESOS_1


def _dv(base: str, pesos) -> int:
    soma = sum((ord(c) - 48) * p for c, p in zip(base, pesos))
    r = soma % 11
    return 0 if r < 2 else 11 - r


def cnpj_valido(c: str) -> bool:
    c = c.upper()
    if not re.fullmatch(r"[0-9A-Z]{12}\d{2}", c) or len(set(c)) == 1:
        return False
    d1 = _dv(c[:12], _PESOS_1)
    d2 = _dv(c[:12] + str(d1), _PESOS_2)
    return c[12:] == f"{d1}{d2}"


def cpf_valido(c: str) -> bool:
    if not re.fullmatch(r"\d{11}", c) or len(set(c)) == 1:
        return False
    for n in (9, 10):
        soma = sum(int(c[i]) * (n + 1 - i) for i in range(n))
        if ((soma * 10) % 11) % 10 != int(c[n]):
            return False
    return True


def limpar_documento(bruto) -> Optional[str]:
    if not bruto:
        return None
    s = re.sub(r"[^0-9A-Za-z]", "", str(bruto)).upper()
    if len(s) == 14 and cnpj_valido(s):
        return s
    if len(s) == 11 and cpf_valido(s):
        return s
    return None


EMPRESA_DOC = limpar_documento(EMPRESA_CNPJ) or ""


def doc_bloqueado(doc: Optional[str]) -> bool:
    if not doc or not EMPRESA_DOC:
        return False
    if doc == EMPRESA_DOC:
        return True
    return BLOQUEAR_TODAS_AS_FILIAIS and len(doc) == 14 and doc[:8] == EMPRESA_DOC[:8]


_S = r"\s*"
RE_CNPJ_FMT = re.compile(
    r"(?<![0-9A-Za-z])[0-9A-Z]{2}" + _S + r"\." + _S + r"[0-9A-Z]{3}" + _S + r"\." + _S
    + r"[0-9A-Z]{3}" + _S + r"/" + _S + r"[0-9A-Z]{4}" + _S + r"-" + _S + r"\d{2}(?![0-9A-Za-z])"
)
RE_CPF_FMT = re.compile(
    r"(?<!\d)\d{3}" + _S + r"\." + _S + r"\d{3}" + _S + r"\." + _S + r"\d{3}" + _S
    + r"-" + _S + r"\d{2}(?!\d)"
)
RE_BRUTO = re.compile(r"(?<!\d)(?:\d{14}|\d{11})(?!\d)")


def documentos_no_texto(texto: str) -> List[str]:
    def achar(regexes):
        achados = []
        for rx in regexes:
            for m in rx.finditer(texto or ""):
                d = limpar_documento(m.group())
                if d:
                    achados.append((m.start(), d))
        achados.sort()
        saida = []
        for _, d in achados:
            if d not in saida:
                saida.append(d)
        return saida
    return achar([RE_CNPJ_FMT, RE_CPF_FMT]) or achar([RE_BRUTO])


# ============================================================
# LIMPEZA DO NOME DO FORNECEDOR (FILTROS ANTI-LIXO)
# ============================================================

class PrestadorEhEmpresaPropria(Exception):
    """Nota emitida pela própria empresa (tomadora)."""


PRIMEIRA_PALAVRA_INVALIDA = {
    "rua", "r", "avenida", "av", "travessa", "tv", "alameda", "praca", "estrada",
    "rodovia", "largo", "beco", "viela", "logradouro", "bairro", "cidade",
    "estado", "municipio", "cep", "endereco", "numero", "complemento", "quadra",
    "telefone", "fone", "email", "e-mail", "inscricao", "indicador", "cnpj",
    "cpf", "nif", "nome", "razao", "prestador", "fornecedor", "tomador",
    "emitente", "destinatario", "adquirente", "empresa", "optante", "simples",
    "regime", "codigo", "sigla", "uf", "data", "valor", "local",
    "municipal", "estadual", "federal", "nota", "fiscal", "nfse", "danfse",
    "nfs", "nf", "danfe", "dps", "documento", "servico", "servicos",
    "empresarial", "social",
}
PALAVRA_INVALIDA_EM_QUALQUER_LUGAR = {"bairro", "cep", "complemento", "logradouro", "endereco", "danfse", "nfse", "empresarial", "social"}
RE_TELEFONE = re.compile(r"\(?\d{2}\)?\s*9?\d{4}-?\d{4}")


def _eh_empresa_propria_bruto(s: str) -> bool:
    """Bloqueio por nome, configurado via EMPRESA_TERMOS_BLOQUEIO (sem termos fixos no código)."""
    n = normalizar_texto(s).replace(" ", "")
    for grupo in EMPRESA_TERMOS_BLOQUEIO.split("|"):
        termos = [t.strip().lower() for t in grupo.split("+") if t.strip()]
        if termos and all(t in n for t in termos):
            return True
    return False


def limpar_nome(raw, confiavel=False) -> Optional[str]:
    if not raw:
        return None
    s = re.sub(r"\s+", " ", remover_acentos(raw)).strip(" :;-|")
    if not s or "@" in s:
        return None

    # MEI: "12.345.678 FULANO DE TAL", "12,345,678 FULANO", "12345678 FULANO"
    s = re.sub(r"^\d{2}[\.,]?\d{3}[\.,]?\d{3}\s*", "", s).strip()
    s = re.sub(r"^\d{8}\s*", "", s).strip()
    s = re.sub(r"\s*\d{11}\s*$", "", s).strip()

    if _eh_empresa_propria_bruto(s):
        raise PrestadorEhEmpresaPropria(s)

    palavras = re.findall(r"[A-Za-z0-9\-\.]+", s.lower())
    if not palavras:
        return None

    if not confiavel:
        primeira = palavras[0].strip(".-")
        if primeira in PRIMEIRA_PALAVRA_INVALIDA:
            return None
        if any(p.strip(".,-") in PALAVRA_INVALIDA_EM_QUALQUER_LUGAR for p in palavras):
            return None

    if re.search(r",\s*\d+", s):
        return None
    if documentos_no_texto(s):
        return None

    limpo = re.sub(r'[\\/:*?"<>|]', "", s)
    limpo = re.sub(r"[^A-Za-z0-9 ]", " ", limpo)
    limpo = re.sub(r"\s+", " ", limpo).strip()
    compacto = limpo.replace(" ", "")
    if len(compacto) < 3:
        return None
    if sum(c.isdigit() for c in compacto) > len(compacto) * 0.30:
        return None

    nome = limpo.upper().replace(" ", SEPARADOR_PALAVRAS)
    return nome[:NOME_MAX_CARACTERES].strip("_")


# ============================================================
# EXTRAÇÃO - PDF (GEOMETRIA DE COLUNAS + OCR TEXTRACT)
# ============================================================

def _titulo_prestador(linha: str) -> bool:
    return bool(re.search(r"\b(prestador|fornecedor|emitente)\b", normalizar_texto(linha)))


def _titulo_tomador(linha: str) -> bool:
    t = normalizar_texto(linha)
    return bool(re.search(r"\b(tomador|adquirente|destinatario)\b", t)) or "destinatario da operacao" in t


def secoes_com_plano_b(textos: List[str]):
    secoes = []
    for i, t in enumerate(textos):
        if _titulo_prestador(t):
            fim = len(textos)
            for j in range(i + 1, len(textos)):
                if _titulo_tomador(textos[j]):
                    fim = j
                    break
            secoes.append((i, fim))
    if secoes:
        return secoes
    fim_pre = next((j for j, t in enumerate(textos) if _titulo_tomador(t)), None)
    if fim_pre:
        return [(0, fim_pre)]
    return []


def agrupar_linhas(words, tol=5):
    linhas = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if linhas and abs(w["top"] - linhas[-1]["top"]) <= tol:
            linhas[-1]["words"].append(w)
        else:
            linhas.append({"top": w["top"], "bottom": w["bottom"], "words": [w]})
    for l in linhas:
        l["words"].sort(key=lambda w: w["x0"])
        l["bottom"] = max(w["bottom"] for w in l["words"])
        l["texto"] = " ".join(w["text"] for w in l["words"])
    return linhas


def celulas(linha, gap=12):
    cs = []
    for w in linha["words"]:
        if cs and w["x0"] - cs[-1]["x1"] <= gap:
            cs[-1]["text"] += " " + w["text"]
            cs[-1]["x1"] = w["x1"]
        else:
            cs.append({"text": w["text"], "x0": w["x0"], "x1": w["x1"]})
    return cs


RE_ROTULO_NOME = re.compile(
    r"^(?:n\s*ome\s*[/|1iI\\-]?\s*nome\s*empresarial|"
    r"n\s*ome\s*[/|1iI\\-]?\s*razao\s*social|"
    r"nome\s*empresarial|"
    r"razao\s*social|"
    r"nome\s*do\s*prestador|"
    r"nome\s*do\s*fornecedor|"
    r"nome\s*do\s*emitente|"
    r"n\s*ome)"
    r"\s*(?::\s*(.*))?$",
    re.IGNORECASE,
)

RE_NOME_MESMA_LINHA = re.compile(
    r"^\s*n\s*ome\s*[/|1iI\\-]?\s*(?:razao\s*social|nome\s*empresarial)\s*:\s*(.+)$",
    re.IGNORECASE,
)

RE_ROTULO_SOZINHO = re.compile(r"^(?:emitente|prestador|fornecedor)\s*:\s*$", re.IGNORECASE)
RE_PREFIXO_DOC = re.compile(
    r"^\s*(?:\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{3}\.\d{3}\.\d{3}-\d{2})\s*[-–]\s*"
)


def _sem_ruido(linha_texto: str) -> str:
    toks = [t for t in linha_texto.split() if "@" not in t]
    return RE_TELEFONE.sub(" ", " ".join(toks))


def nome_na_secao(linhas, ini, fim):
    for i in range(ini, fim):
        m_linha = RE_NOME_MESMA_LINHA.match(remover_acentos(linhas[i]["texto"]))
        if m_linha:
            n = limpar_nome(m_linha.group(1), confiavel=True)
            if n:
                return n, "mesma linha do rótulo"

        if RE_ROTULO_SOZINHO.match(remover_acentos(linhas[i]["texto"]).strip()):
            valor = " ".join(linhas[k]["texto"] for k in range(i + 1, min(i + 3, fim)))
            valor = RE_PREFIXO_DOC.sub("", valor).strip()
            n = limpar_nome(valor, confiavel=True)
            if n:
                return n, "valor abaixo de Emitente"

        if "nota fiscal de servicos eletronica" in normalizar_texto(linhas[i]["texto"]):
            for k in range(i + 1, min(i + 3, fim)):
                n = limpar_nome(_sem_ruido(linhas[k]["texto"]))
                if n:
                    return n, "abaixo do título NFS-e (sem rótulo)"

        cs = celulas(linhas[i])
        for c, cel in enumerate(cs):
            m = RE_ROTULO_NOME.match(remover_acentos(cel["text"]).strip())
            if not m:
                continue
            if m.group(1):
                n = limpar_nome(m.group(1), confiavel=True)
                if n:
                    return n, "mesma célula do rótulo"
                continue

            xl = cel["x0"] - 15
            xr = cs[c + 1]["x0"] - 5 if c + 1 < len(cs) else float("inf")
            for k in range(i + 1, min(i + 4, fim)):
                if linhas[k]["top"] - linhas[i]["bottom"] > 60:
                    break
                raw = " ".join(w["text"] for w in linhas[k]["words"] if xl <= w["x0"] < xr)
                n = limpar_nome(raw, confiavel=True)
                if n:
                    return n, "abaixo do rótulo (mesma coluna)"

            for k in range(i + 1, min(i + 4, fim)):
                n = limpar_nome(_sem_ruido(linhas[k]["texto"]), confiavel=True)
                if n:
                    return n, "linha abaixo do rótulo (plano B)"

    # Fallback na seção do prestador
    for i in range(ini, fim):
        n = limpar_nome(_sem_ruido(linhas[i]["texto"]), confiavel=False)
        if n:
            return n, "linha da seção prestador (fallback)"

    return None, ""


@dataclass
class Resultado:
    nome: Optional[str] = None
    metodo: str = ""
    motivo: str = ""


def extrair_via_aws_textract(bytes_arquivo: bytes) -> List[dict]:
    """OCR nativo em nuvem via Amazon Textract (apenas imagens JPEG/PNG/TIFF)."""
    if not USAR_AWS_TEXTRACT or boto3 is None:
        return []
    # Textract detect_document_text não aceita bytes de PDF diretamente (requer S3 ou imagens)
    if bytes_arquivo.startswith(b"%PDF"):
        return []
    try:
        client = boto3.client("textract")
        response = client.detect_document_text(Document={"Bytes": bytes_arquivo})
        words = []
        for block in response.get("Blocks", []):
            if block.get("BlockType") == "WORD":
                geom = block.get("Geometry", {}).get("BoundingBox", {})
                left = geom.get("Left", 0) * 1000
                top = geom.get("Top", 0) * 1000
                width = geom.get("Width", 0) * 1000
                height = geom.get("Height", 0) * 1000
                words.append({
                    "text": block.get("Text", ""),
                    "x0": left,
                    "x1": left + width,
                    "top": top,
                    "bottom": top + height,
                })
        return words
    except Exception as e:
        logger.warning(f"Falha na chamada do Amazon Textract: {e}")
        return []


def extrair_tokens_pdf_puro(dados_pdf: bytes) -> List[str]:
    """
    Extrai tokens e strings de texto diretamente dos streams do PDF usando zlib padrão.
    Não requer pdfplumber nem dependências externas em C (100% nativo Python).
    """
    tokens = []
    pos = 0
    while True:
        idx = dados_pdf.find(b"stream", pos)
        if idx == -1:
            break
        end_idx = dados_pdf.find(b"endstream", idx)
        if end_idx == -1:
            break

        start_content = idx + 6
        if dados_pdf[start_content : start_content + 2] == b"\r\n":
            start_content += 2
        elif dados_pdf[start_content : start_content + 1] in (b"\n", b"\r"):
            start_content += 1

        stream_raw = dados_pdf[start_content:end_idx]

        decomp = None
        try:
            decomp = zlib.decompress(stream_raw)
        except Exception:
            try:
                d = zlib.decompressobj()
                decomp = d.decompress(stream_raw)
            except Exception:
                pass

        conteudo = decomp if decomp else stream_raw

        for m in re.finditer(rb"\((.*?)\)\s*(?:Tj|\')|\[(.*?)\]\s*TJ", conteudo, re.DOTALL):
            if m.group(1):
                t = m.group(1).decode("latin1", errors="ignore")
                t = t.replace(r"\(", "(").replace(r"\)", ")").replace(r"\\", "\\")
                if t.strip():
                    tokens.append(t.strip())
            elif m.group(2):
                pedacos = re.findall(rb"\((.*?)\)", m.group(2), re.DOTALL)
                t = "".join(p.decode("latin1", errors="ignore") for p in pedacos)
                t = t.replace(r"\(", "(").replace(r"\)", ")").replace(r"\\", "\\")
                if t.strip():
                    tokens.append(t.strip())

        pos = end_idx + 9
    return tokens


def extrair_prestador_de_tokens_puro(tokens: List[str]) -> Tuple[Optional[str], str]:
    """
    Localiza a seção do prestador/emitente nos tokens extraídos do PDF
    e identifica com segurança o nome da empresa/prestador.
    """
    if not tokens:
        return None, ""

    idx_prest = -1
    idx_toma = len(tokens)

    for i, tok in enumerate(tokens):
        norm = normalizar_texto(tok)
        if any(k in norm for k in ["prestador", "fornecedor", "emitente"]):
            if idx_prest == -1:
                idx_prest = i
        if any(k in norm for k in ["tomador", "destinatario", "adquirente"]) and idx_prest != -1:
            idx_toma = i
            break

    secao = tokens[idx_prest:idx_toma] if idx_prest != -1 else tokens[:80]
    texto_secao = "\n".join(secao)

    # 1. Checagem de segurança (própria empresa) na seção do prestador
    docs = documentos_no_texto(texto_secao)
    if docs and doc_bloqueado(docs[0]):
        raise PrestadorEhEmpresaPropria("prestador bloqueado (própria empresa)")

    # 2. Estratégia 1: Mesma string contém rótulo e nome (ex: "Nome/Razão Social: EMPRESA EXEMPLO LTDA")
    for tok in secao:
        m = re.search(r"(?:nome|razao\s*social|nome\s*empresarial)\s*:\s*(.+)", tok, re.IGNORECASE)
        if m:
            n = limpar_nome(m.group(1), confiavel=True)
            if n:
                return n, "rótulo e nome na mesma linha (parser nativo)"

    # 3. Estratégia 2: Rótulo de nome detectado seguido pelos tokens de valor (ex: valor quebrado em vários tokens)
    PALAVRAS_ROTULO = {"nome", "razao", "social", "empresarial", "do", "prestador", "fornecedor", "emitente", "/", "-", ":"}
    for i, tok in enumerate(secao):
        norm = normalizar_texto(tok)
        if norm in ("nome", "razao social", "nome / razao social", "nome empresarial", "nome / nome empresarial"):
            acumulado = []
            for k in range(i + 1, min(i + 15, len(secao))):
                item = secao[k]
                item_norm = normalizar_texto(item)
                if item_norm in PALAVRAS_ROTULO:
                    continue
                if item_norm in ("e-", "email", "e-mail") or "mail" in item_norm or "@" in item or any(div in item_norm for div in ["endereco", "municipio", "telefone", "inscricao", "simples", "tomador", "optante", "regime"]):
                    break
                if re.search(r"^\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", item) or re.search(r"^\d{11,14}$", item):
                    continue
                acumulado.append(item)
            if acumulado:
                tentativa = " ".join(acumulado)
                n = limpar_nome(tentativa, confiavel=False)
                if n:
                    return n, "valor após rótulo (parser nativo)"

    # 4. Estratégia 3: Varredura por candidato válido na seção
    for tok in secao:
        n = limpar_nome(tok, confiavel=False)
        if n:
            return n, "candidato na seção prestador (parser nativo)"

    return None, ""


def extrair_nome_pdf_bytes(conteudo_pdf: bytes) -> Resultado:
    res = Resultado()
    import io

    # 1. Tentativa via pdfplumber (se instalado no ambiente/layer)
    if pdfplumber is not None:
        try:
            with pdfplumber.open(io.BytesIO(conteudo_pdf)) as pdf:
                for numero, pg in enumerate(pdf.pages[:1], 1):
                    words = pg.extract_words(x_tolerance=2, y_tolerance=2)
                    if not words:
                        continue
                    linhas = agrupar_linhas(words, tol=3)
                    textos = [l["texto"] for l in linhas]

                    for ini, fim in secoes_com_plano_b(textos):
                        docs = documentos_no_texto("\n".join(textos[ini:fim]))
                        if docs and doc_bloqueado(docs[0]):
                            res.motivo = "nota emitida pela própria empresa (prestador bloqueado)"
                            return res
                        try:
                            nome, como = nome_na_secao(linhas, ini, fim)
                        except PrestadorEhEmpresaPropria:
                            res.motivo = "nota emitida pela própria empresa (prestador bloqueado)"
                            return res
                        if nome:
                            res.nome, res.metodo = nome, f"{como} (pdfplumber)"
                            return res
        except Exception as e:
            logger.info(f"pdfplumber não processou: {e}")

    # 2. Tentativa via Parser Nativo Python Puro (zlib + re, 100% nativo para AWS Lambda sem layers)
    if not res.nome:
        try:
            tokens_pdf = extrair_tokens_pdf_puro(conteudo_pdf)
            if tokens_pdf:
                try:
                    nome_puro, metodo_puro = extrair_prestador_de_tokens_puro(tokens_pdf)
                    if nome_puro:
                        res.nome, res.metodo = nome_puro, metodo_puro
                        return res
                except PrestadorEhEmpresaPropria:
                    res.motivo = "nota emitida pela própria empresa (prestador bloqueado)"
                    return res
        except Exception as e:
            logger.info(f"Parser nativo de PDF: {e}")

    # 3. Fallback via Amazon Textract (OCR para imagens)
    if not res.nome:
        words_ocr = extrair_via_aws_textract(conteudo_pdf)
        if words_ocr:
            linhas_ocr = agrupar_linhas(words_ocr, tol=8)
            textos_ocr = [l["texto"] for l in linhas_ocr]
            for ini, fim in secoes_com_plano_b(textos_ocr):
                docs = documentos_no_texto("\n".join(textos_ocr[ini:fim]))
                if docs and doc_bloqueado(docs[0]):
                    res.motivo = "nota emitida pela própria empresa (prestador bloqueado)"
                    return res
                try:
                    nome, como = nome_na_secao(linhas_ocr, ini, fim)
                except PrestadorEhEmpresaPropria:
                    res.motivo = "nota emitida pela própria empresa (prestador bloqueado)"
                    return res
                if nome:
                    res.nome, res.metodo = nome, f"{como} (AWS Textract OCR)"
                    return res

    if not res.nome and not res.motivo:
        res.motivo = "nome do fornecedor não encontrado com segurança"
    return res


# ============================================================
# EXTRAÇÃO - XML (ABRASF, layouts municipais, Tabelas, FastReport)
# ============================================================

def tag_local(tag) -> str:
    return (tag or "").split("}")[-1].lower()


BLOCOS_PRESTADOR = [
    "emit", "emitente", "emissor",
    "prestadorservico", "prestadorservicos", "identificacaoprestador",
    "prest", "prestador", "fornecedor",
    "dadosprestador", "dados_prestador", "infprestador",
    "dadosemitente", "dados_emitente", "infemitente",
]
TAGS_NOME = [
    "xnome", "razaosocial", "nomeempresarial", "nomerazaosocial",
    "nome", "xfant", "nomefantasia",
]
TAGS_DOC = {"cnpj", "cpf", "cnpjcpf", "cpfcnpj", "documento"}
TAGS_TOMADOR = {
    "toma", "tomador", "tomadorservico", "dest", "destinatario",
    "adquirente", "interm", "intermediario", "intermediarioservico",
    "dadostomador", "dados_tomador", "inftomador",
    "dadosdestinatario", "dados_destinatario", "infdestinatario",
}


def iter_sem_tomador(el):
    yield el
    for filho in el:
        if tag_local(filho.tag) in TAGS_TOMADOR:
            continue
        yield from iter_sem_tomador(filho)


def _nome_no_elemento(el):
    achados = {}
    for x in iter_sem_tomador(el):
        t = tag_local(x.tag)
        if t in TAGS_NOME and x.text and x.text.strip() and t not in achados:
            achados[t] = x.text
    for t in TAGS_NOME:
        if t in achados:
            n = limpar_nome(achados[t], confiavel=True)
            if n:
                return n
    return None


def extrair_nome_xml_bytes(conteudo_xml: bytes) -> Resultado:
    res = Resultado()
    try:
        root = ET.fromstring(conteudo_xml)
    except Exception as e:
        res.motivo = f"XML inválido: {e}"
        return res

    pai = {filho: p for p in root.iter() for filho in p}

    # 1. Padrão NFS-e / NF-e
    try:
        for tag_bloco in BLOCOS_PRESTADOR:
            for bloco in root.iter():
                if tag_local(bloco.tag) != tag_bloco:
                    continue
                for x in iter_sem_tomador(bloco):
                    if tag_local(x.tag) in TAGS_DOC and x.text:
                        d = limpar_documento(x.text)
                        if d and doc_bloqueado(d):
                            res.motivo = "nota emitida pela própria empresa (emitente bloqueado)"
                            return res

                nome = _nome_no_elemento(bloco)
                como = f"bloco <{tag_bloco}>"
                if not nome and bloco in pai:
                    nome = _nome_no_elemento(pai[bloco])
                    como = f"vizinho de <{tag_bloco}>"
                if nome:
                    res.nome, res.metodo = nome, como
                    return res
    except PrestadorEhEmpresaPropria:
        res.motivo = "nota emitida pela própria empresa"
        return res

    # 2. Caso XML em formato de tabela
    try:
        for table in root.iter():
            if tag_local(table.tag) == "table":
                rows = [r for r in table if tag_local(r.tag) == "row"]
                for r_idx, r in enumerate(rows):
                    cells = [c.text.strip() for c in r if c.text and c.text.strip()]
                    if any(re.search(r"\b(tomador|adquirente|destinatario)\b", normalizar_texto(c)) for c in cells):
                        break
                    for c_idx, text in enumerate(cells):
                        if re.match(r"^n\s*ome\s*[/|1iI\\-]?\s*(?:nome\s*empresarial|razao\s*social)", normalizar_texto(text)):
                            if r_idx + 1 < len(rows):
                                next_cells = [c.text.strip() for c in rows[r_idx + 1] if c.text and c.text.strip()]
                                if c_idx < len(next_cells):
                                    n = limpar_nome(next_cells[c_idx], confiavel=True)
                                    if n:
                                        res.nome, res.metodo = n, "tabela XML"
                                        return res
    except PrestadorEhEmpresaPropria:
        res.motivo = "nota emitida pela própria empresa"
        return res

    # 3. Caso FastReport APS XML (<Text>)
    try:
        textos = [
            (el.text or "").strip()
            for el in root.iter()
            if (el.text or "").strip() and tag_local(el.tag) not in ("tag", "brush", "size", "point", "font")
        ]
        full_str = " ".join(textos)
        idx_tomador = -1
        for m in re.finditer(r"\b(tomador|adquirente|destinatario)\b", full_str, re.IGNORECASE):
            idx_tomador = m.start()
            break
        secao_prest = full_str[:idx_tomador] if idx_tomador != -1 else full_str
        m_rotulo = re.search(
            r"nome\s*[/|1iI\\-]?\s*(?:nome\s*empresarial|razao\s*social)\s*(?:[0-9\.\,\-\s]{8,20})?\s*([A-Za-z0-9\s]{3,60}?)(?=\s*(?:e\s*-\s*mail|endereco|municipio|simples|cpf|cnpj|telefone|$))",
            secao_prest,
            re.IGNORECASE,
        )
        if m_rotulo:
            n = limpar_nome(m_rotulo.group(1), confiavel=True)
            if n:
                res.nome, res.metodo = n, "layout XML (<Text>)"
                return res
    except PrestadorEhEmpresaPropria:
        res.motivo = "nota emitida pela própria empresa"
        return res

    res.motivo = "nome do emitente não encontrado no XML"
    return res


# ============================================================
# EXTRAÇÃO - CSV
# ============================================================

def _coluna_nome_prestador(cabecalho: str) -> bool:
    n = normalizar_texto(cabecalho)
    if any(x in n for x in ("tomador", "adquirente", "destinatario", "intermediario")):
        return False
    if "cnpj" in n or "cpf" in n:
        return False
    if n in ("prestador", "fornecedor", "emitente"):
        return True
    return any(x in n for x in ("prestador", "fornecedor", "emitente")) and (
        "razao" in n or "nome" in n
    )


def extrair_nome_csv_bytes(conteudo_csv: bytes) -> Resultado:
    res = Resultado()
    texto = None
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            texto = conteudo_csv.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if not texto:
        res.motivo = "CSV ilegível"
        return res

    linhas = [l for l in csv.reader(texto.splitlines(), delimiter=";") if l]
    if len(linhas) < 2:
        res.motivo = "CSV vazio"
        return res

    if len(linhas) == 2 and linhas[1] and str(linhas[1][0]).strip().lower() == "total":
        res.motivo = "relatório de NFTS sem movimentação (total 0)"
        return res

    for i, linha in enumerate(linhas[:30]):
        indices = [k for k, c in enumerate(linha) if _coluna_nome_prestador(c)]
        if not indices:
            continue
        nomes = []
        try:
            for dados in linhas[i + 1:]:
                for k in indices:
                    if k < len(dados):
                        n = limpar_nome(dados[k])
                        if n and n not in nomes:
                            nomes.append(n)
        except PrestadorEhEmpresaPropria:
            res.motivo = "nota emitida pela própria empresa"
            return res

        if len(nomes) == 1:
            res.nome, res.metodo = nomes[0], "coluna do prestador no CSV"
        elif len(nomes) > 1:
            res.motivo = f"CSV com múltiplos prestadores ({len(nomes)})"
        else:
            res.motivo = "coluna do prestador sem nome válido"
        return res

    res.motivo = "coluna do prestador não encontrada"
    return res


# ============================================================
# CLIENTE MICROSOFT GRAPH API (SHAREPOINT)
# ============================================================

def obter_token_microsoft_graph() -> str:
    """Obtém o token de acesso OAuth2 usando o App Registration do Azure."""
    if not AZURE_TENANT_ID or not AZURE_CLIENT_ID or not AZURE_CLIENT_SECRET or "SEU_" in AZURE_TENANT_ID:
        raise ValueError("Credenciais do Azure AD / Microsoft Graph não configuradas no Lambda.")

    url = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}/oauth2/v2.0/token"
    dados = urllib.parse.urlencode({
        "client_id": AZURE_CLIENT_ID,
        "client_secret": AZURE_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }).encode("utf-8")

    req = urllib.request.Request(url, data=dados, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")

    with urllib.request.urlopen(req, timeout=15) as resp:
        resultado = json.loads(resp.read().decode("utf-8"))
        return resultado["access_token"]


def resolver_site_e_drive_id(token: str) -> Tuple[str, str]:
    """Resolve automaticamente o Site ID e Drive ID a partir da URL do SharePoint."""
    global SHAREPOINT_SITE_ID, SHAREPOINT_DRIVE_ID
    if SHAREPOINT_SITE_ID and SHAREPOINT_DRIVE_ID:
        return SHAREPOINT_SITE_ID, SHAREPOINT_DRIVE_ID

    if not SHAREPOINT_SITE_ID and not SHAREPOINT_SITE_URL:
        raise ValueError("Configure SHAREPOINT_SITE_URL (ou SHAREPOINT_SITE_ID) nas variáveis de ambiente.")

    parsed = urllib.parse.urlparse(SHAREPOINT_SITE_URL)
    host = parsed.netloc
    site_path = parsed.path.rstrip("/")

    # 1. Busca o site pelo hostname e caminho
    if not SHAREPOINT_SITE_ID:
        url_site = f"https://graph.microsoft.com/v1.0/sites/{host}:{site_path}"
        req = urllib.request.Request(url_site, method="GET")
        req.add_header("Authorization", f"Bearer {token}")
        with urllib.request.urlopen(req, timeout=15) as resp:
            site_data = json.loads(resp.read().decode("utf-8"))
            SHAREPOINT_SITE_ID = site_data["id"]

    # 2. Busca o Drive (Biblioteca de Documentos)
    if not SHAREPOINT_DRIVE_ID:
        nome_biblioteca_alvo = SHAREPOINT_FOLDER_PATH.strip("/").split("/")[0].lower()
        url_drives = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drives"
        req_drives = urllib.request.Request(url_drives, method="GET")
        req_drives.add_header("Authorization", f"Bearer {token}")
        try:
            with urllib.request.urlopen(req_drives, timeout=15) as resp:
                drives_data = json.loads(resp.read().decode("utf-8")).get("value", [])
                for d in drives_data:
                    d_nome = d.get("name", "").lower()
                    if nome_biblioteca_alvo and d_nome == nome_biblioteca_alvo:
                        SHAREPOINT_DRIVE_ID = d["id"]
                        logger.info(f"Drive resolvido pela biblioteca '{d.get('name')}': {SHAREPOINT_DRIVE_ID}")
                        break
        except Exception as e:
            logger.warning(f"Não foi possível listar /sites/drives: {e}")

        # Se não encontrou biblioteca específica, obtém a biblioteca padrão
        if not SHAREPOINT_DRIVE_ID:
            url_drive = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drive"
            req_drive = urllib.request.Request(url_drive, method="GET")
            req_drive.add_header("Authorization", f"Bearer {token}")
            with urllib.request.urlopen(req_drive, timeout=15) as resp:
                drive_data = json.loads(resp.read().decode("utf-8"))
                SHAREPOINT_DRIVE_ID = drive_data["id"]
                logger.info(f"Drive resolvido pela biblioteca padrão: {SHAREPOINT_DRIVE_ID}")

    logger.info(f"Site ID: {SHAREPOINT_SITE_ID} | Drive ID: {SHAREPOINT_DRIVE_ID}")
    return SHAREPOINT_SITE_ID, SHAREPOINT_DRIVE_ID


def baixar_conteudo_sharepoint(drive_id: str, item_id: str, token: str) -> bytes:
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    req = urllib.request.Request(url, method="GET")
    req.add_header("Authorization", f"Bearer {token}")
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read()


def listar_itens_pasta_sharepoint(drive_id: str, item_id_ou_path: str, token: str, por_caminho: bool = False) -> List[dict]:
    """Lista os arquivos e subpastas de uma pasta no SharePoint com suporte a paginação."""
    if por_caminho:
        caminho_enc = urllib.parse.quote(item_id_ou_path.strip("/"))
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{caminho_enc}:/children"
    else:
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id_ou_path}/children"

    itens_totais = []
    while url:
        try:
            dados = _graph_get_json(url, token)
            itens_totais.extend(dados.get("value", []))
            url = dados.get("@odata.nextLink")
        except Exception as e:
            logger.error(f"Erro ao listar pasta '{item_id_ou_path}': {e}")
            break
    return itens_totais


def renomear_arquivo_sharepoint(drive_id: str, item_id: str, novo_nome: str, token: str) -> dict:
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
    payload = json.dumps({"name": novo_nome}).encode("utf-8")
    req = urllib.request.Request(url, data=payload, method="PATCH")
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Content-Type", "application/json")
    with urllib.request.urlopen(req, timeout=15) as resp:
        return json.loads(resp.read().decode("utf-8"))


def calcular_destino_sem_colisao(nome_base: str, ext: str, nomes_existentes: List[str], nome_atual: str) -> Optional[str]:
    n = 1
    while True:
        candidato = f"{nome_base}{ext}" if n == 1 else f"{nome_base}_{n}{ext}"
        if candidato.lower() == nome_atual.lower():
            return None
        if candidato.lower() not in [x.lower() for x in nomes_existentes]:
            return candidato
        n += 1


# ============================================================
# DYNAMODB — Idempotência / histórico
# ============================================================

def _item_ja_processado(item_id: str, etag: str, nome_atual: str = "", forcar: bool = False) -> bool:
    if forcar:
        return False
    tabela = _tabela_dynamodb()
    if not tabela:
        return False
    try:
        resp = tabela.get_item(Key={"item_id": item_id})
        reg = resp.get("Item")
        if not reg:
            return False
        # Se o arquivo já foi renomeado e o nome atual coincide com o nome já registrado,
        # significa que a notificação foi gerada pelo próprio renomeio! Não reprocessa.
        if reg.get("status") in ("RENOMEADO", "JA_NO_PADRAO"):
            if not nome_atual or reg.get("nome_arquivo", "").lower() == nome_atual.lower():
                logger.info(f"Item {item_id} ('{nome_atual}') já está no padrão/renomeado. Ignorando re-trigger.")
                return True
        # Se o status anterior era MANTIDO ou ERRO mas o arquivo ainda está com o nome bruto,
        # permite reprocessar para dar uma nova chance com as regras atualizadas.
    except Exception as e:
        logger.warning(f"DynamoDB get_item falhou: {e}")
    return False


def _registrar_processamento(
    item_id: str,
    nome: str,
    etag: str,
    status: str,
    motivo: str = "",
    metodo: str = "",
):
    tabela = _tabela_dynamodb()
    if not tabela:
        return
    try:
        tabela.put_item(Item={
            "item_id": item_id,
            "nome_arquivo": nome,
            "etag": etag,
            "status": status,
            "motivo": motivo,
            "metodo": metodo,
        })
    except Exception as e:
        logger.warning(f"DynamoDB put_item falhou: {e}")


# ============================================================
# GRAPH — Metadados do item (webhook)
# ============================================================

def obter_item_sharepoint(drive_id: str, item_id: str, token: str) -> dict:
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
    req = urllib.request.Request(url, method="GET")
    req.add_header("Authorization", f"Bearer {token}")
    with urllib.request.urlopen(req, timeout=20) as resp:
        return json.loads(resp.read().decode("utf-8"))


def listar_nomes_irmaos(drive_id: str, parent_id: str, token: str) -> List[str]:
    """Lista todos os arquivos irmãos na pasta, com suporte a paginação para evitar conflitos (409)."""
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{parent_id}/children"
    nomes = []
    while url:
        try:
            dados = _graph_get_json(url, token)
            for it in dados.get("value", []):
                if "file" in it and it.get("name"):
                    nomes.append(it["name"])
            url = dados.get("@odata.nextLink")
        except Exception as e:
            logger.warning(f"Erro listando irmãos: {e}")
            break
    logger.info(f"listar_nomes_irmaos: {len(nomes)} arquivos irmãos carregados na pasta.")
    return nomes


def _graph_get_json(url: str, token: str) -> dict:
    req = urllib.request.Request(url, method="GET")
    req.add_header("Authorization", f"Bearer {token}")
    with urllib.request.urlopen(req, timeout=20) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _parse_graph_datetime(valor: str) -> Optional[datetime]:
    if not valor:
        return None
    try:
        return datetime.fromisoformat(valor.replace("Z", "+00:00"))
    except ValueError:
        return None


def _paths_para_varredura() -> List[str]:
    """Paths explícitos (override). Vazio = usar varredura dinâmica ano/mês/tipo."""
    if SHAREPOINT_WATCH_PATHS.strip():
        return [p.strip().strip("/") for p in SHAREPOINT_WATCH_PATHS.split(",") if p.strip()]
    return []


def _tipos_pasta_varredura() -> set:
    tipos = SHAREPOINT_TIPOS_PASTA.strip() or "PDF,XML,CSV"
    return {t.strip().upper() for t in tipos.split(",") if t.strip()}


def _anos_para_varredura() -> List[str]:
    if SHAREPOINT_WATCH_YEARS.strip():
        return [a.strip() for a in SHAREPOINT_WATCH_YEARS.split(",") if a.strip()]
    return [str(datetime.now(timezone.utc).year)]


def _listar_children_pasta(drive_id: str, token: str, caminho: str) -> List[dict]:
    caminho_limpo = caminho.strip().strip("/")
    if caminho_limpo:
        caminho_enc = urllib.parse.quote(caminho_limpo)
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{caminho_enc}:/children"
    else:
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"

    itens_totais = []
    pagina = 1
    while url:
        try:
            dados = _graph_get_json(url, token)
            itens = dados.get("value", [])
            itens_totais.extend(itens)
            url = dados.get("@odata.nextLink")
            if url:
                pagina += 1
                logger.info(f"Paginando pasta '{caminho_limpo}': carregada página {pagina} ({len(itens_totais)} itens acumulados)")
        except Exception as e:
            logger.warning(f"Erro listando pasta '{caminho_limpo or '/'}': {e}")
            break

    return itens_totais


def _arquivo_modificado_recentemente(item: dict, limite: Optional[datetime]) -> bool:
    if "file" not in item:
        return False
    nome = item.get("name", "")
    ext = os.path.splitext(nome)[1].lower()
    if ext not in EXTENSOES_VALIDAS or nome.startswith("~$"):
        return False
    if limite is None:
        return True
    mod = _parse_graph_datetime(item.get("lastModifiedDateTime", ""))
    return bool(mod and mod >= limite)


def _coletar_arquivos_de_pasta(
    drive_id: str,
    token: str,
    caminho: str,
    limite: Optional[datetime],
    vistos: set,
    encontrados: List[Tuple[str, str]],
) -> None:
    itens = _listar_children_pasta(drive_id, token, caminho)
    logger.info(f"Pasta '{caminho}': total de {len(itens)} itens no SharePoint")
    for item in itens:
        item_id = item.get("id")
        if not item_id or item_id in vistos:
            continue
        if not _arquivo_modificado_recentemente(item, limite):
            continue
        vistos.add(item_id)
        encontrados.append((drive_id, item_id))
        logger.info(
            f"Arquivo para processar: '{item.get('name')}' em '{caminho}' "
            f"(modificado {item.get('lastModifiedDateTime')})"
        )


def _determinar_prefixo_pasta(drive_id: str, token: str, anos: List[str]) -> str:
    """
    Descobre se a pasta do ano está na raiz do Drive (quando o drive já é a biblioteca de notas)
    ou se está dentro do caminho configurado em SHAREPOINT_FOLDER_PATH.
    """
    primeiro_ano = anos[0] if anos else str(datetime.now(timezone.utc).year)

    # 1. Checa se o ano já está na raiz do drive
    itens_raiz = _listar_children_pasta(drive_id, token, "")
    nomes_raiz = [it.get("name") for it in itens_raiz if "folder" in it]
    if primeiro_ano in nomes_raiz:
        logger.info(f"Pastas de ano encontradas diretamente na raiz do Drive.")
        return ""

    # 2. Checa dentro de SHAREPOINT_FOLDER_PATH
    caminho_base = SHAREPOINT_FOLDER_PATH.strip().strip("/")
    if caminho_base:
        itens_base = _listar_children_pasta(drive_id, token, caminho_base)
        nomes_base = [it.get("name") for it in itens_base if "folder" in it]
        if primeiro_ano in nomes_base:
            logger.info(f"Pastas de ano encontradas dentro de '{caminho_base}'.")
            return caminho_base

    return caminho_base


def _varredura_dinamica_ano_mes_tipo(
    drive_id: str, token: str, minutos: int
) -> List[Tuple[str, str]]:
    """
    Varre {prefixo}/{ano}/{mês}/{PDF|XML|CSV} automaticamente.
    Não exige atualizar env var a cada mês novo.
    minutos <= 0: processa todos os arquivos encontrados sem filtro de data.
    """
    limite = datetime.now(timezone.utc) - timedelta(minutes=minutos) if minutos > 0 else None
    tipos = _tipos_pasta_varredura()
    anos = _anos_para_varredura()
    encontrados: List[Tuple[str, str]] = []
    vistos: set = set()

    prefixo = _determinar_prefixo_pasta(drive_id, token, anos)

    logger.info(
        f"Varredura dinâmica: prefixo='{prefixo}' anos={anos} tipos={sorted(tipos)} "
        f"filtro={f'últimos {minutos} min' if minutos > 0 else 'TODOS os arquivos'}"
    )

    for ano in anos:
        caminho_ano = f"{prefixo}/{ano}".strip("/") if prefixo else ano
        for mes in _listar_children_pasta(drive_id, token, caminho_ano):
            if "folder" not in mes:
                continue
            mes_path = f"{caminho_ano}/{mes['name']}"
            for tipo_pasta in _listar_children_pasta(drive_id, token, mes_path):
                if "folder" not in tipo_pasta:
                    continue
                if tipo_pasta["name"].upper() not in tipos:
                    continue
                tipo_path = f"{mes_path}/{tipo_pasta['name']}"
                _coletar_arquivos_de_pasta(
                    drive_id, token, tipo_path, limite, vistos, encontrados
                )

    if not encontrados:
        logger.info("Varredura dinâmica: nenhum arquivo recente encontrado")

    return encontrados


def _extrair_ids_do_resource(resource: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Retorna (drive_id, item_id, tipo).
    tipo: drive_item | list_item | root | None
    """
    if not resource:
        return None, None, None

    res = "/" + resource.lstrip("/")

    # Formato: /drives/{drive-id}/items/{item-id} ou /drives('{drive-id}')/items('{item-id}')
    m = re.search(r"/drives(?:/|\('?)([^/'\)]+)(?:'|\))?/items(?:/|\('?)([^/'\)?]+)", res)
    if m:
        return m.group(1), m.group(2), "drive_item"

    # Formato SharePoint lista: /sites/.../lists/{list-id}/items/{item-id}
    m = re.search(r"/sites/(.+?)/lists/([^/]+)/items/([^/?]+)", res)
    if m:
        return None, m.group(3), "list_item"

    # Formato: /sites/.../drive/items/{item-id}
    m = re.search(r"/sites/(.+?)/drive/items/([^/?]+)", res)
    if m:
        return None, m.group(2), "drive_item"

    # Subscription na raiz — Graph avisa sem id do arquivo:
    # Ex: /drives/{drive-id}/root, /drives('...')/root, /drive/root, drives/.../root, etc.
    if re.search(r"/(?:drive|drives(?:/|\('?)([^/'\)]+)(?:'|\))?)/root(?:/.*)?$", res) or res.rstrip("/").endswith("/root"):
        m_drv = re.search(r"/drives(?:/|\('?)([^/'\)]+)(?:'|\))?/root", res)
        d_id = m_drv.group(1) if m_drv else None
        return d_id, None, "root"

    # Formato genérico: último /items/{id} no path (desde que não seja root)
    m = re.search(r"/items/([^/?]+)", res)
    if m:
        return None, m.group(1), "drive_item"

    return None, None, None


def _extrair_site_e_list_do_resource(resource: str) -> Tuple[Optional[str], Optional[str]]:
    m = re.search(r"/sites/(.+?)/lists/([^/]+)/items/", resource)
    if m:
        return m.group(1), m.group(2)
    return None, None


def resolver_list_item_para_drive(
    site_id: str, list_id: str, list_item_id: str, token: str
) -> Tuple[Optional[str], Optional[str]]:
    """Converte id de item de lista SharePoint em drive item."""
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}"
        f"/lists/{list_id}/items/{list_item_id}/driveItem"
    )
    try:
        data = _graph_get_json(url, token)
        drive_id = data.get("parentReference", {}).get("driveId")
        item_id = data.get("id")
        if drive_id and item_id:
            return drive_id, item_id
    except Exception as e:
        logger.warning(f"Falha ao resolver list item {list_item_id}: {e}")
    return None, None


def resolver_item_pelo_resource_url(resource: str, token: str) -> Tuple[Optional[str], Optional[str]]:
    """Tenta GET direto no resource ou no driveItem derivado."""
    if not resource:
        return None, None

    path = resource if resource.startswith("/") else f"/{resource}"
    if path.rstrip("/").endswith("/root"):
        return None, None

    urls = [f"https://graph.microsoft.com/v1.0{path}"]
    if "/lists/" in path and "/items/" in path and not path.endswith("/driveItem"):
        urls.append(f"https://graph.microsoft.com/v1.0{path}/driveItem")

    for url in urls:
        try:
            data = _graph_get_json(url, token)
            item_id = data.get("id")
            if not item_id:
                continue
            drive_id = data.get("parentReference", {}).get("driveId")
            if drive_id and ("file" in data or "folder" in data):
                return drive_id, item_id
        except Exception:
            continue
    return None, None


def listar_arquivos_modificados_recentemente(
    drive_id: str, token: str, minutos: int
) -> List[Tuple[str, str]]:
    """Fallback: arquivos alterados nos últimos N minutos (ou todos se minutos <= 0)."""
    paths = _paths_para_varredura()
    if not paths:
        return _varredura_dinamica_ano_mes_tipo(drive_id, token, minutos)

    limite = datetime.now(timezone.utc) - timedelta(minutes=minutos) if minutos > 0 else None
    encontrados: List[Tuple[str, str]] = []
    vistos: set = set()

    logger.info(f"Varredura por paths explícitos: {paths}")
    for path in paths:
        _coletar_arquivos_de_pasta(drive_id, token, path, limite, vistos, encontrados)

    return encontrados


def resolver_itens_da_notificacao(
    notif: dict,
    token: str,
    site_id: str,
    drive_id_padrao: str,
) -> List[Tuple[str, str]]:
    """
    Resolve uma notificação Graph em uma ou mais tuplas (drive_id, item_id).
    Cobre: drive items, list items, resourceData e fallback /drive/root.
    """
    resource = (notif.get("resource") or "").strip()
    logger.info(f"resource completo recebido: {resource}")

    resource_data = notif.get("resourceData") or {}
    rd_id = resource_data.get("id")
    rd_type = (resource_data.get("@odata.type") or "").lower()

    if rd_id and "driveitem" in rd_type:
        return [(drive_id_padrao, rd_id)]

    drive_id, item_id, tipo = _extrair_ids_do_resource(resource)
    d_id_final = drive_id or drive_id_padrao

    if tipo == "list_item":
        list_site_id, list_id = _extrair_site_e_list_do_resource(resource)
        site_ref = list_site_id or site_id
        if site_ref and list_id and item_id:
            d_id, i_id = resolver_list_item_para_drive(site_ref, list_id, item_id, token)
            if d_id and i_id:
                return [(d_id, i_id)]

    if tipo == "drive_item" and item_id:
        return [(d_id_final, item_id)]

    if tipo == "root":
        logger.info(
            f"Notificação de raiz ({resource}) no drive {d_id_final} — varredura dos últimos "
            f"{WEBHOOK_FALLBACK_MINUTES} min"
        )
        return listar_arquivos_modificados_recentemente(
            d_id_final, token, WEBHOOK_FALLBACK_MINUTES
        )

    # Último recurso: GET direto no resource
    d_id, i_id = resolver_item_pelo_resource_url(resource, token)
    if d_id and i_id:
        return [(d_id, i_id)]

    if rd_id:
        return [(d_id_final, rd_id)]

    logger.warning(f"Não foi possível resolver item de: {resource}")
    return []


# ============================================================
# PROCESSAMENTO DE UM ARQUIVO (webhook)
# ============================================================

def processar_um_arquivo(drive_id: str, item_id: str, token: str, forcar: bool = False) -> Dict[str, Any]:
    item = obter_item_sharepoint(drive_id, item_id, token)
    nome = item.get("name", "")
    etag = item.get("eTag", "")

    if "folder" in item:
        return {"arquivo": nome, "status": "IGNORADO", "motivo": "é pasta, não arquivo"}

    ext = os.path.splitext(nome)[1].lower()
    if ext not in EXTENSOES_VALIDAS or nome.startswith("~$"):
        return {"arquivo": nome, "status": "IGNORADO", "motivo": "extensão inválida"}

    if nome.lower() in ARQUIVOS_SEM_AVISO:
        return {"arquivo": nome, "status": "IGNORADO", "motivo": "documento interno"}

    if _item_ja_processado(item_id, etag, nome, forcar=forcar):
        return {"arquivo": nome, "status": "PULADO", "motivo": "já processado (DynamoDB)"}

    try:
        conteudo = baixar_conteudo_sharepoint(drive_id, item_id, token)

        if ext == ".pdf":
            res = extrair_nome_pdf_bytes(conteudo)
        elif ext == ".xml":
            res = extrair_nome_xml_bytes(conteudo)
        else:
            res = extrair_nome_csv_bytes(conteudo)

        if not res.nome:
            logger.info(f"MANTIDO | {nome} | {res.motivo}")
            _registrar_processamento(item_id, nome, etag, "MANTIDO", res.motivo)
            return {"arquivo": nome, "status": "MANTIDO", "motivo": res.motivo}

        parent_id = item.get("parentReference", {}).get("id")
        nomes_irmaos = listar_nomes_irmaos(drive_id, parent_id, token) if parent_id else []
        novo_nome = calcular_destino_sem_colisao(res.nome, ext, nomes_irmaos, nome)

        if novo_nome is None:
            _registrar_processamento(item_id, nome, etag, "JA_NO_PADRAO", metodo=res.metodo)
            return {"arquivo": nome, "status": "JA_NO_PADRAO", "metodo": res.metodo}

        if DRY_RUN:
            logger.info(f"[SIMULAÇÃO] {nome} -> {novo_nome} ({res.metodo})")
            _registrar_processamento(item_id, nome, etag, "SIMULADO", metodo=res.metodo)
            return {
                "arquivo_antigo": nome,
                "arquivo_novo": novo_nome,
                "status": "SIMULADO",
                "metodo": res.metodo,
            }

        # Renomeio no SharePoint com resolução automática de colisão (409 Conflict)
        nome_para_tentar = novo_nome
        sufixo = 1
        m_suf = re.search(r"_(\d+)\.[^.]+$", novo_nome)
        if m_suf:
            sufixo = int(m_suf.group(1))

        resultado_rename = None
        for tentativa in range(15):
            try:
                resultado_rename = renomear_arquivo_sharepoint(drive_id, item_id, nome_para_tentar, token)
                break
            except urllib.error.HTTPError as e:
                if e.code == 409:
                    sufixo += 1
                    nome_para_tentar = f"{res.nome}_{sufixo}{ext}"
                    logger.warning(
                        f"Conflito 409 no SharePoint para '{nome}'. "
                        f"Tentando próximo sufixo '{nome_para_tentar}'..."
                    )
                    continue
                raise

        if not resultado_rename:
            raise RuntimeError(f"Não foi possível renomear '{nome}' após tentativas de resolução de conflito 409.")

        novo_nome = nome_para_tentar
        novo_etag = resultado_rename.get("eTag", etag)
        logger.info(f"RENOMEADO | {nome} -> {novo_nome} ({res.metodo})")
        _registrar_processamento(item_id, novo_nome, novo_etag, "RENOMEADO", metodo=res.metodo)
        return {
            "arquivo_antigo": nome,
            "arquivo_novo": novo_nome,
            "status": "RENOMEADO",
            "metodo": res.metodo,
        }

    except Exception as err:
        logger.error(f"Erro ao processar '{nome}': {err}")
        return {"arquivo": nome, "status": "ERRO", "motivo": str(err)}


# ============================================================
# WEBHOOK — Microsoft Graph Change Notifications
# ============================================================

def _resposta_validacao(validation_token: str) -> Dict[str, Any]:
    return {
        "statusCode": 200,
        "headers": {"Content-Type": "text/plain; charset=utf-8"},
        "body": validation_token,
    }


def _resposta_json(status: int, payload: dict) -> Dict[str, Any]:
    return {
        "statusCode": status,
        "headers": {"Content-Type": "application/json"},
        "body": json.dumps(payload, ensure_ascii=False),
    }


def _query_params(event: Dict[str, Any]) -> Dict[str, str]:
    params = event.get("queryStringParameters") or {}
    if isinstance(params, dict):
        return {k: v for k, v in params.items() if v is not None}
    return {}


def _corpo_json(event: Dict[str, Any]) -> dict:
    # Se o próprio evento já contiver chaves diretas (ex: teste manual no Console AWS)
    if isinstance(event, dict) and ("manual_scan" in event or "value" in event):
        return event

    body = event.get("body") if isinstance(event, dict) else None
    if isinstance(body, dict):
        return body
    if event.get("isBase64Encoded") and isinstance(body, str):
        body = base64.b64decode(body).decode("utf-8")
    if not body:
        return event if isinstance(event, dict) else {}
    if isinstance(body, str):
        try:
            return json.loads(body)
        except json.JSONDecodeError:
            return {}
    return event if isinstance(event, dict) else {}


def lambda_handler(event: Dict[str, Any], context: Any) -> Dict[str, Any]:
    """
    Webhook Microsoft Graph via Lambda Function URL.
    - GET/POST ?validationToken=... → validação da subscription
    - POST com notificações → processa arquivo(s) notificados
    """
    # Não loga o evento completo (pode conter clientState, IDs e dados do tenant)
    logger.info(f"Evento recebido (chaves): {sorted(event.keys()) if isinstance(event, dict) else type(event)}")

    params = _query_params(event)
    validation_token = params.get("validationToken")
    if validation_token:
        logger.info("Validação de subscription (validationToken)")
        return _resposta_validacao(validation_token)

    corpo = _corpo_json(event)

    # Suporte a teste manual via Console AWS (ex: {"manual_scan": true, "minutos": 60, "forcar": true})
    if corpo.get("manual_scan"):
        # Varredura manual só por invocação direta (Console/SDK). Via Function URL (HTTP) é bloqueada,
        # senão qualquer pessoa com a URL conseguiria disparar varreduras com forcar=true.
        if "requestContext" in event:
            logger.warning("manual_scan recusado: chamada veio pela Function URL")
            return _resposta_json(403, {"erro": "manual_scan só é permitido por invocação direta"})
        minutos = int(corpo.get("minutos", 60))
        forcar = bool(corpo.get("forcar", False))
        logger.info(f"Modo TESTE MANUAL acionado — varrendo últimos {minutos} minutos (forçar={forcar})...")
        token = obter_token_microsoft_graph()
        site_id, drive_id_padrao = resolver_site_e_drive_id(token)
        itens = listar_arquivos_modificados_recentemente(drive_id_padrao, token, minutos)
        logger.info(f"Arquivos encontrados no teste manual: {len(itens)}")
        resultados = [processar_um_arquivo(d, i, token, forcar=forcar) for d, i in itens]
        return _resposta_json(200, {
            "mensagem": "Varredura manual concluída",
            "total": len(resultados),
            "detalhes": resultados,
        })

    notificacoes = corpo.get("value") or []

    if not notificacoes:
        return _resposta_json(200, {"mensagem": "Nenhuma notificação no corpo"})

    if not GRAPH_CLIENT_STATE:
        logger.error("GRAPH_CLIENT_STATE não configurado — webhook rejeitado")
        return _resposta_json(500, {"erro": "configuração incompleta"})

    try:
        token = obter_token_microsoft_graph()
        site_id, drive_id_padrao = resolver_site_e_drive_id(token)
        resultados = []

        for notif in notificacoes:
            client_state = notif.get("clientState", "")
            if not hmac.compare_digest(str(client_state).encode("utf-8"), GRAPH_CLIENT_STATE.encode("utf-8")):
                logger.warning("clientState inválido — notificação ignorada")
                resultados.append({"status": "IGNORADO", "motivo": "clientState inválido"})
                continue

            itens = resolver_itens_da_notificacao(notif, token, site_id, drive_id_padrao)
            if not itens:
                resultados.append({
                    "status": "IGNORADO",
                    "motivo": "item não resolvido a partir do resource",
                    "resource": notif.get("resource", ""),
                })
                continue

            for drive_id, item_id in itens:
                logger.info(f"Processando notificação: drive={drive_id} item={item_id}")
                resultados.append(processar_um_arquivo(drive_id, item_id, token))

        return _resposta_json(202, {
            "mensagem": "Notificações processadas",
            "total": len(resultados),
            "detalhes": resultados,
        })

    except Exception as e:
        logger.error(f"Erro fatal no webhook: {e}")
        return _resposta_json(500, {"erro": "erro interno ao processar notificações"})
