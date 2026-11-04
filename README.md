# Renomeador de Notas Fiscais — AWS Lambda + SharePoint

Automação em **Python 3.12+** que identifica o fornecedor/prestador responsável pela emissão de notas fiscais e renomeia os arquivos diretamente no SharePoint.

A solução processa arquivos **PDF, XML e CSV**, recebe notificações do Microsoft Graph por uma AWS Lambda Function URL e utiliza o DynamoDB para evitar o reprocessamento de arquivos já tratados.

## Visão geral

O objetivo é reduzir o trabalho manual de identificação e organização de notas fiscais recebidas pelo setor financeiro, mantendo os documentos nas pastas em que já são armazenados no SharePoint.

### Fluxo de funcionamento

1. O Power Automate continua recebendo os documentos enviados por e-mail e salvando-os nas pastas correspondentes do SharePoint.
2. Uma subscription do Microsoft Graph monitora alterações e novos arquivos e envia notificações para a Function URL da AWS Lambda.
3. O Airflow é utilizado no fluxo de infraestrutura para renovar a subscription antes do vencimento. A renovação é uma etapa externa à função Lambda apresentada neste repositório.
4. A Lambda valida a notificação e identifica o arquivo que precisa ser processado.
5. O arquivo é baixado do SharePoint por meio da Microsoft Graph API.
6. O sistema identifica o nome do emitente/prestador, aplicando validações para evitar confundir o fornecedor com o tomador ou destinatário.
7. Quando a emissão é identificada como sendo da própria SUA EMPRESA, o arquivo não é renomeado.
8. Se o nome do fornecedor for identificado com segurança, a Lambda renomeia o arquivo no próprio SharePoint.
9. O DynamoDB registra o processamento para evitar que notificações repetidas provoquem um novo processamento desnecessário.

Quando a notificação não informa diretamente o ID do arquivo e aponta para a raiz do Drive, a função pode realizar uma varredura de fallback nas pastas configuradas, considerando arquivos modificados recentemente.

## Funcionalidades

- **PDF:** extração de texto e análise da posição dos elementos para localizar o nome do prestador. Também possui um parser nativo em Python como alternativa ao `pdfplumber`.
- **PDF escaneado:** integração opcional com o Amazon Textract para OCR de imagens compatíveis. O código não envia diretamente PDFs ao método `detect_document_text`; PDFs que não permitem extração textual podem permanecer sem renomeação se o OCR não conseguir processá-los.
- **XML:** identificação do emitente em diferentes estruturas, incluindo padrões de NFS-e/NF-e, tabelas e layouts FastReport.
- **CSV:** identificação do fornecedor a partir de colunas relacionadas a prestador, fornecedor ou emitente.
- **Validação de documentos:** validação de CPF e CNPJ antes de considerar os documentos encontrados no conteúdo.
- **Bloqueio da SUA EMPRESA:** impede a renomeação de notas identificadas como emitidas pela própria empresa, incluindo filiais que compartilhem a raiz do CNPJ configurado.
- **Limpeza de nomes:** remove acentos e caracteres inválidos, filtra textos que aparentam ser endereços, documentos ou informações sem relação com o nome do fornecedor e limita o nome gerado a 20 caracteres.
- **Prevenção de conflitos:** verifica os nomes existentes na pasta e utiliza sufixos numéricos quando necessário.
- **Idempotência:** consulta o DynamoDB para evitar reprocessar arquivos já renomeados ou reconhecidos como estando no padrão.
- **Modo de simulação:** `DRY_RUN=true` registra o nome que seria aplicado sem alterar o arquivo no SharePoint.
- **Varredura dinâmica:** permite percorrer a estrutura de ano, mês e tipo de arquivo sem precisar atualizar um caminho a cada mês.

## Tecnologias utilizadas

| Tecnologia | Finalidade |
|---|---|
| Python 3.12+ | Linguagem da função de processamento |
| AWS Lambda | Execução da automação e recebimento do webhook |
| AWS DynamoDB | Registro de arquivos processados e controle de idempotência |
| Amazon Textract | OCR opcional para documentos de imagem compatíveis |
| Microsoft Graph API | Autenticação, leitura, listagem e renomeação de arquivos |
| Microsoft Entra ID | App Registration e autenticação OAuth 2.0 |
| SharePoint | Armazenamento dos documentos |
| Power Automate | Recebimento e salvamento dos anexos no SharePoint |
| Apache Airflow | Renovação programada da subscription do Microsoft Graph |
| `pdfplumber` | Extração de texto e coordenadas em PDFs, quando instalado |

## Arquitetura simplificada

```text
Outlook
   |
   v
Power Automate
   |
   v
SharePoint (pastas de notas fiscais)
   |
   v
Microsoft Graph Subscription
   |
   v
AWS Lambda Function URL
   |
   +---- Microsoft Graph API ----> Baixa / renomeia o arquivo
   |
   +---- Parser PDF / XML / CSV -> Identifica o emitente
   |
   +---- Amazon Textract --------> OCR (opcional)
   |
   +---- DynamoDB ---------------> Histórico e idempotência

Apache Airflow
   |
   +---- Renova a subscription do Microsoft Graph
```

## Estrutura esperada das pastas

Por padrão, a varredura dinâmica considera pastas organizadas por ano, mês e tipo de arquivo:

```text
Financeiro - Notas/
└── 2026/
    ├── Janeiro/
    │   ├── PDF/
    │   ├── XML/
    │   └── CSV/
    ├── Fevereiro/
    │   ├── PDF/
    │   ├── XML/
    │   └── CSV/
    └── ...
```

A estrutura real pode variar. O caminho-base é definido por `SHAREPOINT_FOLDER_PATH`. Se o Drive já corresponder à biblioteca de notas, a função também consegue detectar as pastas de ano diretamente na raiz.

## Configuração

As configurações são lidas por variáveis de ambiente na AWS Lambda. Para desenvolvimento local, crie um arquivo `.env` com base no exemplo abaixo; o código apresentado utiliza `os.environ` e não carrega automaticamente o arquivo `.env`.

### Variáveis de ambiente

| Variável | Descrição |
|---|---|
| `AZURE_TENANT_ID` | ID do tenant do Microsoft Entra ID |
| `AZURE_CLIENT_ID` | ID do App Registration |
| `AZURE_CLIENT_SECRET` | Segredo do App Registration |
| `SHAREPOINT_SITE_URL` | URL do site do SharePoint |
| `SHAREPOINT_FOLDER_PATH` | Caminho-base da pasta de notas |
| `SHAREPOINT_SITE_ID` | ID do site; opcional quando pode ser resolvido pela URL |
| `SHAREPOINT_DRIVE_ID` | ID do Drive/biblioteca; opcional quando pode ser resolvido automaticamente |
| `SUA EMPRESA CNPJ` | CNPJ usado como referência para bloquear notas emitidas pela SUA EMPRESA |
| `DRY_RUN` | `true` simula; `false` permite renomear de verdade |
| `USAR_AWS_TEXTRACT` | Ativa ou desativa a integração com Textract |
| `GRAPH_CLIENT_STATE` | Valor secreto compartilhado para validar notificações do webhook |
| `DYNAMODB_TABLE` | Nome da tabela DynamoDB usada no controle de processamento |
| `WEBHOOK_FALLBACK_MINUTES` | Janela, em minutos, para a varredura de fallback |
| `SHAREPOINT_WATCH_PATHS` | Lista opcional de caminhos fixos separados por vírgula |
| `SHAREPOINT_WATCH_YEARS` | Anos da varredura separados por vírgula; vazio utiliza o ano atual |
| `SHAREPOINT_TIPOS_PASTA` | Tipos de subpastas monitoradas; padrão `PDF,XML,CSV` |

O limite de 20 caracteres (`NOME_MAX_CARACTERES`), o separador de palavras e a lista de documentos internos ignorados são definidos diretamente no código apresentado.

### Exemplo de `.env.example`

Copie o conteúdo para `.env` no ambiente local ou configure os valores em **AWS Lambda → Configuration → Environment variables**. Todos os valores abaixo são fictícios.

```dotenv
# Microsoft Entra ID — App Registration
AZURE_TENANT_ID=00000000-0000-0000-0000-000000000000
AZURE_CLIENT_ID=00000000-0000-0000-0000-000000000000
AZURE_CLIENT_SECRET=substitua-por-um-segredo

# SharePoint
SHAREPOINT_SITE_URL=https://SUA-EMPRESA.sharepoint.com/sites/SEU-SITE
SHAREPOINT_FOLDER_PATH=Financeiro - Notas
SHAREPOINT_SITE_ID=
SHAREPOINT_DRIVE_ID=

# Varredura de pastas (opcional)
SHAREPOINT_WATCH_PATHS=
SHAREPOINT_WATCH_YEARS=
SHAREPOINT_TIPOS_PASTA=PDF,XML,CSV

# Regras de negócio
SUAEMPRESA_CNPJ=00.000.000/0001-00

# Webhook Microsoft Graph
# Gere um valor longo, aleatório e difícil de adivinhar.
GRAPH_CLIENT_STATE=SUBSTITUA_POR_UM_VALOR_SECRETO

# AWS e comportamento
DYNAMODB_TABLE=notas-processadas
USAR_AWS_TEXTRACT=true
DRY_RUN=true
WEBHOOK_FALLBACK_MINUTES=60
```

> **Importante:** os nomes das variáveis deste exemplo correspondem aos utilizados no código. Não publique credenciais, CNPJs reais, IDs internos ou valores reais de `GRAPH_CLIENT_STATE`.

## Execução e implantação

A função principal de entrada da AWS Lambda é:

```python
lambda_handler(event, context)
```

### Implantação na AWS

1. Crie ou selecione uma função AWS Lambda com runtime Python 3.12 ou compatível.
2. Configure o arquivo do projeto como código da função e defina o handler conforme o nome do módulo. Por exemplo, se o arquivo for `lambda_function.py`, utilize `lambda_function.lambda_handler`.
3. Configure as variáveis de ambiente descritas anteriormente.
4. Configure a tabela DynamoDB indicada em `DYNAMODB_TABLE`, com `item_id` como chave de partição do tipo String.
5. Configure a role IAM da Lambda com as permissões necessárias para acessar a tabela DynamoDB, utilizar o Textract quando habilitado e gravar logs no CloudWatch.
6. Configure o App Registration no Microsoft Entra ID com as permissões mínimas necessárias para consultar e alterar os arquivos do site SharePoint.
7. Crie uma Function URL e configure a subscription do Microsoft Graph para enviar as notificações ao endpoint.
8. Configure o processo externo no Airflow para renovar a subscription antes do vencimento.

As permissões exatas devem ser definidas de acordo com a política do ambiente e o modelo de autenticação adotado. Não conceda acesso amplo sem necessidade.

### Dependências

O código utiliza módulos da biblioteca padrão do Python e integrações opcionais:

- `boto3`: acesso ao DynamoDB e ao Amazon Textract.
- `pdfplumber`: extração de texto e coordenadas de PDFs.

Na AWS Lambda, confirme a disponibilidade das dependências no runtime ou inclua-as no pacote de implantação ou em uma Lambda Layer. O código possui alternativas quando `pdfplumber` não está instalado; já o acesso ao DynamoDB e ao Textract depende de `boto3` estar disponível e das permissões AWS correspondentes.

## Teste seguro

Antes de permitir alterações reais nos documentos:

1. Configure `DRY_RUN=true`.
2. Configure as credenciais e os acessos necessários.
3. Envie uma notificação de teste ou execute uma varredura manual pelo Console/SDK da Lambda.
4. Confira os logs no CloudWatch e valide os nomes que seriam aplicados.
5. Depois de validar os resultados, altere `DRY_RUN` para `false` para habilitar a renomeação real.

### Varredura manual

O código aceita um evento de teste manual com os campos abaixo:

```json
{
  "manual_scan": true,
  "minutos": 60,
  "forcar": false
}
```

- `minutos`: quantidade de minutos considerada na busca. Valor `0` ou negativo remove o filtro de data e permite varrer todos os arquivos dos caminhos configurados.
- `forcar`: quando `true`, ignora a verificação de idempotência do DynamoDB para os arquivos encontrados.

**Atenção:** a varredura manual deve ser executada por invocação direta da Lambda, pelo Console ou SDK. Não disponibilize esse modo pela Function URL. A implementação atual reconhece `manual_scan` no evento recebido; portanto, proteja o endpoint e não permita que chamadas externas acionem esse comportamento.

## Status de processamento

| Status | Significado |
|---|---|
| `RENOMEADO` | Arquivo renomeado no SharePoint |
| `SIMULADO` | Nome calculado, mas nenhuma alteração foi realizada por causa do `DRY_RUN` |
| `JA_NO_PADRAO` | Arquivo já possui o nome esperado |
| `MANTIDO` | Nome do fornecedor não foi identificado com segurança ou não foi possível aplicar as regras |
| `PULADO` | Arquivo já havia sido processado, conforme o DynamoDB |
| `IGNORADO` | Item não é um arquivo válido para processamento ou está na lista de documentos ignorados |
| `ERRO` | Ocorreu uma falha durante o processamento do arquivo |

## Segurança e LGPD

Notas fiscais podem conter dados pessoais e informações comerciais. Por isso:

- Não armazene credenciais, tokens, CNPJs reais ou dados de fornecedores no repositório.
- Mantenha `.env` fora do controle de versão e disponibilize apenas um `.env.example` com valores fictícios.
- Utilize um `GRAPH_CLIENT_STATE` longo, aleatório e mantido em segredo. Notificações com valor diferente são descartadas quando a validação está configurada.
- Restrinja o acesso à Function URL e à função Lambda. O endpoint deve receber apenas as notificações esperadas do Microsoft Graph.
- Não exponha a varredura `manual_scan` pela Function URL.
- Conceda somente as permissões necessárias ao App Registration e à role IAM.
- Restrinja o acesso aos logs do CloudWatch e evite registrar conteúdo de notas fiscais ou dados pessoais.
- Utilize `DRY_RUN=true` durante testes e validações.
- Revise periodicamente os acessos ao SharePoint, à AWS e aos registros de processamento.

## Observações técnicas

- A função trabalha com arquivos PDF, XML e CSV; outros formatos são ignorados.
- A identificação do fornecedor depende do conteúdo e do layout do documento. Quando não há segurança suficiente, o arquivo é mantido com o nome original.
- A integração de OCR apresentada utiliza `detect_document_text` do Textract para bytes de imagens compatíveis. O próprio código não converte páginas de PDF em imagens nem envia PDFs diretamente ao Textract.
- O DynamoDB utiliza `item_id` como chave de consulta. A tabela deve ser criada com essa chave de partição.
- O Airflow e o Power Automate fazem parte do fluxo operacional, mas não são executados pelo código da função Lambda deste arquivo.

---

**Projeto:** Renomeador de Notas Fiscais  
**Runtime:** Python 3.12+  
**Execução:** AWS Lambda  
**Integração documental:** Microsoft Graph API e SharePoint
