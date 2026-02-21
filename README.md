# BD Votos Homologacao (n8n + Python)

Workflow separado para:
- rodar 1 vez por dia;
- varrer a pasta do Drive `12X55vaaTQx-8xMu00Xmbo31rm18r80sK`;
- processar PDFs novos (ultimas 24h);
- extrair campos juridicos;
- gerar planilha `.xlsx` com colunas ajustadas;
- subir a planilha na mesma pasta do Drive.

## Arquivos
- `workflow_n8n_drive_votos.json`: workflow para importar no n8n.
- `extrair_votos_pdf.py`: script Python de extracao e geracao do XLSX.

## Campos extraidos
- NOME DO ARQUIVO
- PROMOTORIA DE JUSTICA
- NUMERO DO PROCEDIMENTO
- TIPO DO PROCEDIMENTO
- OBJETO
- RELATOR
- EMENTA
- CASO EM EXAME
- CONCLUSAO DO RELATOR
- NOME DO PROMOTOR DE JUSTICA QUE ARQUIVOU
- DATA DE HOMOLOGACAO
- LINK CURTO PARA O ARQUIVO

## Pre-requisitos no servidor n8n
1. Python 3 instalado.
2. Dependencias Python instaladas:
   - `pip install pypdf openpyxl`
3. `Execute Command` habilitado no n8n:
   - `N8N_ENABLE_EXECUTE_COMMAND=true`
4. Credencial Google Drive OAuth2 configurada no n8n.

## Como configurar
1. Copie `extrair_votos_pdf.py` para o servidor n8n em:
   - `/opt/n8n/scripts/extrair_votos_pdf.py`
2. Torne executavel:
   - `chmod +x /opt/n8n/scripts/extrair_votos_pdf.py`
3. Importe `workflow_n8n_drive_votos.json` no n8n.
4. No workflow importado, ajuste:
   - `credentials.googleDriveOAuth2Api.id`
   - `credentials.googleDriveOAuth2Api.name`
5. Opcional: altere horario do no `Cron Diario`.
6. Execute manualmente uma vez e valide a planilha criada no Drive.
7. Ative o workflow.

## Funcionalidades do script Python (`extrair_votos_pdf.py`)

O script expõe dois subcomandos via linha de comando:

### Subcomando `extract`
Extrai campos jurídicos de um único arquivo PDF e imprime um objeto JSON com os dados estruturados.

```
python extrair_votos_pdf.py extract --pdf <caminho_do_pdf> [--file-id <id_drive>] [--file-name <nome_original>]
```

**Etapas internas:**
1. **Leitura do PDF** — todas as páginas são lidas com `pypdf` e o texto é concatenado, normalizando espaços e quebras de linha.
2. **Extração de campos** — os seguintes campos são identificados no texto extraído:
   - `NOME DO ARQUIVO`: nome original do arquivo (ou `--file-name` se informado).
   - `PROMOTORIA DE JUSTICA`: primeira linha não vazia do documento.
   - `NUMERO DO PROCEDIMENTO`: número no formato `NNNNN.NNN.NNN-NNNN` localizado na segunda linha ou em qualquer parte do texto.
   - `TIPO DO PROCEDIMENTO`: identificado por expressões regulares que reconhecem *Inquérito Civil*, *Procedimento Preparatório*, *Notícia de Fato* e *Procedimento Administrativo / PA*.
   - `OBJETO`: texto entre a segunda linha e o cabeçalho `RELATOR / RELATORA`.
   - `RELATOR`: conteúdo após o cabeçalho `RELATOR` ou `RELATORA`.
   - `EMENTA`: conteúdo após o cabeçalho `EMENTA`; caso não encontrado na mesma linha, extrai o bloco até o próximo cabeçalho relevante.
   - `CASO EM EXAME`: parágrafo imediatamente após o cabeçalho `CASO EM EXAME` ou `CASO EXAMINADO`.
   - `CONCLUSÃO DO RELATOR`: parágrafo imediatamente após o cabeçalho `CONCLUSAO DO RELATOR`.
   - `NOME DO PROMOTOR DE JUSTIÇA QUE ARQUIVOU`: nome extraído por expressão regular que identifica "Promotor(a) de Justiça" seguido ou precedido de nome próprio.
   - `DATA DE HOMOLOGAÇÃO`: data no formato `DD/MM/AAAA` localizada após as expressões "DATA DE HOMOLOGACAO" ou "HOMOLOGACAO EM".
   - `LINK CURTO PARA O ARQUIVO`: URL `https://drive.google.com/file/d/<file-id>/view` montada a partir do parâmetro `--file-id`.
3. **Saída JSON** — o registro é impresso em `stdout` como objeto JSON com chave para cada coluna.

### Subcomando `build`
Recebe uma lista de registros JSON (codificada em Base64) e gera uma planilha `.xlsx`.

```
python extrair_votos_pdf.py build --rows-b64 <lista_json_em_base64> --output <caminho_saida.xlsx>
```

**Etapas internas:**
1. **Decodificação** — o argumento `--rows-b64` é decodificado de Base64 e interpretado como lista JSON.
2. **Geração do XLSX** — uma planilha é criada com `openpyxl`; a primeira linha contém os cabeçalhos (colunas definidas em `COLUMNS`); as demais linhas recebem os valores de cada registro.
3. **Ajuste automático de colunas** — a largura de cada coluna é ajustada ao conteúdo (mínimo 16, máximo 80 caracteres).
4. **Saída** — o arquivo `.xlsx` é salvo no caminho indicado por `--output` e o caminho absoluto é impresso em `stdout`.

### Utilitários internos
| Função | Descrição |
|---|---|
| `read_pdf_text` | Extrai e normaliza o texto de todas as páginas do PDF. |
| `non_empty_lines` | Divide o texto em linhas não vazias já normalizadas. |
| `find_line_index` | Localiza o índice da linha que começa com um dos cabeçalhos informados (comparação sem acentos, case-insensitive). |
| `extract_line_after_colon` | Retorna o valor após `:` na linha do cabeçalho, ou a linha seguinte. |
| `paragraph_after_heading` | Retorna a linha imediatamente após o cabeçalho localizado. |
| `extract_block_after_heading` | Extrai um bloco multilínea entre o cabeçalho e o próximo cabeçalho conhecido. |
| `extract_first` | Aplica uma lista de expressões regulares e retorna o primeiro grupo capturado. |
| `normalize_token` | Remove acentos e converte para maiúsculas para comparações robustas. |
| `autosize_columns` | Ajusta a largura das colunas da planilha ao conteúdo. |

## Observacoes
- O filtro de incremento usa janela de 24h (`Filtrar Incrementos (24h)`).
- O parser depende do texto extraivel do PDF. PDFs escaneados sem OCR podem vir com campos vazios.
- O link curto usa formato: `https://drive.google.com/file/d/<fileId>/view`.
