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

## Observacoes
- O filtro de incremento usa janela de 24h (`Filtrar Incrementos (24h)`).
- O parser depende do texto extraivel do PDF. PDFs escaneados sem OCR podem vir com campos vazios.
- O link curto usa formato: `https://drive.google.com/file/d/<fileId>/view`.
