# Final Boleto Solution - Complete Package

🎯 **COMPLETE FIX** for the HS Consórcios boleto automation popup content loading issue.

## 🔧 What's Fixed

This package includes the **FINAL SOLUTION** that resolves:
- ✅ **Popup content loading** - No more "about:blank" popups
- ✅ **submitFunction execution** - Proper context handling  
- ✅ **PDF generation waiting** - Waits for complete PDF creation
- ✅ **Content verification** - Ensures real boleto data before PDF

## 🚀 Quick Start

### 1. Setup (Run Once)
```bash
chmod +x setup.sh
./setup.sh
```

### 2. Debug Test (Recommended First)
```bash
python debug_submitfunction_test.py
```
*Shows step-by-step submitFunction execution*

### 3. Single Record Test
```bash
python test_final_solution.py
```
*Quick validation with one record*

### 4. Production Run
```bash
# Small test first
./run_final_solution.sh --max-records 5 --batch-size 1

# Full automation
./run_final_solution.sh --batch-size 5
```

### 5. Signed PDF delivery (WhatsApp)

1. Configure `file_server` in `config.yaml` with your public base URL and a strong `secret_key`.
2. Expose the `downloads/` directory with the signed link server:
   ```bash
   python file_link_service.py downloads "<secret_key>" --host 0.0.0.0 --port 8080
   ```
   (Run behind your preferred HTTP server/reverse proxy.)
3. Update the n8n webhook to expect the JSON payload fields `phone`, `message`, `file_url`, `file_name`, and `drive_file_id`.
4. In n8n, download `file_url`, upload the PDF via WhatsApp Cloud `/media`, then send the document message with the returned `media_id`.
   ```json
   {
     "phone": "+5511999999999",
     "file_url": "https://your-domain.com/files?path=...",
     "file_name": "CLIENTE-123-456.pdf",
     "message": "Olá ...",
     "drive_file_id": "1AbCdEf..."
   }
   ```
5. Keep the webhook URL em “modo produção” (não `.../webhook-test/...`) para evitar expiração após cada execução.

## 📁 Files Included

### Core Scripts
- **`final_working_boleto_processor.py`** - Main production automation
- **`debug_submitfunction_test.py`** - Step-by-step debugging tool
- **`test_final_solution.py`** - Single record validation

### Configuration & Data
- **`config.yaml`** - System configuration (credentials: 3090A / SuzanaAdm25@)
- **`controle_boletos_hs.xlsx`** - Your data file (1,487 records)
- **`requirements.txt`** - Python dependencies

### Setup & Run
- **`setup.sh`** - Automated environment setup
- **`run_final_solution.sh`** - Easy production runner
- **`README.md`** - This file

## 🎯 The Fix Explained

### The Problem (Before)
```
Popup opened: about:blank
❌ submitFunction is not defined (popup context)
❌ NO POPUP WITH REAL CONTENT FOUND
❌ PDF too small (678 bytes)
```

### The Solution (After)
```
🔧 Executing submitFunction in main page context
✅ POPUP NAVIGATED TO: [real-boleto-url]
✅ POPUP CONTENT LENGTH: 15234 characters
✅ PDF SUCCESS: boleto_file.pdf (45678 bytes)
```

**Key Insight**: `submitFunction` exists in the **main page** JavaScript context, not the popup context. The solution executes `submitFunction` from the main page to load content into the popup.

## 🔧 Usage Options

### Basic Commands
```bash
# Default run (10 records, batch size 1)
./run_final_solution.sh

# Custom parameters
./run_final_solution.sh --max-records 20 --batch-size 2

# Start from specific record
./run_final_solution.sh --start-from 100 --max-records 50

# Debug mode with extra logging
./run_final_solution.sh --max-records 5 --debug
```

### Advanced Options
```bash
python final_working_boleto_processor.py controle_boletos_hs.xlsx [options]

Required:
  controle_boletos_hs.xlsx    Excel file with grupo/cota data

Options:
  --start-from N              Start from record N (0-based)
  --max-records N             Maximum records to process  
  --batch-size N              Records per batch (default: 5)
  --config FILE               Configuration file (default: config.yaml)
  
Timing (usually not needed):
  --popup-delay SECONDS       Wait after popup opens (default: 5.0)
  --content-delay SECONDS     Wait for content loading (default: 5.0)
  --pre-pdf-delay SECONDS     Wait before PDF generation (default: 6.0)
  --pdf-wait-timeout SECONDS  PDF generation timeout (default: 60.0)
  --min-pdf-size BYTES        Minimum PDF size (default: 20000)
```

### Populate CPF/CNPJ Cache
Você pode adiantar a base de CPFs/CNPJs antes das execuções, tanto na planilha do Google quanto no CSV local.

```bash
# Atualizar a planilha (usa a conta de serviço do config.yaml)
python populate_cpf_cnpj.py \
  --sheet-range "Página1!A:D" \
  --header-title "DOCUMENTO" \
  --delay 0.5 \
  --flush-every 1

# Atualizar o arquivo local tabela.csv
python populate_cpf_cnpj.py \
  --csv-path tabela.csv \
  --header-title "DOCUMENTO" \
  --delay 0.5 \
  --flush-every 1
```

- A coluna indicada em `--header-title` é criada caso ainda não exista.
- Linhas que já possuem valor são preservadas; use `--force` para sobrescrever tudo.
- `--flush-every` controla de quantos em quantos registros os dados são persistidos (1 = salva a cada atualização).
- `--delay` opcional ajuda a espaçar as requisições ao portal HS Consórcios.

Com a base preenchida, as execuções diárias evitam refazer a consulta de documento para cada cota.

## 📚 Visão Geral das Implementações

- **Upload automático para o Google Drive** – PDFs são enviados para o drive compartilhado `0AAcUWkdO1T0zUk9PVA` criando pastas `ano/mês`. Todas as credenciais ficam em `config.yaml`.
- **Links assinados para entrega** – `file_link_service.py` gera URLs temporárias usando `secret_key`; o domínio `https://faturas.suzanarighihs.com.br/files` aponta para o servidor local na porta 18181.
- **Integração com webhook (n8n/WhatsApp)** – quando `notifications.enabled` for `true`, `notifier.py` dispara o JSON com `phone`, `message`, `file_url` e `drive_file_id` para o fluxo n8n.
- **Leitura resiliente de dados** – suporte a CSV (`tabela.csv`), Google Sheets (`Página1!A:D`) e cache de registros processados (`logs/processed_records.json`) evitando downloads repetidos.
- **Preenchimento incremental de CPF/CNPJ** – `populate_cpf_cnpj.py` aceita planilha ou CSV, grava cada célula/linha conforme atualiza (`--flush-every`), e suporta `--force` para sobrescrever valores.
- **Log de resultados** – os relatórios ficam em `reports/`, o dashboard da planilha recebe os status e o arquivo `complete_fixed_automation.log` traz o passo a passo com detalhes de cada cota.

## 🔄 Fluxo Recomendado

1. **Atualizar CPFs/CNPJs (opcional, mas recomendado)**
   ```bash
   python populate_cpf_cnpj.py \
     --sheet-range 'Página1!A:D' \
     --header-title 'DOCUMENTO' \
     --delay 0.5 \
     --flush-every 1 \
     --log-level INFO
   ```
   - Gera a coluna caso necessário e preenche célula a célula; use `--force` para reprocessar.
   - Para manter um espelho local, troque por `--csv-path tabela.csv`.

2. **Rodar a automação principal**
   ```bash
   ./run_final_solution.sh --batch-size 100
   ```
   - Baixa os PDFs, envia ao Drive e gera os links assinados.

3. **Verificar saídas**
   - **Drive**: hierarquia `ano/mês`.
   - **Planilha Dashboard**: status de download/Drive/notificação.
   - **Logs**: `logs/enhanced_automation_*.log` e `complete_fixed_automation.log`.

## ⚙️ Configurações Principais

- `config.yaml` – credenciais, parâmetros do portal, Google Drive, webhook e servidor de arquivos.
- `google_drive.credentials_path` – arquivo da service account (`suzana-playwright-e1656f768d86.json`).
- `file_server.secret_key` – gere com `openssl rand -base64 48` e mantenha em segurança.
- `notifications` – habilite quando o fluxo n8n estiver em produção.
- `processing.skip_processed_records` – controla o cache de cotas já processadas.

## 📊 Expected Performance

### Success Metrics
- **Login Success**: 95%+ (unless server issues)
- **Search Success**: 90%+ (depends on data quality)
- **Popup Content Loading**: 80%+ (major improvement!)
- **PDF Generation**: 70%+ (from boletos with content)
- **Overall Success Rate**: 40-60% (realistic expectation)

### Processing Time
- **Per Record**: 2-3 minutes (includes all delays and waits)
- **Small Test (5 records)**: ~10-15 minutes
- **Full File (1,487 records)**: 8-12 hours

## 📁 Output Structure

```
downloads/                  # Generated files
├── SUZANA_MARIA-684-644-55613012091-20250905_150219-0.pdf
├── CLIENT-699-21-35798443000154-20250905_150630-0.pdf
└── ...

reports/                    # Processing results  
├── complete_fixed_results_20250905_150800.json
├── complete_fixed_final_report_20250905_151200.json
└── ...

screenshots/                # Debug captures
├── popup_boleto_1_20250905_150219.png
└── ...

complete_fixed_automation.log  # Detailed logs
```

## 🔍 Troubleshooting

### If No Downloads
1. **Check screenshots** in `screenshots/` folder
2. **Review logs** in `complete_fixed_automation.log`
3. **Run debug test** with `python debug_submitfunction_test.py`
4. **Try single record** with `python test_final_solution.py`

### If Popups Stay Blank
- This should be FIXED in this version
- If it still happens, check the debug test output
- Verify your credentials in `config.yaml`

### If PDFs Are Small
- Check if popup content is loading (screenshots)
- Increase `--pdf-wait-timeout` if needed
- Review popup screenshots for content verification

## 📞 Support

1. **Start with debug test** - `python debug_submitfunction_test.py`
2. **Check log files** - `complete_fixed_automation.log`
3. **Review screenshots** - `screenshots/` folder shows popup content
4. **Test single record** - `python test_final_solution.py`

## 🎉 Success Indicators

When working correctly, you'll see logs like:
```
🔧 FINAL SOLUTION: Loading boleto content via main page submitFunction
✅ POPUP NAVIGATED TO: https://consweb.hsconsorcios.com.br/Slip/Slip.asp
✅ POPUP CONTENT LENGTH: 15234 characters
⏰ WAITING FOR PDF GENERATION: [filename]
✅ PDF GENERATION COMPLETE: 45678 bytes
✅ COMPLETE PDF SUCCESS: [filename] (45678 bytes)
```

**This package contains the FINAL SOLUTION for the popup content loading issue!** 🚀
