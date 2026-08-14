# Finclass Monitor

> 📊 Automação em Python para acessar a Carteira Finclass, coletar ativos recomendados, gerar arquivos Excel, manter histórico diário e enviar notificações por e-mail quando a carteira for alterada.

**Script principal:** `finclass_carteira.py`

---

## 📋 Índice

1. [O que o script faz](#-o-que-o-script-faz)
2. [Requisitos](#-requisitos)
3. [Estrutura do projeto](#-estrutura-recomendada)
4. [Setup inicial](#-setup-inicial)
5. [Configuração](#-configuração)
6. [Execução](#-execução)
7. [Modos auxiliares](#-modos-auxiliares)
8. [Arquivos gerados](#-arquivos-gerados)
9. [Operação automática](#-operação-automática)
10. [Troubleshooting](#-solução-de-problemas)
11. [Comandos principais](#-comandos-principais)

---

## 🔄 O que o script faz

Em uma execução normal, o script:

- ✅ Carrega as credenciais e configurações do arquivo `.env`
- 🌐 Abre o Chromium com Playwright
- 📍 Acessa https://app.finclass.com/
- 🔑 Faz login quando necessário
- 📂 Abre a área Carteira
- 🏷️ Percorre todas as categorias encontradas na área Ativos
- 📊 Coleta todas as linhas das `recommendation-table`
- 📄 Gera o arquivo da carteira atual
- 📑 Gera um catálogo com nomes dos ativos, fundos e empresas
- 📈 Atualiza o histórico diário
- 🔍 Compara a carteira atual com o último snapshot anterior
- ⚠️ Detecta alterações
- 📧 Envia um e-mail HTML quando houver mudanças
- 🚫 Evita reenviar a mesma notificação várias vezes no mesmo dia

### Fluxo do processo

```
Finclass
    ↓
Playwright (Chromium)
    ↓
Carteira
    ↓
Todas as recommendation-table
    ↓
┌─────────────────────────┐
│ finclass_carteira.xlsx  │
│ finclass_ativos.xlsx    │
│ finclass_historico.xlsx │
└─────────────────────────┘
    ↓
Comparação com snapshot anterior
    ↓
Houve alteração?
├── Não  → Encerra
└── Sim  → Envia e-mail
```

## ✅ Requisitos

| Recurso          | Especificação                        |
| ---------------- | ------------------------------------ |
| **SO**           | Linux/Ubuntu (recomendado)           |
| **Python**       | 3.8 ou superior                      |
| **Ambiente**     | venv                                 |
| **Navegador**    | Chromium (instalado pelo Playwright) |
| **Finclass**     | Conta válida                         |
| **E-mail**       | SMTP (ex: Gmail)                     |
| **Dependências** | playwright, openpyxl, python-dotenv  |

## 📁 Estrutura recomendada

```
finclass/
├── .env                          # Credenciais e configuração (NUNCA fazer commit)
├── .gitignore                    # Ignorar arquivos sensibilidade
├── requirements.txt              # Dependências do projeto
├── README.md                     # Este arquivo
├── finclass_carteira.py          # Script principal
├── finclass_storage_state.json   # Sessão Playwright (auto-gerado)
├── finclass_email_state.json     # Controle de e-mails (auto-gerado)
├── finclass_carteira.xlsx        # Carteira atual (auto-gerado)
├── finclass_ativos.xlsx          # Catálogo de ativos (auto-gerado)
└── finclass_historico.xlsx       # Histórico acumulado (auto-gerado)
```

> **Nota:** Arquivos `.xlsx`, `.json` e `.log` são criados/atualizados automaticamente.

## 🚀 Setup inicial

### Passo 1: Criar ambiente virtual

```bash
cd /home/cmopr/investimentos/finclass

python3 -m venv .venv

source .venv/bin/activate

python -m pip install --upgrade pip
```

### Passo 2: Instalar dependências

```bash
python -m pip install -r requirements.txt
```

Ou diretamente:

```bash
python -m pip install playwright openpyxl python-dotenv
```

### Passo 3: Instalar Chromium

```bash
python -m playwright install chromium
```

Se necessário instalar dependências do sistema:

```bash
python -m playwright install-deps chromium
```

Para algumas instalações:

```bash
sudo .venv/bin/python -m playwright install-deps chromium
```

> ⚠️ Não use `sudo` para instalar o navegador no cache do usuário, apenas para dependências de sistema.

## ⚙️ Configuração

### Arquivo `.env`

O arquivo `.env` deve ficar na mesma pasta do script principal.

#### Exemplo completo:

```env
# ======================================
# Finclass
# ======================================

FINCLASS_EMAIL=seu_email@gmail.com
FINCLASS_PASSWORD="sua_senha_finclass"

# Primeira execução: false
# Após validar a sessão: true
FINCLASS_HEADLESS=false

# ======================================
# Notificações por e-mail
# ======================================

EMAIL_ENABLED=true

SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_SECURITY=starttls

SMTP_USER=seu_email@gmail.com
SMTP_PASSWORD="senha_de_app_do_google"

EMAIL_FROM=seu_email@gmail.com
EMAIL_FROM_NAME="Finclass Monitor"

# Vários destinatários: separados por vírgula
EMAIL_TO=seu_email@gmail.com
```

#### Variáveis disponíveis

| Variável            | Obrigatória    | Função                                                   |
| ------------------- | -------------- | -------------------------------------------------------- |
| `FINCLASS_EMAIL`    | ✅             | E-mail para login na Finclass                            |
| `FINCLASS_PASSWORD` | ✅             | Senha da Finclass                                        |
| `FINCLASS_HEADLESS` | ❌             | `true`: executa sem interface; `false`: mostra navegador |
| `EMAIL_ENABLED`     | ❌             | Liga/desliga notificações                                |
| `SMTP_HOST`         | ✅ (se e-mail) | Servidor SMTP                                            |
| `SMTP_PORT`         | ✅ (se e-mail) | Porta SMTP                                               |
| `SMTP_SECURITY`     | ✅ (se e-mail) | `starttls`, `ssl` ou `none`                              |
| `SMTP_USER`         | ✅ (se e-mail) | Usuário SMTP                                             |
| `SMTP_PASSWORD`     | ✅ (se e-mail) | Senha SMTP / Senha de app                                |
| `EMAIL_FROM`        | ❌             | Remetente; padrão: `SMTP_USER`                           |
| `EMAIL_FROM_NAME`   | ❌             | Nome exibido no e-mail                                   |
| `EMAIL_TO`          | ✅ (se e-mail) | Destinatário(s)                                          |

### Configuração para Gmail

Para Gmail, use:

```env
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_SECURITY=starttls

SMTP_USER=usuario@gmail.com
SMTP_PASSWORD="senha_de_app"
```

> 🔐 **Importante:** Use uma [Senha de app do Google](https://myaccount.google.com/apppasswords), não a senha normal da conta.

### 🔒 Segurança

**Nunca faça commit do arquivo `.env`!**

Exemplo de `.gitignore`:

```
.venv/
**/__pycache__/
*.pyc

.env

finclass_storage_state.json
finclass_email_state.json

finclass_carteira.xlsx
finclass_historico.xlsx
finclass_ativos.xlsx

*.log
```

⚠️ O `finclass_storage_state.json` pode conter dados de sessão do navegador e deve ser tratado como credencial.

## ▶️ Execução

### Execução normal

Com o `.venv` ativado:

```bash
python finclass_carteira.py
```

Saída esperada:

```
Abrindo Carteira...
Coletando: A - AÇÕES
Coletando: R - REAL ESTATE (FUNDOS IMOBILIÁRIOS)
Coletando: C - CAIXA (RENDA FIXA)
Coletando: A - ALTERNATIVOS

Concluído.
Categorias coletadas: 4
Itens coletados: 37
Snapshot histórico: 37 itens
Carteira atual: .../finclass_carteira.xlsx
Catálogo de ativos: .../finclass_ativos.xlsx
Nomes no catálogo: 37
Nomes únicos: ...
Histórico: .../finclass_historico.xlsx
```

### Primeira execução

Na primeira execução, recomenda-se:

```env
FINCLASS_HEADLESS=false
```

```bash
python finclass_carteira.py
```

Se a Finclass solicitar CAPTCHA, confirmação ou outro passo manual, conclua no navegador.

Após autenticação bem-sucedida, o Playwright salva a sessão em: `finclass_storage_state.json`

Para automação posterior, teste:

```env
FINCLASS_HEADLESS=true
```

Após validar que funciona sem interface gráfica, está pronto para agendamento automático.

## 🛠️ Modos auxiliares

### Testar e-mail (SMTP)

```bash
python finclass_carteira.py --test-email
```

Este modo:

- ✅ Não acessa a Finclass
- ✅ Não altera o histórico
- ✅ Não altera o estado de notificações
- ✅ Envia um e-mail de teste ao `EMAIL_TO`

Saída esperada:

```
Testando SMTP smtp.gmail.com:587 (starttls)...
E-mail de teste enviado com sucesso para: usuario@gmail.com
```

Recomendação: Execute este teste **antes de ativar notificações automáticas**.

### Simular uma alteração

```bash
python finclass_carteira.py --simulate-change
```

Este modo:

- ✅ Simula alterações como: `% ALTERADA`, `ADICIONADO`, `REMOVIDO`
- ✅ Envia e-mail com o mesmo layout da notificação real
- ✅ Não acessa a Finclass
- ✅ Não altera arquivos Excel
- ✅ Não altera o estado de notificações

Saída esperada:

```
Enviando simulação de alteração para: usuario@gmail.com
E-mail de simulação enviado com sucesso.
Nenhum arquivo Excel ou estado de notificação foi alterado.
```

### Ver opções disponíveis

```bash
python finclass_carteira.py --help
```

## 📊 Arquivos gerados

### `finclass_carteira.xlsx`

Representa a **carteira atual** (snapshot do dia).

- 🔄 Recriado a cada execução
- 📋 Uma aba `Consolidado`
- 📑 Uma aba para cada categoria encontrada

**Exemplos de categorias:**

- A - AÇÕES
- R - REAL ESTATE (FUNDOS IMOBILIÁRIOS)
- C - CAIXA (RENDA FIXA)
- A - ALTERNATIVOS

> ⚠️ Não é histórico. Representa apenas o último estado coletado.

### `finclass_ativos.xlsx`

Catálogo atual de ativos/fundos/empresas.

- 🔄 Recriado a cada execução
- 📋 Aba **Todos**: lista completa de itens
- 📑 Aba **Nomes únicos**: remove duplicidades

#### Aba "Todos"

Contém uma linha para cada item coletado:

| Categoria | Nome do ativo/fundo/empresa | Código/Ticker |
| --------- | --------------------------- | ------------- |
| A - AÇÕES | BR PARTNERS                 | BRBI11        |
| A - AÇÕES | GRUPO MATEUS                | GMAT3         |

#### Aba "Nomes únicos"

Remove duplicidades usando: `Categoria + Nome + Código`

Serve como catálogo de nomes distintos na carteira atual.

### `finclass_historico.xlsx`

O **arquivo principal** de acompanhamento. Não perde dados anteriores.

#### Aba "Histórico"

Registra para cada execução:

| Data       | Categoria | % da Carteira | Nome         | Código |
| ---------- | --------- | ------------- | ------------ | ------ |
| 14/08/2026 | A - AÇÕES | 1,00%         | BR PARTNERS  | BRBI11 |
| 14/08/2026 | A - AÇÕES | 2,50%         | GRUPO MATEUS | GMAT3  |
| 15/08/2026 | A - AÇÕES | 1,50%         | BR PARTNERS  | BRBI11 |

**Execução repetida no mesmo dia:**

- Existe apenas um snapshot por data
- Se executado várias vezes no mesmo dia, substitui apenas o snapshot desse dia
- Dias anteriores são preservados

#### Aba "Alterações"

Registra mudanças detectadas:

| Data       | Ativo  | Alteração  | % anterior | % atual |
| ---------- | ------ | ---------- | ---------- | ------- |
| 15/08/2026 | BRBI11 | % ALTERADA | 1,00%      | 1,50%   |
| 15/08/2026 | TEST3  | ADICIONADO | —          | 2,00%   |
| 15/08/2026 | TEST11 | REMOVIDO   | 3,00%      | —       |

**Tipos de alteração:**

- `ADICIONADO`: novo ativo na carteira
- `REMOVIDO`: ativo saiu da carteira
- `% ALTERADA`: percentual mudou
- `NOME ALTERADO`: nome do ativo foi alterado

> 💡 O primeiro snapshot é tratado como base inicial e não marca todos os ativos como "ADICIONADO".

## 📧 Notificações

### Quando o e-mail é enviado

```
Coleta
  ↓
Atualiza histórico
  ↓
Compara com snapshot anterior
  ↓
Alterações detectadas?
├── 0 alterações → Nenhum e-mail
└── > 0 alterações → Envia e-mail
```

**Se não houver alteração:**

```
Alterações detectadas: 0
E-mail: nenhuma alteração; nenhum e-mail enviado.
```

**Se houver alteração:**
O e-mail apresenta uma tabela com:

- Tipo de alteração
- Categoria
- Nome do ativo/fundo/empresa
- Ticker/código
- Percentual anterior
- Percentual atual

### Prevenção de e-mails duplicados

O arquivo `finclass_email_state.json` guarda a **assinatura da última notificação enviada**.

**Exemplo:**

```
08:00 → alteração A → envia
10:00 → mesma alteração A → não envia novamente
13:00 → nova situação A+B → envia (assinatura diferente)
```

> 💡 A assinatura só é gravada após o SMTP confirmar o envio com sucesso.

## ✔️ Integridade dos dados

### Validação da quantidade de itens

O script possui verificações de **consistência automáticas**.

Se a Finclass retornar:

```
Itens coletados: 37
```

O histórico deve apresentar:

```
Snapshot histórico: 37 itens
```

E o catálogo completo:

```
Nomes no catálogo: 37
```

**Se os números não coincidirem:**

- ❌ O script gera erro e **não grava** um histórico incompleto
- ✅ Isso é proposital para **evitar perda de dados**

---

## 🤖 Operação automática

### Modo headless

Depois de validar login e sessão:

```env
FINCLASS_HEADLESS=true
```

Teste:

```bash
python finclass_carteira.py
```

Se funcionar sem abrir interface gráfica, está pronto para automação/scheduler.

### Agendamento com cron

Para executar **todos os dias às 08:00**:

```bash
crontab -e
```

Adicione:

```bash
0 8 * * * cd /home/cmopr/investimentos/finclass && /home/cmopr/investimentos/finclass/.venv/bin/python /home/cmopr/investimentos/finclass/finclass_carteira.py >> /home/cmopr/investimentos/finclass/finclass.log 2>&1
```

Verifique o log:

```bash
tail -f /home/cmopr/investimentos/finclass/finclass.log
```

> ⚠️ Cron local só funciona enquanto a máquina estiver **ligada**.  
> Para automação 24/7, hospede em um servidor/VPS ou ambiente cloud.

## 🔧 Solução de problemas

### ModuleNotFoundError: No module named 'dotenv'

Ative o `.venv`:

```bash
source .venv/bin/activate
```

Instale:

```bash
python -m pip install python-dotenv
```

Ou:

```bash
python -m pip install -r requirements.txt
```

---

### Erro: "Não foi possível resolver a importação 'playwright.sync_api'"

Confirme o ambiente:

```bash
which python
python -c "import sys; print(sys.executable)"
python -c "import playwright.sync_api; print(playwright.sync_api.__file__)"
```

**No VS Code:**

1. `Ctrl + Shift + P`
2. Python: Select Interpreter
3. Selecione `.venv/bin/python`
4. Developer: Reload Window

---

### Chromium não instalado

Teste:

```bash
python -m playwright install chromium
```

Se necessário, com timeout aumentado:

```bash
PLAYWRIGHT_DOWNLOAD_CONNECTION_TIMEOUT=120000 python -m playwright install chromium
```

---

### EACCES: permission denied em ~/.cache/ms-playwright

Corrija o proprietário:

```bash
sudo chown -R "$USER":"$(id -gn)" ~/.cache/ms-playwright
chmod -R u+rwX ~/.cache/ms-playwright
```

Depois:

```bash
python -m playwright install chromium
```

---

### Problemas de login da Finclass

Altere temporariamente:

```env
FINCLASS_HEADLESS=false
```

Execute:

```bash
python finclass_carteira.py
```

Conclua manualmente qualquer confirmação (CAPTCHA, 2FA, etc).

Depois teste novamente com:

```env
FINCLASS_HEADLESS=true
```

---

### Problemas de e-mail

Primeiro execute:

```bash
python finclass_carteira.py --test-email
```

Se falhar, revise no `.env`:

- `SMTP_HOST`
- `SMTP_PORT`
- `SMTP_SECURITY`
- `SMTP_USER`
- `SMTP_PASSWORD`
- `EMAIL_FROM`
- `EMAIL_TO`

**Para Gmail, confirme:**

```env
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_SECURITY=starttls
SMTP_PASSWORD="senha_de_app_válida"
```

## 📚 Comandos principais

| Comando                                         | Função                  |
| ----------------------------------------------- | ----------------------- |
| `source .venv/bin/activate`                     | Ativar ambiente virtual |
| `python finclass_carteira.py`                   | Execução normal         |
| `python finclass_carteira.py --test-email`      | Testar SMTP             |
| `python finclass_carteira.py --simulate-change` | Simular mudança         |
| `python finclass_carteira.py --help`            | Ver opções disponíveis  |
| `python -m pip install -r requirements.txt`     | Instalar dependências   |
| `python -m playwright install chromium`         | Instalar Chromium       |
| `tail -f finclass.log`                          | Ver log em tempo real   |
| `crontab -e`                                    | Editar agendamento cron |

---

## ✨ Resultado esperado

Uma execução normal bem-sucedida deve manter:

| Arquivo                       | Propósito                                  |
| ----------------------------- | ------------------------------------------ |
| `finclass_carteira.xlsx`      | Estado atual completo                      |
| `finclass_ativos.xlsx`        | Catálogo atual de ativos/fundos/empresas   |
| `finclass_historico.xlsx`     | Snapshots acumulados + alterações          |
| `finclass_storage_state.json` | Sessão Playwright (não expor publicamente) |
| `finclass_email_state.json`   | Controle de notificações já enviadas       |

### Objetivo operacional

```
Executar diariamente
       ↓
Coletar 100% da Carteira Finclass
       ↓
Preservar o histórico
       ↓
Identificar qualquer alteração
       ↓
Enviar notificação somente quando necessário
```

---

## 📝 Notas finais

- ✅ Sempre use um ambiente virtual (`.venv`)
- 🔒 Nunca faça commit de `.env` ou arquivos de estado
- 📊 O histórico é acumulativo e nunca é perdido
- 📧 E-mails duplicados são evitados automaticamente
- 🚀 Após validar, está pronto para automação com cron/scheduler
- 💾 Faça backup do `finclass_historico.xlsx` periodicamente

---

**Versão:** finclass_carteira.py  
**Última atualização:** Agosto 2026
