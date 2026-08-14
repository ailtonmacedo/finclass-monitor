Finclass Monitor

Automação em Python para acessar a Carteira Finclass, coletar os ativos recomendados, gerar arquivos Excel, manter um histórico diário e enviar notificações por e-mail quando a carteira for alterada.

Script documentado: finclass_carteira.py
A versão atual também pode estar com o nome finclass_carteira_v8_catalogo_ativos.py.

1. O que o script faz

Em uma execução normal, o script:

Carrega as credenciais e configurações do arquivo .env.

Abre o Chromium com Playwright.

Acessa <https://app.finclass.com/>.

Faz login quando necessário.

Abre a área Carteira.

Percorre todas as categorias encontradas na área Ativos.

Coleta todas as linhas das recommendation-table.

Gera o arquivo da carteira atual.

Gera um catálogo com os nomes dos ativos, fundos e empresas.

Atualiza o histórico diário.

Compara a carteira atual com o último snapshot anterior.

Detecta alterações.

Envia um e-mail HTML quando houver mudanças.

Evita reenviar a mesma notificação várias vezes no mesmo dia.

Fluxo resumido:

Finclass
↓
Playwright
↓
Carteira
↓
Todas as recommendation-table
↓
┌───────────────────────────────┐
│ finclass_carteira.xlsx │
│ finclass_ativos.xlsx │
│ finclass_historico.xlsx │
└───────────────────────────────┘
↓
Comparação com snapshot anterior
↓
Houve alteração?
├── Não → encerra
└── Sim → envia e-mail

2. Requisitos

Recomendado:

Linux/Ubuntu

Python 3.8 ou superior

venv

Chromium instalado pelo Playwright

Conta Finclass válida

Conta de e-mail SMTP, por exemplo Gmail

3. Estrutura recomendada

finclass/
├── .env
├── .gitignore
├── requirements.txt
├── finclass_carteira.py
├── finclass_storage_state.json
├── finclass_email_state.json
├── finclass_carteira.xlsx
├── finclass_ativos.xlsx
└── finclass_historico.xlsx

Os arquivos .xlsx, finclass_storage_state.json e finclass_email_state.json são criados/atualizados automaticamente.

4. Criando o ambiente virtual

Entre na pasta do projeto:

cd /home/cmopr/investimentos/finclass

Crie o ambiente virtual:

python3 -m venv .venv

Ative:

source .venv/bin/activate

Atualize o pip:

python -m pip install --upgrade pip

5. Dependências

Um requirements.txt mínimo:

playwright
openpyxl
python-dotenv

Instale:

python -m pip install -r requirements.txt

Ou diretamente:

python -m pip install playwright openpyxl python-dotenv

Instale o Chromium:

python -m playwright install chromium

Se o Linux informar que faltam bibliotecas do sistema:

python -m playwright install-deps chromium

Em algumas instalações pode ser necessário:

sudo .venv/bin/python -m playwright install-deps chromium

Não use sudo para instalar o navegador no cache do usuário, salvo quando o comando for especificamente para dependências de sistema.

6. Configuração do .env

O .env deve ficar na mesma pasta do script.

Exemplo:

# ------------------------------------------------------------

# Finclass

# ------------------------------------------------------------

FINCLASS_EMAIL=seu_email@gmail.com
FINCLASS_PASSWORD="sua_senha_finclass"

# Primeira execução: false.

# Após validar a sessão, pode ser true para automação.

FINCLASS_HEADLESS=false

# ------------------------------------------------------------

# Notificações por e-mail

# ------------------------------------------------------------

EMAIL_ENABLED=true

SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_SECURITY=starttls

SMTP_USER=seu_email@gmail.com
SMTP_PASSWORD="senha_de_app_do_google"

EMAIL_FROM=seu_email@gmail.com
EMAIL_FROM_NAME="Finclass Monitor"

# Vários destinatários podem ser separados por vírgula.

EMAIL_TO=seu_email@gmail.com

Variáveis disponíveis

Variável

Obrigatória

Função

FINCLASS_EMAIL

Sim

E-mail usado no login da Finclass

FINCLASS_PASSWORD

Sim

Senha da Finclass

FINCLASS_HEADLESS

Não

true executa o Chromium sem interface gráfica

EMAIL_ENABLED

Não

Liga/desliga notificações de alterações

SMTP_HOST

Sim para e-mail

Servidor SMTP

SMTP_PORT

Sim para e-mail

Porta SMTP

SMTP_SECURITY

Sim para e-mail

starttls, ssl ou none

SMTP_USER

Sim para e-mail

Usuário SMTP

SMTP_PASSWORD

Sim para e-mail

Senha SMTP / Senha de app

EMAIL_FROM

Não

Remetente; padrão é SMTP_USER

EMAIL_FROM_NAME

Não

Nome apresentado no e-mail

EMAIL_TO

Sim para e-mail

Destinatário(s), separados por vírgula

7. Gmail

Para Gmail, use:

SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_SECURITY=starttls

No SMTP_PASSWORD, utilize uma Senha de app do Google, e não a senha normal da conta.

Exemplo:

SMTP_USER=usuario@gmail.com
SMTP_PASSWORD="senha_de_app"

A conta precisa estar configurada adequadamente para gerar Senhas de app.

8. Segurança

Nunca faça commit do .env.

Exemplo de .gitignore:

.venv/
**pycache**/
\*.pyc

.env

finclass_storage_state.json
finclass_email_state.json

finclass_carteira.xlsx
finclass_historico.xlsx
finclass_ativos.xlsx

\*.log

O finclass_storage_state.json pode conter dados de sessão do navegador e deve ser tratado como credencial.

Uso

9. Execução normal

Com o .venv ativo:

python finclass_carteira.py

Ou:

.venv/bin/python finclass_carteira.py

O script deve apresentar algo semelhante a:

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

10. Primeira execução

Na primeira execução, recomenda-se:

FINCLASS_HEADLESS=false

Execute:

python finclass_carteira.py

Caso a Finclass solicite CAPTCHA, confirmação ou outro passo manual, conclua no navegador.

Após uma autenticação bem-sucedida, o Playwright salva a sessão em:

finclass_storage_state.json

Para uma execução automática posterior, teste:

FINCLASS_HEADLESS=true

e execute novamente:

python finclass_carteira.py

A execução deve funcionar sem abrir uma janela gráfica.

Modos auxiliares

11. Testar o e-mail

O modo:

python finclass_carteira.py --test-email

faz somente o teste SMTP.

Ele:

não acessa a Finclass;

não altera o histórico;

não altera o estado de notificações;

envia um e-mail de teste ao EMAIL_TO.

Saída esperada:

Testando SMTP smtp.gmail.com:587 (starttls)...
E-mail de teste enviado com sucesso para: usuario@gmail.com

Esse é o primeiro teste a fazer antes de ativar notificações automáticas.

12. Simular uma alteração

Use:

python finclass_carteira.py --simulate-change

Esse modo envia um e-mail utilizando o mesmo layout da notificação real e simula alterações como:

% ALTERADA;

ADICIONADO;

REMOVIDO.

O modo de simulação:

não acessa a Finclass;

não altera finclass_carteira.xlsx;

não altera finclass_historico.xlsx;

não altera finclass_email_state.json.

Saída esperada:

Enviando simulação de alteração para: usuario@gmail.com
E-mail de simulação enviado com sucesso.
Nenhum arquivo Excel ou estado de notificação foi alterado.

Arquivos gerados

13. finclass_carteira.xlsx

Representa a carteira atual.

É recriado a cada execução.

Contém:

uma aba Consolidado;

uma aba para cada categoria encontrada na Finclass.

Exemplos de categorias:

A - AÇÕES
R - REAL ESTATE (FUNDOS IMOBILIÁRIOS)
C - CAIXA (RENDA FIXA)
A - ALTERNATIVOS

Esse arquivo não é histórico. Ele representa somente o último estado coletado.

14. finclass_ativos.xlsx

Catálogo atual de ativos/fundos/empresas.

É recriado a cada execução.

Aba Todos

Contém uma linha para cada item coletado:

Categoria

Nome do ativo / fundo / empresa

Código / Ticker

A - AÇÕES

BR PARTNERS

BRBI11

A - AÇÕES

GRUPO MATEUS

GMAT3

A quantidade desta aba deve ser compatível com Itens coletados.

Aba Nomes únicos

Remove duplicidades utilizando:

Categoria + Nome + Código

Serve como catálogo de nomes distintos disponíveis na carteira atual.

15. finclass_historico.xlsx

Esse é o arquivo principal de acompanhamento.

Ele não perde os dias anteriores.

Possui as abas:

Histórico
Alterações

Aba Histórico

Registra:

Data
Categoria
% da Carteira
Nome da empresa/ativo/fundo
Código/Ticker

Exemplo:

Data

Categoria

% da Carteira

Nome

Código

14/08/2026

A - AÇÕES

1,00%

BR PARTNERS

BRBI11

14/08/2026

A - AÇÕES

2,50%

GRUPO MATEUS

GMAT3

15/08/2026

A - AÇÕES

1,50%

BR PARTNERS

BRBI11

Execução repetida no mesmo dia

Existe apenas um snapshot por data.

Se o script for executado novamente no mesmo dia:

08:00 → snapshot do dia
13:00 → substitui somente o snapshot desse mesmo dia
20:00 → substitui novamente somente esse dia

Os dias anteriores são preservados.

Exemplo:

14/08 → preservado
15/08 → preservado
16/08 → atualizado pela última execução de 16/08

16. Comparação histórica

Cada novo snapshot é comparado com a última data anterior existente.

Exemplo:

14/08 → cria base inicial
15/08 → compara com 14/08
16/08 → compara com 15/08

Se o script não executar no dia 15:

14/08 → base
16/08 → compara com 14/08

A comparação não exige dias consecutivos.

17. Alterações detectadas

A aba Alterações pode registrar:

ADICIONADO
REMOVIDO
% ALTERADA
NOME ALTERADO

Exemplo:

Data

Ativo

Alteração

% anterior

% atual

15/08/2026

BRBI11

% ALTERADA

1,00%

1,50%

15/08/2026

TEST3

ADICIONADO

—

2,00%

15/08/2026

TEST11

REMOVIDO

3,00%

—

O primeiro snapshot é tratado como base inicial e não gera falsamente todos os ativos como ADICIONADO.

Notificações

18. Quando o e-mail é enviado

No modo normal:

Coleta
↓
Atualiza histórico
↓
Compara com snapshot anterior
↓
Alterações?
├── 0 → nenhum e-mail
└── > 0 → envia e-mail

Se não houver alteração:

Alterações detectadas: 0
E-mail: nenhuma alteração; nenhum e-mail enviado.

Se houver alteração, o e-mail apresenta uma tabela com:

tipo de alteração;

categoria;

nome do ativo/fundo/empresa;

ticker/código;

percentual anterior;

percentual atual.

19. Prevenção de e-mails duplicados

O arquivo:

finclass_email_state.json

guarda a assinatura da última notificação enviada.

Exemplo:

08:00 → alteração A → envia
10:00 → mesma alteração A → não envia novamente
13:00 → nova situação A+B → envia

A assinatura só é gravada depois que o SMTP confirma o envio.

Integridade dos dados

20. Validação da quantidade de itens

O script possui verificações de consistência.

Se a Finclass retornar:

Itens coletados: 37

o histórico também deve apresentar:

Snapshot histórico: 37 itens

O catálogo completo também deve ter:

Nomes no catálogo: 37

Se os números não coincidirem, o script gera erro em vez de gravar silenciosamente um histórico incompleto.

Isso é proposital para evitar perda de dados.

Operação automática

21. Execução em modo headless

Depois de confirmar que login e sessão estão funcionando:

FINCLASS_HEADLESS=true

Teste:

python finclass_carteira.py

Se funcionar, o script está preparado para scheduler/cron.

22. Agendamento local com cron

Exemplo: executar todos os dias às 08:00.

Abra:

crontab -e

Adicione:

0 8 \* \* \* cd /home/cmopr/investimentos/finclass && /home/cmopr/investimentos/finclass/.venv/bin/python /home/cmopr/investimentos/finclass/finclass_carteira.py >> /home/cmopr/investimentos/finclass/finclass.log 2>&1

Verifique o log:

tail -f /home/cmopr/investimentos/finclass/finclass.log

O cron local só funciona enquanto essa máquina estiver ligada.

Para executar com o computador desligado, hospede o projeto em um servidor/VPS ou outro ambiente cloud que fique ativo continuamente.

Solução de problemas

23. ModuleNotFoundError: No module named 'dotenv'

Ative o .venv:

source .venv/bin/activate

Instale:

python -m pip install python-dotenv

Ou:

python -m pip install -r requirements.txt

24. Não foi possível resolver a importação "playwright.sync_api"

Confirme o ambiente:

which python
python -c "import sys; print(sys.executable)"
python -c "import playwright.sync_api; print(playwright.sync_api.**file**)"

No VS Code:

Ctrl + Shift + P
→ Python: Select Interpreter
→ .venv/bin/python

Depois:

Developer: Reload Window

25. Chromium não instalado

Teste:

python -m playwright install chromium

Se necessário:

PLAYWRIGHT_DOWNLOAD_CONNECTION_TIMEOUT=120000 \
python -m playwright install chromium

26. Erro de permissão em ~/.cache/ms-playwright

Exemplo:

EACCES: permission denied

Corrija o proprietário da pasta:

sudo chown -R "$USER":"$(id -gn)" ~/.cache/ms-playwright
chmod -R u+rwX ~/.cache/ms-playwright

Depois:

python -m playwright install chromium

27. Problemas de login da Finclass

Troque temporariamente:

FINCLASS_HEADLESS=false

Execute:

python finclass_carteira.py

Conclua manualmente qualquer confirmação apresentada.

Depois teste novamente em:

FINCLASS_HEADLESS=true

28. Problemas de e-mail

Primeiro execute:

python finclass_carteira.py --test-email

Se o teste falhar, revise:

SMTP_HOST
SMTP_PORT
SMTP_SECURITY
SMTP_USER
SMTP_PASSWORD
EMAIL_FROM
EMAIL_TO

Para Gmail:

SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_SECURITY=starttls

Use uma Senha de app válida.

Comandos principais

29. Resumo rápido

Ativar ambiente

source .venv/bin/activate

Execução normal

python finclass_carteira.py

Testar SMTP

python finclass_carteira.py --test-email

Simular mudança

python finclass_carteira.py --simulate-change

Ver opções disponíveis

python finclass_carteira.py --help

Instalar dependências

python -m pip install -r requirements.txt

Instalar Chromium

python -m playwright install chromium

Resultado esperado

Uma execução normal bem-sucedida deve manter:

finclass_carteira.xlsx
→ estado atual completo

finclass_ativos.xlsx
→ catálogo atual de ativos/fundos/empresas

finclass_historico.xlsx
→ snapshots acumulados + alterações

finclass_storage_state.json
→ sessão Playwright

finclass_email_state.json
→ controle de notificações já enviadas

O objetivo operacional é:

Executar diariamente
↓
Coletar 100% da Carteira Finclass
↓
Preservar o histórico
↓
Identificar qualquer alteração
↓
Enviar notificação somente quando necessário

Executar diariamente
       ↓
Coletar 100% da Carteira Finclass
       ↓
Preservar o histórico
       ↓
Identificar qualquer alteração
       ↓
Enviar notificação somente quando necessário
