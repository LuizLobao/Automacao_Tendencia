Ferramenta para Automatização das Tendências Diárias

Objetivo:
 - Criar ferramenta visual para facilitar o processo diário de cálculo e atualização das tendências
 
Funcionalidades a Desenvolver:
 - [ ] Layout em GUI para facilitar o processo de atualização
 - [ ] Puxar a data/hora de atualização dos arquivos necessários
    - [X] No Monitor de Cargas (site interno)
    - [X] No Diretório de Rede
    - [ ] No Sharepoint
 - [X] Preparar Arquivo EXCEL para área cliente
 - [X] Enviar Excel por e-mail
 - [X] Executar sequencia de queries SQL
 

Funcionalidades Futuras:
 - [ ] Log de data/hora de execução de cada rotina (inicio e fim)
 - [ ] Comparação das bases de D0 vs D-1 antes de seguir o processo
 - [ ] Ao colocar a verificação de Data em LOOP, perguntar se quer continuar o LOOP ou sair - caso nao tenha resposta em X segundos, continuar o Loop.

 Pré-Requisitos:
  - no Python:
      - Pandas (pip install pandas)
      - Playwright (pip install pytest-playwright  /  playwright install)
      - TQDM (pip install tqdm)
      - PIL (pip install Pillow)
      - Openpyxl (pip install openpyxl)
      - Requests (pip install requests)
      - PyODBC (pip install pyodbc)
      - PyWin23 (pip install pywin32)
      -
      -

  - diretórios de rede mapeados:
      - \\naspc01\ger_desempenho_operacional$\Report\Report_Demonstrativo_FTTH ---> Y:\
      - \\netprd01\Plan_Vendas_InfoGer ----> S:\


