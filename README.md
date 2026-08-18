# Automação de Criação de Títulos - Senior ERP

Esta automação foi desenvolvida em Python para automatizar o cadastro e lançamento de títulos a receber no formulário **F301TCR** do **Senior ERP (Sapiens)**.

---

## Bibliotecas Utilizadas

Abaixo está a relação detalhada das bibliotecas utilizadas no projeto e a responsabilidade de cada uma na automação:

| Biblioteca | Tipo | Função no Projeto |
|------------|------|-------------------|
| **`PyAutoGUI`** | Terceiros | **Automação de Interface (GUI/RPA):** Responsável pelos cliques do mouse em coordenadas específicas, envio de atalhos e teclas (`Tab`, `Enter`, `Alt+A`), capturas de tela para verificação visual por pixels (`pyautogui.screenshot`) e mecanismo de emergência (*Fail-Safe*). |
| **`Pyperclip`** | Terceiros | **Manipulação do Clipboard:** Utilizado para copiar textos via `pyperclip.copy()` e colar com `Ctrl+V`. Essa estratégia garante a inserção 100% perfeita de caracteres especiais e acentuação da língua portuguesa (`ç`, `ã`, `é`, `º`), evitando falhas comuns na digitação direta por layouts de teclado ABNT. |
| **`OpenPyXL`** | Terceiros | **Relatórios em Excel (`.xlsx`):** Utilizado no módulo de saída para criar e atualizar as planilhas estilizadas em `outputs/` (`titulos_processados_*.xlsx` e `titulos_pulados_*.xlsx`), aplicando cores de cabeçalho, negrito, colunas ajustadas e indicação visual por cores de status. |
| **`Pandas` / `xlrd`** | Terceiros | **Processamento de Dados:** Utilizados no script `scripts/processar_planilha.py` para ler a planilha Excel original de títulos, realizar a limpeza dos dados, tratamento de tipos e converter a lista de registros para o formato `titulos_tratados.json`. |
| **`JSON`** | Nativa | **Estrutura de Dados e Checkpoints:** Gerencia a leitura das configurações (`config/config.json`), a fila de dados a serem processados (`titulos_tratados.json`) e o estado de progresso (`outputs/progresso.json`) para retomada automática. |
| **`DateTime` / `Time`** | Nativa | **Controle de Tempo e Datas:** Formatação de datas nos padrões aceitos pelo ERP (`DDMMYYYY` / `MMDDYYYY`), geração de timestamps em logs e gerenciamento de pausas (*delays*) integradas ao Fail-Safe. |
| **`Pathlib` / `OS` / `SYS`** | Nativa | **Gerenciamento de Arquivos e Pastas:** Garante a manipulação de caminhos absolutos/relativos de forma portável entre sistemas operacionais Windows. |
| **`CSV`** | Nativa | **Backup Universal de Dados:** Garante a gravação dos relatórios de baixas e títulos pulados no formato CSV (`utf-8-sig` com separador `;`), servindo como alternativa caso o pacote Excel não esteja instalado. |

---

## Estrutura do Projeto

```text
automacao-criacao-titulos/
├── config/
│   └── config.json           # Mapeamento de coordenadas (X, Y), filiais e delays
├── data/
│   ├── titulos_tratados.json # Fila de títulos extraída da planilha original
│   └── *.xlsx                # Planilha de entrada com os dados dos títulos
├── outputs/                  # Pasta gerada automaticamente para os resultados
│   ├── log_execucao_*.txt    # Logs de execução com timestamps
│   ├── titulos_processados_* # Relatório de títulos concluídos/processados (Excel e CSV)
│   ├── titulos_pulados_*     # Relatório de títulos pulados (ex: cliente duplicado)
│   └── progresso.json        # Arquivo de checkpoint para retomada automática
├── scripts/
│   ├── get_mouse_position.py # Utilitário para captura de coordenadas do mouse
│   └── processar_planilha.py # Script para conversão da planilha Excel em JSON
├── main.py                   # Script principal de automação
└── README.md                 # Documentação do projeto
```

---

## Pré-Requisitos e Instalação

Certifique-se de utilizar o Python 3.8+ e instalar as bibliotecas necessárias:

```bash
pip install pyautogui pyperclip openpyxl pandas xlrd
```

---

## Como Executar

1. **Preparação dos Dados:**
   Caso haja uma nova planilha de títulos, execute o script de tratamento:
   ```bash
   python scripts/processar_planilha.py
   ```
2. **Execução da Automação:**
   Inicie a automação principal:
   ```bash
   python main.py
   ```
3. **Mude o foco:** Você terá **5 segundos** para alternar para a janela do **Senior ERP**.

---

## Fluxo de Funcionamento Detalhado

### 1. Troca Inteligente de Filial
- O robô verifica a filial (projeto) do título atual.
- Se a filial for diferente da anterior:
  - Clica na foto/menu do usuário -> `Trocar Filial`.
  - Clica no início da lista de filiais e navega até a posição da filial correspondente.
  - Confirma com `Enter` e navega pelos menus até a tela de títulos.
- Se a filial for a mesma, o robô abre diretamente a tela de **Títulos/Manutenção**.

---

### 2. Preenchimento do Formulário (F301TCR)

A automação executa a inserção dos dados no formulário seguindo a ordem abaixo:

| # | Campo / Etapa | Ação Executada |
|---|---------------|----------------|
| **1** | Nº Título | Digita o número do título + `Tab` |
| **2** | Tipo Título | Digita o tipo (ex: `SGE`) + `Tab` |
| **3** | Transação | Digita o código da transação (ex: `90335`) + `Tab` |
| **4** | Data Emissão | Digita a data no formato padrão `DDMMYYYY` + `Tab` |
| **5** | Data Entrada | Digita a data no formato padrão `DDMMYYYY` + `Tab` |
| **6** | Pesquisa de Cliente | Clica no botão `pesquisa_cliente` |
| **7** | Filtro de Cliente | Pressiona `Alt+A`, insere o **CPF/CNPJ** do cliente (obtido do campo `cpf_cnpj` da planilha) e clica em `filtrar_cliente` |
| **8** | **Validação Prévia & Duplicidade** | Inspeciona os pixels pretos na grade em `(X: 60, Y: 240)`. Se a grade estiver **vazia (0 pixels pretos)**, o cliente **não existe**: pressiona `ESC` (2x) e marca o título como `PULADO - CLIENTE NÃO ENCONTRADO`. Se houver múltiplos registros iguais, clica em cancelar e marca como `PULADO - CLIENTE DUPLICADO`. |
| **9** | Confirmação do Cliente | Executa um **duplo clique** em `confirmar_cliente` |
| **10** | Navegação | Pressiona **4x `Tab`** para alcançar o campo de observações |
| **11** | Observação | Cola o texto de observação via clipboard (`Ctrl+V`) |
| **12** | Pular Natureza | Pressiona **2x `Tab`** |
| **13** | Projeto | Digita o código do projeto + `Tab` |
| **14** | Fase | Digita a fase (ex: `1`) + `Tab` |
| **15** | Conta Financeira | Digita o código da conta financeira + `Tab` |
| **16** | Centro de Custo | Digita o código do centro de custo + `Tab` |
| **17** | Data Vencimento | Digita no formato brasileiro padrão `DDMMYYYY` + `Tab` |
| **18** | Valor Original | Digita o valor formatado com vírgula (ex: `269,05`) |
| **19** | Navegação | Pressiona **8x `Tab`** para alcançar o campo Data Provável |
| **20** | Data Provável | Digita a data no formato brasileiro padrão `DDMMYYYY` + `Tab` |
| **21** | **Finalização** | Pressiona **26x `Tab`** restantes para efetivar a gravação do registro |

---

## Segurança, Resiliência e Logs

- **Fail-Safe (Parada de Emergência):** Mover o mouse para o **canto superior esquerdo da tela** (`X <= 15, Y <= 15`) interrompe a automação imediatamente.
- **Mecanismo de Retomada (Checkpoint):** Salva o progresso no arquivo `outputs/progresso.json`. Caso a automação pare (por parada manual ou erro), a próxima execução **retoma exatamente de onde parou**.
- **Relatórios em `outputs/`:**
  - **Logs em TXT:** Registro detalhado com timestamps das operações.
  - **Planilha Excel/CSV:** Grava o status de cada título (`SUCESSO`, `PULADO - CLIENTE NÃO ENCONTRADO` ou `PULADO - CLIENTE DUPLICADO`).

---

## Controle de Revisão da Documentação

| Versão | Data | Autor | Descrição das Alterações |
|--------|------|-------|--------------------------|
| 1.0 | 14/08/2026 | Vinícius Cisneiros | Documentação completa do fluxo de automação do Senior ERP, mapeamento de campos e bibliotecas |
| 1.1 | 18/08/2026 | Vinícius Cisneiros | Atualização para busca de cliente por CPF/CNPJ, validação prévia de existência na grade (`X: 60, Y: 240`), formatação de datas no padrão `DDMMYYYY` e registro de títulos pulados por cliente não encontrado |
