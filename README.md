# Sistema de Relatórios

Este projeto oferece uma interface gráfica simples para gerar relatórios de produção a partir de planilhas Excel. Os scripts estão organizados em módulos dentro do pacote `relatorios` e são acionados por `main.py`.

## Objetivo

Concentrar rotinas de geração de relatórios para facilitar o processamento de dados de produção e tecidos. O sistema executa diferentes processadores que transformam arquivos de entrada em relatórios consolidados no formato `.xlsx`.

## Execução do `main.py`

1. Instale as dependências do arquivo `requirements.txt`:

   ```bash
   pip install -r requirements.txt
   ```

2. Execute o programa principal:

   ```bash
   python main.py
   ```

   Opcionalmente, utilize `--test-duration SEG` para executar a janela por um tempo limitado (em segundos), recurso útil para testes automatizados.

Ao iniciar, uma janela exibirá os relatórios disponíveis. Escolha o desejado e siga as instruções da interface.

## Relatórios Disponíveis

- **Produção de Camas** (`relatorios.producao.processador`)
  - Processa um arquivo ZIP contendo planilhas de produção.
  - Gera um arquivo Excel com abas separadas para `BOX` e `BAU`.

- **Tecidos por Cores** (`relatorios.tecido.processador`)
  - Lê o relatório de produção gerado e calcula médias diárias e mensais por tipo e cor de tecido.
  - Cria três planilhas: médias diárias, médias mensais e média geral.

- **Média de Produção** (`relatorios.media_producao.processador`)
  - Calcula médias de produção por tamanho de cama a partir das planilhas de produção.
  - Exporta planilhas de médias diárias e mensais para cada categoria.

Consulte os módulos dentro de `relatorios/` para detalhes de cada processamento.
