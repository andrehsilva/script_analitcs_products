# Analytics por Produtos V3

Este repositório contém o código utilizado para análise e tratamento de dados relacionados a usuários, turmas e produtos do sistema LEX. A solução processa planilhas e arquivos CSV, realiza merges de dados e cria relatórios automatizados, otimizando o trabalho com grandes volumes de informações educacionais.

## Funcionalidades

- **Leitura e Tratamento de Dados**: Carregamento de arquivos CSV e Excel, remoção de caracteres especiais e padronização de texto (caixa alta/baixa).
- **Integração de Dados**: Combinação de múltiplas tabelas utilizando merges.
- **Filtragem e Agrupamento**: Criação de DataFrames específicos por produto, escola e tipo de ensino.
- **Exportação Automatizada**: Geração de relatórios em Excel organizados em abas para cada produto.

## Estrutura do Código

1. **Instalação de Dependências**:
   - `xlsxwriter` é utilizado para criar relatórios Excel.
   
2. **Funções Principais**:
   - `up(data)`: Converte os textos das colunas e dados para maiúsculas.
   - `lower(data)`: Converte os textos das colunas e dados para minúsculas.
   - `remover_caracteres_especiais(texto)`: Remove caracteres especiais de strings.
   - `carregar_e_concatenar_csv(caminho_input, padrao_arquivo)`: Carrega múltiplos arquivos CSV e os concatena em um único DataFrame.

3. **Listas e Configurações**:
   - Listas de módulos e produtos para filtrar e organizar os dados.
   
4. **Processamento de Dados**:
   - Leitura e tratamento de planilhas principais (`LEX-Usuarios_com_suas_turmas` e `LEX-Turma_X_Produtos`).
   - Integração com tabelas auxiliares, como grades e informações de clientes.

5. **Geração de Relatórios**:
   - Relatórios por produto, escola, tipo de ensino e grau são criados automaticamente com o uso de loops.
   - Os arquivos Excel gerados contêm múltiplas abas organizadas.

## Estrutura de Pastas

- **Input**: Diretório para armazenar os arquivos de entrada (planilhas e CSVs).
- **Output**: Diretório onde os relatórios processados são salvos.

## Como Usar

1. **Configurar os Caminhos**:
   Atualize as variáveis `caminho_input` e `caminho_output` com os diretórios desejados.

2. **Instalar Dependências**:
   Execute o comando abaixo para instalar o `xlsxwriter`:
   ```bash
   pip install xlsxwriter

3. **Executar o Script: Rode o script em seu ambiente Python ou no Google Colab.**

4. **Gerar Relatórios: Os arquivos Excel serão salvos no diretório definido em caminho_output, organizados por data e produto.**

5. **Pré-requisitos**
   - Python 3.7+
   - Pacotes Python: pandas, numpy, xlsxwriter, re, glob.

  
6. **Personalização**

   - Atualize as listas modulo_comunicacao, educacross e produtos para refletir os produtos específicos do seu ambiente.
   - Modifique os filtros e agrupamentos conforme necessário.


**Licença**
   - Este projeto é distribuído sob a licença MIT. Consulte o arquivo LICENSE para mais informações.
