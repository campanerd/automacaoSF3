# Automação de Ocorrências SAC – SF3

Este projeto foi desenvolvido para automatizar o processo diário de coleta, filtragem e envio de ocorrências SAC, eliminando tarefas manuais e reduzindo erros operacionais. A automação realiza a integração entre um servidor SFTP, arquivos Excel e o Microsoft Outlook, garantindo que as ocorrências do dia sejam tratadas e encaminhadas de forma padronizada e confiável.

O fluxo foi pensado para uso corporativo, com foco em rastreabilidade, clareza das informações e facilidade de manutenção do código.

---

## 🎯 Objetivo da Automação

O principal objetivo deste projeto é automatizar o processo que anteriormente dependia de etapas manuais, como:

- Download diário de planilhas de ocorrências em servidor externo
- Abertura e filtragem manual de dados no Excel
- Identificação de contratos únicos
- Criação de e-mail com tabela formatada
- Anexação do arquivo filtrado e envio aos responsáveis

Com esta automação, todo esse processo é executado de forma automática, padronizada e segura, bastando apenas executar o script principal.

---

## ⚙️ Funcionamento Geral

Ao executar o projeto, a automação segue a seguinte sequência lógica:

Primeiro, o sistema se conecta a um servidor SFTP utilizando credenciais configuradas em variáveis de ambiente. O arquivo de ocorrências do mês corrente é baixado automaticamente e salvo em uma pasta local do projeto.

Em seguida, o arquivo Excel é lido utilizando a biblioteca Pandas. A automação acessa uma aba específica da planilha e realiza o tratamento dos dados, removendo registros inválidos, filtrando apenas as ocorrências referentes à data atual e eliminando contratos duplicados.

Após o processamento, é gerado um novo arquivo Excel contendo apenas as ocorrências válidas do dia. Paralelamente, esses mesmos dados são convertidos em uma tabela HTML, formatada para ser exibida corretamente no corpo de um e-mail.

Por fim, a automação cria e envia um e-mail pelo Microsoft Outlook, contendo:
- Um texto padrão explicativo
- A tabela HTML com as ocorrências do dia
- O arquivo Excel filtrado em anexo

Caso não existam ocorrências para a data atual, o e-mail ainda é enviado, informando explicitamente que não há contratos para tratamento naquele dia.

---

## 🗂️ Estrutura do Projeto

O projeto é organizado de forma modular, separando responsabilidades e facilitando manutenção e evolução do código.

- `main.py`  
  Arquivo principal responsável por orquestrar toda a execução da automação.

- `ftp_service.py`  
  Contém a lógica de conexão com o servidor SFTP e o download do arquivo de ocorrências.

- `filtre_service.py`  
  Responsável pela leitura do Excel, tratamento dos dados, filtragem por data, remoção de duplicidades e geração do arquivo final.

- `email_service.py`  
  Responsável pela geração da tabela HTML e pelo envio do e-mail via Outlook.

- `src/downloads/`  
  Diretório utilizado para armazenar temporariamente os arquivos baixados e gerados durante a execução.

- `.env`  
  Arquivo de configuração contendo credenciais e parâmetros sensíveis (não versionado).

---

## 🧰 Tecnologias Utilizadas

O projeto foi desenvolvido utilizando as seguintes tecnologias e bibliotecas:

- Python 3
- Pandas para manipulação de dados
- Paramiko para conexão SFTP
- win32com para integração com Microsoft Outlook
- python-dotenv para gerenciamento de variáveis de ambiente
- Excel como formato de entrada e saída de dados

---

## 🛠️ Configuração do Ambiente

Antes de executar o projeto, é necessário configurar o ambiente.

Instale as dependências do projeto:

```bash
pip install pandas python-dotenv paramiko pywin32 openpyxl xlrd
