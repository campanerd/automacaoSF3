# Automação de Ocorrências SAC – SF3

Este projeto automatiza o processo diário de coleta, validação, rastreio e envio de ocorrências SAC, eliminando tarefas manuais, prevenindo envios duplicados e garantindo rastreabilidade completa das informações.

A automação integra um servidor SFTP, planilhas Excel e o Microsoft Outlook, processando dados históricos e atuais de forma confiável, mesmo em cenários onde o arquivo não é atualizado diariamente.

O fluxo foi projetado para uso corporativo, com foco em segurança operacional, controle de duplicidade, auditoria e facilidade de manutenção.

---

## 🎯 Objetivo da Automação

Automatizar integralmente um processo que antes dependia de múltiplas etapas manuais, como:

Download de planilhas de ocorrências via servidor SFTP

Abertura e análise manual de dados no Excel

Identificação de contratos válidos e únicos por dia

Controle de contratos já enviados anteriormente

Criação de e-mail com tabela formatada

Anexação de arquivo Excel apenas quando aplicável

Com esta automação, todo o processo ocorre de forma automática, padronizada e segura, bastando executar o script principal (ou deixá-lo agendado)

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

O projeto é organizado de forma modular, separando responsabilidades e facilitando manutenção, entendimento e evolução do código.

- `main.py`
Arquivo principal responsável por orquestrar toda a execução da automação. Realiza a chamada dos serviços de download, processamento dos dados e envio do e-mail, funcionando como ponto central do fluxo.

- `ftp_service.py`
Contém a lógica de conexão com o servidor SFTP, autenticação por credenciais externas e download automático do arquivo de ocorrências do mês corrente.

- `filtre_service.py`
Responsável pela leitura da planilha Excel, tratamento e normalização dos dados, criação de chaves únicas, controle de histórico, filtragem de ocorrências válidas, remoção de contratos duplicados no mesmo dia e geração do arquivo Excel final.

- `email_service.py`
Responsável pela geração da tabela HTML a partir dos dados filtrados e pelo envio do e-mail via Microsoft Outlook. O e-mail é enviado mesmo quando não há ocorrências, informando explicitamente a ausência de contratos.

- `src/downloads/`
Diretório utilizado para armazenar os arquivos baixados via SFTP, os arquivos Excel gerados pela automação e o histórico de ocorrências já processadas.

- `.env`
Arquivo de configuração contendo credenciais, endereços de e-mail e demais parâmetros sensíveis utilizados pela automação (não versionado).

---

## 🧰 Tecnologias Utilizadas

O projeto foi desenvolvido utilizando as seguintes tecnologias e bibliotecas:

- Python 3
- Pandas para manipulação de dados
- Paramiko para conexão SFTP
- win32com para integração com Microsoft Outlook
- python-dotenv para gerenciamento de variáveis de ambiente
- Excel como formato de entrada e saída de dados
