# Projeto de Cadastro de Cliente em Python

## Descrição do Projeto

Este projeto visa estimular a habilidade de criação de cadastros de clientes em Python, utilizando a biblioteca CustomTkinter para a interface gráfica. O armazenamento de dados é feito em um arquivo Excel, que serve como banco de dados simples.

## Funcionalidades

- Cadastro de novos clientes
- Edição de informações de clientes existentes
- Exclusão de clientes
- Visualização e busca de clientes cadastrados

## Requisitos

- Python 3.x instalado
- Biblioteca CustomTkinter instalada (`pip install customtkinter`)
- Microsoft Excel para armazenamento de dados (um arquivo Excel será criado no diretório do projeto)

## Como Executar

1. Clone este repositório para o seu ambiente local.
2. Instale as dependências com o comando: `pip install -r requirements.txt`
3. Execute o script principal: `python main.py`

Certifique-se de ter permissões para escrever no diretório do projeto, pois é onde o arquivo Excel será gerado.

## Estrutura do Projeto

- `main.py`: Script principal que inicia a aplicação.
- `customtkinter_utils.py`: Funções utilitárias para facilitar a integração da CustomTkinter.
- `excel_database.py`: Módulo para interação com o banco de dados Excel.
- `cadastro_cliente.py`: Módulo contendo a lógica de cadastro de clientes.
- `gui/`: Diretório contendo os arquivos relacionados à interface gráfica.

## Aviso Legal

Este projeto é destinado exclusivamente para fins educacionais. Você é livre para utilizar e modificar o código conforme necessário para seus próprios propósitos de aprendizado. No entanto, o autor não se responsabiliza por qualquer uso indevido ou ilegal do código.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues, propor melhorias ou enviar pull requests.

## Licença

Este projeto está licenciado sob a licença MIT - consulte o arquivo [LICENSE](LICENSE) para obter detalhes.
