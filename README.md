Sorteador de Lista e Grupos
Este projeto é uma aplicação em Python com interface gráfica (PyQt6) que realiza sorteios automáticos a partir de listas, com opção de gerar grupos aleatórios. O sistema permite importar listas, realizar sorteios, visualizar resultados e salvar as informações.

Funcionalidades
Importação de listas de participantes (ex: via arquivo Excel ou entrada manual).

Sorteio aleatório de nomes da lista.

Criação de grupos aleatórios a partir da lista.

Interface gráfica amigável usando PyQt6.

Visualização gráfica dos resultados usando matplotlib.

Salvar e carregar dados em formatos estruturados (ex: JSON, Excel).

Registro de data e hora do sorteio.

Tecnologias utilizadas
Python 3.8+

PyQt6 para interface gráfica

matplotlib para visualização gráfica

openpyxl para manipulação de arquivos Excel

JSON para persistência de dados

datetime para controle de data e hora

Como usar
Clone o repositório:

bash
Copiar
Editar
git clone https://github.com/seu-usuario/sorteador-lista-grupos.git
cd sorteador-lista-grupos
Instale as dependências:

bash
Copiar
Editar
pip install PyQt6 matplotlib openpyxl
Execute o programa:

bash
Copiar
Editar
python seu_arquivo_principal.py
Na interface gráfica, importe a lista de participantes, selecione o tipo de sorteio (individual ou grupos), e execute o sorteio.

Estrutura do Código
main.py — arquivo principal que inicializa a interface.

ui/ — módulos relacionados à interface gráfica com PyQt6.

logic/ — funções que realizam o sorteio e manipulação dos dados.

data/ — pasta onde os arquivos importados e gerados são armazenados.

utils/ — funções auxiliares, como leitura/escrita de arquivos
