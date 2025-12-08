## -->> Verificando a instalação do Python <<--
python --version

## -->> Criação do ambiente virtual <<--
python -m venv venv

## -->> Ativando ambiente virtual <<--
.\venv\Scripts\Activate

## -->> Problema de bloqueio Windows <<--
Todos os pip install só rodam passando o parâmetro específico do Python:
python.exe -m pip install --upgrade pip
python -m pip install pandas openpyxl
python -m pip install pywin32

Se tentar rodar apenas pip install pandas openpyxl será bloqueado pelo SISGAG!!!

## -->> Listando meus pacotes instalados dentro do venv <<--
python -m pip list
python -m pip freeze

## -->> Gerenado meu arquivo requirements.txt <<--
python -m pip freeze > requirements.txt

## -->> Instalando pacotes através de um arquivo requirements existente <<--
python -m pip install -r requirements.txt

## -->> Rodando o main dentro do SISGAG:
python -m src.Shrinkage.app

## -->> Estrutura dos meus arquivos <<--
DesenvolvimentoBackEnd/
│
├── src/
│   └── teste.py
│
├── tests/
│   └── (vazio por enquanto)
│
├── venv/                  ✔ ignorar no Git
│
├── .gitignore             ✔ ótimo
├── README.md              ✔ documentação
├── requirements.txt       ✔ dependências

