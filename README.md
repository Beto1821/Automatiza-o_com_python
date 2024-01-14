# Geração de relatório com imagens de teste VSWR de propagação em cabos de RF(FEEDER) projeto Hospital ALbert Einsten Claro/VIVO
Exemplo de automatização usando Python

# Crie o ambiente virtual para o projeto
python3 -m venv .venv && source .venv/bin/activate

# Instale a dependência Pandas
pip install pandas openpyxl Pillow

# Execução com diretorio onde as imagens estão localizadas pré-defino
python3 import_os.py 