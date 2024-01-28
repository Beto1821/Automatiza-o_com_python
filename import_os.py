import os
from PIL import Image
import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage

# Função para extrair dados da imagem
def extract_image_data(image_path):
    img = Image.open(image_path)
    width, height = 9.33 * 9, 7.0 * 9  # Aumenta o tamanho em 900%
    return {'Image Name': os.path.basename(image_path),
            'Width': width, 'Height': height}

# Pasta contendo as imagens
image_folder = '/Users/adalbertoramosribeiro/Desktop/jog'

# Arquivo Excel
excel_file = 'image_data_large_spacing.xlsx'

# Obter uma lista ordenada de arquivos de imagem na pasta
image_files = sorted([f for f in os.listdir(image_folder) if f.endswith('.jpg')])

# Inicializar um DataFrame vazio
data = pd.DataFrame(columns=['Image Name', 'Width', 'Height'])

# Inicializar o Workbook do openpyxl
wb = Workbook()
ws = wb.active

# Definir espaçamento entre as imagens (1 linha)
spacing = 1  # 1 linha

# Iniciar a partir da segunda linha após o cabeçalho
row_number = 2

# Loop através de cada arquivo de imagem e extrair dados
for image_file in image_files:
    image_path = os.path.join(image_folder, image_file)
    image_data = extract_image_data(image_path)

    # Adicionar dados ao DataFrame
    data = pd.concat([data, pd.DataFrame([image_data])], ignore_index=True)

    # Adicionar a imagem à planilha com espaçamento
    img = ExcelImage(image_path)
    img.width = 90 * 9  # Ajuste conforme necessário
    img.height = 70 * 9  # Ajuste conforme necessário

    # Calcular a próxima célula na coluna B
    next_cell = f'B{row_number}'

    # Adicionar a imagem na célula
    ws.add_image(img, next_cell)

    # Atualizar a próxima linha para 32
    # (31 linhas da imagem + 1 linha de espaço)
    row_number += 32

# Salvar a planilha com imagens
wb.save(excel_file)

print(f"Planilha Excel '{excel_file}' criada com sucesso, "
      f"incluindo imagens aumentadas em 900% e espaçamento "
      f"de 1 linha entre elas.")
