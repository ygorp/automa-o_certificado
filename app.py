import openpyxl
from PIL import Image, ImageDraw, ImageFont
arquivo = openpyxl.load_workbook('planilha_alunos.xlsx')

pagina = arquivo['Sheet1']

for linha in pagina.iter_rows(min_row=2):
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horaria = linha[5].value
    data_emissao = linha[6].value
    
    font_nome = ImageFont.truetype('tahomabd.ttf', 90)
    font_geral = ImageFont.truetype('tahoma.ttf', 80)
    font_datas = ImageFont.truetype('tahoma.ttf', 60)
    
    imagem = Image.open('certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(imagem)
    
    desenhar.text((1050, 820), nome_participante, fill=(0, 0, 0), font=font_nome)
    desenhar.text((1080, 950), nome_curso, fill=(0, 0, 0), font=font_geral)
    desenhar.text((1450, 1070), tipo_participacao, fill=(0, 0, 0), font=font_geral)
    desenhar.text((1500, 1190), str(carga_horaria) + ' horas', fill=(0, 0, 0), font=font_geral)
    desenhar.text((740, 1770), str(data_inicio), fill=(0, 0, 0), font=font_datas)
    desenhar.text((740, 1930), str(data_final), fill=(0, 0, 0), font=font_datas)
    desenhar.text((2210, 1920), str(data_inicio), fill=(0, 0, 0), font=font_datas)
    
    imagem.save(f'./{nome_participante}.jpg')