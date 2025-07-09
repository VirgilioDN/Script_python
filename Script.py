from openpyxl import Workbook
import pymupdf

#converte o pdf para txt
doc = pymupdf.open("dados inscritos.pdf") #colocar na pasta o pdf para ele encontrar 
with open("output.txt", "wb") as out: #cria um txt para facilitar a extração
    for page in doc:
        text = page.get_text().encode("utf8")
        out.write(text)
        out.write(bytes((12,)))  

#extrai os dados do texto
with open("output.txt", "r", encoding="utf-8", errors="ignore") as saida:
    linhas = saida.readlines()
dados_gerais = []
dados = {}

for i in range(len(linhas)):
    linha = linhas[i].strip()

    if "CPF:" in linha:
        if dados:
            dados_gerais.append(dados)
            dados = {}
        dados["CPF"] = linha.replace("CPF: ", "")
    elif "Nome:" in linha:
        dados["Nome"] = linha.replace("Nome: ", "")
    elif "Sexo:" in linha:
        dados["Sexo"] = linha.replace("Sexo: ", "")
    elif "Email:" in linha:
        dados["Email"] = linha.replace("Email: ", "")
    elif "Data de Nascimento:" in linha:
        dados["Data de Nascimento"] = linha.replace("Data de Nascimento: ", "")
    elif "Raça:" in linha:
        dados["Raça"] = linha.replace("Raça: ", "")
    elif "Nome da Mãe:" in linha:
        dados["Nome da Mãe"] = linha.replace("Nome da Mãe: ", "")
    elif "Nome do Pai:" in linha:
        dados["Nome do Pai"] = linha.replace("Nome do Pai: ", "")
    elif "Bairro:" in linha:
        dados["Bairro"] = linha.replace("Bairro: ", "")
    elif "Município:" in linha:
        valor = linha.replace("Município: ", "").strip()
        if "Município (Naturalidade)" not in dados:
            dados["Município (Naturalidade)"] = valor
        dados["Município (Endereço)"] = valor
    elif "UF:" in linha:
        valor = linha.replace("UF: ", "").strip()
        if "UF (Naturalidade)" not in dados:
            dados["UF (Naturalidade)"] = valor
        dados["UF (Endereço)"] = valor
    elif "Curso de Graduação" in linha:
        dados["Curso"] = linhas[i + 1].strip()
    elif "Carta de Recomendação" in linha:
        dados["Carta de Recomendação"] = linhas[i + 1].strip()
    elif "Você é candidato a bolsa de estudos?" in linha:
        if "x" in linha.lower():
            dados["Bolsa?"] = linhas[i].strip()
        else:
            dados["Bolsa?"] = linhas[i + 1].strip()
    elif "Qual o seu conhecimento de Inglês" in linha:
        for j in range(i+1, i+9):
            if "(X)" in linhas[j] or "(x)" in linhas[j]:
                dados["Ingles"] = linhas[j].replace("(X)", "").replace("(x)", "").strip()
                break
    elif "Para qual linha de pesquisa deseja se candidatar?" in linha:
        for j in range(i+1, i+5):
            if "(X)" in linhas[j] or "(x)" in linhas[j]:
                dados["Linha de Pesquisa"] = linhas[j].replace("(X)", "").replace("(x)", "").strip()
                break
    elif "Você já tem orientador(a) definido? Se sim, quem?" in linha:
        dados["Orientador"] = linhas[i + 1]

if dados:
    dados_gerais.append(dados)

#escreve na planilha
wb = Workbook()
sheet = wb.active
cabecalhos = list(dados_gerais[0].keys())
sheet.append(cabecalhos)
for pessoa in dados_gerais:
    linha = [pessoa.get(campo, "") for campo in cabecalhos]
    sheet.append(linha)

wb.save("dados_extraidos.xlsx") #trocar o nome da planilha aqui
