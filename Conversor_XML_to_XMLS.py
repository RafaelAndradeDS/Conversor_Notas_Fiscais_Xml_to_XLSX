import os
import xml.etree.ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Função para selecionar pasta
def selecionar_pasta():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal do Tkinter
    pasta = filedialog.askdirectory(title="Selecione a pasta com os arquivos XML")
    return pasta

# Função para extrair informações de um arquivo XML
def extrair_dados(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    
    # Namespace da NF-e
    ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
    
    try:
        numero_nota = root.find(".//ns:infNFe/ns:ide/ns:nNF", ns).text
        data_emissao = root.find(".//ns:infNFe/ns:ide/ns:dhEmi", ns).text
        nome_cliente = root.find(".//ns:infNFe/ns:dest/ns:xNome", ns).text
        
        endereco = root.find(".//ns:infNFe/ns:dest/ns:enderDest", ns)
        rua = endereco.find("ns:xLgr", ns).text if endereco is not None else ""
        numero = endereco.find("ns:nro", ns).text if endereco is not None else ""
        municipio = endereco.find("ns:xMun", ns).text if endereco is not None else ""
        
        peso_bruto = root.find(".//ns:infNFe/ns:transp/ns:vol/ns:pesoB", ns)
        peso_bruto = peso_bruto.text if peso_bruto is not None else "0"
        
        return [numero_nota, data_emissao, nome_cliente, rua, numero, municipio, peso_bruto]
    except AttributeError:
        return None

# Função para processar todos os arquivos XML de um diretório
def processar_arquivos_xml(pasta):
    dados = []
    
    for arquivo in os.listdir(pasta):
        if arquivo.endswith(".xml"):
            caminho_arquivo = os.path.join(pasta, arquivo)
            dados_nota = extrair_dados(caminho_arquivo)
            if dados_nota:
                dados.append(dados_nota)
    
    return dados

# Função principal para salvar em Excel
def main():
    pasta = selecionar_pasta()
    if not pasta:
        print("Nenhuma pasta selecionada.")
        return
    
    dados_extraidos = processar_arquivos_xml(pasta)
    
    if not dados_extraidos:
        print("Nenhum dado extraído.")
        return
    
    colunas = ["Número da Nota", "Data de Emissão", "Nome do Cliente", "Rua", "Número", "Município", "Peso Bruto"]
    df = pd.DataFrame(dados_extraidos, columns=colunas)
    
    caminho_saida = os.path.join(pasta, "notas_fiscais.xlsx")
    df.to_excel(caminho_saida, index=False)
    
    print(f"Arquivo Excel salvo em: {caminho_saida}")

if __name__ == "__main__":
    main()
