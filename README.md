# **EXTRAÇÃO DE DADOS DE NOTAS FISCAIS XML PARA EXCEL**

## 📝 **DESCRIÇÃO**
Este projeto consiste em um script em *Python* que processa em massa arquivos *XML* de notas fiscais eletrônicas (*NF-e*) e extrai informações específicas para um arquivo *Excel*. 

🔹 **Bibliotecas Utilizadas:**
- `xml.etree.ElementTree` → Leitura dos arquivos XML.
- `pandas` → Manipulação de dados e geração do arquivo Excel.
- `tkinter` → Interface gráfica para seleção de pastas.

---

## ✅ **REQUISITOS DO SCRIPT**
📥 **Entrada:** Vários arquivos XML de notas fiscais (*NF-e*).
📤 **Saída:** Um arquivo Excel (`.xlsx`) contendo os seguintes dados extraídos:
- **Número da Nota**
- **Data de Emissão da Nota**
- **Nome do Cliente**
- **Endereço Completo do Cliente**, com os seguintes campos:
  - Rua
  - Número
  - Município
- **Peso Bruto**

---

## 📦 **DEPENDÊNCIAS**
Antes de executar o script, instale as bibliotecas necessárias:
```bash
pip install pandas openpyxl
```
Essas bibliotecas são usadas para manipulação de dados e geração do arquivo Excel.

---

## 🚀 **COMO EXECUTAR O SCRIPT**

1️⃣ **Clone este repositório:**
```bash
git clone https://github.com/seu-usuario/nome-do-repositorio.git
cd nome-do-repositorio
```

2️⃣ **Execute o script:**
```bash
python Conversor_XML_to_Excel.py
```

3️⃣ **Selecione a pasta contendo os arquivos XML** quando solicitado.

4️⃣ **O script irá processar os arquivos e salvar o resultado** em um arquivo chamado `notas_fiscais.xlsx` na mesma pasta onde estão os arquivos XML.

---

## 📂 **ESTRUTURA DO PROJETO**

📌 O código foi desenvolvido para processar múltiplos arquivos *XML* de notas fiscais eletrônicas (*NF-e*) e extrair informações relevantes para um arquivo *Excel*.

🔹 **Tecnologias Utilizadas:**
- `xml.etree.ElementTree` → Para lidar com XML.
- `pandas` → Manipulação de dados.
- `tkinter` → Interface gráfica de seleção de pastas.

---

## 🔍 **EXPLICAÇÃO DO CÓDIGO**

### 1️⃣ **Importação das Bibliotecas**
```python
import os
import xml.etree.ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter import filedialog
```
📌 **Explicação:**
- `os` → Manipula arquivos e diretórios.
- `xml.etree.ElementTree` → Faz a leitura e extração de dados do XML.
- `pandas` → Cria e manipula tabelas de dados (*DataFrame*) e salva o resultado em Excel.
- `tkinter` → Cria uma janela de diálogo para o usuário selecionar a pasta com os arquivos XML.

---

### 2️⃣ **Seleção da Pasta com XMLs**
```python
def selecionar_pasta():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal do Tkinter
    pasta = filedialog.askdirectory(title="Selecione a pasta com os arquivos XML")
    return pasta
```
📌 **Explicação:**
- Cria uma janela para o usuário escolher a pasta onde estão os arquivos XML.
- `withdraw()` esconde a janela principal do Tkinter, exibindo apenas a caixa de diálogo.

---

### 3️⃣ **Extração de Dados do XML**
```python
def extrair_dados(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    
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
```
📌 **Explicação:**
- Lê e analisa o XML (`ET.parse(xml_file)`).
- Usa `Namespace (ns)` pois os arquivos *NF-e* possuem namespaces.
- Busca os dados relevantes como **número da nota, data de emissão, cliente, endereço e peso bruto**.
- **Tratamento de Erro**: Se algum campo estiver ausente, a função retorna `None`.

---

### 4️⃣ **Processamento de Múltiplos Arquivos XML**
```python
def processar_arquivos_xml(pasta):
    dados = []
    
    for arquivo in os.listdir(pasta):
        if arquivo.endswith(".xml"):
            caminho_arquivo = os.path.join(pasta, arquivo)
            dados_nota = extrair_dados(caminho_arquivo)
            if dados_nota:
                dados.append(dados_nota)
    
    return dados
```
📌 **Explicação:**
- Lista todos os arquivos XML da pasta.
- Chama `extrair_dados()` para cada arquivo e adiciona os resultados a uma lista.

---

### 5️⃣ **Criação do Arquivo Excel**
```python
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
```
📌 **Explicação:**
- Cria um `DataFrame` com os dados extraídos.
- Salva o arquivo `notas_fiscais.xlsx` na pasta selecionada.

---

### 6️⃣ **Executando o Script**
```python
if __name__ == "__main__":
    main()
```
📌 **Isso garante que o código seja executado apenas quando rodado diretamente.**
