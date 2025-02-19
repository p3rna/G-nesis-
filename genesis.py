import os
import cv2
import numpy as np
import logging
from datetime import datetime
import zipfile
import pandas as pd
from tkinterdnd2 import TkinterDnD, DND_FILES
from pdf2image import convert_from_path
from pytesseract import image_to_string
import customtkinter as ctk
from tkinter import filedialog, messagebox
import re
from pytesseract import pytesseract
import requests
from PIL import Image
from fpdf import FPDF

# Caminhos dos binários
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPLER_PATH = r"C:\poppler\bin"

# URLs para download
TESSERACT_URL = "https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-v5.3.0.20221214.exe"
POPPLER_URL = "https://github.com/oschwartz10612/poppler-windows/releases/download/v23.01.0/Release-23.01.0.zip"

# Caminho da base de dados externa
BASE_DADOS_PATH = r"C:\Users\NicolasAndré\Wedo Contabilidade e Solu&ccedil;&otilde;es Empresariais\W E D O - W E D O - DEPARTAMENTOS\T.I\Projetos Gênesis\Base_Data\BASE DE DADOS.xlsx"

# Teste de leitura do arquivo
try:
    base_dados = pd.read_excel(BASE_DADOS_PATH)
    print(f"Base de dados carregada com sucesso. Contém {base_dados.shape[0]} linhas e {base_dados.shape[1]} colunas.")
    print("Colunas presentes:", base_dados.columns)

    # Validação das colunas necessárias
    colunas_necessarias = ["Histórico", "Cód. Conta Debito", "Cód. Conta Credito", "Cód. Histórico"]
    if not all(coluna in base_dados.columns for coluna in colunas_necessarias):
        raise ValueError(f"A base de dados deve conter as colunas: {', '.join(colunas_necessarias)}")

except FileNotFoundError:
    print(f"Arquivo não encontrado: {BASE_DADOS_PATH}")
except Exception as e:
    print(f"Erro ao carregar a base de dados: {e}")

# Configuração de logs
logging.basicConfig(filename="processamento.log", level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")

def verificar_e_instalar_tesseract():
    """
    Verifica se o Tesseract está instalado. Se não estiver, baixa e instala automaticamente.
    """
    if not os.path.isfile(TESSERACT_PATH):
        print("Tesseract não encontrado. Baixando e instalando...")
        instalador_path = "tesseract-installer.exe"
        baixar_arquivo(TESSERACT_URL, instalador_path)
        os.system(f'"{instalador_path}" /SILENT')

        if not os.path.isfile(TESSERACT_PATH):
            raise Exception("Falha ao instalar o Tesseract.")
        print("Tesseract instalado com sucesso.")

def verificar_e_instalar_poppler():
    """
    Verifica se o Poppler está instalado. Se não estiver, baixa e configura automaticamente.
    """
    if not os.path.isdir(POPLER_PATH):
        print("Poppler não encontrado. Baixando e configurando...")
        zip_path = "poppler.zip"
        baixar_arquivo(POPPLER_URL, zip_path)

        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall("poppler_temp")
        os.rename("poppler_temp/Library/bin", POPLER_PATH)
        os.remove(zip_path)
        os.rmdir("poppler_temp")
        print("Poppler configurado com sucesso.")

def baixar_arquivo(url, destino):
    """
    Faz o download de um arquivo da URL especificada para o destino.
    """
    resposta = requests.get(url, stream=True)
    if resposta.status_code == 200:
        with open(destino, "wb") as f:
            for chunk in resposta.iter_content(chunk_size=1024):
                f.write(chunk)
        print(f"Download concluído: {destino}")
    else:
        raise Exception(f"Falha ao baixar o arquivo de {url} (status {resposta.status_code}).")

# Configuração do Tesseract
pytesseract.tesseract_cmd = TESSERACT_PATH

def is_gray(rgb):
    """
    Verifica se a cor é cinza.
    """
    r, g, b = rgb
    return r == g == b

def extrair_dados_ocr(caminho_pdf):
    """
    Extrai os dados do PDF utilizando OCR, incluindo colunas padrão.
    """
    try:
        imagens = convert_from_path(caminho_pdf, poppler_path=POPLER_PATH, dpi=300)
        linhas_relevantes = []

        # Regex ajustada para ignorar o campo 'Nr.Doc'
        padrao_linha = re.compile(
            r"(\d{2}/\d{2}/\d{4})\s+(.*?)\s+([\d.]+,\d{2}\s?[CD]?)"
        )

        for imagem in imagens:
            imagem = imagem.convert('L')  # Converter para escala de cinza
            texto = image_to_string(imagem, lang="por", config="--psm 6")
            texto = re.sub(r"\s{2,}", " ", texto)

            for linha in texto.split("\n"):
                linha = linha.strip()
                match = padrao_linha.search(linha)
                if match:
                    data_mov = match.group(1)
                    historico = re.sub(r"[^a-zA-Z\s]", "", match.group(2)).strip()
                    valor = match.group(3).replace(" ", "")
                    linhas_relevantes.append([data_mov, historico, valor])
                else:
                    logging.warning(f"Linha ignorada (não corresponde ao padrão): {linha}")

        # Criar DataFrame com os dados extraídos, incluindo colunas padrão
        df = pd.DataFrame(linhas_relevantes, columns=["Data Mov.", "Histórico", "Valor"])
        df["Cód. Conta Debito"] = None
        df["Cód. Conta Credito"] = None
        df["Cód. Histórico"] = None

        # Remover linhas onde o campo "Histórico" contém "SALDO"
        df = df[~df["Histórico"].str.contains("SALDO", case=False, na=False)]

        # Remover linhas onde o campo "Valor" é igual a 0
        df = df[df["Valor"] != 0]


        # Adicionar colunas personalizadas
        df = adicionar_colunas_personalizadas(df)

        return df
    
        # Log do conteúdo extraído
        logging.info(f"Dados extraídos do PDF {caminho_pdf}: {df.head()}")

        return df
    
    except Exception as e:
        logging.error(f"Erro ao processar PDF com OCR: {e}")
        return pd.DataFrame()

    except Exception as e:
        logging.error(f"Erro ao processar PDF com OCR: {e}")
        return pd.DataFrame()
def remover_numeros_inicio_historico(df):
    """
    Remove números e espaços no início da coluna 'Histórico', mantendo apenas o restante da string.
    """
    df["Histórico"] = df["Histórico"].str.replace(r"^\s*\d+\s*", "", regex=True)
    return df

def formatar_valor(df):
    """
    Formata os valores extraídos.
    """
def ajustar_valor(valor):
    """
    Converte valores terminados em 'C' para positivos e 'D' para negativos.
    Mantém o separador decimal correto e processa números com separador de milhares.
    """
    if isinstance(valor, str):  # Garantir que é uma string antes de manipular
        valor = valor.strip().replace(" ", "")  # Remove espaços extras

        # Expressão regular para identificar números corretamente
        match = re.match(r"^(-?[\d\.]+),(\d{2})([CD]?)$", valor)

        if match:
            numero = match.group(1).replace(".", "") + "." + match.group(2)  # Corrige formato decimal
            numero = float(numero)  # Converte para número
            
            if match.group(3) == "D":  # Se termina com "D", torna negativo
                return -numero
            return numero  # Se termina com "C" ou nada, mantém positivo

    return valor  # Retorna o valor original caso não precise ser alterado

def adicionar_colunas_personalizadas(df):
    """
    Adiciona as colunas Cód. Conta Débito, Cód. Conta Crédito e Cód. Histórico ao DataFrame
    com base na base de dados externa.
    """
    try:
        # Carregar a base de dados externa
        base_dados = pd.read_excel(BASE_DADOS_PATH)

        # Normalizar os valores para evitar diferenças de capitalização ou espaços
        base_dados["Histórico"] = base_dados["Histórico"].str.strip().str.upper()
        df["Histórico"] = df["Histórico"].str.strip().str.upper()

        # Criar mapeamento para as colunas
        mapeamento_debito = base_dados.set_index("Histórico")["Cód. Conta Debito"].to_dict()
        mapeamento_credito = base_dados.set_index("Histórico")["Cód. Conta Credito"].to_dict()
        mapeamento_historico = base_dados.set_index("Histórico")["Cód. Histórico"].to_dict()

        # Adicionar colunas ao DataFrame com os valores mapeados
        df["Cód. Conta Debito"] = df["Histórico"].map(mapeamento_debito)
        df["Cód. Conta Credito"] = df["Histórico"].map(mapeamento_credito)
        df["Cód. Histórico"] = df["Histórico"].map(mapeamento_historico)

        # Verificação de correspondência
        valores_nao_encontrados = df[~df["Histórico"].isin(base_dados["Histórico"])]["Histórico"].unique()
        if len(valores_nao_encontrados) > 0:
            logging.warning(f"Os seguintes valores do 'Histórico' não foram encontrados na base de dados: {valores_nao_encontrados}")
            print(f"Valores não encontrados na base de dados: {valores_nao_encontrados}")

        return df
    except Exception as e:
        logging.error(f"Erro ao adicionar colunas personalizadas: {e}")
        return df



def adicionar_coluna_historico(df):
    """
    Adiciona a coluna "Código" ao DataFrame com base na base de dados externa.
    """
    try:
        # Carregar a base de dados externa
        base_dados = pd.read_excel(BASE_DADOS_PATH)

        # Garantir que as colunas necessárias estão presentes na base
        if "Histórico" not in base_dados.columns or "Código" not in base_dados.columns:
            raise ValueError("A base de dados deve conter as colunas 'Histórico' e 'Código'.")

        # Criar um dicionário para mapear histórico -> código
        mapeamento_codigo = dict(zip(base_dados["Histórico"], base_dados["Código"]))

        # Adicionar depuração
        print("Mapeamento de 'Histórico' para 'Código':")
        print(mapeamento_codigo)

        # Adicionar a coluna "Código" ao DataFrame com base no mapeamento
        df["Código"] = df["Histórico"].map(mapeamento_codigo)

        return df
    except Exception as e:
        logging.error(f"Erro ao adicionar coluna 'Código': {e}")
        return df

def salvar_excel_formatado(df, caminho_pdf, diretorio_saida):
    """
    Salva os dados extraídos em um arquivo Excel com nome baseado no PDF.
    """
    try:
        nome_pdf = os.path.splitext(os.path.basename(caminho_pdf))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        caminho_excel = os.path.join(diretorio_saida, f"{nome_pdf}_{timestamp}.xlsx")

        # Adicionar depuração antes de salvar
        print("Colunas no DataFrame antes de salvar no Excel:")
        print(df.columns)
        print("Exemplo de dados no DataFrame antes de salvar:")
        print(df.head())

        df.to_excel(caminho_excel, index=False)
        logging.info(f"Dados salvos com sucesso em: {caminho_excel}")
    except Exception as e:
        logging.error(f"Erro ao salvar Excel: {e}")

def salvar_txt_formatado(df, caminho_pdf, diretorio_saida):
    """
    Salva os dados extraídos em um arquivo TXT com campos reorganizados.
    """
    try:
        nome_pdf = os.path.splitext(os.path.basename(caminho_pdf))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        caminho_txt = os.path.join(diretorio_saida, f"{nome_pdf}_{timestamp}.txt")

        # Adicionar as novas colunas com valores padrão
        df["Lote"] = "1"  # Coluna com valor padrão
        df["Filial"] = ""  # Coluna vazia
        df["Centro de Custo"] = ""  # Coluna vazia

        # Reorganizar as colunas para o TXT
        colunas_reorganizadas = [
            "Data Mov.", "Cód. Conta Credito", "Cód. Conta Debito",
            "Valor", "Cód. Histórico", "Histórico", "Lote"
        ]
        df = df[colunas_reorganizadas]

        # Garantir que os valores das colunas de código sejam inteiros e depois convertidos para strings
        colunas_numericas = ["Cód. Conta Credito", "Cód. Conta Debito", "Cód. Histórico"]
        for coluna in colunas_numericas:
            df[coluna] = df[coluna].fillna(0).astype(int).astype(str)

        # Ajustar a formatação dos valores
        def formatar_valor(valor):
            return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        df["Valor"] = df["Valor"].apply(formatar_valor)

        # Formatar o conteúdo de cada linha no estilo desejado
        linhas_formatadas = []
        for _, row in df.iterrows():
            linha = (
                f"{row['Data Mov.']};{row['Cód. Conta Credito']};{row['Cód. Conta Debito']};"
                f"{row['Valor']};{row['Cód. Histórico']};{row['Histórico']};{row['Lote']};;;"
            )
            linhas_formatadas.append(linha)

        # Salvar as linhas formatadas no arquivo TXT
        with open(caminho_txt, "w", encoding="utf-8") as arquivo_txt:
            arquivo_txt.write("\n".join(linhas_formatadas))

        logging.info(f"Dados salvos com sucesso em: {caminho_txt}")
        print(f"Arquivo TXT salvo com sucesso: {caminho_txt}")
    except Exception as e:
        logging.error(f"Erro ao salvar TXT: {e}")
        print(f"Erro ao salvar TXT: {e}")

class AppInterface:
    """
    Interface gráfica do GÊNESIS.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("GÊNESIS")
        self.root.geometry("800x700")

        # Configuração inicial do tema
        self.tema_atual = "dark"  # Define o tema inicial como escuro
        ctk.set_appearance_mode(self.tema_atual)
        ctk.set_default_color_theme("dark-blue")

        

        # Configuração de transparência (alpha)
        self.root.attributes("-alpha", 0.9)  # Define a opacidade da janela (90%)

        # Título da aplicação
        self.titulo = ctk.CTkLabel(root, text="GÊNESIS", font=("Arial", 28, "bold"))
        self.titulo.pack(pady=10)

        self.subtitulo = ctk.CTkLabel(root, text="WEDO CONTABILIDADE DIGITAL", font=("Arial", 16, "italic"))
        self.subtitulo.pack(pady=5)

        # Frame principal
        self.frame = ctk.CTkFrame(root, corner_radius=15)
        self.frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Texto de status
        self.texto_status = ctk.CTkTextbox(self.frame, width=750, height=200, font=("Arial", 12), corner_radius=10)
        self.texto_status.pack(pady=10, padx=10)

        # Barra de progresso
        self.progress_bar = ctk.CTkProgressBar(self.frame, width=700)
        self.progress_bar.pack(pady=10)

        # Botões
        self.botoes_frame = ctk.CTkFrame(self.frame, fg_color="transparent")
        self.botoes_frame.pack(pady=10)

        self.btn_selecionar = ctk.CTkButton(self.botoes_frame, text="Selecionar PDFs",
                                            command=self.selecionar_pdfs, corner_radius=10, width=180)
        self.btn_selecionar.grid(row=0, column=0, padx=10, pady=5)

        self.btn_diretorio = ctk.CTkButton(self.botoes_frame, text="Selecionar Diretório",
                                           command=self.selecionar_diretorio, corner_radius=10, width=180)
        self.btn_diretorio.grid(row=0, column=1, padx=10, pady=5)

        self.btn_processar = ctk.CTkButton(self.botoes_frame, text="Processar e Salvar",
                                           command=self.processar_e_salvar, corner_radius=10, width=180)
        self.btn_processar.grid(row=0, column=2, padx=10, pady=5)

        self.btn_processar_excel = ctk.CTkButton(
            self.botoes_frame, 
            text="Processar Excel", 
            command=self.processar_excel, 
            corner_radius=10, 
            width=180
        )
        self.btn_processar_excel.grid(row=1, column=1, padx=10, pady=5)

        self.btn_processar_excel = ctk.CTkButton(
            self.botoes_frame, 
            text="Processar Excel", 
            command=self.processar_excel,  
            corner_radius=10, 
            width=180
        )
        self.btn_processar_excel.grid(row=1, column=1, padx=10, pady=5)



        # Botão para alternar temas
        self.btn_tema = ctk.CTkButton(root, text="Alternar Tema", command=self.alternar_tema,
                                      corner_radius=10, width=180)
        self.btn_tema.pack(pady=10)

    def alternar_tema(self):
        """
        Alterna entre os temas claro e escuro.
        """
        if self.tema_atual == "dark":
            self.tema_atual = "light"
        else:
            self.tema_atual = "dark"

        ctk.set_appearance_mode(self.tema_atual)
        self.texto_status.insert("end", f"Tema alterado para: {self.tema_atual}\n")

    def selecionar_pdfs(self):
        arquivos = filedialog.askopenfilenames(filetypes=[("Arquivos PDF", "*.pdf")])
        if arquivos:
            self.arquivos_pdf = arquivos
            self.texto_status.insert("end", f"{len(arquivos)} arquivos selecionados para processamento.\n")
            logging.info(f"{len(arquivos)} arquivos selecionados.")

    def selecionar_diretorio(self):
        diretorio = filedialog.askdirectory()
        if diretorio:
            self.diretorio_saida = diretorio
            self.texto_status.insert("end", f"Diretório de saída selecionado: {diretorio}\n")
            logging.info(f"Diretório de saída selecionado: {diretorio}")

    def processar_e_salvar(self):
        """
        Processa e salva os arquivos PDF selecionados.
        """
        if not self.arquivos_pdf:
            messagebox.showwarning("Aviso", "Nenhum arquivo PDF foi selecionado.")
            return
        if not self.diretorio_saida:
            messagebox.showwarning("Aviso", "Nenhum diretório de saída foi selecionado.")
            return

        self.texto_status.insert("end", "Processando e salvando arquivos...\n")
        os.makedirs(self.diretorio_saida, exist_ok=True)

        for index, caminho_pdf in enumerate(self.arquivos_pdf, start=1):
            try:
                df = extrair_dados_ocr(caminho_pdf)
                if not df.empty:
                    df = formatar_valor(df)
                    df = adicionar_colunas_personalizadas(df)
                    df = adicionar_coluna_historico(df)  # Adiciona a coluna "Código"

                    salvar_excel_formatado(df, caminho_pdf, self.diretorio_saida)  # Salva o Excel
                    salvar_txt_formatado(df, caminho_pdf, self.diretorio_saida)    # Salva o TXT

                    self.texto_status.insert("end", f"Arquivo {index}/{len(self.arquivos_pdf)} processado e salvo.\n")
                else:
                    self.texto_status.insert("end", f"Falha ao processar arquivo: {caminho_pdf}\n")

            except ValueError as e:
                self.texto_status.insert("end", f"Erro: {e} ao processar o arquivo: {caminho_pdf}\n")
            except Exception as e:
                self.texto_status.insert("end", f"Erro inesperado ao processar {caminho_pdf}: {e}\n")

        self.texto_status.insert("end", "Todos os arquivos foram processados e salvos no diretório selecionado.\n")

    def processar_excel(self):
        """
        Permite selecionar um arquivo Excel, processa os valores terminados em 'C' ou 'D',
        insere um cabeçalho correto na linha 1 e salva as alterações.
        """
        arquivo_excel = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])

        if not arquivo_excel:
            messagebox.showwarning("Aviso", "Nenhum arquivo Excel foi selecionado.")
            return

        try:
            df = pd.read_excel(arquivo_excel, dtype=str, header=None)  # Carregar SEM definir cabeçalho

            print("Primeiras linhas antes da modificação:\n", df.head())  # Depuração
            self.texto_status.insert("end", f"Primeiras linhas antes: {df.head()}\n")

            # 🔹 Criar um novo DataFrame para o cabeçalho
            colunas_novas = pd.DataFrame([["Data", "Lançamento", "Valor"]])

            # 🔹 Concatenar o cabeçalho com os dados originais, deslocando tudo para baixo
            df = pd.concat([colunas_novas, df], ignore_index=True)

            # Aplicar a formatação dos valores na planilha
            df = df.applymap(ajustar_valor)

            # Salvar novamente o Excel
            df.to_excel(arquivo_excel, index=False, header=False)  # Salva sem cabeçalho extra

            messagebox.showinfo("Sucesso", "O arquivo Excel foi processado com sucesso!")

            self.texto_status.insert("end", f"Arquivo Excel '{os.path.basename(arquivo_excel)}' processado com sucesso!\n")
            logging.info(f"Arquivo Excel '{arquivo_excel}' processado e salvo.")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar o arquivo: {e}")
            logging.error(f"Erro ao processar Excel: {e}")

if __name__ == "__main__":
    try:
        verificar_e_instalar_tesseract()
        verificar_e_instalar_poppler()
        print("Todas as dependências foram instaladas.")

        root = ctk.CTk()
        app = AppInterface(root)
        root.mainloop()
    except Exception as e:
        print(f"Erro ao configurar dependências: {e}")
