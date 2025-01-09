import os
import logging
from datetime import datetime
import pandas as pd
from pdf2image import convert_from_path
from pytesseract import image_to_string, pytesseract
import customtkinter as ctk
from tkinter import filedialog, messagebox
import re
import requests
import zipfile

# Caminhos dos binários
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPLER_PATH = r"C:\poppler\bin"

# URLs para download
TESSERACT_URL = "https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-v5.3.0.20221214.exe"
POPPLER_URL = "https://github.com/oschwartz10612/poppler-windows/releases/download/v23.01.0/Release-23.01.0.zip"


# Caminho da base de dados externa
BASE_DADOS_PATH = r"C:\Users\NicolasAndré\Wedo Contabilidade e Soluções Empresariais\W E D O - W E D O - DEPARTAMENTOS\T.I\Projetos Gênesis\Base_Data\BASE DE DADOS.xlsx"

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


def extrair_dados_ocr(caminho_pdf):
    """
    Extrai os dados do PDF utilizando OCR, ignorando a coluna 'Nr.Doc'.
    """
    try:
        imagens = convert_from_path(caminho_pdf, poppler_path=POPLER_PATH)
        linhas_relevantes = []

        padrao_linha = re.compile(
            r"(\d{2}/\d{2}/\d{4})\s+(.*?)\s+([\d.]+,\d{2}\s?[CD]?)"
        )

        for imagem in imagens:
            texto = image_to_string(imagem, lang="por", config="--psm 6")
            texto = re.sub(r"\s{2,}", " ", texto)

            for linha in texto.split("\n"):
                linha = linha.strip()
                match = padrao_linha.search(linha)
                if match:
                    data_mov = match.group(1)
                    historico = match.group(2).strip()
                    valor = match.group(3).replace(" ", "")
                    linhas_relevantes.append([data_mov, historico, valor])
                else:
                    logging.warning(f"Linha ignorada (não corresponde ao padrão): {linha}")

        df = pd.DataFrame(linhas_relevantes, columns=["Data Mov.", "Histórico", "Valor"])

        logging.info(f"Dados extraídos do PDF {caminho_pdf}: {df.head()}")

        df = df[~df["Histórico"].str.contains("SALDO DIA", case=False, na=False)]
        df = df[df["Valor"].str.strip() != ""]
        df = remover_numeros_inicio_historico(df)

        return df
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
        try:
            valor = valor.replace(".", "").replace(",", ".")
            if "D" in valor:
                return -float(valor.replace("D", ""))
            elif "C" in valor:
                return float(valor.replace("C", ""))
            else:
                return float(valor)
        except Exception as e:
            logging.warning(f"Erro ao formatar valor: {valor} - {e}")
            return None

    df["Valor"] = df["Valor"].apply(ajustar_valor)
    return df


def adicionar_colunas_personalizadas(df):
    """
    Adiciona as colunas Cód. Conta Débito, Cód. Conta Crédito e Cód. Histórico ao DataFrame.
    """
    mapeamento = {
        "CR-COM-SIL": (9, 5, 10),
        "AZCX MC CD": (9, 5, 10),
        "AZCX EL CD": (9, 5, 10),
        "AZCX VS CD": (9, 5, 10),
        "DEB ISSQN": (215, 9, 10),
        "DP DIN LOT": (289, 9, 10),
    }

    def mapear_codigos(historico):
        if historico in mapeamento:
            return mapeamento[historico]
        else:
            return (None, None, None)

    df[["Cód. Conta Débito", "Cód. Conta Crédito", "Cód. Histórico"]] = df["Histórico"].apply(
        lambda x: pd.Series(mapear_codigos(x))
    )
    return df


def adicionar_coluna_codigo(df):
    """
    Adiciona a coluna "Código" ao DataFrame com base na base de dados externa.
    """
    try:
        base_dados = pd.read_excel(BASE_DADOS_PATH)

        if "Histórico" not in base_dados.columns or "Código" not in base_dados.columns:
            raise ValueError("A base de dados deve conter as colunas 'Histórico' e 'Código'.")

        mapeamento_codigo = dict(zip(base_dados["Histórico"], base_dados["Código"]))
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
        df.to_excel(caminho_excel, index=False)
        logging.info(f"Dados salvos com sucesso em: {caminho_excel}")
    except Exception as e:
        logging.error(f"Erro ao salvar Excel: {e}")


def salvar_txt_formatado(df, caminho_pdf, diretorio_saida):
    """
    Salva os dados extraídos em um arquivo TXT com campos separados por ";".
    """
    try:
        colunas_codigo = ["Cód. Conta Débito", "Cód. Conta Crédito", "Cód. Histórico"]
        for coluna in colunas_codigo:
            if coluna in df.columns:
                df[coluna] = df[coluna].fillna(0).astype(int)

        nome_pdf = os.path.splitext(os.path.basename(caminho_pdf))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        caminho_txt = os.path.join(diretorio_saida, f"{nome_pdf}_{timestamp}.txt")
        df.to_csv(caminho_txt, sep=";", index=False, encoding="utf-8")
        logging.info(f"Dados salvos com sucesso em: {caminho_txt}")
        print(f"Arquivo TXT salvo com sucesso: {caminho_txt}")
    except Exception as e:
        logging.error(f"Erro ao salvar TXT: {e}")
        print(f"Erro ao salvar TXT: {e}")


class AppInterface:
    def __init__(self, root):
        self.root = root
        self.root.title("GÊNESIS LOTÉRICOS")
        self.root.geometry("800x700")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.diretorio_saida = ""

        titulo = ctk.CTkLabel(root, text="GÊNESIS LOTÉRICOS",
                              font=("Arial", 28, "bold"), text_color="#FFC107")
        titulo.pack(pady=10)

        subtitulo = ctk.CTkLabel(root, text="WEDO CONTABILIDADE DIGITAL",
                                 font=("Arial", 16, "italic"), text_color="#FFD700")
        subtitulo.pack(pady=5)

        frame = ctk.CTkFrame(root, corner_radius=15, fg_color="#2C2C2C")
        frame.pack(pady=20, padx=20, fill="both", expand=True)

        self.texto_status = ctk.CTkTextbox(frame, width=750, height=200, font=("Arial", 12),
                                           corner_radius=10, fg_color="#333333", text_color="#FFC107")
        self.texto_status.pack(pady=10, padx=10)

        self.progress_bar = ctk.CTkProgressBar(frame, width=700, fg_color="#444444", progress_color="#FFC107")
        self.progress_bar.pack(pady=10)

        botoes_frame = ctk.CTkFrame(frame, fg_color="transparent")
        botoes_frame.pack(pady=10)

        btn_selecionar = ctk.CTkButton(botoes_frame, text="Selecionar PDFs",
                                       command=self.selecionar_pdfs, corner_radius=10, width=180,
                                       fg_color="#FFC107", text_color="#2C2C2C")
        btn_selecionar.grid(row=0, column=0, padx=10, pady=5)

        btn_diretorio = ctk.CTkButton(botoes_frame, text="Selecionar Diretório",
                                      command=self.selecionar_diretorio, corner_radius=10, width=180,
                                      fg_color="#FF6347", text_color="#2C2C2C")
        btn_diretorio.grid(row=0, column=1, padx=10, pady=5)

        btn_processar = ctk.CTkButton(botoes_frame, text="Processar e Salvar",
                                      command=self.processar_e_salvar, corner_radius=10, width=180,
                                      fg_color="#FFD700", text_color="#2C2C2C")
        btn_processar.grid(row=0, column=2, padx=10, pady=5)

        self.arquivos_pdf = []

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
        if not self.arquivos_pdf:
            messagebox.showwarning("Aviso", "Nenhum arquivo PDF foi selecionado.")
            return
        if not self.diretorio_saida:
            messagebox.showwarning("Aviso", "Nenhum diretório de saída foi selecionado.")
            return

        self.texto_status.insert("end", "Processando e salvando arquivos...\n")
        os.makedirs(self.diretorio_saida, exist_ok=True)

        for index, caminho_pdf in enumerate(self.arquivos_pdf, start=1):
            df = extrair_dados_ocr(caminho_pdf)
            if not df.empty:
                df = formatar_valor(df)
                df = adicionar_colunas_personalizadas(df)
                df = adicionar_coluna_codigo(df)
                salvar_excel_formatado(df, caminho_pdf, self.diretorio_saida)
                salvar_txt_formatado(df, caminho_pdf, self.diretorio_saida)
                self.texto_status.insert("end", f"Arquivo {index}/{len(self.arquivos_pdf)} processado e salvo.\n")
            else:
                self.texto_status.insert("end", f"Falha ao processar arquivo: {caminho_pdf}\n")

            self.progress_bar.set(index / len(self.arquivos_pdf))
            self.root.update()

        self.texto_status.insert("end", "Todos os arquivos foram processados e salvos no diretório selecionado.\n")


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
