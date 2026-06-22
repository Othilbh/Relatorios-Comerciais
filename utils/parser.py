# utils/parser.py
import re
import subprocess
import tempfile
import os
from pathlib import Path

import pytesseract
from PIL import Image
from pdf2image import convert_from_path

VENDEDOR_MAP = {
    "RONISTONIS": "Roni",
    "RONI": "Roni",
    "ADILSON": "Dora",
    "ADILSON-DORA": "Dora",
    "ADILSON - DORA": "Dora",
    "DORA": "Dora",
    "REGIS": "Reginaldo",
    "REGINALDO": "Reginaldo",
    "FALEY": "Farley",
    "FARLEY": "Farley",
    "AFANAIS": "Afanais",
    "LUCIANO": "Luciano",
    "JULIANA": "Juliana",
    "CLAUDIA": "Claudia",
    "CLÁUDIA": "Claudia",
}

VENDEDORES_ATIVOS = ["Farley", "Dora", "Afanais", "Roni", "Reginaldo", "Luciano", "Juliana", "Claudia"]

def normalizar_vendedor(nome: str) -> str:
    upper = nome.strip().upper().replace("-", " ").replace("  ", " ")
    for chave, canonico in VENDEDOR_MAP.items():
        if chave.upper() in upper:
            return canonico
    return nome.strip().title()

def pdf_para_texto(pdf_path: str) -> str:
    try:
        resultado = subprocess.run(
            ["pdftotext", "-layout", pdf_path, "-"],
            capture_output=True, text=True, timeout=60
        )
        texto = resultado.stdout.strip()
        if len(texto) > 100:
            return texto
    except Exception:
        pass
    return pdf_ocr(pdf_path)

def pdf_ocr(pdf_path: str) -> str:
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            subprocess.run(
                ["pdftoppm", "-jpeg", "-r", "200", pdf_path, os.path.join(tmpdir, "page")],
                capture_output=True, timeout=120, check=True
            )
        except Exception:
            images = convert_from_path(pdf_path, dpi=200)
            for i, img in enumerate(images):
                img.save(os.path.join(tmpdir, f"page-{i+1:03d}.jpg"), "JPEG")

        pages = sorted(Path(tmpdir).glob("*.jpg"))
        textos = []
        for pg in pages:
            img = Image.open(pg)
            t = pytesseract.image_to_string(img, lang="por", config="--psm 6")
            textos.append(t)
        return "\n".join(textos)

def limpar_numero(s: str) -> float:
    s = s.strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0

def _remover_sufixo_vendedor(nome: str) -> str:
    for suf in sorted(VENDEDOR_MAP.keys(), key=len, reverse=True):
        if nome.upper().endswith(suf.upper()):
            nome = nome[:-len(suf)].strip()
            break
    return nome

def extrair_dados_pdf(pdf_path: str) -> dict:
    texto = pdf_para_texto(pdf_path)
    linhas = texto.splitlines()

    vendedor = ""
    data_ref = ""
    clientes = []
    cliente_atual = None
    produtos_atual = []
    total_fat_rodape = 0.0
    total_custo_rodape = 0.0

    re_vendedor = re.compile(r"Vendedor[:\s]+(.+)", re.IGNORECASE)
    re_data = re.compile(r"Per[ií]odo[:\s]+(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
    re_cliente = re.compile(r"^Cliente[:\s]+(.+)", re.IGNORECASE)
    re_totais_cliente = re.compile(
        r"Totais\s+do\s+Cliente\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)", re.IGNORECASE)
    re_totais_vendedor = re.compile(
        r"Totais\s+do\s+Vendedor\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)", re.IGNORECASE)
    re_total_geral = re.compile(
        r"Total\s+Geral\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)", re.IGNORECASE)
    re_produto = re.compile(
        r"(\d+)\s+(.+?)\s+([\d.,]+)\s+CX\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)")

    for linha in linhas:
        linha_strip = linha.strip()
        if not linha_strip:
            continue

        m = re_vendedor.match(linha_strip)
        if m and not vendedor:
            vendedor = normalizar_vendedor(m.group(1))
            continue

        m = re_data.search(linha_strip)
        if m and not data_ref:
            data_ref = m.group(1)
            continue

        m = re_cliente.match(linha_strip)
        if m:
            if cliente_atual:
                cliente_atual["produtos"] = produtos_atual
                clientes.append(cliente_atual)
            nome_cli = _remover_sufixo_vendedor(m.group(1).strip())
            cliente_atual = {
                "cliente": nome_cli,
                "volume": 0.0,
                "faturamento": 0.0,
                "custo": 0.0,
                "produtos": []
            }
            produtos_atual = []
            continue

        m = re_totais_cliente.search(linha_strip)
        if m and cliente_atual:
            cliente_atual["volume"] = limpar_numero(m.group(1))
            cliente_atual["faturamento"] = limpar_numero(m.group(2))