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
                "volume": 0.0, "faturamento": 0.0, "custo": 0.0, "produtos": []
            }
            produtos_atual = []
            continue
        m = re_totais_cliente.search(linha_strip)
        if m and cliente_atual:
            cliente_atual["volume"] = limpar_numero(m.group(1))
            cliente_atual["faturamento"] = limpar_numero(m.group(2))
            cliente_atual["custo"] = limpar_numero(m.group(3))
            continue
        m = re_totais_vendedor.search(linha_strip)
        if not m:
            m = re_total_geral.search(linha_strip)
        if m:
            total_fat_rodape = limpar_numero(m.group(2))
            total_custo_rodape = limpar_numero(m.group(3))
            continue
        m = re_produto.search(linha_strip)
        if m and cliente_atual:
            prod = {
                "descricao": m.group(2).strip(),
                "qtd": limpar_numero(m.group(3)),
                "faturamento": limpar_numero(m.group(4)),
                "custo_unit": limpar_numero(m.group(5)),
                "custo_total": limpar_numero(m.group(6)),
            }
            produtos_atual.append(prod)

    if cliente_atual:
        cliente_atual["produtos"] = produtos_atual
        clientes.append(cliente_atual)

    return {
        "vendedor": vendedor, "data": data_ref, "clientes": clientes,
        "total_faturamento": total_fat_rodape, "total_custo": total_custo_rodape,
    }

def validar_totais(dados: dict) -> tuple:
    soma_fat = sum(c["faturamento"] for c in dados["clientes"])
    soma_custo = sum(c["custo"] for c in dados["clientes"])
    rodape_fat = dados["total_faturamento"]
    rodape_custo = dados["total_custo"]
    ok_fat = abs(soma_fat - rodape_fat) < 1.0
    ok_custo = abs(soma_custo - rodape_custo) < 1.0
    if ok_fat and ok_custo:
        return True, "✅ Totais validados com sucesso"
    else:
        msg = f"⚠️ Divergência: Fat R${soma_fat:,.2f} vs rodapé R${rodape_fat:,.2f} | Custo R${soma_custo:,.2f} vs rodapé R${rodape_custo:,.2f}"
        return False, msg


# ══════════════════════════════════════════════════════════════════════════
# PARSER ESPECÍFICO: Lucratividade por Vendedor-Faturamento no Previsão
# Usado na página de Metas Semanais
# ══════════════════════════════════════════════════════════════════════════

def extrair_vendas_por_vendedor(pdf_path: str) -> dict:
    """
    Extrai do PDF 'Lucratividade por Vendedor-Faturamento no Previsão':
    Retorna dict {vendedor_canonico: {descricao_produto: qtd_vendida}}
    Ignora LUCA automaticamente.
    """
    texto = pdf_para_texto(pdf_path)
    linhas = texto.splitlines()

    resultado = {}
    vendedor_atual = None

    re_vendedor_secao = re.compile(
        r"Vendedor\s*:\s*\d+\s+(.+)", re.IGNORECASE
    )
    re_total_vendedor = re.compile(
        r"Total\s+por\s+Vendedor", re.IGNORECASE
    )
    # Linha de produto: código descrição CX qtd dev total val.unit custo lucro lucro_unit lucro%
    re_linha_produto = re.compile(
        r"^\S+\s+(.+?)\s+CX\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)"
    )

    for linha in linhas:
        linha_strip = linha.strip()
        if not linha_strip:
            continue

        # Nova seção de vendedor
        m = re_vendedor_secao.match(linha_strip)
        if m:
            nome_raw = m.group(1).strip()
            vendedor_canonico = normalizar_vendedor(nome_raw)
            # Ignora Luca
            if vendedor_canonico not in VENDEDORES_ATIVOS:
                vendedor_atual = None
            else:
                vendedor_atual = vendedor_canonico
                if vendedor_atual not in resultado:
                    resultado[vendedor_atual] = {}
            continue

        # Fim da seção
        if re_total_vendedor.search(linha_strip):
            vendedor_atual = None
            continue

        # Linha de produto
        if vendedor_atual:
            m = re_linha_produto.match(linha_strip)
            if m:
                descricao = m.group(1).strip()
                qtd = limpar_numero(m.group(2))
                # Remove sufixo do vendedor colado na descrição (ex: "-FARLEY - FARLEY")
                descricao = re.sub(r'\s*-\s*\w+\s*-\s*\w+\s*$', '', descricao).strip()
                if qtd > 0:
                    if descricao not in resultado[vendedor_atual]:
                        resultado[vendedor_atual][descricao] = 0.0
                    resultado[vendedor_atual][descricao] += qtd

    return resultado