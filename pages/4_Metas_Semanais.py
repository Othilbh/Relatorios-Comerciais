# pages/4_Metas_Semanais.py
import streamlit as st
from datetime import date, timedelta
import sys, os, tempfile, math, io, json, base64, requests
import pandas as pd
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

from utils.parser import extrair_vendas_por_vendedor, extrair_estoque_por_vendedor, VENDEDORES_ATIVOS
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm

st.set_page_config(page_title="Metas Semanais · OTHIL", page_icon="🎯", layout="wide")
st.title("🎯 Metas Semanais")
st.caption("Relatório diário por vendedor + metas semanais. Luca excluído automaticamente.")
st.divider()

PERCENTUAIS = {
    "Farley":0.175,"Dora":0.175,"Afanais":0.250,"Roni":0.250,
    "Reginaldo":0.225,"Juliana":0.075,"Claudia":0.075,"Luciano":0.075,
}

PRODUTOS_DEFAULT = [
    "Portuguesa Sabor 55/60","Portuguesa 60/70","Portuguesa 70/80",
    "Forelle","Ercoline","Pera Asiática","Gala Santa Carol","Gala Azaleia",
    "Fuji Expressa 180","Fuji Azaleia","Fuji Hiragami","Fuji Suprema",
    "Thompson Vitace","Thompson Seedless","Uva Crimson","Uva Isis",
    "Uva Jubilee","Uva Itália","Maçã Argentina","Maçã Chilena",
    "Maçã Pink Lady","Maçã Granny Smith","Maçã Red Globe",
    "Mamão Havai","Mamão Formoso","Goiaba","Melão Amarelo","Melão Gaia",
    "Melão Cantaloupe","Tangerina Cumbuca","Tangerina Ponkan",
    "Ameixa","Pêssego","Nectarina","Morango","Mirtilo","Abacaxi",
    "Manga Palmer","Manga Tommy","Laranja","Limão Tahiti","Limão Siciliano","Tomate Roma",
]

MAPA_PRODUTO = {
    "Portuguesa Sabor 55/60":["PORTUGUESA SABOR","PORTUGUESA 55","02050032"],
    "Portuguesa 60/70":["PORTUGUESA 60"],"Portuguesa 70/80":["PORTUGUESA 70"],
    "Forelle":["FORELLE","020502701"],"Ercoline":["ERCOLINE"],
    "Pera Asiática":["ASIATICA","ASIÁTICA"],"Gala Santa Carol":["GALA SANTA","SANTA CAROL"],
    "Gala Azaleia":["GALA AZALEIA"],"Fuji Expressa 180":["FUJI EXPRESSA","FUJI 180"],
    "Fuji Azaleia":["FUJI AZALEIA","702145135","702145165"],"Fuji Hiragami":["HIRAGAMI"],
    "Fuji Suprema":["FUJI SUPREMA"],"Thompson Vitace":["THOMPSON VITACE"],
    "Thompson Seedless":["THOMPSON SEEDLESS"],"Uva Crimson":["CRINSON","CRIMSON"],
    "Uva Isis":["ISIS"],"Uva Jubilee":["JUBILEE"],"Uva Itália":["ITALIA","ITÁLIA"],
    "Maçã Argentina":["ARGENTINA"],"Maçã Chilena":["CHILENA"],
    "Maçã Pink Lady":["PINK LADY"],"Maçã Granny Smith":["GRAN SMITH","GRANNY SMITH"],
    "Maçã Red Globe":["RED GLOBE"],"Mamão Havai":["MAMAO HAVAI","MAMÃO HAVAI"],
    "Mamão Formoso":["MAMAO FORMOSO","MAMÃO FORMOSO"],
    "Goiaba":["GOIABA","300200203","300200208"],
    "Melão Amarelo":["MELAO AMARELO","MELÃO AMARELO"],
    "Melão Gaia":["MELAO GAIA","MELÃO GAIA","MELAO GALIA","MELÃO GALIA","3102006"],
    "Melão Cantaloupe":["CANTALOUPE"],"Tangerina Cumbuca":["CUMBUCA","830100903"],
    "Tangerina Ponkan":["PONKAN"],"Ameixa":["AMEIXA"],
    "Pêssego":["PESSEGO","PÊSSEGO"],"Nectarina":["NECTARINA"],
    "Morango":["MORANGO"],"Mirtilo":["MIRTILO","BLUEBERRY"],"Abacaxi":["ABACAXI"],
    "Manga Palmer":["MANGA PALMER"],"Manga Tommy":["MANGA TOMMY"],"Laranja":["LARANJA"],
    "Limão Tahiti":["LIMAO TAHITI","LIMÃO TAHITI"],
    "Limão Siciliano":["LIMAO SICILIANO","LIMÃO SICILIANO"],
    "Tomate Roma":["TOMATE ROMA","ROMA"],
}

RODAPE = [
    ("É DE RESPONSABILIDADE DO VENDEDOR:",True),
    ("AVALIAR DIARIAMENTE A QUALIDADE E ARMAZENAGEM DE CADA PRODUTO DE SUA RESPONSABILIDADE.",False),
    ("CONFERIR O QUE ESTA EM CADA PAVILHÃO",False),
    ("CONFERIR O QUE ESTA NA VENDA FUTURA E ACOMPANHAR DIARIAMENTE",False),
    ("CONFERIR O QUE ESTA ARMAZENADO EM OUTROS FRIGORIFICOS",False),
    ("VENDER ATÉ A ÚLTIMA CAIXA",False),
    ("DEVOLUCAO SO SE FOR NO MESMO DIA",False),
    ("MERCADORIAS NO SOL",False),
    ("CAMINHOES REFRIGERADOS SEMPRE FECHADOS",False),
]

CORES_VEND = {
    "Farley":   (colors.HexColor("#1A3A5C"),colors.HexColor("#D6E4F0"),colors.HexColor("#1A3A5C")),
    "Dora":     (colors.HexColor("#1B5E20"),colors.HexColor("#DCEDC8"),colors.HexColor("#1B5E20")),
    "Afanais":  (colors.HexColor("#4A148C"),colors.HexColor("#E8DAEF"),colors.HexColor("#4A148C")),
    "Roni":     (colors.HexColor("#BF360C"),colors.HexColor("#FDEBD0"),colors.HexColor("#BF360C")),
    "Reginaldo":(colors.HexColor("#B71C1C"),colors.HexColor("#FADBD8"),colors.HexColor("#B71C1C")),
    "Luciano":  (colors.HexColor("#006064"),colors.HexColor("#D0ECE7"),colors.HexColor("#006064")),
    "Juliana":  (colors.HexColor("#880E4F"),colors.HexColor("#FDEDEC"),colors.HexColor("#880E4F")),
    "Claudia":  (colors.HexColor("#795548"),colors.HexColor("#D7CCC8"),colors.HexColor("#795548")),
}

COR_H     = colors.HexColor("#1A3A5C")
COR_CINZA = colors.HexColor("#D9D9D9")
COR_ALT   = colors.HexColor("#F5F5F5")
COR_TOTAL = colors.HexColor("#BDD7EE")
COR_V = colors.HexColor("#C6EFCE"); COR_VF = colors.HexColor("#276221")
COR_A = colors.HexColor("#FFEB9C"); COR_AF = colors.HexColor("#7D6608")
COR_R = colors.HexColor("#FFC7CE"); COR_RF = colors.HexColor("#9C0006")
BRANCO = colors.white; PRETO = colors.black

def mapear_produto(desc):
    desc_u = desc.upper()
    for prod_meta, keywords in MAPA_PRODUTO.items():
        for kw in keywords:
            if kw in desc_u: return prod_meta
    return None

def cor_pct(pct):
    if pct >= 100: return COR_V, COR_VF
    if pct >= 50:  return COR_A, COR_AF
    return COR_R, COR_RF

def brl(val):
    return f"R$ {val:,.2f}".replace(",","X").replace(".",",").replace("X",".")

def tabela_titulo(texto, bg=None, fg=BRANCO, sz=11, largura=19):
    bg = bg or COR_H
    t = Table([[texto]], colWidths=[largura*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),bg),("TEXTCOLOR",(0,0),(-1,-1),fg),
        ("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),sz),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
        ("BOX",(0,0),(-1,-1),1.5,bg),
    ]))
    return t
    # ── GitHub ────────────────────────────────────────────────────────────────
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN","")
GITHUB_REPO  = st.secrets.get("GITHUB_REPO","Othilbh/Relatorios-Comerciais")
GITHUB_FILE_METAS    = "data/metas_semana.json"
GITHUB_FILE_PRODUTOS = "data/produtos_extra.json"

def github_get(filepath):
    if not GITHUB_TOKEN: return None, None
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{filepath}"
    r = requests.get(url, headers={"Authorization":f"token {GITHUB_TOKEN}"})
    if r.status_code == 200:
        data = r.json()
        return json.loads(base64.b64decode(data["content"]).decode("utf-8")), data["sha"]
    return None, None

def github_save(filepath, content, sha=None):
    if not GITHUB_TOKEN: return False
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{filepath}"
    body = {"message":f"Atualiza {filepath}",
            "content":base64.b64encode(json.dumps(content,ensure_ascii=False,indent=2).encode()).decode()}
    if sha: body["sha"] = sha
    r = requests.put(url, headers={"Authorization":f"token {GITHUB_TOKEN}"}, json=body)
    return r.status_code in [200,201]

if "produtos_meta" not in st.session_state:
    dados, _ = github_get(GITHUB_FILE_METAS)
    st.session_state.produtos_meta = dados or []
if "produtos_extra" not in st.session_state:
    dados, _ = github_get(GITHUB_FILE_PRODUTOS)
    st.session_state.produtos_extra = dados or []
if "vendido" not in st.session_state: st.session_state.vendido = {}
if "estoque" not in st.session_state: st.session_state.estoque = {}
if "sem_vendedor" not in st.session_state: st.session_state.sem_vendedor = []

PRODUTOS_LISTA = PRODUTOS_DEFAULT + [p for p in st.session_state.produtos_extra if p not in PRODUTOS_DEFAULT]

# ══════════════════════════════════════════════════════════════════════════
# PDF VENDEDOR
# ══════════════════════════════════════════════════════════════════════════
def gerar_pdf_vendedor(vendedor, data_ref, itens_estoque, metas, vendido_v):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=1*cm, rightMargin=1*cm, topMargin=1*cm, bottomMargin=1*cm)
    story = []
    story.append(tabela_titulo(f"Vendedor : {vendedor.upper()}      Data: {data_ref}", sz=12))
    story.append(Spacer(1,8))

    est_header = [["Produto","Complemento","Dt.Entrada","Saldo","Qtde Vend","Custo Unit","Md Venda"]]
    est_data = []; est_style = [
        ("BACKGROUND",(0,0),(-1,0),COR_CINZA),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),8),("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("ALIGN",(0,1),(0,-1),"LEFT"),("ALIGN",(5,1),(6,-1),"RIGHT"),
        ("GRID",(0,0),(-1,-1),0.5,PRETO),
        ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),
        ("LEFTPADDING",(0,0),(-1,-1),4),
    ]
    for i, item in enumerate(itens_estoque, 1):
        bg = BRANCO if i%2==1 else COR_ALT
        est_style.append(("BACKGROUND",(0,i),(-1,i),bg))
        if item["qtd_vendida"] == 0: est_style.append(("TEXTCOLOR",(4,i),(4,i),COR_RF))
        if item["md_venda"] == 0: est_style.append(("TEXTCOLOR",(6,i),(6,i),COR_RF))
        est_data.append([item["descricao"],item["complemento"],item["data_entrada"],
            str(item["saldo_atual"]),str(item["qtd_vendida"]),brl(item["custo"]),brl(item["md_venda"])])
    est_t = Table(est_header+est_data, colWidths=[5.5*cm,2.5*cm,2.2*cm,1.5*cm,1.8*cm,2.5*cm,2.5*cm])
    est_t.setStyle(TableStyle(est_style))
    story.append(est_t)
    story.append(Spacer(1,10))

    story.append(tabela_titulo(f"METAS SEMANAIS — {vendedor.upper()} — {data_ref}", sz=11))
    meta_header = [["Produto","Meta (cx)","Vendido (cx)","Falta (cx)","%"]]
    meta_data = []; meta_style = [
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#2E6DA4")),
        ("TEXTCOLOR",(0,0),(-1,0),BRANCO),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),9),("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("ALIGN",(0,1),(0,-1),"LEFT"),("GRID",(0,0),(-1,-1),0.5,PRETO),
        ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
        ("LEFTPADDING",(0,0),(-1,-1),4),
    ]
    total_meta = total_vend = 0
    for i, item in enumerate(metas, 1):
        prod=item["produto"]; est=item["estoque"]
        meta=math.ceil(est*PERCENTUAIS.get(vendedor,0))
        vend=vendido_v.get(prod,0.0)
        falta=max(meta-vend,0); pct=(vend/meta*100) if meta>0 else 0.0
        total_meta+=meta; total_vend+=vend
        bg_m,fg_m=cor_pct(pct)
        meta_data.append([prod,str(meta),str(round(vend,1)),str(round(falta,1)),f"{pct:.1f}%"])
        meta_style+=[ ("BACKGROUND",(4,i),(4,i),bg_m),("TEXTCOLOR",(4,i),(4,i),fg_m),
                      ("FONTNAME",(4,i),(4,i),"Helvetica-Bold") ]
        if vend==0: meta_style.append(("TEXTCOLOR",(2,i),(2,i),COR_RF))

    pct_t=(total_vend/total_meta*100) if total_meta>0 else 0
    bg_t,fg_t=cor_pct(pct_t); ti=len(metas)+1
    meta_data.append(["TOTAL",str(total_meta),str(round(total_vend,1)),
                      str(round(max(total_meta-total_vend,0),1)),f"{pct_t:.1f}%"])
    meta_style+=[ ("BACKGROUND",(0,ti),(3,ti),COR_TOTAL),
                  ("BACKGROUND",(4,ti),(4,ti),bg_t),("TEXTCOLOR",(4,ti),(4,ti),fg_t),
                  ("FONTNAME",(0,ti),(-1,ti),"Helvetica-Bold") ]
    meta_t = Table(meta_header+meta_data, colWidths=[7*cm,3*cm,3*cm,3*cm,3*cm])
    meta_t.setStyle(TableStyle(meta_style))
    story.append(meta_t)
    story.append(Spacer(1,10))

    bg_kpi,fg_kpi=cor_pct(pct_t)
    kpi_t = Table([[
        f"Meta Total\n{int(total_meta)} cx",f"Vendido\n{int(total_vend)} cx",
        f"Falta\n{int(max(total_meta-total_vend,0))} cx",f"% Atingido\n{pct_t:.1f}%",
    ]], colWidths=[4.75*cm]*4)
    kpi_t.setStyle(TableStyle([
        ("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),12),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("TOPPADDING",(0,0),(-1,-1),10),("BOTTOMPADDING",(0,0),(-1,-1),10),
        ("BACKGROUND",(0,0),(0,0),colors.HexColor("#2E6DA4")),("TEXTCOLOR",(0,0),(0,0),BRANCO),
        ("BACKGROUND",(1,0),(1,0),COR_V),("TEXTCOLOR",(1,0),(1,0),COR_VF),
        ("BACKGROUND",(2,0),(2,0),COR_R),("TEXTCOLOR",(2,0),(2,0),COR_RF),
        ("BACKGROUND",(3,0),(3,0),bg_kpi),("TEXTCOLOR",(3,0),(3,0),fg_kpi),
        ("BOX",(0,0),(0,0),1,PRETO),("BOX",(1,0),(1,0),1,PRETO),
        ("BOX",(2,0),(2,0),1,PRETO),("BOX",(3,0),(3,0),1,PRETO),
    ]))
    story.append(kpi_t)
    story.append(Spacer(1,12))

    for texto, bold in RODAPE:
        st_r = ParagraphStyle("rod", fontSize=9 if not bold else 10,
            textColor=COR_RF if bold else PRETO,
            fontName="Helvetica-Bold" if bold else "Helvetica")
        story.append(Paragraph(texto, st_r))
        story.append(Spacer(1,2))

    doc.build(story)
    buf.seek(0)
    return buf

# ══════════════════════════════════════════════════════════════════════════
# PDF DASHBOARD
# ══════════════════════════════════════════════════════════════════════════
def gerar_pdf_dashboard(metas, vendido_todos, data_ref, periodo):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=1*cm, rightMargin=1*cm, topMargin=1*cm, bottomMargin=1*cm)
    story = []
    story.append(tabela_titulo(f"OTHIL — DASHBOARD DE METAS  |  {periodo}", sz=13))
    story.append(Spacer(1,8))

    total_meta_g=sum(math.ceil(m["estoque"]*PERCENTUAIS.get(v,0)) for m in metas for v in VENDEDORES_ATIVOS)
    total_vend_g=sum(vendido_todos.get(v,{}).get(m["produto"],0) for m in metas for v in VENDEDORES_ATIVOS)
    pct_g=(total_vend_g/total_meta_g*100) if total_meta_g>0 else 0
    n_crit=sum(1 for m in metas if
        sum(vendido_todos.get(v,{}).get(m["produto"],0) for v in VENDEDORES_ATIVOS)/
        max(sum(math.ceil(m["estoque"]*PERCENTUAIS.get(v,0)) for v in VENDEDORES_ATIVOS),1)*100<50)
    bg_g,fg_g=cor_pct(pct_g)

    kpi_t=Table([[
        f"Meta Total\n{int(total_meta_g):,} cx",f"Vendido\n{int(total_vend_g):,} cx",
        f"Falta\n{int(max(total_meta_g-total_vend_g,0)):,} cx",f"% Atingido\n{pct_g:.1f}%",
        f"Criticos <50%\n{n_crit}",f"Produtos\n{len(metas)}",
    ]], colWidths=[3.17*cm]*6)
    kpi_t.setStyle(TableStyle([
        ("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),10),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("TOPPADDING",(0,0),(-1,-1),10),("BOTTOMPADDING",(0,0),(-1,-1),10),
        ("BACKGROUND",(0,0),(0,0),colors.HexColor("#2E6DA4")),("TEXTCOLOR",(0,0),(0,0),BRANCO),
        ("BACKGROUND",(1,0),(1,0),COR_V),("TEXTCOLOR",(1,0),(1,0),COR_VF),
        ("BACKGROUND",(2,0),(2,0),COR_R),("TEXTCOLOR",(2,0),(2,0),COR_RF),
        ("BACKGROUND",(3,0),(3,0),bg_g),("TEXTCOLOR",(3,0),(3,0),fg_g),
        ("BACKGROUND",(4,0),(4,0),colors.HexColor("#C62828")),("TEXTCOLOR",(4,0),(4,0),BRANCO),
        ("BACKGROUND",(5,0),(5,0),COR_H),("TEXTCOLOR",(5,0),(5,0),BRANCO),
        ("BOX",(0,0),(0,0),1,PRETO),("BOX",(1,0),(1,0),1,PRETO),
        ("BOX",(2,0),(2,0),1,PRETO),("BOX",(3,0),(3,0),1,PRETO),
        ("BOX",(4,0),(4,0),1,PRETO),("BOX",(5,0),(5,0),1,PRETO),
    ]))
    story.append(kpi_t)
    story.append(Spacer(1,12))

    story.append(tabela_titulo("RANKING DE VENDEDORES",bg=colors.HexColor("#2E6DA4"),sz=10))
    rank_header=[["#","Vendedor","Meta (cx)","Vendido (cx)","Falta (cx)","% Atingido","Status","% Meta"]]
    rank_data=[]; rank_style=[
        ("BACKGROUND",(0,0),(-1,0),COR_CINZA),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),9),("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("ALIGN",(1,1),(1,-1),"LEFT"),("GRID",(0,0),(-1,-1),0.5,PRETO),
        ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
    ]
    ranking=[]
    for v in VENDEDORES_ATIVOS:
        mv=sum(math.ceil(m["estoque"]*PERCENTUAIS.get(v,0)) for m in metas)
        vv=sum(vendido_todos.get(v,{}).get(m["produto"],0) for m in metas)
        pv=(vv/mv*100) if mv>0 else 0
        ranking.append((v,mv,round(vv,1),max(mv-vv,0),round(pv,1)))
    ranking.sort(key=lambda x:x[4],reverse=True)
    for pos,(v,mv,vv,fv,pv) in enumerate(ranking,1):
        bg_r,fg_r=cor_pct(pv)
        st_txt="Atingiu" if pv>=100 else "Andamento" if pv>=50 else "Abaixo"
        bg_row=BRANCO if pos%2==0 else COR_ALT
        rank_data.append([f"{pos}°",v,str(mv),str(round(vv,1)),str(round(fv,1)),
                          f"{pv:.1f}%",st_txt,f"{int(PERCENTUAIS.get(v,0)*100)}%"])
        i=pos
        rank_style+=[ ("BACKGROUND",(0,i),(-1,i),bg_row),
                      ("BACKGROUND",(5,i),(5,i),bg_r),("TEXTCOLOR",(5,i),(5,i),fg_r),
                      ("FONTNAME",(5,i),(5,i),"Helvetica-Bold"),
                      ("BACKGROUND",(6,i),(6,i),bg_r),("TEXTCOLOR",(6,i),(6,i),fg_r),
                      ("TEXTCOLOR",(1,i),(1,i),COR_H) ]
        if vv==0: rank_style.append(("TEXTCOLOR",(3,i),(3,i),COR_RF))
    rank_t=Table(rank_header+rank_data,
        colWidths=[1.2*cm,3*cm,2.5*cm,2.5*cm,2.5*cm,2.5*cm,2.5*cm,2.3*cm])
    rank_t.setStyle(TableStyle(rank_style))
    story.append(rank_t)
    story.append(Spacer(1,12))

    story.append(tabela_titulo("PRODUTOS CRITICOS — ABAIXO DE 50%",
        bg=colors.HexColor("#C62828"),sz=10))
    crit_header=[["Produto","Meta Total","Vendido","Falta","% Geral","Melhor Vendedor"]]
    crit_data=[]; crit_style=[
        ("BACKGROUND",(0,0),(-1,0),COR_CINZA),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),9),("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("ALIGN",(0,1),(0,-1),"LEFT"),("GRID",(0,0),(-1,-1),0.5,PRETO),
        ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
    ]
    criticos=[]
    for m in metas:
        prod=m["produto"]; est=m["estoque"]
        mt=sum(math.ceil(est*PERCENTUAIS.get(v,0)) for v in VENDEDORES_ATIVOS)
        vt=sum(vendido_todos.get(v,{}).get(prod,0) for v in VENDEDORES_ATIVOS)
        pt=(vt/mt*100) if mt>0 else 0
        if pt<50:
            melhor=max(VENDEDORES_ATIVOS,
                key=lambda v:vendido_todos.get(v,{}).get(prod,0)/
                max(math.ceil(est*PERCENTUAIS.get(v,0)),1)*100)
            criticos.append((prod,mt,round(vt,1),round(max(mt-vt,0),1),round(pt,1),melhor))
    criticos.sort(key=lambda x:x[4])
    if not criticos:
        crit_data.append(["Todos os produtos acima de 50%!","","","","",""])
        crit_style+=[ ("BACKGROUND",(0,1),(-1,1),COR_V),("TEXTCOLOR",(0,1),(-1,1),COR_VF) ]
    else:
        for i,(prod,mt,vt,ft,pt,melhor) in enumerate(criticos,1):
            bg_c,fg_c=cor_pct(pt)
            crit_data.append([prod,str(mt),str(vt),str(ft),f"{pt:.1f}%",melhor])
            crit_style+=[ ("BACKGROUND",(0,i),(3,i),COR_R),("TEXTCOLOR",(0,i),(3,i),COR_RF),
                          ("BACKGROUND",(4,i),(4,i),bg_c),("TEXTCOLOR",(4,i),(4,i),fg_c),
                          ("FONTNAME",(4,i),(4,i),"Helvetica-Bold"),
                          ("BACKGROUND",(5,i),(5,i),COR_V),("TEXTCOLOR",(5,i),(5,i),COR_VF),
                          ("FONTNAME",(5,i),(5,i),"Helvetica-Bold") ]
    crit_t=Table(crit_header+crit_data,colWidths=[5*cm,2.5*cm,2.5*cm,2.5*cm,2.5*cm,4*cm])
    crit_t.setStyle(TableStyle(crit_style))
    story.append(crit_t)
    doc.build(story)
    buf.seek(0)
    return buf
    # ══════════════════════════════════════════════════════════════════════════
# PDF RESUMO GERAL
# ══════════════════════════════════════════════════════════════════════════
def gerar_pdf_resumo(metas, vendido_todos, data_ref, periodo):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
        leftMargin=0.8*cm, rightMargin=0.8*cm, topMargin=0.8*cm, bottomMargin=0.8*cm)
    story = []

    titulo_t=Table([[f"OTHIL — RESUMO GERAL DE METAS  |  {periodo}  |  {data_ref}"]],
        colWidths=[27.6*cm])
    titulo_t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),COR_H),("TEXTCOLOR",(0,0),(-1,-1),BRANCO),
        ("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),12),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7),
    ]))
    story.append(titulo_t)
    story.append(Spacer(1,6))

    total_meta_g=sum(math.ceil(m["estoque"]*PERCENTUAIS.get(v,0)) for m in metas for v in VENDEDORES_ATIVOS)
    total_vend_g=sum(vendido_todos.get(v,{}).get(m["produto"],0) for m in metas for v in VENDEDORES_ATIVOS)
    pct_g=(total_vend_g/total_meta_g*100) if total_meta_g>0 else 0
    bg_g,fg_g=cor_pct(pct_g)

    kpi_t=Table([[
        f"Estoque Total: {sum(m['estoque'] for m in metas):,} cx",
        f"Meta Total: {int(total_meta_g):,} cx",
        f"Vendido Total: {int(total_vend_g):,} cx",
        f"Falta: {int(max(total_meta_g-total_vend_g,0)):,} cx",
        f"% Geral Atingido: {pct_g:.1f}%",
    ]],colWidths=[5.52*cm]*5)
    kpi_t.setStyle(TableStyle([
        ("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),9),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
        ("BACKGROUND",(0,0),(0,0),colors.HexColor("#455A64")),("TEXTCOLOR",(0,0),(0,0),BRANCO),
        ("BACKGROUND",(1,0),(1,0),COR_H),("TEXTCOLOR",(1,0),(1,0),BRANCO),
        ("BACKGROUND",(2,0),(2,0),COR_V),("TEXTCOLOR",(2,0),(2,0),COR_VF),
        ("BACKGROUND",(3,0),(3,0),COR_R),("TEXTCOLOR",(3,0),(3,0),COR_RF),
        ("BACKGROUND",(4,0),(4,0),bg_g),("TEXTCOLOR",(4,0),(4,0),fg_g),
        ("BOX",(0,0),(0,0),1,PRETO),("BOX",(1,0),(1,0),1,PRETO),
        ("BOX",(2,0),(2,0),1,PRETO),("BOX",(3,0),(3,0),1,PRETO),("BOX",(4,0),(4,0),1,PRETO),
    ]))
    story.append(kpi_t)
    story.append(Spacer(1,8))

    col_prod=4.2*cm; col_est=1.5*cm; col_vv=1.2*cm; col_pp=1.0*cm
    col_tv=1.4*cm; col_tp=1.1*cm
    all_colwidths=[col_prod,col_est]+[col_vv,col_pp]*len(VENDEDORES_ATIVOS)+[col_tv,col_tp]

    h1=["Produto","Estoque"]
    for v in VENDEDORES_ATIVOS: h1+=[v,""]
    h1+=["TOTAL",""]
    h2=["",""]
    for _ in VENDEDORES_ATIVOS: h2+=["Vend","%"]
    h2+=["Vend","%"]

    table_data=[h1,h2]
    table_style=[
        ("FONTNAME",(0,0),(-1,1),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),7.5),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("ALIGN",(0,2),(0,-1),"LEFT"),
        ("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#AAAAAA")),
        ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),
        ("LEFTPADDING",(0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),
        ("BACKGROUND",(0,0),(0,1),COR_CINZA),("BACKGROUND",(1,0),(1,1),COR_CINZA),
        ("SPAN",(0,0),(0,1)),("SPAN",(1,0),(1,1)),
        ("BACKGROUND",(2+len(VENDEDORES_ATIVOS)*2,0),(-1,1),COR_H),
        ("TEXTCOLOR",(2+len(VENDEDORES_ATIVOS)*2,0),(-1,1),BRANCO),
        ("SPAN",(2+len(VENDEDORES_ATIVOS)*2,0),(2+len(VENDEDORES_ATIVOS)*2+1,0)),
        ("BACKGROUND",(2,1),(2+len(VENDEDORES_ATIVOS)*2-1,1),COR_CINZA),
    ]
    for vi,v in enumerate(VENDEDORES_ATIVOS):
        col_s=2+vi*2
        cor_h,cor_d,cor_txt=CORES_VEND.get(v,(COR_H,COR_ALT,COR_H))
        table_style+=[ ("BACKGROUND",(col_s,0),(col_s+1,0),cor_h),
                       ("TEXTCOLOR",(col_s,0),(col_s+1,0),BRANCO),
                       ("SPAN",(col_s,0),(col_s+1,0)) ]

    tot_vend_v={v:0.0 for v in VENDEDORES_ATIVOS}
    tot_meta_v={v:0 for v in VENDEDORES_ATIVOS}

    for row_i,m in enumerate(metas):
        prod=m["produto"]; est=m["estoque"]
        row=[prod,str(est)]
        tot_vend_l=0
        dr=len(table_data)
        bg_row=BRANCO if row_i%2==0 else COR_ALT
        table_style+=[ ("BACKGROUND",(0,dr),(0,dr),bg_row),("BACKGROUND",(1,dr),(1,dr),bg_row) ]
        for vi,v in enumerate(VENDEDORES_ATIVOS):
            cor_h,cor_d,cor_txt=CORES_VEND.get(v,(COR_H,COR_ALT,COR_H))
            mv=math.ceil(est*PERCENTUAIS.get(v,0))
            vv=round(vendido_todos.get(v,{}).get(prod,0.0),1)
            pv=(vv/mv*100) if mv>0 else 0.0
            tot_vend_v[v]+=vv; tot_meta_v[v]+=mv; tot_vend_l+=vv
            row+=[str(vv),f"{pv:.0f}%"]
            cv=2+vi*2; cp=cv+1
            table_style+=[ ("BACKGROUND",(cv,dr),(cv,dr),cor_d),
                           ("TEXTCOLOR",(cv,dr),(cv,dr),PRETO),
                           ("BACKGROUND",(cp,dr),(cp,dr),cor_d),
                           ("TEXTCOLOR",(cp,dr),(cp,dr),cor_txt),
                           ("FONTNAME",(cp,dr),(cp,dr),"Helvetica-Bold") ]
        tot_meta_l=sum(math.ceil(est*PERCENTUAIS.get(v,0)) for v in VENDEDORES_ATIVOS)
        pct_l=(tot_vend_l/tot_meta_l*100) if tot_meta_l>0 else 0
        ctv=2+len(VENDEDORES_ATIVOS)*2; ctp=ctv+1
        row+=[str(round(tot_vend_l,1)),f"{pct_l:.0f}%"]
        table_style+=[ ("BACKGROUND",(ctv,dr),(ctv,dr),COR_TOTAL),
                       ("FONTNAME",(ctv,dr),(ctv,dr),"Helvetica-Bold"),
                       ("TEXTCOLOR",(ctv,dr),(ctv,dr),COR_H),
                       ("BACKGROUND",(ctp,dr),(ctp,dr),COR_TOTAL),
                       ("TEXTCOLOR",(ctp,dr),(ctp,dr),COR_H),
                       ("FONTNAME",(ctp,dr),(ctp,dr),"Helvetica-Bold") ]
        table_data.append(row)

    grand_v=sum(tot_vend_v.values())
    pct_grand=(grand_v/total_meta_g*100) if total_meta_g>0 else 0
    tot_row=["TOTAL",str(sum(m["estoque"] for m in metas))]
    tr=len(table_data)
    for vi,v in enumerate(VENDEDORES_ATIVOS):
        pv=(tot_vend_v[v]/tot_meta_v[v]*100) if tot_meta_v[v]>0 else 0
        tot_row+=[str(round(tot_vend_v[v],1)),f"{pv:.0f}%"]
    tot_row+=[str(round(grand_v,1)),f"{pct_grand:.0f}%"]
    table_style+=[ ("BACKGROUND",(0,tr),(-1,tr),COR_H),
                   ("TEXTCOLOR",(0,tr),(-1,tr),BRANCO),
                   ("FONTNAME",(0,tr),(-1,tr),"Helvetica-Bold"),
                   ("TEXTCOLOR",(2+len(VENDEDORES_ATIVOS)*2+1,tr),(2+len(VENDEDORES_ATIVOS)*2+1,tr),BRANCO),
                   ("BACKGROUND",(2+len(VENDEDORES_ATIVOS)*2+1,tr),(2+len(VENDEDORES_ATIVOS)*2+1,tr),COR_H) ]
    table_data.append(tot_row)

    main_t=Table(table_data,colWidths=all_colwidths,repeatRows=2)
    main_t.setStyle(TableStyle(table_style))
    story.append(main_t)
    doc.build(story)
    buf.seek(0)
    return buf

# ══════════════════════════════════════════════════════════════════════════
# INTERFACE
# ══════════════════════════════════════════════════════════════════════════
st.subheader("1️⃣ Metas da Semana")
col_ini,col_fim=st.columns(2)
with col_ini:
    semana_ini=st.date_input("Início",value=date.today()-timedelta(days=date.today().weekday()))
with col_fim:
    semana_fim=st.date_input("Fim",value=semana_ini+timedelta(days=5))
periodo=f"{semana_ini.strftime('%d/%m/%Y')} a {semana_fim.strftime('%d/%m/%Y')}"

st.caption("Busque ou digite um produto para adicionar:")
produtos_ja=[p["produto"] for p in st.session_state.produtos_meta]
produtos_disp=[p for p in PRODUTOS_LISTA if p not in produtos_ja]

col_b,col_n,col_e,col_btn=st.columns([2,2,1,1])
with col_b:
    prod_sel=st.selectbox("Buscar",options=[""]+produtos_disp,
        format_func=lambda x:"🔍 Selecione da lista..." if x=="" else x,
        label_visibility="collapsed")
with col_n:
    prod_novo=st.text_input("Novo",placeholder="✏️ Ou digite produto novo...",
        label_visibility="collapsed")
with col_e:
    est_input=st.number_input("Est",min_value=0.0,step=1.0,
        label_visibility="collapsed",placeholder="Estoque CX")
with col_btn:
    if st.button("➕ Adicionar",use_container_width=True,type="primary"):
        prod_final=prod_novo.strip() if prod_novo.strip() else prod_sel
        if prod_final and prod_final!="":
            if prod_final not in produtos_ja:
                st.session_state.produtos_meta.append({"produto":prod_final,"estoque":est_input})
                if prod_final not in PRODUTOS_LISTA:
                    st.session_state.produtos_extra.append(prod_final)
                    _,sha=github_get(GITHUB_FILE_PRODUTOS)
                    github_save(GITHUB_FILE_PRODUTOS,st.session_state.produtos_extra,sha)
                    st.toast(f"✅ '{prod_final}' salvo na lista permanente!")
                st.rerun()
        else:
            st.warning("Selecione ou digite um produto.")

if st.session_state.produtos_meta:
    rows=[]
    for item in st.session_state.produtos_meta:
        est=item["estoque"]
        row={"Produto":item["produto"],"Estoque CX":int(est)}
        for v in VENDEDORES_ATIVOS:
            row[f"{v} ({int(PERCENTUAIS[v]*100)}%)"]=math.ceil(est*PERCENTUAIS[v])
        rows.append(row)

    df_edit=st.data_editor(pd.DataFrame(rows),use_container_width=True,hide_index=True,
        disabled=["Produto"]+[f"{v} ({int(PERCENTUAIS[v]*100)}%)" for v in VENDEDORES_ATIVOS],
        key="editor_metas")
    for i,row in df_edit.iterrows():
        if i<len(st.session_state.produtos_meta):
            st.session_state.produtos_meta[i]["estoque"]=row["Estoque CX"]

    col_s,col_a,col_i=st.columns([1,1,3])
    with col_s:
        if st.button("💾 Salvar Metas",type="primary",use_container_width=True):
            _,sha=github_get(GITHUB_FILE_METAS)
            ok=github_save(GITHUB_FILE_METAS,st.session_state.produtos_meta,sha)
            st.success("✅ Metas salvas!") if ok else st.error("❌ Erro ao salvar.")
    with col_a:
        if st.button("🗑️ Apagar Metas",type="secondary",use_container_width=True):
            _,sha=github_get(GITHUB_FILE_METAS)
            github_save(GITHUB_FILE_METAS,[],sha)
            st.session_state.produtos_meta=[]
            st.session_state.vendido={}
            st.session_state.estoque={}
            st.rerun()
    with col_i:
        st.caption(f"💡 Semana: **{periodo}**")

st.divider()
st.subheader("2️⃣ Upload dos PDFs")
col_p1,col_p2=st.columns(2)
with col_p1:
    st.caption("📊 PDF de Vendas Acumuladas")
    pdf_vendas=st.file_uploader("Vendas",type=["pdf"],label_visibility="collapsed")
with col_p2:
    st.caption("📦 PDF de Estoque Físico")
    pdf_estoque=st.file_uploader("Estoque",type=["pdf"],label_visibility="collapsed")

if pdf_vendas:
    with st.spinner("🔍 Lendo vendas..."):
        with tempfile.NamedTemporaryFile(suffix=".pdf",delete=False) as tmp:
            tmp.write(pdf_vendas.read()); tmp_path=tmp.name
        vendas_raw=extrair_vendas_por_vendedor(tmp_path)
        os.unlink(tmp_path)
    vendido_c={v:{} for v in VENDEDORES_ATIVOS}
    for vend,prods in vendas_raw.items():
        if vend not in vendido_c: continue
        for desc,qtd in prods.items():
            pm=mapear_produto(desc)
            if pm: vendido_c[vend][pm]=vendido_c[vend].get(pm,0)+qtd
    st.session_state.vendido=vendido_c
    found=[v for v in vendas_raw if v in VENDEDORES_ATIVOS]
    st.success(f"✅ Vendas: {', '.join(found)}")

if pdf_estoque:
    with st.spinner("🔍 Lendo estoque..."):
        with tempfile.NamedTemporaryFile(suffix=".pdf",delete=False) as tmp:
            tmp.write(pdf_estoque.read()); tmp_path=tmp.name
        est_raw=extrair_estoque_por_vendedor(tmp_path)
        os.unlink(tmp_path)
    st.session_state.estoque={v:est_raw.get(v,[]) for v in VENDEDORES_ATIVOS}
    st.session_state.sem_vendedor=est_raw.get("SEM_VENDEDOR",[])
    total=sum(len(v) for v in st.session_state.estoque.values())
    st.success(f"✅ Estoque: {total} produto(s) distribuídos.")
    if st.session_state.sem_vendedor:
        with st.expander(f"⚠️ {len(st.session_state.sem_vendedor)} produto(s) SEM VENDEDOR",expanded=True):
            df_sv=pd.DataFrame(st.session_state.sem_vendedor)[["codigo","descricao","data_entrada","saldo_atual"]]
            df_sv.columns=["Código","Descrição","Data Entrada","Saldo Atual"]
            st.dataframe(df_sv,use_container_width=True,hide_index=True)

st.divider()
st.subheader("3️⃣ Gerar Relatórios")

if st.button("📋 Gerar Relatórios Completos",use_container_width=True,type="primary"):
    if not st.session_state.estoque:
        st.warning("Faça upload do PDF de Estoque primeiro.")
    else:
        data_ref=date.today().strftime("%d/%m/%Y")
        vendido=st.session_state.vendido
        metas=st.session_state.produtos_meta

        with st.spinner("⚙️ Gerando PDFs..."):
            buf_dash=gerar_pdf_dashboard(metas,vendido,data_ref,periodo)
            buf_resumo=gerar_pdf_resumo(metas,vendido,data_ref,periodo) if metas else None
            bufs_vend={}
            for vendedor in VENDEDORES_ATIVOS:
                bufs_vend[vendedor]=gerar_pdf_vendedor(
                    vendedor,data_ref,
                    st.session_state.estoque.get(vendedor,[]),
                    metas,vendido.get(vendedor,{}))

        st.success("✅ PDFs gerados!")
        col1,col2=st.columns(2)
        with col1:
            st.download_button("⬇️ Dashboard",data=buf_dash,
                file_name=f"Dashboard_{date.today().strftime('%d%m%Y')}.pdf",
                mime="application/pdf",use_container_width=True)
        with col2:
            if buf_resumo:
                st.download_button("⬇️ Resumo Geral",data=buf_resumo,
                    file_name=f"Resumo_Geral_{date.today().strftime('%d%m%Y')}.pdf",
                    mime="application/pdf",use_container_width=True)

        st.divider()
        st.caption("📄 Relatórios individuais por vendedor:")
        cols=st.columns(4)
        for i,vendedor in enumerate(VENDEDORES_ATIVOS):
            with cols[i%4]:
                st.download_button(
                    f"⬇️ {vendedor}",data=bufs_vend[vendedor],
                    file_name=f"Relatorio_{vendedor}_{date.today().strftime('%d%m%Y')}.pdf",
                    mime="application/pdf",use_container_width=True)