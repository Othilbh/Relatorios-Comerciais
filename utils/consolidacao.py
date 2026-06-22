# utils/consolidacao.py
GRUPOS = {
    "MERCANTIL BASTOS": "Mercantil Bastos",
    "NOVO HORIZONTE": "Novo Horizonte",
    "ENTRE RIOS": "Entre Rios",
    "PAIS E FILHOS": "Pais e Filhos",
    "PAIS & FILHOS": "Pais e Filhos",
    "SACOLAO CASTELO": "Sacolão Castelo",
    "SACOLÃO CASTELO": "Sacolão Castelo",
    "NOVA SUISSA": "Dário",
    "NOVA SUÍSSA": "Dário",
    "SANTA LUCIA": "Dário",
    "SANTA LÚCIA": "Dário",
    "SINFRONIONHO BROCHADO": "Dário",
    "SINFRONIO BROCHADO": "Dário",
    "SINFRÔNIO BROCHADO": "Dário",
    "DARIO": "Dário",
    "DÁRIO": "Dário",
    "ADELSON BOY": "Adelson Boy",
    "ADELSON": "Adelson Boy",
    "TEREZINHA": "Terezinha",
    "GOMES E SANTOS": "Gomes & Santos",
    "GOMES & SANTOS": "Gomes & Santos",
    "GOMES SANTOS": "Gomes & Santos",
    "SACOLAO MAXIMO": "Sacolão Máximo",
    "SACOLÃO MÁXIMO": "Sacolão Máximo",
    "SACOLAO MÁXIMO": "Sacolão Máximo",
    "ANDRE MACHACALIS": "Andre Machacalis",
    "MACHACALIS": "Andre Machacalis",
    "EXPERFRUT": "Experfrut",
    "EXPERFRUT GOURMET": "Experfrut",
    "ABC CAMPEAO": "ABC Campeão",
    "ABC CAMPEÃO": "ABC Campeão",
    "BRIXX": "Brixx",
}

def normalizar_cliente(nome: str) -> str:
    if not nome:
        return nome
    upper = nome.strip().upper()
    if upper in GRUPOS:
        return GRUPOS[upper]
    for chave, canonico in GRUPOS.items():
        if chave in upper:
            return canonico
    return nome.strip()

def consolidar_clientes(lista_clientes: list) -> dict:
    consolidado = {}
    for item in lista_clientes:
        nome = normalizar_cliente(item.get("cliente", ""))
        fat = item.get("faturamento", 0.0)
        custo = item.get("custo", 0.0)
        vol = item.get("volume", 0.0)
        if nome not in consolidado:
            consolidado[nome] = {"faturamento": 0.0, "custo": 0.0, "volume": 0.0}
        consolidado[nome]["faturamento"] += fat
        consolidado[nome]["custo"] += custo
        consolidado[nome]["volume"] += vol
    return consolidado

def calcular_margem(faturamento: float, custo: float) -> float:
    if faturamento == 0:
        return 0.0
    return (faturamento - custo) / faturamento * 100