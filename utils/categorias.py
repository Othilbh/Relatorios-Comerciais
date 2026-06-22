# utils/categorias.py
CATEGORIAS = {
    "MAÇÃ IMPORTADA": [
        "ARGENTINA", "MACA ARGENTINA", "MAÇÃ ARGENTINA",
        "CHILENA", "MACA CHILENA", "MAÇÃ CHILENA",
        "PINK LADY", "MACA PINK", "MAÇÃ PINK",
        "GRAN SMITH", "GRANNY SMITH", "MACA SMITH", "MAÇÃ SMITH",
        "RED GLOBE", "MACA RED", "MAÇÃ RED GLOBE",
        "MACA IMPORTADA", "MAÇÃ IMPORTADA",
    ],
    "CAROÇO": [
        "AMEIXA", "PESSEGO", "PÊSSEGO", "PESSEGO NACIONAL",
        "PÊSSEGO NACIONAL", "NECTARINA",
    ],
    "MAMÃO": [
        "MAMAO HAVAI", "MAMÃO HAVAI", "MAMAO HAVAÍ", "MAMÃO HAVAÍ",
        "MAMAO FORMOSO", "MAMÃO FORMOSO",
        "MAMAO", "MAMÃO",
    ],
    "UVA VERMELHA": [
        "UVA ISIS", "ISIS",
        "UVA CRINSON", "CRINSON", "CRIMSON",
        "UVA NUBIA", "NUBIA", "NUBIANA",
        "UVA JUBILEE", "JUBILEE",
        "UVA VERMELHA",
    ],
    "PERA FORELLE/ERCOLINE": [
        "PERA FORELLE", "FORELLE",
        "PERA ERCOLINE", "ERCOLINE",
        "PERA ASIATICA", "PERA ASIÁTICA", "ASIATICA", "ASIÁTICA",
        "PERA WILLIAMS", "WILLIAMS",
    ],
    "ROMA/MIRTILO": [
        "ROMA", "TOMATE ROMA",
        "MIRTILO", "BLUEBERRY",
    ],
    "MAÇÃ NACIONAL": [
        "MACA GALA", "MAÇÃ GALA", "GALA",
        "MACA FUJI", "MAÇÃ FUJI", "FUJI",
        "MACA NACIONAL", "MAÇÃ NACIONAL",
    ],
    "UVA VERDE": [
        "UVA VERDE", "UVA ITALIA", "ITALIA", "ITÁLIA",
        "UVA THOMPSON", "THOMPSON",
        "UVA RUBI", "RUBI",
    ],
    "MELÃO": [
        "MELAO AMARELO", "MELÃO AMARELO",
        "MELAO GALIA", "MELÃO GALIA", "MELÃO GAÍA",
        "MELAO CANTALOUPE", "CANTALOUPE",
        "MELAO", "MELÃO",
    ],
    "TANGERINA": [
        "TANGERINA", "MEXERICA", "PONKAN",
        "TANGERINA CUMBUCA", "CUMBUCA",
    ],
    "GOIABA": ["GOIABA"],
    "LARANJA": ["LARANJA"],
    "LIMÃO": [
        "LIMAO", "LIMÃO", "LIMAO TAHITI", "LIMÃO TAHITI",
        "LIMAO SICILIANO", "LIMÃO SICILIANO",
    ],
    "MORANGO": ["MORANGO"],
    "ABACAXI": ["ABACAXI"],
    "MANGA": ["MANGA", "MANGA PALMER", "MANGA TOMMY"],
    "UVAS DIVERSAS": ["UVA", "UVAS"],
}

ORDEM_COLUNAS = [
    "MAÇÃ IMPORTADA", "MAÇÃ NACIONAL", "CAROÇO", "MAMÃO",
    "UVA VERMELHA", "UVA VERDE", "UVAS DIVERSAS",
    "PERA FORELLE/ERCOLINE", "ROMA/MIRTILO", "MELÃO",
    "TANGERINA", "GOIABA", "MORANGO", "LARANJA",
    "LIMÃO", "ABACAXI", "MANGA",
]

def categorizar_produto(descricao: str) -> str:
    if not descricao:
        return "OUTROS"
    upper = descricao.strip().upper()
    for categoria, keywords in CATEGORIAS.items():
        for kw in keywords:
            if kw in upper:
                return categoria
    return "OUTROS"

def margem_real(margem_rel: float) -> float:
    return margem_rel + 15.0