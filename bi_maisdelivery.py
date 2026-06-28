#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
============================================================================
  MAIS DELIVERY — BI direto na API  (sem Docker / sem n8n)
============================================================================
  Puxa pedidos de todos os estabelecimentos (nrOperacao 2), guarda em cache
  SQLite, calcula as métricas de negócio e gera:
     • dashboard.html         — painel interativo (tema escuro)
     • bi_maisdelivery.xlsx   — uma aba por métrica
     • cache.sqlite           — persistência local consultável
     • _amostra_pedido.json   — 1 pedido cru, pra mapear os campos reais

  Requisitos:
     pip install requests pandas openpyxl

  Uso:
     python bi_maisdelivery.py                  # fetch real + dashboard
     python bi_maisdelivery.py --mock           # dados sintéticos (preview)
     python bi_maisdelivery.py --no-fetch       # só recalcula do cache
     python bi_maisdelivery.py --desde 2026-04-01 --ate 2026-06-28
     python bi_maisdelivery.py --estabs 82229,86759
     python bi_maisdelivery.py --ciclo 35 --ocupacao 0.90
============================================================================
"""

import argparse
import json
import math
import os
import sqlite3
import sys
import time
from collections import defaultdict
from datetime import datetime, timedelta

try:
    from zoneinfo import ZoneInfo
    TZ = ZoneInfo("America/Sao_Paulo")
except Exception:  # fallback se tzdata não estiver disponível
    TZ = None

# ── dependências externas com mensagem amigável ────────────────────────────
try:
    import requests
except ImportError:
    sys.exit("❌  Falta a lib 'requests'.  Rode:  pip install requests pandas openpyxl")
try:
    import pandas as pd
except ImportError:
    sys.exit("❌  Falta a lib 'pandas'.  Rode:  pip install requests pandas openpyxl")


# ════════════════════════════════════════════════════════════════════════
#  CONFIG
# ════════════════════════════════════════════════════════════════════════

API_URL = "https://maisdeliveryapp.com.br/oapi/v2/controller/operacao_oapi_v2_publica.php"
ID_ORIGEM = 11
LIMITE_API = 500
PAUSA = 0.4
DEFAULT_DIAS = 180          # janela da 1ª rodada (depois o cache cobre o resto)
OVERLAP_DIAS = 3            # re-puxa os últimos N dias p/ capturar updates de status

# Mapa de situações (texto é o que vale; número é fallback) — fonte: convenções do projeto
STATUS_MAP = {
    1: "Verificando", 2: "Na fila de preparo", 3: "Em andamento",
    4: "Pronto pra Entrega", 5: "Saiu para entrega", 6: "Entrega realizada",
    7: "Entregador na sua porta", 9: "Cancelado",
}
ID_CANCELADO = 9
ID_CONCLUIDO = 6

# Estabelecimentos (id = idRestaurante). Lista herdada do projeto.
ESTABELECIMENTOS = [
    {"id": 62875, "nome": "Bar do Edmar"}, {"id": 64781, "nome": "Villa Burguer"},
    {"id": 65351, "nome": "Guinho Pizzaria"}, {"id": 75693, "nome": "Faluá Lanches"},
    {"id": 75757, "nome": "Pedi Aqui - Castelinho"}, {"id": 76021, "nome": "Parrillaria Goumert"},
    {"id": 76295, "nome": "Dom Henrique"}, {"id": 76351, "nome": "Tradição Mineira"},
    {"id": 76535, "nome": "Ice Dream sonho gelado"}, {"id": 77071, "nome": "Capitão Açaí"},
    {"id": 77285, "nome": "Bora de Açaí"}, {"id": 77803, "nome": "Sorveteria Frut Bom"},
    {"id": 77899, "nome": "Divino Sabor"}, {"id": 78024, "nome": "Bar do Edmar - Rest. e Mercearia"},
    {"id": 78152, "nome": "Restaurante - Pedi Aqui - Castelinho"}, {"id": 78693, "nome": "Tá Na Chapa"},
    {"id": 79325, "nome": "Nawiki"}, {"id": 79589, "nome": "DISTRIBUIDORA DO BROOOW"},
    {"id": 79965, "nome": "Angela-Cakes"}, {"id": 80291, "nome": "Aladdin esfiharia e pizzaria"},
    {"id": 80493, "nome": "Empadão do Tatá"}, {"id": 80495, "nome": "Casa de Carnes Chapéu de Sol"},
    {"id": 80499, "nome": "VENDA DO NECA"}, {"id": 80733, "nome": "Point do Açaí"},
    {"id": 81247, "nome": "Padaria Macaúbas"}, {"id": 81253, "nome": "Bom Doce - Bolos & Cia"},
    {"id": 81369, "nome": "Vivi Cones Trufados Atelie"}, {"id": 82045, "nome": "Ling Lig Box Delivery"},
    {"id": 82227, "nome": "Casa de Carnes Dini"}, {"id": 82229, "nome": "Açaí C&M"},
    {"id": 82339, "nome": "Street Company"}, {"id": 82469, "nome": "RODRIGO SPORTS BAR"},
    {"id": 82527, "nome": "Bola de Grude"}, {"id": 82691, "nome": "Uai Digital"},
    {"id": 83001, "nome": "Nair Salgados e Quitandas"}, {"id": 83309, "nome": "Baixada Supermercado"},
    {"id": 83813, "nome": "MilkShow"}, {"id": 84263, "nome": "Elo Comunicação"},
    {"id": 84427, "nome": "Amigos do Espetinho"}, {"id": 84533, "nome": "Açougue Canarinho"},
    {"id": 84865, "nome": "Mercado Bela Vista"}, {"id": 85367, "nome": "Atacado Rural MLV"},
    {"id": 85745, "nome": "Chef Lau"}, {"id": 86211, "nome": "FA Salgados"},
    {"id": 86759, "nome": "Kolesterol"}, {"id": 87193, "nome": "Café no Ponto"},
    {"id": 87269, "nome": "Docês Isa Confeitaria"}, {"id": 87383, "nome": "Inês Salgados"},
    {"id": 88165, "nome": "Canecão - Jacaré"}, {"id": 88351, "nome": "CacauShow- revenda"},
    {"id": 88889, "nome": "Batata Recheada da Grazi"}, {"id": 89175, "nome": "Distribuidora Silva Oliveira"},
    {"id": 89453, "nome": "Casa das Embalagens"}, {"id": 89743, "nome": "DSG Popular 24 Horas"},
    {"id": 90337, "nome": "Amor de Pipoca"}, {"id": 90395, "nome": "Boca Doce"},
    {"id": 91553, "nome": "Muzaervas"}, {"id": 91661, "nome": "Bar São José"},
    {"id": 92727, "nome": "Seo Zé Burguer Delivery"}, {"id": 93097, "nome": "teste api"},
    {"id": 93235, "nome": "Zilda Doces e Afetos"}, {"id": 94279, "nome": "Faz de Conta - Papelaria E Presentes"},
    {"id": 95803, "nome": "Adega 035"}, {"id": 95919, "nome": "Marmitex Souza"},
    {"id": 95953, "nome": "DaNina Salgados"}, {"id": 95955, "nome": "Rei das Bebidas"},
    {"id": 97329, "nome": "Top Burguer"}, {"id": 97343, "nome": "Moriah Atacado e Varejo"},
    {"id": 97651, "nome": "Restaurante Bom D Mais"}, {"id": 98213, "nome": "Maria Baiana"},
    {"id": 98711, "nome": "Pastelaria do Dilsão"}, {"id": 98941, "nome": "Nice Bolos Decorados"},
    {"id": 100703, "nome": "Pizzaria Salena"}, {"id": 101333, "nome": "Bolos da Ana"},
    {"id": 101559, "nome": "Esfirras Montanna"}, {"id": 101577, "nome": "Flor de belém"},
    {"id": 102125, "nome": "Sacola Delivery"}, {"id": 102627, "nome": "Come Come"},
    {"id": 102653, "nome": "Salena Pizzaria"}, {"id": 102655, "nome": "Salena Pizzaria - Muzambinho"},
    {"id": 102733, "nome": "Delícias da Tata"}, {"id": 102941, "nome": "Divina Gula"},
    {"id": 103805, "nome": "Restaurante Villa Cintra"}, {"id": 103933, "nome": "Dsg Farma"},
    {"id": 103937, "nome": "DSG Poupe Brasil"}, {"id": 103939, "nome": "UltraFarma"},
    {"id": 104945, "nome": "Mamitaria e Lanchonete Diniz"}, {"id": 104965, "nome": "Camadas Geladas"},
    {"id": 105617, "nome": "Bliss Gourmet"}, {"id": 106207, "nome": "OBoticário Revendedor"},
    {"id": 106209, "nome": "Natura Revendedor"}, {"id": 106487, "nome": "Remi Modas"},
    {"id": 106551, "nome": "Mais Gostoso Lanchonete"}, {"id": 107867, "nome": "Candido Doceria"},
    {"id": 108489, "nome": "Ponta de Lápis Papelaria"}, {"id": 108797, "nome": "Hot Show"},
    {"id": 108863, "nome": "Dona Palmira Confeitaria"}, {"id": 108887, "nome": "Remí Modas"},
]

# Nomes candidatos por campo lógico (detecção automática do shape real do JSON)
CANDIDATOS = {
    "data":         ["dtPedido", "dtCriacao", "dtCadastro", "data_pedido", "dt_pedido", "dataPedido", "created_at"],
    "valor":        ["vlPedido", "vlTotal", "vlTotalPedido", "vlPedidoTotal", "valor", "vlValor"],
    "status_txt":   ["dsSituacao", "nmSituacao", "situacao", "dsStatus", "status"],
    "status_id":    ["idSituacao", "idStatus"],
    "cliente_nome": ["dsNome", "dsCliente", "nmCliente", "nomeCliente"],
    "cliente_fone": ["nrTelefone", "telefone", "nrCelular", "celular"],
    "ddd":          ["nrDDD", "ddd"],
    "bairro":       ["dsBairro", "bairro"],
    "cidade":       ["dsCidade", "cidade", "dsMunicipio"],
    "pagamento":    ["dsPagamento", "formaPagamento", "dsFormaPagamento"],
    "pix":          ["flgPix"],
    "motivo":       ["dsMotivo", "dsMotivoCancelamento"],
    "produtos":     ["produtos", "itens", "items", "produto"],
}
# dentro de cada produto:
CAND_PROD = {
    "nome": ["dsNome", "dsProduto", "nome", "dsItem"],
    "valor": ["vlItem", "vlProduto", "vlUnitario", "valor", "vlTotal"],
    "qtd":   ["quantidade", "qtItem", "qtde", "qt", "quantidadeItem"],
}


# ════════════════════════════════════════════════════════════════════════
#  CREDENCIAIS (_env tolerante a encoding + prefixo "Basic ")
# ════════════════════════════════════════════════════════════════════════

def _ler_texto(path):
    """Lê arquivo tentando vários encodings (Windows costuma salvar diferente)."""
    for enc in ("utf-8-sig", "utf-8", "utf-16", "latin-1"):
        try:
            with open(path, "r", encoding=enc) as f:
                return f.read()
        except (UnicodeError, UnicodeDecodeError):
            continue
    with open(path, "rb") as f:  # último recurso
        return f.read().decode("latin-1", errors="replace")


def _parse_valor(raw):
    """Extrai o valor: respeita aspas e remove comentário inline (# ...)."""
    raw = raw.strip()
    if raw[:1] in ('"', "'"):
        q = raw[0]
        fim = raw.find(q, 1)
        return raw[1:fim] if fim != -1 else raw[1:]
    return raw.split("#", 1)[0].strip()


def _diag_env(caminho, env):
    """Quando as chaves não são lidas, mostra exatamente o que há no arquivo."""
    print("\n🔧  Diagnóstico do arquivo de credenciais:")
    print(f"    Arquivo lido    : {caminho}")
    try:
        raw = open(caminho, "rb").read()
        print(f"    Tamanho         : {len(raw)} bytes")
        print(f"    Primeiros bytes : {raw[:60]!r}")
        if len(raw) == 0:
            print("    →  O arquivo está VAZIO. Cole as 4 linhas de credenciais nele.")
        elif raw[:2] in (b"\xff\xfe", b"\xfe\xff") or b"\x00" in raw[:40]:
            print("    →  Parece UTF-16 (PowerShell '>' salva assim). Salve como UTF-8.")
    except Exception as ex:
        print(f"    (não consegui ler os bytes: {ex})")
    print(f"    Chaves lidas    : {list(env.keys()) or '(nenhuma)'}")
    print("    Esperado        : ACCESS_TOKEN, BASIC_AUTH, ID_EMPRESA_INTEGRADA, SANDBOX")
    print("\n    Para recriar limpo no PowerShell (UTF-8):")
    print("      @'\nACCESS_TOKEN = \"SEU_TOKEN\"\nBASIC_AUTH = \"Basic SEU_BASIC\"\n"
          "ID_EMPRESA_INTEGRADA = 289\nSANDBOX = False\n'@ | Set-Content -Encoding utf8 _env")


def load_env(base_dir):
    """Procura o arquivo de credenciais na pasta do script e na CWD.
    Aceita _env, .env e as variantes .txt que o Bloco de Notas costuma criar."""
    nomes = ["_env", ".env", "_env.txt", ".env.txt", "env.txt", "env"]
    pastas = list(dict.fromkeys([base_dir, os.getcwd()]))
    candidatos = [os.path.join(d, n) for d in pastas for n in nomes]
    caminho = next((c for c in candidatos if os.path.isfile(c)), None)
    if not caminho:
        locais = "\n     ".join(pastas)
        sys.exit("❌  Arquivo de credenciais não encontrado.\n"
                 "    Procurei por _env / .env (e variantes .txt) em:\n     " + locais +
                 "\n    Dica: no Windows o Bloco de Notas pode ter salvo como '_env.txt'. "
                 "Renomeie para '_env' ou deixe na mesma pasta do script.")
    env = {}
    for linha in _ler_texto(caminho).splitlines():
        linha = linha.strip()
        if not linha or linha.startswith("#") or "=" not in linha:
            continue
        chave, _, valor = linha.partition("=")
        env[chave.strip()] = _parse_valor(valor)
    return env, caminho


def build_headers(env):
    token = env.get("ACCESS_TOKEN", "")
    basic = env.get("BASIC_AUTH", "").strip()
    if not token or not basic:
        sys.exit("❌  ACCESS_TOKEN ou BASIC_AUTH ausentes no _env.")
    if not basic.lower().startswith("basic "):   # normaliza o prefixo
        basic = "Basic " + basic
    return {"Accesstoken": token, "Authorization": basic, "Content-Type": "application/json"}


# ════════════════════════════════════════════════════════════════════════
#  API
# ════════════════════════════════════════════════════════════════════════

def _extrair_lista(dados):
    if isinstance(dados, list):
        return dados
    if isinstance(dados, dict):
        if dados.get("retorno") == "NOK":
            return []
        for k in ("pedidos", "data", "resultado", "retorno_dados", "items", "results", "records"):
            if isinstance(dados.get(k), list):
                return dados[k]
    return []


def api_post(body, headers, sandbox):
    body["idOrigem"] = ID_ORIGEM
    body["sandBox"] = sandbox            # atenção: 'B' maiúsculo
    last_err = None
    for tent in range(3):
        try:
            r = requests.post(API_URL, headers=headers, json=body, timeout=45)
            if r.status_code == 429:
                time.sleep(int(r.headers.get("Retry-After", 3 * (tent + 1))))
                continue
            if r.status_code in (401, 403):
                raise RuntimeError(
                    f"HTTP {r.status_code} — autenticação rejeitada. "
                    f"Confira ACCESS_TOKEN e BASIC_AUTH no _env. Resposta: {r.text[:160]}")
            r.raise_for_status()
            try:
                return r.json()
            except ValueError:
                raise RuntimeError(f"Resposta não-JSON da API (HTTP {r.status_code}): {r.text[:160]}")
        except requests.exceptions.ConnectionError as e:
            last_err = e
            if tent == 2:
                raise
            time.sleep(2 ** tent)
    if last_err:
        raise last_err
    return {}


def fetch_estab(tenant_id, id_empresa, desde, ate, headers, sandbox):
    """Busca pedidos de 1 estabelecimento com chunking semanal + subdivisão no limite."""
    todos, vistos = [], set()
    cursor = desde
    while cursor < ate:
        proximo = min(cursor + timedelta(days=7), ate)
        body = {
            "nrOperacao": 2, "idEmpresaIntegrada": id_empresa, "idRestaurante": tenant_id,
            "idSituacao": [], "idPedido": [],
            "dtInicio": cursor.strftime("%Y-%m-%d %H:%M:%S"),
            "dtFim": proximo.strftime("%Y-%m-%d %H:%M:%S"),
        }
        lote = _extrair_lista(api_post(body, headers, sandbox))
        if len(lote) >= LIMITE_API:   # bateu limite -> divide em dias
            sub = cursor
            lote = []
            while sub < proximo:
                sp = min(sub + timedelta(days=1), proximo)
                body["dtInicio"] = sub.strftime("%Y-%m-%d %H:%M:%S")
                body["dtFim"] = sp.strftime("%Y-%m-%d %H:%M:%S")
                lote.extend(_extrair_lista(api_post(body, headers, sandbox)))
                sub = sp
                time.sleep(0.1)
        for p in lote:
            if not isinstance(p, dict):
                continue
            pid = p.get("idPedido") or p.get("nrPedido")
            if pid in vistos:
                continue
            if pid is not None:
                vistos.add(pid)
            todos.append(p)
        cursor = proximo
        time.sleep(0.1)
    return todos


# ════════════════════════════════════════════════════════════════════════
#  PARSE / DETECÇÃO DE CAMPOS
# ════════════════════════════════════════════════════════════════════════

FORMATOS_DATA = ["%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%S.%f",
                 "%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d-%m-%Y %H:%M:%S"]


def parse_data(raw):
    if not raw:
        return None
    raw = str(raw).strip()
    for fmt in FORMATOS_DATA:
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(raw.replace("Z", "+00:00")).replace(tzinfo=None)
    except Exception:
        return None


def br_to_float(v):
    """'10,00' -> 10.0 | '1.234,56' -> 1234.56 | 10 -> 10.0"""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def detect_fields(amostra):
    """Dada uma amostra de pedidos, escolhe o nome real de cada campo lógico."""
    freq = defaultdict(lambda: defaultdict(int))
    for p in amostra:
        if not isinstance(p, dict):
            continue
        for logico, nomes in CANDIDATOS.items():
            for n in nomes:
                if n in p and p[n] not in (None, ""):
                    freq[logico][n] += 1
    campos = {}
    for logico in CANDIDATOS:
        if freq[logico]:
            campos[logico] = max(freq[logico], key=freq[logico].get)
        else:
            campos[logico] = None
    # campos dentro do produto
    prod_key = campos.get("produtos")
    campos["_prod"] = {}
    if prod_key:
        amostra_itens = []
        for p in amostra:
            arr = p.get(prod_key) if isinstance(p, dict) else None
            if isinstance(arr, list):
                amostra_itens.extend(x for x in arr if isinstance(x, dict))
        pf = defaultdict(lambda: defaultdict(int))
        for it in amostra_itens[:500]:
            for logico, nomes in CAND_PROD.items():
                for n in nomes:
                    if n in it and it[n] not in (None, ""):
                        pf[logico][n] += 1
        for logico in CAND_PROD:
            campos["_prod"][logico] = max(pf[logico], key=pf[logico].get) if pf[logico] else None
    return campos


def normalize_order(p, campos, tenant_id, estab_nome):
    g = lambda k: p.get(campos[k]) if campos.get(k) else None
    dt = parse_data(g("data"))
    sid = g("status_id")
    try:
        sid = int(sid) if sid is not None else None
    except (ValueError, TypeError):
        sid = None
    stxt = g("status_txt") or (STATUS_MAP.get(sid) if sid is not None else None)
    stxt_l = (stxt or "").lower()
    is_cancel = ("cancel" in stxt_l) or (sid == ID_CANCELADO)
    is_concl = (stxt_l == "entrega realizada") or (sid == ID_CONCLUIDO)
    fone = g("cliente_fone")
    fone = "".join(ch for ch in str(fone) if ch.isdigit()) if fone else None
    pix_raw = str(g("pix") or "0").lower()
    return {
        "id_pedido": p.get("idPedido") or p.get("nrPedido"),
        "tenant_id": tenant_id,
        "estab_nome": estab_nome,
        "dt_pedido": dt.strftime("%Y-%m-%d %H:%M:%S") if dt else None,
        "nm_situacao": stxt,
        "id_situacao": sid,
        "is_cancelado": int(is_cancel),
        "is_concluido": int(is_concl),
        "vl_pedido": br_to_float(g("valor")),
        "cliente_nome": (str(g("cliente_nome")).strip() or None) if g("cliente_nome") else None,
        "cliente_fone": fone or None,
        "bairro": (str(g("bairro")).strip() or None) if g("bairro") else None,
        "cidade": (str(g("cidade")).strip() or None) if g("cidade") else None,
        "forma_pgto": (str(g("pagamento")).strip() or None) if g("pagamento") else None,
        "is_pix": int(pix_raw in ("1", "true", "t", "sim")),
        "motivo_cancel": (str(g("motivo")).strip() or None) if g("motivo") else None,
        "raw": json.dumps(p, ensure_ascii=False),
        "fetched_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


# ════════════════════════════════════════════════════════════════════════
#  CACHE SQLITE
# ════════════════════════════════════════════════════════════════════════

DDL = """
CREATE TABLE IF NOT EXISTS pedidos(
  id_pedido INTEGER PRIMARY KEY, tenant_id INTEGER, estab_nome TEXT,
  dt_pedido TEXT, nm_situacao TEXT, id_situacao INTEGER,
  is_cancelado INTEGER, is_concluido INTEGER, vl_pedido REAL,
  cliente_nome TEXT, cliente_fone TEXT, bairro TEXT, cidade TEXT,
  forma_pgto TEXT, is_pix INTEGER, motivo_cancel TEXT, raw TEXT, fetched_at TEXT
);
CREATE INDEX IF NOT EXISTS ix_ped_tenant_dt ON pedidos(tenant_id, dt_pedido);
CREATE INDEX IF NOT EXISTS ix_ped_situacao  ON pedidos(nm_situacao);
CREATE TABLE IF NOT EXISTS watermark(tenant_id INTEGER PRIMARY KEY, ultimo_dt TEXT);
"""
COLS = ["id_pedido", "tenant_id", "estab_nome", "dt_pedido", "nm_situacao", "id_situacao",
        "is_cancelado", "is_concluido", "vl_pedido", "cliente_nome", "cliente_fone",
        "bairro", "cidade", "forma_pgto", "is_pix", "motivo_cancel", "raw", "fetched_at"]


def cache_init(path):
    con = sqlite3.connect(path)
    con.executescript(DDL)
    con.commit()
    return con


def cache_upsert(con, rows):
    rows = [r for r in rows if r.get("id_pedido") is not None]
    if not rows:
        return 0
    ph = ",".join("?" * len(COLS))
    upd = ",".join(f"{c}=excluded.{c}" for c in COLS if c != "id_pedido")
    sql = (f"INSERT INTO pedidos ({','.join(COLS)}) VALUES ({ph}) "
           f"ON CONFLICT(id_pedido) DO UPDATE SET {upd}")
    con.executemany(sql, [[r.get(c) for c in COLS] for r in rows])
    con.commit()
    return len(rows)


def get_watermark(con, tenant_id):
    row = con.execute("SELECT ultimo_dt FROM watermark WHERE tenant_id=?", (tenant_id,)).fetchone()
    return parse_data(row[0]) if row and row[0] else None


def set_watermark(con, tenant_id, dt_str):
    con.execute("INSERT INTO watermark(tenant_id, ultimo_dt) VALUES(?,?) "
                "ON CONFLICT(tenant_id) DO UPDATE SET ultimo_dt=excluded.ultimo_dt",
                (tenant_id, dt_str))
    con.commit()


# ════════════════════════════════════════════════════════════════════════
#  MÉTRICAS (pandas)
# ════════════════════════════════════════════════════════════════════════

def agora_local():
    if TZ:
        return datetime.now(TZ).replace(tzinfo=None)
    return datetime.now()


def _prev_month(d):
    return d.replace(year=d.year - 1, month=12, day=1) if d.month == 1 else d.replace(month=d.month - 1, day=1)


def kpis(df):
    now = agora_local()
    d0 = now.replace(hour=0, minute=0, second=0, microsecond=0)
    w0 = d0 - timedelta(days=now.isoweekday() - 1)
    m0 = d0.replace(day=1)
    y0 = d0.replace(month=1, day=1)
    val = df[~df["is_cancelado"].astype(bool)]
    val_liq = df[df["is_concluido"].astype(bool)]

    def jan(ini, fim):
        s = val[(val["dt"] >= ini) & (val["dt"] < fim)]
        rec = s["vl_pedido"].sum(min_count=1)
        return len(s), (None if pd.isna(rec) else round(float(rec), 2))

    def jan_liq(ini, fim):
        s = val_liq[(val_liq["dt"] >= ini) & (val_liq["dt"] < fim)]
        rec = s["vl_pedido"].sum(min_count=1)
        return None if pd.isna(rec) else round(float(rec), 2)

    def pct(cur, prev):
        return None if not prev else round(100.0 * (cur - prev) / prev, 1)

    out = {}
    for nome, b0, prev0 in [("hoje", d0, d0 - timedelta(days=1)),
                            ("semana", w0, w0 - timedelta(days=7)),
                            ("mes", m0, _prev_month(m0)),
                            ("ano", y0, y0.replace(year=y0.year - 1))]:
        ped, rec = jan(b0, now)
        ped_ant, _ = jan(prev0, b0)
        out[f"pedidos_{nome}"] = ped
        out[f"receita_{nome}"] = rec
        out[f"ticket_{nome}"] = (None if not ped or rec is None else round(rec / ped, 2))
        out[f"cresc_{nome}"] = pct(ped, ped_ant)
    out["receita_total"] = (None if val["vl_pedido"].dropna().empty else round(float(val["vl_pedido"].sum()), 2))
    out["pedidos_total"] = int(len(val))
    out["receita_liquida_mes"] = jan_liq(m0, now)
    out["receita_liquida_total"] = (None if val_liq["vl_pedido"].dropna().empty else round(float(val_liq["vl_pedido"].sum()), 2))
    return out


def metrics_for(df, ciclo_min, ocupacao, com_ranking=False):
    val = df[~df["is_cancelado"].astype(bool)].copy()
    por_hora = [int((val["hora"] == h).sum()) for h in range(24)]
    por_dow = [int((val["dow"] == d).sum()) for d in range(1, 8)]

    mes = (val.groupby("ano_mes")
              .agg(pedidos=("id_pedido", "size"), receita=("vl_pedido", "sum"))
              .reset_index().sort_values("ano_mes"))
    por_mes = [{"mes": r.ano_mes, "pedidos": int(r.pedidos),
                "receita": (None if pd.isna(r.receita) else round(float(r.receita), 2))}
               for r in mes.itertuples()]

    heatmap = [[int(((val["dow"] == d) & (val["hora"] == h)).sum()) for h in range(24)] for d in range(1, 8)]

    if any(por_hora):
        hp = max(range(24), key=lambda h: por_hora[h])
        nz = [h for h in range(24) if por_hora[h] > 0]
        hf = min(nz, key=lambda h: por_hora[h])
        pico = {"hora_pico": hp, "qtd_pico": por_hora[hp], "hora_fraca": hf, "qtd_fraca": por_hora[hf]}
    else:
        pico = {"hora_pico": None, "qtd_pico": 0, "hora_fraca": None, "qtd_fraca": 0}

    fds = val[val["fim_semana"]]
    uteis = val[~val["fim_semana"]]
    util_fds = {
        "util": {"pedidos": int(len(uteis)), "media_dia": round(len(uteis) / max(uteis["dia"].nunique(), 1), 1)},
        "fds": {"pedidos": int(len(fds)), "media_dia": round(len(fds) / max(fds["dia"].nunique(), 1), 1)},
    }

    # dimensionamento (modelo de fila)
    dim = []
    if not val.empty:
        g = val.groupby(["dia", "hora"]).size().reset_index(name="qt")
        media_h = g.groupby("hora")["qt"].mean()
        for h in range(24):
            if h in media_h.index:
                m = float(media_h[h])
                moto = max(0, math.ceil(m * (ciclo_min / 60.0) / max(ocupacao, 0.01)))
                dim.append({"hora": h, "media": round(m, 2), "motoboys": moto})

    # cancelamentos por hora
    cancel_hora = []
    for h in range(24):
        tot = int((df["hora"] == h).sum())
        canc = int(((df["hora"] == h) & (df["is_cancelado"].astype(bool))).sum())
        if tot:
            cancel_hora.append({"hora": h, "cancel": canc, "total": tot,
                                "taxa": round(100.0 * canc / tot, 1)})
    taxa_cancel = round(100.0 * int(df["is_cancelado"].sum()) / max(len(df), 1), 1)

    res = {
        "kpis": kpis(df), "por_hora": por_hora, "por_dow": por_dow, "por_mes": por_mes,
        "heatmap": heatmap, "pico": pico, "util_fds": util_fds,
        "dimensionamento": dim, "cancel_hora": cancel_hora, "taxa_cancel": taxa_cancel,
        "top_produtos": top_produtos(df), "clientes": clientes_metrics(val),
    }
    if com_ranking:
        res["ranking"] = ranking_estabs(df)
    return res


def ranking_estabs(df):
    out = []
    for tid, g in df.groupby("tenant_id"):
        v = g[~g["is_cancelado"].astype(bool)]
        rec = v["vl_pedido"].sum(min_count=1)
        rec = None if pd.isna(rec) else round(float(rec), 2)
        out.append({
            "id": int(tid), "nome": g["estab_nome"].iloc[0],
            "pedidos": int(len(v)),
            "receita": rec,
            "ticket": (None if not len(v) or rec is None else round(rec / len(v), 2)),
            "cancelados": int(g["is_cancelado"].sum()),
            "taxa_cancel": round(100.0 * int(g["is_cancelado"].sum()) / max(len(g), 1), 1),
        })
    out.sort(key=lambda x: x["pedidos"], reverse=True)
    return out


def top_produtos(df, n=15):
    """Só funciona se os campos de produto tiverem sido detectados (ver _PRODFIELDS)."""
    pk = _PRODFIELDS.get("produtos")
    nk, vk, qk = _PRODFIELDS.get("nome"), _PRODFIELDS.get("valor"), _PRODFIELDS.get("qtd")
    if not pk or not nk:
        return None
    acc = defaultdict(lambda: {"qtd": 0.0, "receita": 0.0})
    for raw in df[~df["is_cancelado"].astype(bool)]["raw"]:
        try:
            obj = json.loads(raw)
        except Exception:
            continue
        for it in (obj.get(pk) or []):
            if not isinstance(it, dict):
                continue
            nome = str(it.get(nk, "")).strip() or "(sem nome)"
            q = br_to_float(it.get(qk)) or 1 if qk else 1
            v = br_to_float(it.get(vk)) or 0 if vk else 0
            acc[nome]["qtd"] += q
            acc[nome]["receita"] += v * q
    if not acc:
        return None
    top = sorted(acc.items(), key=lambda kv: kv[1]["qtd"], reverse=True)[:n]
    return [{"nome": k, "qtd": round(v["qtd"], 1), "receita": round(v["receita"], 2)} for k, v in top]


def clientes_metrics(val, n=15):
    if "cliente_fone" not in val or val["cliente_fone"].dropna().empty:
        return None
    c = val.dropna(subset=["cliente_fone"])
    if c.empty:
        return None
    por = c.groupby("cliente_fone").agg(
        pedidos=("id_pedido", "size"),
        receita=("vl_pedido", "sum"),
        nome=("cliente_nome", "first")).reset_index()
    recorrentes = int((por["pedidos"] > 1).sum())
    top = por.sort_values("pedidos", ascending=False).head(n)
    return {
        "total": int(len(por)),
        "recorrentes": recorrentes,
        "taxa_recorrencia": round(100.0 * recorrentes / max(len(por), 1), 1),
        "ticket_medio": (None if por["receita"].dropna().empty
                         else round(float(c["vl_pedido"].sum(min_count=1) or 0) / max(len(c), 1), 2)),
        "top": [{"nome": (r.nome or r.cliente_fone), "pedidos": int(r.pedidos),
                 "receita": (None if pd.isna(r.receita) else round(float(r.receita), 2))}
                for r in top.itertuples()],
    }


_PRODFIELDS = {}  # preenchido em build_payload a partir da detecção


def df_from_cache(con):
    df = pd.read_sql_query("SELECT * FROM pedidos WHERE dt_pedido IS NOT NULL", con)
    if df.empty:
        return df
    df["dt"] = pd.to_datetime(df["dt_pedido"], errors="coerce")
    df = df.dropna(subset=["dt"])
    df["hora"] = df["dt"].dt.hour
    df["dow"] = df["dt"].dt.dayofweek + 1   # 1=seg .. 7=dom (equivale a isoweekday)
    df["dia"] = df["dt"].dt.date.astype(str)
    df["ano_mes"] = df["dt"].dt.strftime("%Y-%m")
    df["fim_semana"] = df["dow"] >= 6
    return df


def build_payload(con, campos, desde, ate, ciclo_min, ocupacao):
    global _PRODFIELDS
    _PRODFIELDS = {"produtos": campos.get("produtos"), **campos.get("_prod", {})}
    df = df_from_cache(con)
    estabs = (df.groupby(["tenant_id", "estab_nome"]).size()
                .reset_index(name="pedidos").sort_values("pedidos", ascending=False))
    lista = [{"id": int(r.tenant_id), "nome": r.estab_nome, "pedidos": int(r.pedidos)}
             for r in estabs.itertuples()]

    dados = {"todos": metrics_for(df, ciclo_min, ocupacao, com_ranking=True)}
    for e in lista:
        sub = df[df["tenant_id"] == e["id"]]
        dados[str(e["id"])] = metrics_for(sub, ciclo_min, ocupacao, com_ranking=False)

    detectados = {k: v for k, v in campos.items() if k not in ("_prod",)}
    return {
        "meta": {
            "gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "desde": desde.strftime("%d/%m/%Y"), "ate": ate.strftime("%d/%m/%Y"),
            "total_pedidos": int(len(df)), "total_estabs": len(lista),
            "campos_detectados": detectados, "prod_fields": campos.get("_prod", {}),
            "ciclo_min": ciclo_min, "ocupacao": ocupacao,
            "tem_valor": campos.get("valor") is not None,
            "tem_produtos": dados["todos"]["top_produtos"] is not None,
            "tem_clientes": dados["todos"]["clientes"] is not None,
        },
        "estabelecimentos": lista,
        "dados": dados,
    }


# ════════════════════════════════════════════════════════════════════════
#  EXCEL
# ════════════════════════════════════════════════════════════════════════

def write_excel(con, payload, path):
    df = df_from_cache(con)
    t = payload["dados"]["todos"]
    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        pd.DataFrame([payload["dados"]["todos"]["kpis"]]).T.rename(columns={0: "valor"}).to_excel(xl, sheet_name="Métricas Gerais")
        pd.DataFrame(payload["estabelecimentos"]).to_excel(xl, sheet_name="Estabelecimentos", index=False)
        pd.DataFrame(t["ranking"]).to_excel(xl, sheet_name="Ranking", index=False)
        pd.DataFrame({"hora": range(24), "pedidos": t["por_hora"]}).to_excel(xl, sheet_name="Por Hora", index=False)
        pd.DataFrame({"dia_semana": ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"],
                      "pedidos": t["por_dow"]}).to_excel(xl, sheet_name="Por Dia Semana", index=False)
        pd.DataFrame(t["por_mes"]).to_excel(xl, sheet_name="Por Mês", index=False)
        pd.DataFrame(t["dimensionamento"]).to_excel(xl, sheet_name="Dimensionamento", index=False)
        pd.DataFrame(t["cancel_hora"]).to_excel(xl, sheet_name="Cancelamentos", index=False)
        if t["top_produtos"]:
            pd.DataFrame(t["top_produtos"]).to_excel(xl, sheet_name="Top Produtos", index=False)
        if t["clientes"]:
            pd.DataFrame(t["clientes"]["top"]).to_excel(xl, sheet_name="Top Clientes", index=False)
        # base bruta (limitada p/ não estourar)
        cols = [c for c in df.columns if c not in ("raw",)]
        df[cols].head(50000).to_excel(xl, sheet_name="Pedidos (base)", index=False)


# ════════════════════════════════════════════════════════════════════════
#  DASHBOARD HTML
# ════════════════════════════════════════════════════════════════════════

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="pt-BR"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Mais Delivery — BI</title>
<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
<style>
:root{--bg:#0d0e14;--surface:#13141c;--surface2:#1a1b26;--surface3:#21222f;--border:rgba(255,255,255,.07);
--border2:rgba(255,255,255,.13);--accent:#ff5c1a;--accent2:#ff7a3d;--text:#e2e4f0;--muted:#6b7099;
--green:#22c97a;--red:#e74c3c;--amber:#f0a500;--blue:#4a90d9;--purple:#8b5cf6;
--sans:'Outfit',sans-serif;--mono:'JetBrains Mono',monospace;}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:var(--sans);padding:22px;font-size:14px}
h1{font-size:20px;font-weight:800;letter-spacing:-.5px}
.muted{color:var(--muted)}.mono{font-family:var(--mono)}
.topbar{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-bottom:6px}
.logo{width:30px;height:30px;border-radius:8px;background:var(--accent);display:grid;place-items:center;font-weight:800;color:#fff}
select{background:var(--surface3);color:var(--text);border:1px solid var(--border2);border-radius:8px;
padding:8px 12px;font-family:var(--sans);font-size:13px;font-weight:600;cursor:pointer;margin-left:auto}
.sub{font-size:12px;color:var(--muted);font-family:var(--mono);margin-bottom:18px}
.chips{display:flex;gap:6px;flex-wrap:wrap;margin:10px 0 18px}
.chip{font-size:11px;font-family:var(--mono);padding:3px 9px;border-radius:20px;border:1px solid var(--border)}
.chip.ok{background:rgba(34,201,122,.1);color:var(--green);border-color:rgba(34,201,122,.25)}
.chip.no{background:rgba(231,76,60,.1);color:var(--red);border-color:rgba(231,76,60,.25)}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(155px,1fr));gap:12px;margin-bottom:18px}
.card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:15px 17px}
.card .lbl{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;font-weight:600}
.card .val{font-size:26px;font-weight:800;margin-top:6px;letter-spacing:-1px}
.card .delta{font-size:12px;font-family:var(--mono);margin-top:3px}
.up{color:var(--green)}.down{color:var(--red)}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(340px,1fr));gap:14px}
.panel{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:16px}
.panel h3{font-size:13px;font-weight:700;margin-bottom:12px;display:flex;align-items:center;gap:8px}
.panel h3 .dot{width:7px;height:7px;border-radius:50%;background:var(--accent)}
.full{grid-column:1/-1}
canvas{max-height:260px}
table{width:100%;border-collapse:collapse;font-size:12.5px}
th,td{text-align:left;padding:7px 9px;border-bottom:1px solid var(--border)}
th{color:var(--muted);font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.4px}
td.num,th.num{text-align:right;font-family:var(--mono)}
tbody tr:hover{background:var(--surface2)}
.hm{border-collapse:separate;border-spacing:2px;font-size:10px;font-family:var(--mono)}
.hm td{padding:0;width:15px;height:15px;border:none;border-radius:3px;text-align:center;color:rgba(255,255,255,.55)}
.hm th{padding:2px 4px;font-size:9px;color:var(--muted);border:none}
.note{font-size:11.5px;color:var(--muted);background:var(--surface2);border:1px solid var(--border);
border-radius:10px;padding:10px 12px;margin-top:14px;line-height:1.5}
footer{margin-top:24px;font-size:11px;color:var(--muted);font-family:var(--mono);line-height:1.6}
</style></head><body>
<div class="topbar">
  <div class="logo">M</div>
  <div><h1>Mais Delivery · BI</h1></div>
  <select id="sel"></select>
</div>
<div class="sub" id="sub"></div>
<div class="chips" id="chips"></div>
<div class="cards" id="cards"></div>
<div class="grid">
  <div class="panel"><h3><span class="dot"></span>Pedidos por mês</h3><canvas id="cMes"></canvas></div>
  <div class="panel"><h3><span class="dot"></span>Pedidos por hora</h3><canvas id="cHora"></canvas></div>
  <div class="panel"><h3><span class="dot"></span>Pedidos por dia da semana</h3><canvas id="cDow"></canvas></div>
  <div class="panel"><h3><span class="dot"></span>Motoboys necessários / hora <span class="muted" style="font-weight:400;font-size:11px" id="dimcfg"></span></h3><canvas id="cDim"></canvas></div>
  <div class="panel full"><h3><span class="dot"></span>Heatmap — hora × dia da semana</h3><div style="overflow-x:auto"><table class="hm" id="hm"></table></div></div>
  <div class="panel" id="pRank"><h3><span class="dot"></span>Ranking de estabelecimentos</h3><div style="max-height:340px;overflow:auto"><table id="tRank"></table></div></div>
  <div class="panel" id="pProd"><h3><span class="dot"></span>Top produtos</h3><div id="prodWrap"></div></div>
  <div class="panel" id="pCli"><h3><span class="dot"></span>Clientes</h3><div id="cliWrap"></div></div>
</div>
<div class="note" id="dimnote"></div>
<footer id="ft"></footer>
<script>
const PAYLOAD = __DATA__;
const BRL = v => v==null ? "—" : v.toLocaleString("pt-BR",{style:"currency",currency:"BRL"});
const NUM = v => v==null ? "—" : v.toLocaleString("pt-BR");
const DOW = ["Seg","Ter","Qua","Qui","Sex","Sáb","Dom"];
const C = {grid:"rgba(255,255,255,.06)",tick:"#6b7099",accent:"#ff5c1a",accent2:"#ff7a3d",blue:"#4a90d9",green:"#22c97a"};
Chart.defaults.color=C.tick; Chart.defaults.font.family="Outfit"; Chart.defaults.borderColor=C.grid;
let charts={};
function mkChart(id,cfg){ if(charts[id])charts[id].destroy(); charts[id]=new Chart(document.getElementById(id),cfg); }
const baseScales={x:{grid:{color:C.grid}},y:{grid:{color:C.grid},beginAtZero:true}};

function delta(v){ if(v==null)return ""; const c=v>=0?"up":"down"; const s=v>=0?"▲":"▼"; return `<div class="delta ${c}">${s} ${Math.abs(v)}%</div>`; }

function render(id){
  const d=PAYLOAD.dados[id], k=d.kpis, m=PAYLOAD.meta;
  // cards
  document.getElementById("cards").innerHTML = [
    ["Pedidos hoje",NUM(k.pedidos_hoje),delta(k.cresc_hoje)],
    ["Pedidos semana",NUM(k.pedidos_semana),delta(k.cresc_semana)],
    ["Pedidos mês",NUM(k.pedidos_mes),delta(k.cresc_mes)],
    ["Pedidos ano",NUM(k.pedidos_ano),delta(k.cresc_ano)],
    ["Valor Bruto mês",BRL(k.receita_mes),""],
    ["Valor Líquido mês",BRL(k.receita_liquida_mes),""],
    ["Ticket médio mês",BRL(k.ticket_mes),""],
  ].map(([l,v,dd])=>`<div class="card"><div class="lbl">${l}</div><div class="val">${v}</div>${dd}</div>`).join("");

  mkChart("cMes",{type:"line",data:{labels:d.por_mes.map(x=>x.mes),
    datasets:[{label:"Pedidos",data:d.por_mes.map(x=>x.pedidos),borderColor:C.accent,backgroundColor:"rgba(255,92,26,.12)",fill:true,tension:.3,pointRadius:3}]},
    options:{plugins:{legend:{display:false}},scales:baseScales}});
  mkChart("cHora",{type:"bar",data:{labels:[...Array(24).keys()].map(h=>h+"h"),
    datasets:[{data:d.por_hora,backgroundColor:C.accent,borderRadius:4}]},
    options:{plugins:{legend:{display:false}},scales:baseScales}});
  mkChart("cDow",{type:"bar",data:{labels:DOW,
    datasets:[{data:d.por_dow,backgroundColor:C.blue,borderRadius:4}]},
    options:{plugins:{legend:{display:false}},scales:baseScales}});
  mkChart("cDim",{type:"bar",data:{labels:d.dimensionamento.map(x=>x.hora+"h"),
    datasets:[{data:d.dimensionamento.map(x=>x.motoboys),backgroundColor:C.green,borderRadius:4}]},
    options:{plugins:{legend:{display:false}},scales:baseScales}});
  document.getElementById("dimcfg").textContent=`(ciclo ${m.ciclo_min}min · ocup. ${Math.round(m.ocupacao*100)}%)`;

  // heatmap
  const flat=d.heatmap.flat(), mx=Math.max(1,...flat);
  let h="<tr><th></th>"+[...Array(24).keys()].map(x=>`<th>${x}</th>`).join("")+"</tr>";
  d.heatmap.forEach((row,i)=>{ h+=`<tr><th>${DOW[i]}</th>`+row.map(v=>{
    const a=v/mx; const bg=`rgba(255,92,26,${(.08+a*.92).toFixed(2)})`;
    return `<td style="background:${v?bg:'rgba(255,255,255,.03)'}" title="${DOW[i]} ${''}h: ${v}">${v||""}</td>`;
  }).join("")+"</tr>"; });
  document.getElementById("hm").innerHTML=h;

  // ranking (só em "todos")
  const pRank=document.getElementById("pRank");
  if(d.ranking){ pRank.style.display="";
    document.getElementById("tRank").innerHTML="<thead><tr><th>#</th><th>Estabelecimento</th><th class='num'>Pedidos</th><th class='num'>Receita</th><th class='num'>Ticket</th><th class='num'>Cancel.</th></tr></thead><tbody>"+
      d.ranking.map((r,i)=>`<tr><td class="mono">${i+1}</td><td>${r.nome}</td><td class="num">${NUM(r.pedidos)}</td><td class="num">${BRL(r.receita)}</td><td class="num">${BRL(r.ticket)}</td><td class="num">${r.taxa_cancel}%</td></tr>`).join("")+"</tbody>";
  } else pRank.style.display="none";

  // produtos
  const pProd=document.getElementById("pProd");
  if(d.top_produtos){ pProd.style.display="";
    document.getElementById("prodWrap").innerHTML="<div style='max-height:340px;overflow:auto'><table><thead><tr><th>Produto</th><th class='num'>Qtd</th><th class='num'>Receita</th></tr></thead><tbody>"+
      d.top_produtos.map(p=>`<tr><td>${p.nome}</td><td class="num">${NUM(p.qtd)}</td><td class="num">${BRL(p.receita)}</td></tr>`).join("")+"</tbody></table></div>";
  } else { pProd.style.display=""; document.getElementById("prodWrap").innerHTML="<div class='note' style='margin:0'>Campo de produtos não encontrado no JSON da API. Vamos mapear pelo <span class='mono'>_amostra_pedido.json</span>.</div>"; }

  // clientes
  const pCli=document.getElementById("pCli");
  if(d.clientes){ pCli.style.display="";
    const c=d.clientes;
    document.getElementById("cliWrap").innerHTML=
      `<div class="mono" style="font-size:12px;color:var(--muted);margin-bottom:10px">${NUM(c.total)} clientes · ${NUM(c.recorrentes)} recorrentes (${c.taxa_recorrencia}%) · ticket ${BRL(c.ticket_medio)}</div>`+
      "<div style='max-height:300px;overflow:auto'><table><thead><tr><th>Cliente</th><th class='num'>Pedidos</th><th class='num'>Receita</th></tr></thead><tbody>"+
      c.top.map(x=>`<tr><td>${x.nome}</td><td class="num">${NUM(x.pedidos)}</td><td class="num">${BRL(x.receita)}</td></tr>`).join("")+"</tbody></table></div>";
  } else { pCli.style.display=""; document.getElementById("cliWrap").innerHTML="<div class='note' style='margin:0'>Dados de cliente (telefone) não encontrados no JSON. Mapear depois.</div>"; }
}

// init
(function(){
  const m=PAYLOAD.meta;
  document.getElementById("sub").textContent=`${m.desde} → ${m.ate}  ·  ${NUM(m.total_pedidos)} pedidos  ·  ${m.total_estabs} estabelecimentos  ·  gerado ${m.gerado_em}`;
  const chips=[["valor (receita)",m.tem_valor],["produtos",m.tem_produtos],["clientes",m.tem_clientes]];
  document.getElementById("chips").innerHTML=chips.map(([l,ok])=>`<span class="chip ${ok?'ok':'no'}">${ok?'✓':'✗'} ${l}</span>`).join("");
  const sel=document.getElementById("sel");
  sel.innerHTML=`<option value="todos">🏙️ Todos os estabelecimentos</option>`+
    PAYLOAD.estabelecimentos.map(e=>`<option value="${e.id}">${e.nome} (${e.pedidos})</option>`).join("");
  sel.onchange=()=>render(sel.value);
  document.getElementById("dimnote").innerHTML="⚠️ <b>Dimensionamento é um modelo</b>, não dado medido: a API legada não traz o ciclo real de entrega por motoboy. Estima-se a partir do volume/hora × ciclo médio ÷ ocupação alvo. Para números reais, é preciso capturar aceite/retirada/entrega por entregador.";
  document.getElementById("ft").innerHTML=`Mais Delivery BI · cache local SQLite · ${NUM(m.total_pedidos)} pedidos · campos detectados: ${Object.entries(m.campos_detectados).filter(([,v])=>v).map(([k,v])=>k+'='+v).join(', ')}`;
  render("todos");
})();
</script></body></html>"""


def render_html(payload, path):
    html = HTML_TEMPLATE.replace("__DATA__", json.dumps(payload, ensure_ascii=False))
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)


# ════════════════════════════════════════════════════════════════════════
#  MOCK (preview sem API)
# ════════════════════════════════════════════════════════════════════════

def mock_orders():
    import random
    random.seed(7)
    prods = ["X-Burguer", "X-Salada", "Açaí 500ml", "Pizza Calabresa", "Coca 2L",
             "Marmita P", "Esfiha Carne", "Pastel Queijo", "Milkshake", "Batata Frita"]
    bairros = ["Centro", "Jardim", "Vila Nova", "São José", "Industrial"]
    pgtos = ["PIX", "Cartão Crédito", "Dinheiro", "Cartão Débito"]
    base = datetime.now() - timedelta(days=120)
    out = []
    pid = 60000000
    estabs = ESTABELECIMENTOS[:10]
    for e in estabs:
        n = random.randint(120, 900)
        for _ in range(n):
            dia = base + timedelta(days=random.randint(0, 119))
            # picos de almoço/janta
            hora = random.choices(range(24),
                weights=[1,1,1,1,1,1,2,3,4,5,7,12,16,11,6,4,4,6,11,17,15,9,4,2])[0]
            dt = dia.replace(hour=hora, minute=random.randint(0, 59), second=0)
            cancel = random.random() < 0.06
            itens = []
            for _ in range(random.randint(1, 4)):
                p = random.choice(prods)
                itens.append({"dsNome": p, "vlItem": f"{random.uniform(8,45):.2f}".replace(".", ","),
                              "quantidade": random.randint(1, 3)})
            total = sum(br_to_float(i["vlItem"]) * i["quantidade"] for i in itens)
            pid += 1
            out.append({
                "idPedido": pid, "dtPedido": dt.strftime("%Y-%m-%d %H:%M:%S"),
                "idSituacao": 9 if cancel else random.choice([6, 6, 6, 5, 3]),
                "vlPedido": f"{total:.2f}".replace(".", ","),
                "dsNome": f"Cliente {random.randint(1, 220)}",
                "nrTelefone": f"3599{random.randint(1000000, 9999999)}",
                "dsBairro": random.choice(bairros), "dsCidade": "Muzambinho",
                "dsPagamento": random.choice(pgtos), "flgPix": "1" if random.random() < .4 else "0",
                "produtos": itens, "_estab": e,
            })
    return estabs, out


# ════════════════════════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════════════════════════

def main():
    ap = argparse.ArgumentParser(description="Mais Delivery — BI direto na API")
    ap.add_argument("--desde", help="YYYY-MM-DD (padrão: hoje - %d dias)" % DEFAULT_DIAS)
    ap.add_argument("--ate", help="YYYY-MM-DD (padrão: hoje)")
    ap.add_argument("--estabs", help="filtra IDs separados por vírgula")
    ap.add_argument("--out", default=".", help="pasta de saída (padrão: atual)")
    ap.add_argument("--db", default=None, help="caminho do cache sqlite")
    ap.add_argument("--ciclo", type=float, default=40, help="ciclo médio de entrega (min)")
    ap.add_argument("--ocupacao", type=float, default=0.85, help="taxa de ocupação alvo (0..1)")
    ap.add_argument("--mock", action="store_true", help="usa dados sintéticos (sem API)")
    ap.add_argument("--no-fetch", action="store_true", help="só recalcula do cache existente")
    args = ap.parse_args()

    base_dir = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.abspath(args.out)
    os.makedirs(out_dir, exist_ok=True)
    db_path = args.db or os.path.join(out_dir, "cache.sqlite")
    con = cache_init(db_path)

    ate = datetime.strptime(args.ate, "%Y-%m-%d") if args.ate else datetime.now()
    desde = (datetime.strptime(args.desde, "%Y-%m-%d") if args.desde
             else ate - timedelta(days=DEFAULT_DIAS))

    print(f"\n🛵  MAIS DELIVERY — BI")
    print("=" * 60)
    print(f"📅  Janela : {desde:%d/%m/%Y} → {ate:%d/%m/%Y}")
    print(f"📂  Saída  : {out_dir}")
    print(f"⚙️   Modelo : ciclo {args.ciclo:.0f}min · ocupação {args.ocupacao:.0%}\n")

    campos = None

    if args.mock:
        print("🧪  MODO MOCK — gerando dados sintéticos...")
        estabs, pedidos = mock_orders()
        campos = detect_fields(pedidos[:300])
        rows = [normalize_order(p, campos, p["_estab"]["id"], p["_estab"]["nome"]) for p in pedidos]
        cache_upsert(con, rows)
        print(f"   {len(rows)} pedidos sintéticos em {len(estabs)} estabelecimentos.")
    elif not args.no_fetch:
        env, cam = load_env(base_dir)
        if not env.get("ACCESS_TOKEN") or not env.get("BASIC_AUTH"):
            _diag_env(cam, env)
            sys.exit("\n❌  Não consegui ler ACCESS_TOKEN/BASIC_AUTH do arquivo acima. "
                     "Corrija o arquivo (veja o diagnóstico) e rode de novo.")
        headers = build_headers(env)
        id_empresa = int(env.get("ID_EMPRESA_INTEGRADA", "0") or "0")
        sandbox = str(env.get("SANDBOX", "false")).lower() == "true"
        print(f"🔑  _env: {cam}  ·  empresa {id_empresa}  ·  sandbox={sandbox}\n")

        alvos = ESTABELECIMENTOS
        if args.estabs:
            ids = {int(x) for x in args.estabs.split(",")}
            alvos = [e for e in ESTABELECIMENTOS if e["id"] in ids]

        # amostra p/ detectar campos (primeiro estab com pedidos)
        amostra = []
        try:
            for e in alvos:
                amostra = fetch_estab(e["id"], id_empresa, ate - timedelta(days=30), ate, headers, sandbox)
                if amostra:
                    break
        except Exception as ex:
            sys.exit(f"❌  Falha ao consultar a API (amostra): {ex}")
        if not amostra:
            print("⚠️  Nenhum pedido retornado na amostra (últimos 30 dias). "
                  "Pode ser janela sem pedidos — seguindo mesmo assim.")
        campos = detect_fields(amostra) if amostra else detect_fields([])
        # salva 1 pedido cru p/ inspeção
        if amostra:
            with open(os.path.join(out_dir, "_amostra_pedido.json"), "w", encoding="utf-8") as f:
                json.dump(amostra[0], f, ensure_ascii=False, indent=2)

        total = 0
        for i, e in enumerate(alvos, 1):
            wm = get_watermark(con, e["id"])
            ini = max(desde, wm - timedelta(days=OVERLAP_DIAS)) if wm else desde
            print(f"  [{i:02d}/{len(alvos)}] {e['nome'][:34]:34}", end=" ", flush=True)
            try:
                ped = fetch_estab(e["id"], id_empresa, ini, ate, headers, sandbox)
                rows = [normalize_order(p, campos, e["id"], e["nome"]) for p in ped]
                n = cache_upsert(con, rows)
                datas = [r["dt_pedido"] for r in rows if r["dt_pedido"]]
                if datas:
                    set_watermark(con, e["id"], max(datas))
                total += n
                print(f"✅ {n} pedidos" if n else "— sem novos")
            except Exception as ex:
                print(f"❌ {ex}")
            time.sleep(PAUSA)
        print(f"\n💾  {total} pedidos gravados/atualizados no cache.")
    else:
        print("↩️   --no-fetch: recalculando do cache existente.")
        # detecta campos a partir do raw já salvo
        sample = [json.loads(r[0]) for r in con.execute("SELECT raw FROM pedidos LIMIT 300").fetchall()]
        campos = detect_fields(sample)

    # ── relatório de campos detectados ──
    print("\n🔎  Campos detectados no JSON da API:")
    for k in ("data", "valor", "status_txt", "status_id", "cliente_nome", "cliente_fone",
              "bairro", "cidade", "pagamento", "produtos"):
        v = campos.get(k)
        print(f"     {k:14} → {v if v else '— não encontrado'}")
    if campos.get("_prod"):
        print(f"     produto.itens  → {campos['_prod']}")

    # ── payload + saídas ──
    print("\n📊  Calculando métricas...")
    payload = build_payload(con, campos, desde, ate, args.ciclo, args.ocupacao)
    if payload["meta"]["total_pedidos"] == 0:
        print("⚠️  Cache vazio — nada a renderizar. Rode sem --no-fetch (ou com --mock).")
        return

    html_path = os.path.join(out_dir, "dashboard.html")
    xlsx_path = os.path.join(out_dir, "bi_maisdelivery.xlsx")
    render_html(payload, html_path)
    try:
        write_excel(con, payload, xlsx_path)
    except Exception as ex:
        print(f"⚠️  Excel falhou ({ex}). Instale openpyxl: pip install openpyxl")
        xlsx_path = None

    print("\n" + "=" * 60)
    print(f"✅  Dashboard : {html_path}")
    if xlsx_path:
        print(f"✅  Excel     : {xlsx_path}")
    print(f"✅  Cache     : {db_path}")
    if not args.mock and not args.no_fetch:
        print(f"✅  Amostra   : {os.path.join(out_dir, '_amostra_pedido.json')}")
    print("=" * 60)
    print("\n👉  Abra o dashboard.html no navegador.\n")


if __name__ == "__main__":
    main()