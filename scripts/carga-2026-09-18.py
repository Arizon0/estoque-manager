# -*- coding: utf-8 -*-
import json, os

HOJE = "2026-09-18"

# (id, nome, sku, categoria, contagem_fisica, envio_ml_full)
ITENS = [
    ("anel-8126-050",   "Anéis 8126 0,50",                    "8126 0,50", "Anéis",        3,   0),
    ("anel-9200-std",   "Anéis 9200 STD",                     "9200 STD",  "Anéis",        18,  0),
    ("anel-9403-std",   "Anéis 9403 STD",                     "9403 STD",  "Anéis",        20,  0),
    ("anel-8126-std",   "Anel 8126 STD",                      "8126 STD",  "Anéis",        1,   0),
    ("anel-6631-050",   "Anel SDA 6631 0,50",                 "6631 0,50", "Anéis",        130, 30),
    ("anel-7092",       "Anel SDA 7092",                      "7092",      "Anéis",        25,  0),
    ("bronzina-1035j",  "Bronzina de Biela SBB 1035-J 0,25",  "1035-J",    "Bronzina",     24,  0),
    ("jogo-7224-050",   "Jogo de Anel 7224 0,50",             "7224 0,50", "Jogo de Anel", 54,  0),
    ("jogo-7224-std",   "Jogo de Anel 7224 STD",              "7224 STD",  "Jogo de Anel", 43,  0),
    ("ret-1135",        "Retentor 1135",                      "1135",      "Retentor",     118, 0),
    ("ret-1942",        "Retentor 1942",                      "1942",      "Retentor",     675, 0),
    ("ret-2075",        "Retentor 2075",                      "2075",      "Retentor",     124, 0),
    ("ret-2178",        "Retentor 2178",                      "2178",      "Retentor",     652, 300),
    ("ret-2283",        "Retentor 2283",                      "2283",      "Retentor",     274, 0),
    ("ret-2317",        "Retentor 2317",                      "2317",      "Retentor",     72,  0),
    ("ret-2370",        "Retentor 2370",                      "2370",      "Retentor",     58,  0),
    ("ret-2371",        "Retentor 2371 (Shopee)",             "2371",      "Retentor",     359, 0),
    ("ret-2373",        "Retentor 2373",                      "2373",      "Retentor",     101, 0),
    ("ret-2374",        "Retentor 2374",                      "2374",      "Retentor",     42,  40),
    ("ret-2400",        "Retentor 2400",                      "2400",      "Retentor",     1,   0),
    ("ret-2525",        "Retentor 2525",                      "2525",      "Retentor",     600, 0),
    ("ret-2539",        "Retentor 2539",                      "2539",      "Retentor",     65,  30),
    ("ret-2544",        "Retentor 2544",                      "2544",      "Retentor",     224, 0),
    ("ret-3044",        "Retentor 3044",                      "3044",      "Retentor",     598, 100),
    ("ret-5159",        "Retentor 5159",                      "5159",      "Retentor",     911, 0),
    ("ret-5245",        "Retentor 5245",                      "5245",      "Retentor",     9,   0),
    ("ret-5266",        "Retentor 5266",                      "5266",      "Retentor",     53,  30),
    ("ret-5338",        "Retentor 5338",                      "5338",      "Retentor",     102, 0),
    ("ret-5502",        "Retentor 5502",                      "5502",      "Retentor",     45,  0),
    ("ret-5601",        "Retentor 5601",                      "5601",      "Retentor",     23,  15),
    ("ret-5699",        "Retentor 5699",                      "5699",      "Retentor",     49,  20),
    ("ret-5702",        "Retentor 5702",                      "5702",      "Retentor",     0,   0),   # não contado
    ("ret-5772",        "Retentor 5772",                      "5772",      "Retentor",     172, 0),
    ("ret-5801",        "Retentor 5801",                      "5801",      "Retentor",     29,  0),
    ("ved-9203",        "Vedadores 9203",                     "9203",      "Vedador",      475, 200),
    ("item-8566-std",   "8566 STD",                           "8566 STD",  "",             20,  0),
    ("item-2900",       "2900",                               "2900",      "",             50,  0),
]

# vendas registradas nas listas de contagem (fotos 2 e 3 + lista sem data)
VENDAS_LISTAS = {
    "ret-5338": 10, "ret-2178": 19, "ret-5601": 2,  "ret-5772": 5,  "ved-9203": 20,
    "ret-5266": 6,  "ret-2370": 16, "anel-9200-std": 7, "jogo-7224-std": 8,
    "ret-5159": 7,  "ret-2373": 3,  "ret-1942": 58, "ret-3044": 8,  "anel-6631-050": 1,
    "ret-2544": 3,  "ret-5245": 9,  "ret-2374": 5,  "anel-7092": 2, "ret-2283": 3,
    "anel-8126-050": 3, "ret-1135": 1, "jogo-7224-050": 1,
}
VENDAS_16 = {
    "ved-9203": 4, "ret-2178": 3, "ret-5338": 2, "ret-1942": 9, "anel-9200-std": 2,
    "ret-3044": 2, "ret-2283": 1, "ret-5601": 1, "ret-5772": 1, "bronzina-1035j": 1,
}
VENDAS_17 = {
    "ret-5159": 2, "ret-2370": 3, "ret-5772": 3, "ret-5699": 1,
    "ret-5338": 2, "ret-3044": 2, "ret-2544": 1,
}

def ts(h, m, s=0):
    return "2026-09-18T%02d:%02d:%02d.000Z" % (h, m, s)

docs_itens, lancamentos, relatorio = [], [], []
seq = [0]

def lancar(item_id, nome, sku, tipo, delta, saldo, obs, hora):
    lancamentos.append({
        "itemId": item_id, "nome": nome, "sku": sku, "tipo": tipo,
        "delta": delta, "saldo": saldo, "obs": obs, "em": hora, "origem": "relatorio",
    })

for item_id, nome, sku, categoria, contagem, ml in ITENS:
    saldo = contagem
    lancar(item_id, nome, sku, "ajuste", contagem, saldo,
           "Contagem física do galpão" if item_id != "ret-5702" else "Sem contagem na planilha — confirmar",
           ts(8, 0))
    if ml:
        saldo -= ml
        lancar(item_id, nome, sku, "transferencia", -ml, saldo, "Envio para o galpão ML Full", ts(9, 0))
    for fonte, mapa, hora, rot in (
        ("16/09", VENDAS_16, ts(10, 0), "Vendas de 16/09"),
        ("17/09", VENDAS_17, ts(11, 0), "Vendas de 17/09"),
        ("listas", VENDAS_LISTAS, ts(12, 0), "Vendas conferidas nas listas"),
    ):
        q = mapa.get(item_id, 0)
        if q:
            saldo -= q
            lancar(item_id, nome, sku, "venda", -q, saldo, rot, hora)

    total_vendas = VENDAS_16.get(item_id, 0) + VENDAS_17.get(item_id, 0) + VENDAS_LISTAS.get(item_id, 0)
    docs_itens.append((item_id, {
        "nome": nome, "sku": sku, "categoria": categoria, "unidade": "un",
        "qtd": saldo, "minimo": 0, "custo": 0, "preco": 0,
        "criadoEm": ts(8, 0), "atualizadoEm": ts(12, 0),
    }))
    relatorio.append((nome, sku, contagem, ml, total_vendas, saldo))

os.makedirs("json", exist_ok=True)
for item_id, corpo in docs_itens:
    json.dump(corpo, open("json/item-%s.json" % item_id, "w", encoding="utf-8"), ensure_ascii=False)
json.dump({"dia": HOJE, "atualizadoEm": ts(12, 0), "lancamentos": lancamentos},
          open("json/mov-%s.json" % HOJE, "w", encoding="utf-8"), ensure_ascii=False)

writes = [{"op": "delete", "collection": "itens", "doc_id": "ex%d" % i} for i in range(1, 7)]
writes += [{"op": "delete", "collection": "movimentos", "doc_id": d} for d in ("2026-09-16", "2026-09-17")]
writes += [{"op": "set", "collection": "itens", "doc_id": i,
            "file_path": os.path.abspath("json/item-%s.json" % i)} for i, _ in docs_itens]
writes += [{"op": "set", "collection": "movimentos", "doc_id": HOJE,
            "file_path": os.path.abspath("json/mov-%s.json" % HOJE)}]
json.dump(writes, open("writes.json", "w", encoding="utf-8"), ensure_ascii=False)

print("%-36s %6s %7s %7s %8s" % ("ITEM", "CONT.", "MLFULL", "VENDAS", "SALDO"))
print("-" * 70)
for nome, sku, cont, ml, vend, saldo in relatorio:
    marca = "  <-- NEGATIVO" if saldo < 0 else ("  <-- zerou" if saldo == 0 else "")
    print("%-36s %6d %7s %7s %8d%s" % (nome[:36], cont, ("-%d" % ml) if ml else "", ("-%d" % vend) if vend else "", saldo, marca))
print("-" * 70)
print("itens: %d | lançamentos: %d | writes: %d" % (len(docs_itens), len(lancamentos), len(writes)))
print("total contado: %d | total ML Full: %d | total vendas: %d | saldo final: %d" % (
    sum(r[2] for r in relatorio), sum(r[3] for r in relatorio),
    sum(r[4] for r in relatorio), sum(r[5] for r in relatorio)))
print("tamanho doc movimentos: %d bytes" % os.path.getsize("json/mov-%s.json" % HOJE))
