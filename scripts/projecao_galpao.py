#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Gera a planilha de projeção do estoque do galpão.

Lê a planilha de estoque do galpão (Item, Código, Categoria, Quantidade,
Custo unit., Preço venda) e o relatório "stock_general_full" do Mercado Livre,
para colocar ao lado de cada item o que ele tem no Full e quantos meses o
estoque dura no ritmo de venda do ML.

Uso:
    python3 scripts/projecao_galpao.py ESTOQUE_GALPAO.xlsx RELATORIO_ML.xlsx SAIDA.xlsx DD/MM/AAAA
"""
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from analise_full_ml import (AMARELO, ARIAL, AZUL, BORDA, BRANCO, CINZA, MAPA, MARCA, MOEDA, SUAVE,
                             fonte, injetar_cache, ler_relatorio, validar)

CABECALHO_GALPAO = ["Item", "Código", "Categoria", "Quantidade", "Custo unit. (R$)", "Preço venda (R$)"]
PCT = '0.0%;-0.0%;"—"'


def ler_galpao(caminho):
    ws = load_workbook(caminho, data_only=True).active
    cab = [ws.cell(row=1, column=c).value for c in range(1, 7)]
    if cab != CABECALHO_GALPAO:
        sys.exit("Cabeçalho da planilha do galpão mudou: %r" % cab)
    itens, r = [], 2
    while ws["A%d" % r].value not in (None, "TOTAL"):
        itens.append(dict(nome=ws["A%d" % r].value, codigo=ws["B%d" % r].value, categoria=ws["C%d" % r].value or "",
                          qtd=ws["D%d" % r].value or 0, custo=ws["E%d" % r].value or 0, preco=ws["F%d" % r].value or 0))
        r += 1
    codigos = [i["codigo"] for i in itens]
    if len(codigos) != len(set(codigos)):
        sys.exit("Código repetido na planilha do galpão")
    return itens


def gerar(arq_galpao, arq_ml, saida, data_txt):
    itens = ler_galpao(arq_galpao)
    full = {}
    for l in ler_relatorio(arq_ml):
        full[MAPA[l["sku_ml"]]] = dict(no_full=l["aptas"] + l["transf"] + l["nao"], caminho=l["pendente"],
                                       devol=l["pend_espaco"] - l["pendente"], vendas=l["vendas_un"])
    sem_galpao = sorted(set(full) - {i["codigo"] for i in itens})
    if sem_galpao:
        sys.exit("Itens do Full que não estão na planilha do galpão: %s" % sem_galpao)
    for i in itens:
        f = full.get(i["codigo"], dict(no_full=0, caminho=0, devol=0, vendas=0))
        i.update(f)
        i["cpar"] = round(i["qtd"] * i["custo"], 2)
        i["rec"] = round(i["qtd"] * i["preco"], 2)
        i["total_emp"] = i["qtd"] + i["no_full"] + i["caminho"] + i["devol"]
        i["meses"] = round(i["total_emp"] / float(i["vendas"]), 1) if i["vendas"] else ""
    itens.sort(key=lambda i: (-i["cpar"], i["codigo"]))

    cache = {"Resumo": {}, "Galpão por SKU": {}}
    wb = Workbook()

    # ------------------------------ Galpão por SKU ------------------------------
    det = wb.active
    det.title = "Galpão por SKU"
    cd = cache["Galpão por SKU"]
    colunas = [("nome", "Item", 32), ("cod", "Código", 11), ("cat", "Categoria", 13), ("qtd", "Quantidade no galpão", 11),
               ("cus", "Custo unit. (R$)", 11), ("pre", "Preço venda líquido (R$)", 11), ("cpar", "Custo parado (R$)", 14),
               ("rec", "Receita projetada (R$)", 14), ("mar", "Margem projetada (R$)", 14), ("mpct", "Margem (%)", 9),
               ("part", "% do dinheiro parado no galpão", 11), ("full", "No Full", 8), ("cam", "A caminho do Full", 10),
               ("dev", "Devoluções voltando", 11), ("emp", "Total da empresa", 10), ("vml", "Vendidas 30 dias no ML", 10),
               ("mes", "Meses de estoque", 10)]
    C = {k: i + 1 for i, (k, _, _) in enumerate(colunas)}
    L = {k: get_column_letter(v) for k, v in C.items()}
    for k, t, w in colunas:
        cel = det.cell(row=1, column=C[k], value=t)
        cel.font, cel.fill = fonte(True, BRANCO), PatternFill("solid", fgColor=MARCA)
        cel.alignment = Alignment(vertical="center", wrap_text=True)
        det.column_dimensions[L[k]].width = w
    det.row_dimensions[1].height = 44
    det.freeze_panes = "C2"

    L0 = 2
    ULT = L0 + len(itens) - 1
    TOT = ULT + 1
    total_cpar = round(sum(i["cpar"] for i in itens), 2)
    inteiros = {C[k] for k in ("qtd", "full", "cam", "dev", "emp", "vml")}
    dinheiro = {C[k] for k in ("cus", "pre", "cpar", "rec", "mar")}
    pct = {C["mpct"], C["part"]}

    for n, i in enumerate(itens):
        r = L0 + n
        for k, v in (("nome", i["nome"]), ("cod", i["codigo"]), ("cat", i["categoria"]), ("qtd", i["qtd"]),
                     ("cus", i["custo"]), ("pre", i["preco"] or None), ("full", i["no_full"]), ("cam", i["caminho"]),
                     ("dev", i["devol"]), ("vml", i["vendas"])):
            det.cell(row=r, column=C[k], value=v)

        def f(k, frm, val):
            det.cell(row=r, column=C[k], value=frm)
            cd["%s%d" % (L[k], r)] = val
        mar = round(i["rec"] - i["cpar"], 2) if i["preco"] else ""
        f("cpar", "={q}{r}*{c}{r}".format(q=L["qtd"], c=L["cus"], r=r), i["cpar"])
        f("rec", "={q}{r}*{p}{r}".format(q=L["qtd"], p=L["pre"], r=r), i["rec"])
        f("mar", '=IF({p}{r}>0,{rc}{r}-{cp}{r},"")'.format(p=L["pre"], rc=L["rec"], cp=L["cpar"], r=r), mar)
        f("mpct", '=IF({rc}{r}>0,{m}{r}/{rc}{r},"")'.format(rc=L["rec"], m=L["mar"], r=r),
          (mar / i["rec"]) if i["rec"] else "")
        f("part", "=IF({cp}${t}>0,{cp}{r}/{cp}${t},0)".format(cp=L["cpar"], t=TOT, r=r),
          i["cpar"] / total_cpar if total_cpar else 0)
        f("emp", "={q}{r}+{fu}{r}+{ca}{r}+{de}{r}".format(q=L["qtd"], fu=L["full"], ca=L["cam"], de=L["dev"], r=r),
          i["total_emp"])
        f("mes", '=IF({v}{r}>0,ROUND({e}{r}/{v}{r},1),"")'.format(v=L["vml"], e=L["emp"], r=r), i["meses"])
        for col in range(1, len(colunas) + 1):
            cel = det.cell(row=r, column=col)
            cel.font, cel.border = fonte(), BORDA
            if col in inteiros: cel.number_format = "#,##0"
            if col in dinheiro: cel.number_format = MOEDA
            if col in pct: cel.number_format = PCT
            if col == C["mes"]: cel.number_format = "0.0"
        for k in ("cus", "pre"):
            det.cell(row=r, column=C[k]).fill = PatternFill("solid", fgColor=AMARELO)
            det.cell(row=r, column=C[k]).font = fonte(cor=AZUL)

    det.cell(row=TOT, column=1, value="TOTAL")
    somas = {"qtd": sum(i["qtd"] for i in itens), "cpar": total_cpar, "rec": round(sum(i["rec"] for i in itens), 2),
             "mar": round(sum(i["rec"] - i["cpar"] for i in itens if i["preco"]), 2),
             "full": sum(i["no_full"] for i in itens), "cam": sum(i["caminho"] for i in itens),
             "dev": sum(i["devol"] for i in itens), "emp": sum(i["total_emp"] for i in itens),
             "vml": sum(i["vendas"] for i in itens)}
    for k, v in somas.items():
        det.cell(row=TOT, column=C[k], value="=SUM({c}{a}:{c}{b})".format(c=L[k], a=L0, b=ULT))
        cd["%s%d" % (L[k], TOT)] = v
    det.cell(row=TOT, column=C["mpct"], value='=IF({rc}{t}>0,{m}{t}/{rc}{t},"")'.format(rc=L["rec"], m=L["mar"], t=TOT))
    cd["%s%d" % (L["mpct"], TOT)] = somas["mar"] / somas["rec"]
    det.cell(row=TOT, column=C["part"], value="=SUM({c}{a}:{c}{b})".format(c=L["part"], a=L0, b=ULT))
    cd["%s%d" % (L["part"], TOT)] = sum(cd["%s%d" % (L["part"], r)] for r in range(L0, ULT + 1))
    for col in range(1, len(colunas) + 1):
        cel = det.cell(row=TOT, column=col)
        cel.font, cel.fill = fonte(True), PatternFill("solid", fgColor=CINZA)
        cel.border = Border(top=Side(style="medium", color=MARCA))
        if col in inteiros: cel.number_format = "#,##0"
        if col in dinheiro: cel.number_format = MOEDA
        if col in pct: cel.number_format = PCT

    det.auto_filter.ref = "A1:%s%d" % (L["mes"], ULT)
    vermelho = dict(fill=PatternFill("solid", fgColor="FFFAE6E4"), font=Font(name=ARIAL, size=10, bold=True, color="FFAF2B23"))
    laranja = dict(fill=PatternFill("solid", fgColor="FFFBEEDA"), font=Font(name=ARIAL, size=10, bold=True, color="FF9D5C05"))
    fx = "%s2:%s%d" % (L["mes"], L["mes"], ULT)
    det.conditional_formatting.add(fx, CellIsRule(operator="greaterThan", formula=["12"], stopIfTrue=True, **vermelho))
    det.conditional_formatting.add(fx, CellIsRule(operator="greaterThan", formula=["6"], **laranja))
    det.sheet_view.showGridLines = False

    # ------------------------------ Resumo ------------------------------
    res = wb.create_sheet("Resumo", 0)
    cr = cache["Resumo"]
    for col, w in zip("ABCDEFG", (40, 12, 18, 20, 18, 10, 56)):
        res.column_dimensions[col].width = w
    res["A1"] = "Seu galpão — projeção de %s" % data_txt
    res["A1"].font = fonte(True, MARCA, 15)
    res["A2"] = "Fonte: planilha de estoque do galpão + relatório de estoque Full do Mercado Livre."
    res["A2"].font = fonte(cor=SUAVE, italic=True)

    D = "'Galpão por SKU'!"
    rng = lambda k: "%s%s%d:%s%d" % (D, L[k], L0, L[k], ULT)

    def put(ref, frm, val):
        res[ref] = frm
        cr[ref] = val

    def cab(r, titulos):
        for n, t in enumerate(titulos):
            cel = res.cell(row=r, column=1 + n, value=t)
            cel.font, cel.fill = fonte(True, BRANCO), PatternFill("solid", fgColor=MARCA)
            cel.alignment = Alignment(vertical="center", wrap_text=True)
        res.row_dimensions[r].height = 30

    def estilo(r, negrito=False, cinza=False, fmt_c=MOEDA):
        for col in range(1, 8):
            cel = res.cell(row=r, column=col)
            cel.font = fonte(negrito, cor=(SUAVE if col == 7 else None), size=(9 if col == 7 else 10))
            cel.border = BORDA
            if cinza: cel.fill = PatternFill("solid", fgColor=CINZA)
            if col == 7: cel.alignment = Alignment(wrap_text=True, vertical="top")
        res.cell(row=r, column=2).number_format = "#,##0"
        res.cell(row=r, column=3).number_format = fmt_c
        for col in (4, 5):
            res.cell(row=r, column=col).number_format = MOEDA
        res.cell(row=r, column=6).number_format = PCT

    def margem(r, rec, mar):
        put("E%d" % r, "=D%d-C%d" % (r, r), mar)
        put("F%d" % r, '=IF(D{r}>0,E{r}/D{r},"")'.format(r=r), (mar / rec) if rec else "")

    cab(4, ["Categoria", "Unidades", "Custo parado (R$)", "Receita projetada (R$)", "Margem projetada (R$)",
            "Margem (%)", "Observação"])
    com_preco = [i for i in itens if i["preco"]]
    categorias = sorted({i["categoria"] for i in com_preco})
    r = 5
    for cat in categorias:
        grupo = [i for i in com_preco if i["categoria"] == cat]
        un, cus, rec = sum(i["qtd"] for i in grupo), round(sum(i["cpar"] for i in grupo), 2), round(sum(i["rec"] for i in grupo), 2)
        res["A%d" % r] = cat or "Sem categoria"
        put("B%d" % r, '=SUMIFS(%s,%s,A%d,%s,">0")' % (rng("qtd"), rng("cat"), r, rng("pre")), un)
        put("C%d" % r, '=SUMIFS(%s,%s,A%d,%s,">0")' % (rng("cpar"), rng("cat"), r, rng("pre")), cus)
        put("D%d" % r, '=SUMIFS(%s,%s,A%d)' % (rng("rec"), rng("cat"), r), rec)
        margem(r, rec, round(rec - cus, 2))
        estilo(r)
        r += 1
    SUB = r
    res["A%d" % SUB] = "SUBTOTAL — ITENS COM PREÇO"
    for col in "BCD":
        put("%s%d" % (col, SUB), "=SUM(%s5:%s%d)" % (col, col, SUB - 1),
            round(sum(cr["%s%d" % (col, x)] for x in range(5, SUB)), 2))
    margem(SUB, cr["D%d" % SUB], round(cr["D%d" % SUB] - cr["C%d" % SUB], 2))
    estilo(SUB, True)

    SEM = SUB + 1
    sem = [i for i in itens if not i["preco"]]
    res["A%d" % SEM] = "Sem preço de venda"
    put("B%d" % SEM, "=SUM(%s)-B%d" % (rng("qtd"), SUB), sum(i["qtd"] for i in sem))
    put("C%d" % SEM, "=SUM(%s)-C%d" % (rng("cpar"), SUB), round(sum(i["cpar"] for i in sem), 2))
    res["D%d" % SEM] = 0
    res["G%d" % SEM] = "%s: entram no custo, mas não na receita, até ter preço de venda." % ", ".join(i["codigo"] for i in sem)
    estilo(SEM)

    TG = SEM + 1
    res["A%d" % TG] = "TOTAL DO GALPÃO"
    put("B%d" % TG, "=SUM(%s)" % rng("qtd"), somas["qtd"])
    put("C%d" % TG, "=SUM(%s)" % rng("cpar"), somas["cpar"])
    put("D%d" % TG, "=SUM(%s)" % rng("rec"), somas["rec"])
    margem(TG, somas["rec"], round(somas["rec"] - somas["cpar"], 2))
    res["G%d" % TG] = "A margem total desconta o custo dos itens sem preço de venda."
    estilo(TG, True, True)

    # visão da empresa
    r0 = TG + 2
    res["A%d" % r0] = "Visão da empresa (galpão + Mercado Livre Full)"
    res["A%d" % r0].font = fonte(True, MARCA, 12)
    cab(r0 + 1, ["Onde está", "Unidades", "Custo parado (R$)", "Receita projetada (R$)", "Margem projetada (R$)",
                 "Margem (%)", "Observação"])
    RG = r0 + 2
    res["A%d" % RG] = "Seu galpão"
    for col in "BCDEF":
        put("%s%d" % (col, RG), "=%s%d" % (col, TG), cr["%s%d" % (col, TG)])
    estilo(RG)
    linhas_ml = [("Dentro do Full", "full", "no_full", "Aptas, em transferência entre galpões do ML e não aptas."),
                 ("A caminho do Full", "cam", "caminho", "Saíram do galpão antes da contagem de 18/09."),
                 ("Devoluções voltando ao Full", "dev", "devol", "Podem voltar avariadas: o valor é o máximo.")]
    for n, (rot, k, campo, obs) in enumerate(linhas_ml):
        r = RG + 1 + n
        un = sum(i[campo] for i in itens)
        cus = round(sum(i[campo] * i["custo"] for i in itens), 2)
        rec = round(sum(i[campo] * i["preco"] for i in itens), 2)
        res["A%d" % r] = rot
        put("B%d" % r, "=SUM(%s)" % rng(k), un)
        put("C%d" % r, "=SUMPRODUCT(%s,%s)" % (rng(k), rng("cus")), cus)
        put("D%d" % r, "=SUMPRODUCT(%s,%s)" % (rng(k), rng("pre")), rec)
        margem(r, rec, round(rec - cus, 2))
        res["G%d" % r] = obs
        estilo(r)
    TE = RG + 1 + len(linhas_ml)
    res["A%d" % TE] = "TOTAL DA EMPRESA"
    for col in "BCD":
        put("%s%d" % (col, TE), "=SUM(%s%d:%s%d)" % (col, RG, col, TE - 1),
            round(sum(cr["%s%d" % (col, x)] for x in range(RG, TE)), 2))
    margem(TE, cr["D%d" % TE], round(cr["D%d" % TE] - cr["C%d" % TE], 2))
    res["G%d" % TE] = "Mesmos números da planilha do Full."
    estilo(TE, True, True)

    # giro
    r0 = TE + 2
    res["A%d" % r0] = "Giro — quanto tempo o estoque dura"
    res["A%d" % r0].font = fonte(True, MARCA, 12)
    cab(r0 + 1, ["Indicador", "SKUs", "Custo parado no galpão (R$)", "", "", "", "Observação"])
    giro = [
        ("Mais de 6 meses de estoque", '=COUNTIF(%s,">6")' % rng("mes"), '=SUMIFS(%s,%s,">6")' % (rng("cpar"), rng("mes")),
         [i for i in itens if i["meses"] != "" and i["meses"] > 6],
         "Somando galpão, Full, a caminho e devoluções, no ritmo de venda do ML dos últimos 30 dias."),
        ("Mais de 12 meses de estoque", '=COUNTIF(%s,">12")' % rng("mes"), '=SUMIFS(%s,%s,">12")' % (rng("cpar"), rng("mes")),
         [i for i in itens if i["meses"] != "" and i["meses"] > 12], ""),
        ("Sem venda no ML nos últimos 30 dias", "=COUNTIF(%s,0)" % rng("vml"), "=SUMIFS(%s,%s,0)" % (rng("cpar"), rng("vml")),
         [i for i in itens if not i["vendas"]], "Não têm anúncio no Full. Podem estar vendendo por outros canais."),
    ]
    for n, (rot, f_n, f_c, grupo, obs) in enumerate(giro):
        r = r0 + 2 + n
        res["A%d" % r] = rot
        put("B%d" % r, f_n, len(grupo))
        put("C%d" % r, f_c, round(sum(i["cpar"] for i in grupo), 2))
        res["G%d" % r] = obs
        estilo(r)

    r = r0 + 2 + len(giro) + 1
    res["A%d" % r] = "Como ler"
    res["A%d" % r].font = fonte(True, MARCA, 12)
    notas = [
        "Custo parado = quantidade × custo. Receita projetada = quantidade × seu preço de venda líquido (o que você recebe por peça, já sem as tarifas do ML). Margem = receita − custo.",
        "Custo e preço vêm da sua planilha (células amarelas da aba Galpão por SKU). Se mudar um valor ali, tudo recalcula.",
        "No Full, a caminho e devoluções vêm do relatório do Mercado Livre de %s, os mesmos números da planilha do Full." % data_txt,
        "Meses de estoque = total da empresa ÷ unidades vendidas no ML nos últimos 30 dias. Não conta vendas feitas direto do galpão (Shopee, balcão): itens que também vendem por fora giram mais rápido do que aparece. Laranja acima de 6 meses, vermelho acima de 12.",
    ]
    for t in notas:
        r += 1
        cel = res.cell(row=r, column=1, value="• " + t)
        cel.font, cel.alignment = fonte(), Alignment(wrap_text=True, vertical="top")
        res.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
        res.row_dimensions[r].height = 30
    res.sheet_view.showGridLines = False

    wb.save(saida)
    injetar_cache(saida, cache)
    validar(saida, cache)
    return cache, itens, dict(SUB=SUB, SEM=SEM, TG=TG, RG=RG, TE=TE)


if __name__ == "__main__":
    if len(sys.argv) != 5:
        sys.exit(__doc__)
    cache, itens, R = gerar(*sys.argv[1:])
    c = cache["Resumo"]
    print("Galpão:  custo R$ %10.2f | receita R$ %10.2f | margem R$ %10.2f (%.1f%%)" % (
        c["C%d" % R["TG"]], c["D%d" % R["TG"]], c["E%d" % R["TG"]], 100 * c["F%d" % R["TG"]]))
    print("Com preço: margem R$ %.2f (%.1f%%) | sem preço: %d un, custo R$ %.2f" % (
        c["E%d" % R["SUB"]], 100 * c["F%d" % R["SUB"]], c["B%d" % R["SEM"]], c["C%d" % R["SEM"]]))
    print("Empresa: custo R$ %10.2f | receita R$ %10.2f | margem R$ %10.2f" % (
        c["C%d" % R["TE"]], c["D%d" % R["TE"]], c["E%d" % R["TE"]]))
    for r in range(5, R["SUB"]):
        print("  categoria %-13s un=%5d custo=%10.2f receita=%10.2f margem=%.1f%%" % (
            c.get("A%d" % r, ""), c["B%d" % r], c["C%d" % r], c["D%d" % r], 100 * c["F%d" % r]))
