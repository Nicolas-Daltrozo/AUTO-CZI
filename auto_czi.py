import pandas as pd
from docxtpl import DocxTemplate

ARQUIVO = "EPE023 Lista de Brincos.xlsx"

TEMPLATE = "DRAFT CZI.docx"

df_brincos = pd.read_excel(
    ARQUIVO,
    sheet_name= "Lista",
    usecols=["EPE", "CZIs", "CONSIGNEE / NOTIFY", "ADDRES", "QTDE", "KGS"])

df_czi = pd.read_excel(
    ARQUIVO,
    sheet_name= "Engorda",
    dtype={"IDE FINAL": str},
    usecols=["IDE FINAL", "IDADE", "PROPRIEDADE DE ORIGEM", "COD. ESTABELECIMENTO", "MUNICIPIO/ESTADO", "Municipio - Cod. IBGE"])

cont = 0

for linha in df_brincos.itertuples(index=False):

    quantidade = linha.QTDE #QTDE

    lote = df_czi.iloc[cont:cont + quantidade]

    nome_czi = linha.CZIs
    cosignatario = linha._2.upper() #CONSIGNEE / NOTIFY
    endereco = linha.ADDRES.upper() #ADDRES
    peso = linha.KGS #KGS


    bois_quantidade = [ x for x in range(1, quantidade +1)]
    bois_col2 = [ "Bovina/Bovine" ] * quantidade
    bois_col3 = "\n".join([ "ELECTRONIC EAR TAG" ]* quantidade)
    bois_brincos = "\n".join(lote["IDE FINAL"])
    bois_idades = "\n".join(pd.to_datetime(lote["IDADE"]).dt.strftime("%d/%m/%Y"))
    bois_col6 = [ "Macho / Male" ] * quantidade

    #fazendas
    lista_fazenda = list(
        dict.fromkeys(
            zip(
                lote["PROPRIEDADE DE ORIGEM"],
                lote["COD. ESTABELECIMENTO"],
                lote["MUNICIPIO/ESTADO"]
            )
        )
    )

    #ibge
    ibge_total = "\n".join(
        dict.fromkeys(
            lote["Municipio - Cod. IBGE"].astype(str)
        )
    )

    context = {
        "nome_czi": nome_czi,
        "cosignatario": cosignatario,
        "endereco": endereco,
        "IBGE": ibge_total,
        "lista_fazenda": lista_fazenda,
        "quantidade": quantidade,
        "peso": (f"{peso:_.3f}").replace('.',',').replace('_','.'),
        "bois_quantidade": bois_quantidade,
        "bois_col2": bois_col2,
        "bois_col3": bois_col3,
        "bois_brincos": bois_brincos,
        "bois_idades": bois_idades,
        "bois_col6": bois_col6
    }

    doc = DocxTemplate(TEMPLATE)
    doc.render(context)
    doc.save(f"{nome_czi}.docx")

    cont+=quantidade

    if cont >= len(df_czi):
        break

    # print(lista_fazenda)