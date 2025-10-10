import streamlit as st
import pandas as pd
import re
import numpy as np
from io import BytesIO
# ----------------------------
# Configurações do app
# ----------------------------
st.set_page_config(
    page_title="Validador de Embalagens e Preços",
    layout="wide"
)

st.markdown("<h1 style='text-align: center;'>Validador de Mapeio e Preços</h1>", unsafe_allow_html=True)
st.markdown("""
Para o funcionamento correto da ferramenta, são necessárias colunas que **tenham nomes semelhantes** aos seguintes:
- `Descripcion`, `PROD_NOMBRE_ORIGINAL`, `Nome SKU`
- `Contenido`, `Qtd Conteúdo SKU`
- `Precio KG/LT`, `Preço convertido kg/lt R$`, `Preço kg/lt`
- `Est Mer 7 (Subcategoria)`, `NIVEL1`
- `Imp Vta (Ult.24 Meses)`, `Vendas em volume`

A ferramenta faz:
- Validação da quantidade de embalagem (`QtdEmbalagem` e `QtdEmbalagemGramas`)
- Validação de conteúdo (`ValidacaoContenido`)
- Detecção de preços fora do padrão por subcategoria (`ValidacionPrecio`)
- Detecção adicional de preços extremos baseados na mediana (`ValidacionPrecioMediana`)
- Download do Excel processado
""")

# ----------------------------
# Funções auxiliares
# ----------------------------
def extrair_peso(texto):
    if pd.isna(texto):
        return None, None
    texto = str(texto)
    unidades_intermed = r"(?:UN|UNID|CJ|CX|DS|PCT|FD|SC)?"
    unidade_final = r"(kg|g|gr|ml|l|lt)"

    match_multi = re.search(rf"((?:\d+\s*{unidades_intermed}\s*[xX]\s*)+\d+[.,]?\d*\s*{unidade_final})", texto, re.IGNORECASE)
    if match_multi:
        bloco = match_multi.group(1)
        unidade = match_multi.group(len(match_multi.groups())).lower()
        numeros = [float(n.replace(",", ".")) for n in re.findall(r"\d+[.,]?\d*", bloco)]
        multiplicadores = numeros[:-1] if len(numeros) > 1 else []
        peso = numeros[-1]
        if unidade in ["kg", "lt", "l"]:
            peso *= 1000
        total = peso
        for n in multiplicadores:
            total *= n
        return bloco, int(total)

    match_3d = re.search(rf"(\d+)\s*[xX]\s*(\d+)\s*[xX]\s*(\d+[.,]?\d*)\s*{unidade_final}\b", texto, re.IGNORECASE)
    if match_3d:
        n1 = int(match_3d.group(1))
        n2 = int(match_3d.group(2))
        valor = float(match_3d.group(3).replace(",", "."))
        unidade = match_3d.group(4).lower()
        if unidade in ["kg", "lt", "l"]:
            valor *= 1000
        return match_3d.group(0), int(n1 * n2 * valor)

    match_2d = re.search(rf"(\d+)\s*[xX]\s*(\d+[.,]?\d*)\s*{unidade_final}\b", texto, re.IGNORECASE)
    if match_2d:
        n1 = int(match_2d.group(1))
        valor = float(match_2d.group(2).replace(",", "."))
        unidade = match_2d.group(3).lower()
        if unidade in ["kg", "lt", "l"]:
            valor *= 1000
        return match_2d.group(0), int(n1 * valor)

    match = re.search(rf"(\d+[.,]?\d*)\s*{unidade_final}\b", texto, re.IGNORECASE)
    if match:
        valor = float(match.group(1).replace(",", "."))
        unidade = match.group(2).lower()
        if unidade in ["kg", "lt", "l"]:
            valor *= 1000
        return match.group(0), int(valor)

    return None, None

# ----------------------------
# Validador de preço por categoria (IQR 5-95%)
# ----------------------------
def validar_precio_por_categoria(df, coluna_preco, coluna_categoria):
    df[coluna_preco] = (
        df[coluna_preco]
        .astype(str)                    # garante que é texto
        .str.replace(r"[^\d,.-]", "", regex=True)  # remove "R$", espaços, etc.
        .str.replace(",", ".", regex=False)        # troca vírgula por ponto
        )
    df[coluna_preco] = pd.to_numeric(df[coluna_preco], errors="coerce")

    def marcar_outliers(grupo):
        limite_inferior = grupo.quantile(0.05)
        limite_superior = grupo.quantile(0.95)
        return grupo.apply(lambda x: "OK" if limite_inferior <= x <= limite_superior else "OUTLIER")
    return df.groupby(coluna_categoria)[coluna_preco].transform(marcar_outliers)

# ----------------------------
# Novo validador: outliers com base na mediana (3x acima ou 1/3 abaixo)
# ----------------------------
def validar_precio_mediana(df, coluna_preco, coluna_categoria):
    df[coluna_preco] = (
        df[coluna_preco]
        .astype(str)                    # garante que é texto
        .str.replace(r"[^\d,.-]", "", regex=True)  # remove "R$", espaços, etc.
        .str.replace(",", ".", regex=False)        # troca vírgula por ponto
    )
    df[coluna_preco] = pd.to_numeric(df[coluna_preco], errors="coerce")

    def marcar_por_mediana(grupo):
        mediana = grupo.median()
        limite_inferior = mediana / 3
        limite_superior = mediana * 3
        return grupo.apply(lambda x: "OK" if limite_inferior <= x <= limite_superior else "OUTLIER_MEDIANA")
    return df.groupby(coluna_categoria)[coluna_preco].transform(marcar_por_mediana)

# ----------------------------
# Exportar para Excel
# ----------------------------
def to_excel_com_resumo(df, coluna_vendas):
    from io import BytesIO
    output = BytesIO()

    # ----------------------------
    # Criar resumo
    # ----------------------------
    total_itens = len(df)
    problemas_contenido = (df["ValidacaoContenido"] == "PROBLEMA").sum()
    outliers_quartil = (df["ValidacionPrecio"] == "OUTLIER").sum()
    outliers_mediana = (df["ValidacionPrecioMediana"] == "OUTLIER_MEDIANA").sum()

    outliers_ambos = ((df["ValidacionPrecio"] == "OUTLIER") & 
                      (df["ValidacaoContenido"] != "PROBLEMA") & 
                      (df["ValidacionPrecioMediana"] == "OUTLIER_MEDIANA")).sum()
    outliers_somente_mediana = ((df["ValidacionPrecio"] != "OUTLIER") & 
                                (df["ValidacionPrecioMediana"] == "OUTLIER_MEDIANA")).sum()
    outliers_somente_quartil = ((df["ValidacionPrecio"] == "OUTLIER") & 
                                (df["ValidacionPrecioMediana"] != "OUTLIER_MEDIANA")).sum()

    problemas_valor_bruto = problemas_contenido + outliers_ambos + outliers_somente_mediana + outliers_somente_quartil
    problemas_valor_perc = problemas_valor_bruto / total_itens * 100 if total_itens else 0

    volume_total = df[coluna_vendas].sum()
    df_problemas = df[
        (df["ValidacaoContenido"] == "PROBLEMA") |
        (df["ValidacionPrecio"] == "OUTLIER") |
        (df["ValidacionPrecioMediana"] == "OUTLIER_MEDIANA")
    ]
    volume_problemas = df_problemas[coluna_vendas].sum()
    volume_problemas_perc = volume_problemas / volume_total * 100 if volume_total else 0

    df_resumo = pd.DataFrame({
        "Métrica": [
            "Qtd total de SKUs/Itens",
            "Qtd de problemas de contenido",
            "Qtd de outliers em ambos (mediana e quartil)",
            "Qtd de outliers apenas mediana",
            "Qtd de outliers apenas quartil",
            "Qtd total de SKUs/itens com problema (valor bruto)",
            "Qtd total de SKUs/itens com problema (%)",
            "Volume de vendas total",
            "Volume de vendas com problema (valor bruto)",
            "Volume de vendas com problema (%)"
        ],
        "Valor": [
            total_itens,
            problemas_contenido,
            outliers_ambos,
            outliers_somente_mediana,
            outliers_somente_quartil,
            problemas_valor_bruto,
            round(problemas_valor_perc, 2),
            volume_total,
            volume_problemas,
            round(volume_problemas_perc, 2)
        ]
    })

    # # ----------------------------
    # # Criar Excel com duas abas
    # # ----------------------------
    # with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    #     # Aba com dados detalhados
    #     df.to_excel(writer, index=False, sheet_name="Dados")
        
    #     # Aba resumo
    #     df_resumo.to_excel(writer, index=False, sheet_name="Resumo")
        
    #     # Formatação
    #     workbook  = writer.book
    #     worksheet = writer.sheets["Resumo"]
        
    #     # Formato de porcentagem
    #     percent_fmt = workbook.add_format({'num_format': '0.0%'})
        
    #     # Aplica formato de % apenas nas linhas correspondentes
    #     # df_resumo é 0-indexed: linha 6 → B8, linha 9 → B11
    #     worksheet.write_number(7, 1, df_resumo.loc[6, "Valor"] / 100, percent_fmt)
    #     worksheet.write_number(10, 1, df_resumo.loc[9, "Valor"] / 100, percent_fmt)


    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # --- Aba Detalhes ---
        df.to_excel(writer, index=False, sheet_name="Detalhes")

        # --- Aba Resumo ---
        df_resumo.to_excel(writer, index=False, sheet_name="Resumo", startrow=1)

        workbook = writer.book
        worksheet = writer.sheets["Resumo"]

        # ----------------------------
        # FORMATOS
        # ----------------------------
        header_format = workbook.add_format({
            "bold": True, "align": "center", "valign": "vcenter",
            "bg_color": "#D9D9D9", "border": 1
        })

        normal_format = workbook.add_format({"border": 1})
        bold_format = workbook.add_format({"bold": True, "border": 1})
        orange_bold_format = workbook.add_format({
            "bold": True, "border": 1, "font_color": "#E36C0A"
        })
        gray_format = workbook.add_format({
            "bg_color": "#F2F2F2", "border": 1, "bold": True
        })
        percent_format = workbook.add_format({
            "num_format": "0.0%", "border": 1
        })
        number_format = workbook.add_format({
            "num_format": "#,##0", "border": 1
        })
        empty_format = workbook.add_format()  # para linhas em branco

        # ----------------------------
        # AJUSTE DE LARGURAS
        # ----------------------------
        worksheet.set_column("A:A", 60)
        worksheet.set_column("B:B", 25)

        # ----------------------------
        # CABEÇALHOS
        # ----------------------------
        worksheet.write("A1", "Métrica", header_format)
        worksheet.write("B1", "Números", header_format)

        # ----------------------------
        # APLICA FORMATAÇÃO LINHA A LINHA
        # ----------------------------
        # for i, (metrica, valor) in enumerate(zip(df_resumo["Métrica"], df_resumo["Valor"]), start=2):
        #     # Linhas com % (mesmas do seu código anterior)
        #     if i in [8, 12]:  # B8 e B11 → linhas 8 e 11 (1-based + 1 de header)
        #         worksheet.write_number(i - 1, 1, valor / 100, percent_format)
        #     else:
        #         worksheet.write(i - 1, 1, valor, number_format)
        #     if i == 3:
        #         worksheet.write(i - 1, 0, metrica, normal_format)

        # 1️⃣ Escreve a primeira linha do df_resumo na linha 2
        worksheet.write(1, 0, df_resumo["Métrica"].iloc[0], normal_format)
        worksheet.write(1, 1, df_resumo["Valor"].iloc[0], number_format)

        # 2️⃣ Cabeçalho na linha 3
        worksheet.write(2, 0, "Métrica", header_format)
        worksheet.write(2, 1, "Valor", header_format)

        # Linha inicial para escrever os dados (começando na linha 4, pois linha 3 é cabeçalho)
        linha_inicial = 3

        for i, (metrica, valor) in enumerate(zip(df_resumo["Métrica"], df_resumo["Valor"])):
            linha_atual = linha_inicial + i
            
            # Se chegamos à linha 8 (percentual), aplicamos o formato percentual
            if linha_atual == 8:
                worksheet.write_number(linha_atual - 1, 1, valor / 100, percent_format)
            else:
                worksheet.write(linha_atual - 1, 1, valor, number_format)
            
            # Escreve a métrica na coluna A
            worksheet.write(linha_atual - 1, 0, metrica, normal_format)

        # Inserir linha vazia após a linha 8 (que será a linha 9)
        worksheet.write_blank(8, 0, None, normal_format)
        worksheet.write_blank(8, 1, None, number_format)

        # ----------------------------
        # BLOCOS COLORIDOS
        # ----------------------------
        # 1️⃣ Critérios na linha 3
        worksheet.merge_range("A3:B3", "Critérios de itens com possíveis problemas", gray_format)

        # 2️⃣ Linhas laranja — na mesma linha correta
        worksheet.write("A8", "Qtd total de SKUs/itens com problema (%)", orange_bold_format)
        worksheet.write("B8", df_resumo.loc[6, "Valor"] / 100, percent_format)  # Qtd total problemas

        worksheet.write("A12", "Volume de vendas com problema (%)", orange_bold_format)
        worksheet.write_number("B12", df_resumo.loc[9, "Valor"] / 100, percent_format)

        # 3️⃣ Remove destaque do “Volume de vendas total” e insere linha em branco antes do “com problema”
        worksheet.write("A9", "", empty_format)
        worksheet.write("B9", "", empty_format)

        # 4️⃣ Título inferior (Top 50 SKUs)
        # worksheet.merge_range("A13:B13", "Top 50 Skus - Share Acumulado", gray_format)



    return output.getvalue()

# ----------------------------
# Colorir valores
# ----------------------------
def colorir_valores(val):
    if val == "PROBLEMA":
        return "background-color: #fff3cd"  # amarelo claro
    elif val in ["OUTLIER", "OUTLIER_MEDIANA"]:
        return "background-color: #f8d7da"  # vermelho claro
    return ""

# ----------------------------
# Upload do arquivo
# ----------------------------
uploaded_file = st.file_uploader("Escolha o arquivo Excel ou CSV", type=["xlsx", "csv"])

if uploaded_file is not None:
    st.info("Processando arquivo...")

    # ----------------------------
    # Leitura do arquivo
    # ----------------------------
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, encoding="utf-8", sep=None, engine="python")
    else:
        df = pd.read_excel(uploaded_file, header=0)

    # ----------------------------
    # Normalização de nomes de colunas
    # ----------------------------
    df.columns = df.columns.str.strip().str.lower()

    # ----------------------------
    # Mapeamento flexível de colunas
    # ----------------------------
    mapa_colunas = {
        "descricao": ["descripcion", "prod_nombre_original", "nome sku"],
        "contenido": ["contenido", "qtd conteúdo sku"],
        "preco": ["precio kg/lt", "preço convertido kg/lt r$", "preço kg/lt"],
        "categoria": ["est mer 7 (subcategoria)", "nivel1"],
        "vendas": ["imp vta (ult.24 meses)", "vendas em volume"]
    }

    def encontrar_coluna(possiveis):
        for nome in possiveis:
            if nome in df.columns:
                return nome
        return None

    coluna_descricao = encontrar_coluna(mapa_colunas["descricao"])
    coluna_contenido = encontrar_coluna(mapa_colunas["contenido"])
    coluna_preco = encontrar_coluna(mapa_colunas["preco"])
    coluna_categoria = encontrar_coluna(mapa_colunas["categoria"])
    coluna_vendas = encontrar_coluna(mapa_colunas["vendas"])

    # ----------------------------
    # Validação de colunas obrigatórias
    # ----------------------------
    colunas_necessarias = {
        "Descrição": coluna_descricao,
        "Conteúdo": coluna_contenido,
        "Preço": coluna_preco,
        "Categoria": coluna_categoria,
        "Vendas": coluna_vendas
    }

    colunas_faltando = [nome for nome, valor in colunas_necessarias.items() if valor is None]

    if colunas_faltando:
        st.error(
            f"❌ Não foi possível identificar as seguintes colunas no arquivo: "
            f"{', '.join(colunas_faltando)}"
        )
        st.stop()

    # ----------------------------
    # Conversão da coluna de vendas para numérico e filtragem
    # ----------------------------
    df[coluna_vendas] = (
        df[coluna_vendas]
        .astype(str)
        .str.replace(r"[^\d,.-]", "", regex=True)  # remove símbolos estranhos
        .str.replace(".", "", regex=False)         # remove separador de milhar
        .str.replace(",", ".", regex=False)        # converte vírgula decimal em ponto
    )
    df[coluna_vendas] = pd.to_numeric(df[coluna_vendas], errors="coerce")

    df = df[df[coluna_vendas] > 0]

    # ----------------------------
    # Processamento principal
    # ----------------------------
    df[["QtdEmbalagem", "QtdEmbalagemGramas"]] = df[coluna_descricao].apply(
        lambda x: pd.Series(extrair_peso(x))
    )

    def comparar_contenido(qtd_embalagem_gramas, contenido):
        try:
            if pd.isna(qtd_embalagem_gramas) or pd.isna(contenido):
                return "PROBLEMA"

            contenido_val = float(str(contenido).replace(",", "."))
            return "OK" if abs(qtd_embalagem_gramas - contenido_val) < 1 else "PROBLEMA"
        except Exception:
            return "PROBLEMA"

    df["ValidacaoContenido"] = df.apply(
        lambda row: comparar_contenido(row["QtdEmbalagemGramas"], row[coluna_contenido]),
        axis=1
    )

    # ----------------------------
    # Validações de preço
    # ----------------------------
    df["ValidacionPrecio"] = validar_precio_por_categoria(df, coluna_preco, coluna_categoria)
    df["ValidacionPrecioMediana"] = validar_precio_mediana(df, coluna_preco, coluna_categoria)

    # Cria a coluna 'StatusGeral' indicando se há algum problema
    df["StatusGeral"] = df.apply(
        lambda x: "RISCO"
        if (
            x["ValidacaoContenido"] != "OK"
            or x["ValidacionPrecio"] != "OK"
            or x["ValidacionPrecioMediana"] != "OK"
        )
        else "OK",
        axis=1
    )


    st.success("✅ Processamento concluído com sucesso!")


    # ----------------------------
    # Exibição
    # ----------------------------
    st.dataframe(
        df.style.applymap(colorir_valores, subset=["ValidacaoContenido", "ValidacionPrecio", "ValidacionPrecioMediana", "StatusGeral"]),
        use_container_width=True
    )

###########################################################################
### FUNÇÃO PARA GERAR ABA NO EXCEL COM RESUMO DOS PROBLEMAS ENCONTRADOS ###
###########################################################################
    def gerar_resumo(df, coluna_vendas="vendas", col_validacao_contenido="ValidacaoContenido",
                 col_outlier_quartil="ValidacionPrecio", col_outlier_mediana="ValidacionPrecioMediana"):
        
        # Cria um DataFrame resumo com métricas consolidadas:
        # - Quantidade de SKUs/itens
        # - Problemas de conteúdo
        # - Outliers em comum e exclusivos
        # - Totais e % de problemas em valor e volume

        resumo = {}

        # Total de SKUs/itens
        resumo["Qtd Total SKUs/Itens"] = len(df)

        # Problemas de conteúdo
        resumo["Qtd Problemas Conteúdo"] = (df[col_validacao_contenido] == "PROBLEMA").sum()

        # Outliers
        outlier_quartil = df[col_outlier_quartil] == "OUTLIER"
        outlier_mediana = df[col_outlier_mediana] == "OUTLIER_MEDIANA"

        resumo["Qtd Outliers em comum (Quartil e Mediana)"] = (outlier_quartil & outlier_mediana).sum()
        resumo["Qtd Outliers apenas Mediana"] = (outlier_mediana & ~outlier_quartil).sum()
        resumo["Qtd Outliers apenas Quartil"] = (outlier_quartil & ~outlier_mediana).sum()

        # Totais e %
        resumo["Qtd Total SKUs com problema (conteúdo ou preço)"] = (
            (df[col_validacao_contenido] != "PROBLEMA") | outlier_quartil | outlier_mediana
        ).sum()
        resumo["% SKUs com problema"] = resumo["Qtd Total SKUs com problema (conteúdo ou preço)"] / len(df) * 100

        # Volume de vendas
        if coluna_vendas in df.columns:
            resumo["Volume Total Vendas"] = df[coluna_vendas].sum()
            resumo["Volume Vendas com problema"] = df.loc[
                (df[col_validacao_contenido] == "PROBLEMA") | outlier_quartil | outlier_mediana,
                coluna_vendas
            ].sum()
            resumo["% Volume Vendas com problema"] = resumo["Volume Vendas com problema"] / resumo["Volume Total Vendas"] * 100
        else:
            resumo["Volume Total Vendas"] = np.nan
            resumo["Volume Vendas com problema"] = np.nan
            resumo["% Volume Vendas com problema"] = np.nan

        # Retorna como DataFrame (uma linha)
        return pd.DataFrame([resumo])



    # Após processar o DataFrame df
    df_resumo = gerar_resumo(df, coluna_vendas=coluna_vendas)

    st.download_button(
    label="📥 Baixar Excel Processado com Resumo",
    data=to_excel_com_resumo(df, coluna_vendas),
    file_name="dados_processados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

