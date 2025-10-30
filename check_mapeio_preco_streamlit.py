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
- **Descrição:** `Descripcion`, `PROD_NOMBRE_ORIGINAL`, `Nome SKU`
- **Contenido:** `Contenido`, `Qtd Conteúdo SKU`
- **Preço Kg/Lt:** `Precio KG/LT`, `Preço convertido kg/lt R$`, `Preço kg/lt`
- **Subcategoria:** `Est Mer 7 (Subcategoria)`, `NIVEL1`
- **Venda em volume:** `Imp Vta (Ult.24 Meses)`, `Vendas em volume`

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
import re
import pandas as pd

def extrair_peso(texto):
    if pd.isna(texto):
        return None, None

    texto = str(texto).upper().strip()

    # -------------------------------------------------
    # 1️⃣ Tenta capturar formatos de PESO / VOLUME (Kg, L, etc)
    # -------------------------------------------------
    unidades_intermed = r"(?:UN|UNID|CJ|CX|DS|PCT|FD|SC)?"
    unidade_final = r"(KILOS|KILO|KG|G|GR|GRS|GRAMAS|GRAMA|ML|L|LT|LTS|LITROS|LITRO)"
    
    # Casos compostos tipo 3x200G ou 2x500ML
    match_multi = re.search(rf"((?:\d+\s*{unidades_intermed}\s*[xX]\s*)+\d+[.,]?\d*\s*{unidade_final})", texto, re.IGNORECASE)
    if match_multi:
        bloco = match_multi.group(1)
        unidade = match_multi.group(len(match_multi.groups())).lower()
        numeros = [float(n.replace(",", ".")) for n in re.findall(r"\d+[.,]?\d*", bloco)]
        multiplicadores = numeros[:-1] if len(numeros) > 1 else []
        peso = numeros[-1]
        if unidade in ["kg","kilos", "kilo", "lt", "l", "lts", "litros", "litro"]:
            peso *= 1000
        total = peso
        for n in multiplicadores:
            total *= n
        return bloco, int(total)

    # Casos como 3x4x200ML
    match_3d = re.search(rf"(\d+)\s*[xX]\s*(\d+)\s*[xX]\s*(\d+[.,]?\d*)\s*{unidade_final}\b", texto, re.IGNORECASE)
    if match_3d:
        n1 = int(match_3d.group(1))
        n2 = int(match_3d.group(2))
        valor = float(match_3d.group(3).replace(",", "."))
        unidade = match_3d.group(4).lower()
        if unidade in ["kg","kilos", "kilo", "lt", "l", "lts", "litros", "litro"]:
            valor *= 1000
        return match_3d.group(0), int(n1 * n2 * valor)

    # Casos como 3x200ML
    match_2d = re.search(rf"(\d+)\s*[xX]\s*(\d+[.,]?\d*)\s*{unidade_final}\b", texto, re.IGNORECASE)
    if match_2d:
        n1 = int(match_2d.group(1))
        valor = float(match_2d.group(2).replace(",", "."))
        unidade = match_2d.group(3).lower()
        if unidade in ["kg","kilos", "kilo", "lt", "l", "lts", "litros", "litro"]:
            valor *= 1000
        return match_2d.group(0), int(n1 * valor)

    # Casos simples como "200ML", "1L", "500G"
    match = re.search(rf"(\d+[.,]?\d*)\s*{unidade_final}\b", texto, re.IGNORECASE)
    if match:
        valor = float(match.group(1).replace(",", "."))
        unidade = match.group(2).lower()
        if unidade in ["kg","kilos", "kilo", "lt", "l", "lts", "litros", "litro"]:
            valor *= 1000
        return match.group(0), int(valor)

    # -------------------------------------------------
    # 2️⃣ Caso não tenha achado peso/volume → tenta UNIDADES (robusto, cobre C/XX, C/XXxYY e XXxYY)
    # -------------------------------------------------

    # 1) Padrões com número antes do sufixo: "3x12UN", "24 UN", "12UN", "2x24 UN"
    match_un = re.search(
        r"(?:(\d+)\s*[xX]\s*)?(\d+)\s*(?:UN|UNID|UND|UNIDADE|UNIDADES|CJ|CX|PCT|FD|SC)\b",
        texto,
        re.IGNORECASE
    )
    if match_un:
        mult = int(match_un.group(1)) if match_un.group(1) else 1
        qtd = int(match_un.group(2))
        return match_un.group(0), mult * qtd

    # 2) Padrões tipo "C/3X24", "C 2X6", "C.4X12"
    match_c_pack = re.search(r"C[\s./]?(\d+)\s*[xX]\s*(\d+)\b", texto, re.IGNORECASE)
    if match_c_pack:
        mult = int(match_c_pack.group(1))
        qtd = int(match_c_pack.group(2))
        return match_c_pack.group(0), mult * qtd

    # 3) Padrões simples "C/32", "C 32", "C.32", "C32"
    match_c = re.search(r"C[\s./]?(\d{1,4})\b", texto, re.IGNORECASE)
    if match_c:
        return match_c.group(0), int(match_c.group(1))

    # 4) Padrões "3X12", "2X6", "4X24" sem UN no final
    match_pack = re.search(r"(\d+)\s*[xX]\s*(\d+)\b", texto, re.IGNORECASE)
    if match_pack:
        mult = int(match_pack.group(1))
        qtd = int(match_pack.group(2))
        return match_pack.group(0), mult * qtd

    # 5) Fallback: último número do texto (pode capturar casos residuais)
    nums = re.findall(r"\d+", texto)
    if nums:
        last = int(nums[-1])
        if 0 < last <= 10000:
            return str(last), last

    # -------------------------------------------------
    # 3️⃣ Caso nada encontrado
    # -------------------------------------------------
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
        n = len(grupo)
        if n < 1000:
            limite_inferior = grupo.quantile(0.05)
            limite_superior = grupo.quantile(0.95)
        elif n < 2000:
            limite_inferior = grupo.quantile(0.03)
            limite_superior = grupo.quantile(0.97)
        else:
            limite_inferior = grupo.quantile(0.02)
            limite_superior = grupo.quantile(0.98)
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
                                (df["ValidacaoContenido"] != "PROBLEMA") & 
                                (df["ValidacionPrecioMediana"] == "OUTLIER_MEDIANA")).sum()
    outliers_somente_quartil = ((df["ValidacionPrecio"] == "OUTLIER") & 
                                (df["ValidacaoContenido"] != "PROBLEMA") & 
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
            "1. Skus com possíveis problemas de contenido",
            "2. Outliers idenficados através da mediana (3x) e quartil (5%)",
            "3. Outilers apenas mediana (3x)",
            "4. Outliers apenas quartil (5%)",
            "Qtd de SKUs/itens com possíveis problemas",
            "'%' de SKUs/itens com possíveis problemas",
            "Volume de vendas total",
            "Volume de vendas dos skus com possíveis problemas",
            "% Volume de vendas dos skus com possíveis problemas"
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
        worksheet.write("A8", "'%' de SKUs/itens com possíveis problemas", orange_bold_format)
        worksheet.write("B8", df_resumo.loc[6, "Valor"] / 100, percent_format)  # Qtd total problemas

        worksheet.write("A12", "% Volume de vendas dos skus com possíveis problemas", orange_bold_format)
        worksheet.write_number("B12", df_resumo.loc[9, "Valor"] / 100, percent_format)

        # 3️⃣ Remove destaque do “Volume de vendas total” e insere linha em branco antes do “com problema”
        worksheet.write("A9", "", empty_format)
        worksheet.write("B9", "", empty_format)

        # 4️⃣ Título inferior (Top 50 SKUs)
        # worksheet.merge_range("A13:B13", "Top 50 Skus - Share Acumulado", gray_format)

    return output.getvalue()

# ----------------------------
# Upload do arquivo principal
# ----------------------------
uploaded_file = st.file_uploader("Selecione o arquivo Excel ou CSV **Bruto** com a categoria em questão", type=["xlsx", "csv"])

# 🔹 Novo: Upload da base auxiliar
uploaded_aux = st.file_uploader("📎 Selecione a base validadora (para cruzar por EAN)", type=["xlsx", "csv"])

if uploaded_file is not None:
    st.info("Processando arquivo...")

    # ----------------------------
    # Leitura do arquivo principal
    # ----------------------------
    if uploaded_file.name.endswith(".csv"):
        try:
            df = pd.read_csv(uploaded_file, encoding="latin-1", sep=None, engine="python")
        except UnicodeDecodeError:
            df = pd.read_csv(uploaded_file, encoding="utf-8", sep=None, engine="python")
    else:
        df = pd.read_excel(uploaded_file, header=0)

    # 🔹 Novo: leitura da base auxiliar (se enviada)
    # 🔹 Novo: leitura da base auxiliar (se enviada) - lê especificamente a aba "Planilha Validadora"
    df_aux = None
    if uploaded_aux is not None:
        try:
            filename_aux_lower = uploaded_aux.name.lower()
            # Se for CSV, lemos normalmente (CSV não tem sheets)
            if filename_aux_lower.endswith(".csv"):
                df_aux = pd.read_csv(uploaded_aux, encoding="utf-8", sep=None, engine="python")
                st.info("🔁 Base auxiliar lida como CSV (nenhuma aba disponível).")
            else:
                # Tenta ler a aba "Planilha Validadora"
                try:
                    df_aux = pd.read_excel(uploaded_aux, sheet_name="PLANILHA VALIDADORA")
                    st.info('✅ Aba "Planilha Validadora" encontrada e carregada da base auxiliar.')
                except ValueError:
                    # aba não encontrada — fallback para a primeira aba e aviso
                    try:
                        df_aux = pd.read_excel(uploaded_aux, sheet_name=0)
                        st.warning('⚠️ Aba "Planilha Validadora" não encontrada — carregada a primeira aba como fallback.')
                    except Exception as e_inner:
                        st.warning(f"⚠️ Erro ao ler a base auxiliar (fallback): {e_inner}")
                        df_aux = None
        except Exception as e:
            st.warning(f"⚠️ Não foi possível ler a base auxiliar: {e}")
            df_aux = None


    # ----------------------------
    # Normalização de nomes
    # ----------------------------
    df.columns = df.columns.str.strip().str.lower()
    if df_aux is not None:
        df_aux.columns = df_aux.columns.str.strip().str.lower()
    df.columns = df.columns.str.strip().str.lower()

    # ----------------------------
    # Mapeamento flexível de colunas
    # ----------------------------
    mapa_colunas = {
        "descricao": ["descripcion", "prod_nombre_original", "nome sku"],
        "contenido": ["contenido", "qtd conteúdo sku"],
        "preco": ["precio kg/lt", "preço convertido kg/lt r$", "preço kg/lt"],
        "categoria": ["est mer 7 (subcategoria)", "nivel1", "est mer 7 descripcion"],
        "vendas": ["imp vta (ult.24 meses)", "vendas em volume", "imp vta (ult 24meses)"]
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
    # Conversão da coluna de vendas
    # ----------------------------
    df[coluna_vendas] = (
        df[coluna_vendas]
        .astype(str)
        .str.replace(r"[^\d,.-]", "", regex=True)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
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

    # ===================================================
    # 🔹 NOVO BLOCO: Cruzamento com base auxiliar por EAN
    # ===================================================
    df_final = df.copy()

    if df_aux is not None:
        try:
            # Normaliza nome da coluna de código de barras
            possiveis_ean_df = ["codigo barras", "código barras", "ean"]
            possiveis_ean_aux = ["codigo barras", "código barras", "ean", "codigo_barras"]

            def encontrar_ean(df_ref, lista_nomes):
                for nome in lista_nomes:
                    if nome in df_ref.columns:
                        return nome
                return None

            col_ean_df = encontrar_ean(df, possiveis_ean_df)
            col_ean_aux = encontrar_ean(df_aux, possiveis_ean_aux)

            if col_ean_df and col_ean_aux:
                # Seleciona apenas colunas R, W e Y (independente de maiúsculas/minúsculas)
                colunas_aux_interesse = ["analise preço kg/lt", "mapeio pack", "localiza se há conteúdo no descritivo atual", "fórmula dun -> ean"]
                colunas_aux_existentes = [
                    c for c in df_aux.columns if c.lower() in colunas_aux_interesse
                ]

                if colunas_aux_existentes:
                    # Faz o merge mantendo todas as linhas do df principal
                    df_final = df.merge(
                        df_aux[[col_ean_aux] + colunas_aux_existentes],
                        how="left",
                        left_on=col_ean_df,
                        right_on=col_ean_aux,
                        suffixes=("", "_aux")
                    )

                    # 🔹 Garante que a coluna de EAN original da base principal seja preservada
                    if col_ean_df not in df_final.columns:
                        # Se por algum motivo sumiu, recria a partir do right_on
                        df_final[col_ean_df] = df[col_ean_df]

                    # 🔹 Evita perda de EAN quando as colunas tinham o mesmo nome
                    if col_ean_aux in df_final.columns and col_ean_aux != col_ean_df:
                        # Mantém o EAN da base principal, remove apenas o duplicado
                        df_final = df_final.drop(columns=[col_ean_aux])

                    st.success(
                        f"✅ Bases cruzadas com sucesso por '{col_ean_df}'. "
                        f"Colunas adicionadas: {', '.join(colunas_aux_existentes)}"
                    )
                else:
                    st.warning("⚠️ Nenhuma das colunas de interesse foi encontrada na base auxiliar. "
                            "Verifique os nomes: " + ", ".join(colunas_aux_interesse))
            else:
                st.warning("⚠️ Não foi possível localizar a coluna de EAN em uma das bases.")

        except Exception as e:
            st.warning(f"⚠️ Erro ao cruzar as bases: {e}")
            df_final = df


    # ===================================================
    # Resultado final
    # ===================================================
    st.success("✅ Processamento concluído com sucesso!")
    st.dataframe(df_final.head(20))

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
    df_resumo = gerar_resumo(df_final, coluna_vendas=coluna_vendas)

    st.download_button(
    label="📥 Baixar Excel Processado com Resumo",
    data=to_excel_com_resumo(df_final, coluna_vendas),
    file_name="dados_processados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

