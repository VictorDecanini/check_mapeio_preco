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
# Detectar nome flexível das colunas
# ----------------------------
def encontrar_coluna(df, opcoes):
    for opcao in opcoes:
        for col in df.columns:
            if opcao.lower() in col.lower():
                return col
    return None

# ----------------------------
# Validador de preço por categoria (IQR 5-95%)
# ----------------------------
def validar_precio_por_categoria(df, coluna_preco, coluna_categoria):
    def marcar_outliers(grupo):
        limite_inferior = grupo.quantile(0.05)
        limite_superior = grupo.quantile(0.95)
        return grupo.apply(lambda x: "OK" if limite_inferior <= x <= limite_superior else "OUTLIER")
    return df.groupby(coluna_categoria)[coluna_preco].transform(marcar_outliers)

# ----------------------------
# Novo validador: outliers com base na mediana (3x acima ou 1/3 abaixo)
# ----------------------------
def validar_precio_mediana(df, coluna_preco, coluna_categoria):
    def marcar_por_mediana(grupo):
        mediana = grupo.median()
        limite_inferior = mediana / 3
        limite_superior = mediana * 3
        return grupo.apply(lambda x: "OK" if limite_inferior <= x <= limite_superior else "OUTLIER_MEDIANA")
    return df.groupby(coluna_categoria)[coluna_preco].transform(marcar_por_mediana)

# ----------------------------
# Exportar para Excel
# ----------------------------
def to_excel(df):
    output = BytesIO()
    df.to_excel(output, index=False)
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

    # Detecta tipo de arquivo
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, encoding="utf-8", sep=None, engine="python")
    else:
        df = pd.read_excel(uploaded_file, header=0)

    df.columns = df.columns.str.strip()

    # ----------------------------
    # Identifica colunas de forma flexível
    # ----------------------------
    coluna_vendas = encontrar_coluna(df, ["Imp Vta (Ult.24 Meses)", "Vendas em volume"])
    coluna_descricao = encontrar_coluna(df, ["Descripcion", "PROD_NOMBRE_ORIGINAL", "Nome SKU"])
    coluna_contenido = encontrar_coluna(df, ["Contenido", "Qtd Conteúdo SKU"])
    coluna_preco = encontrar_coluna(df, ["Precio KG/LT", "Preço convertido kg/lt R$", "Preço kg/lt"])
    coluna_categoria = encontrar_coluna(df, ["Est Mer 7 (Subcategoria)", "NIVEL1"])

    if not all([coluna_descricao, coluna_contenido, coluna_preco, coluna_categoria]):
        st.error("Não foi possível identificar todas as colunas necessárias. Verifique os nomes.")
        st.stop()

    # ----------------------------
    # Filtros e processamento
    # ----------------------------
    if coluna_vendas in df.columns:
        # Converte para número de forma segura (erros viram NaN)
        df[coluna_vendas] = pd.to_numeric(df[coluna_vendas], errors="coerce")
        df = df[df[coluna_vendas] > 0]

    df[["QtdEmbalagem", "QtdEmbalagemGramas"]] = df[coluna_descricao].apply(
        lambda x: pd.Series(extrair_peso(x))
    )

    def comparar_contenido(qtd_embalagem_gramas, contenido):
        try:
            if pd.isna(qtd_embalagem_gramas) or pd.isna(contenido):
                return "PROBLEMA"

            # Converte conteúdo para número (float) independentemente do tipo original
            contenido_val = float(str(contenido).replace(",", "."))
            return "OK" if abs(qtd_embalagem_gramas - contenido_val) < 1 else "PROBLEMA"
        except Exception:
            return "PROBLEMA"

    df["ValidacaoContenido"] = df.apply(
        lambda row: comparar_contenido(row["QtdEmbalagemGramas"], row[coluna_contenido]),
        axis=1
    )

    df["ValidacionPrecio"] = validar_precio_por_categoria(df, coluna_preco, coluna_categoria)
    df["ValidacionPrecioMediana"] = validar_precio_mediana(df, coluna_preco, coluna_categoria)

    st.success("✅ Processamento concluído!")

    # ----------------------------
    # Exibição
    # ----------------------------
    st.dataframe(
        df.style.applymap(colorir_valores, subset=["ValidacaoContenido", "ValidacionPrecio", "ValidacionPrecioMediana"]),
        use_container_width=True
    )

    # ----------------------------
    # Download
    # ----------------------------
    st.download_button(
        label="📥 Baixar Excel Processado",
        data=to_excel(df),
        file_name="dados_processados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
