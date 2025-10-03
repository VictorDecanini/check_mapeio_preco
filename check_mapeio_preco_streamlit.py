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
Para o funcionamento correto da ferramente, são necessárias as colunas exatamente com esses nomes:
- `Descripcion`
- `Contenido`
- `Precio KG/LT`
- `Est Mer 7 (Subcategoria)`

Esta ferramenta permite:
- Validar a quantidade de embalagem (`QtdEmbalagem` e `QtdEmbalagemGramas`)
- Validar o conteúdo (`ValidacaoContenido`)
- Detectar preços fora do padrão por subcategoria (`ValidacionPrecio`)
- Baixar o Excel processado
""")

# ----------------------------
# Funções auxiliares
# ----------------------------
def extrair_peso(texto):
    if pd.isna(texto):
        return None, None
    texto = str(texto)
    unidades_intermed = r"(?:UN|UNID|CJ|CX|DS|PCT|FD|SC)?"
    unidade_final = r"(kg|g|gr)"

    match_multi = re.search(rf"((?:\d+\s*{unidades_intermed}\s*[xX]\s*)+\d+[.,]?\d*\s*{unidade_final})", texto, re.IGNORECASE)
    if match_multi:
        bloco = match_multi.group(1)
        unidade = match_multi.group(len(match_multi.groups())).lower()
        numeros = [float(n.replace(",", ".")) for n in re.findall(r"\d+[.,]?\d*", bloco)]
        multiplicadores = numeros[:-1] if len(numeros) > 1 else []
        peso = numeros[-1]
        if unidade in ["kg"]:
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
        if unidade == "kg":
            valor *= 1000
        return match_3d.group(0), int(n1 * n2 * valor)

    match_2d = re.search(rf"(\d+)\s*[xX]\s*(\d+[.,]?\d*)\s*{unidade_final}\b", texto, re.IGNORECASE)
    if match_2d:
        n1 = int(match_2d.group(1))
        valor = float(match_2d.group(2).replace(",", "."))
        unidade = match_2d.group(3).lower()
        if unidade == "kg":
            valor *= 1000
        return match_2d.group(0), int(n1 * valor)

    match = re.search(rf"(\d+[.,]?\d*)\s*{unidade_final}\b", texto, re.IGNORECASE)
    if match:
        valor = float(match.group(1).replace(",", "."))
        unidade = match.group(2).lower()
        if unidade == "kg":
            valor *= 1000
        return match.group(0), int(valor)

    return None, None

def validar_precio_por_categoria(df, coluna_preco, coluna_categoria):
    def marcar_outliers(grupo):
        q1 = grupo[coluna_preco].quantile(0.15)
        q3 = grupo[coluna_preco].quantile(0.85)
        iqr = q3 - q1
        limite_inferior = q1 - 1.5 * iqr
        limite_superior = q3 + 1.5 * iqr
        return grupo[coluna_preco].apply(
            lambda x: "OK" if limite_inferior <= x <= limite_superior else "OUTLIER"
        )
    return df.groupby(coluna_categoria, group_keys=False).apply(marcar_outliers)

def to_excel(df):
    output = BytesIO()
    df.to_excel(output, index=False)
    return output.getvalue()

def colorir_valores(val):
    """Destaca PROBLEMA e OUTLIER em cores claras"""
    if val == "PROBLEMA":
        color = 'background-color: #ffeb99'  # amarelo claro
    elif val == "OUTLIER":
        color = 'background-color: #ffcccc'  # vermelho clarinho
    else:
        color = ''
    return color

# ----------------------------
# Upload do arquivo
# ----------------------------
uploaded_file = st.file_uploader("Escolha o arquivo Excel", type=["xlsx"])
if uploaded_file is not None:
    st.info("Processando arquivo...")
    df = pd.read_excel(uploaded_file)

    # Ajuste das colunas utilizadas
    coluna_descricao = "Descripcion"
    coluna_contenido = "Contenido"
    coluna_preco = "Precio KG/LT"
    coluna_categoria = "Est Mer 7 (Subcategoria)"

    # ----------------------------
    # Processamento
    # ----------------------------
    df[["QtdEmbalagem", "QtdEmbalagemGramas"]] = df[coluna_descricao].apply(
        lambda x: pd.Series(extrair_peso(x))
    )

    df["ValidacaoContenido"] = df.apply(
        lambda row: "OK" if pd.notna(row["QtdEmbalagemGramas"]) 
                              and pd.notna(row[coluna_contenido]) 
                              and int(row["QtdEmbalagemGramas"]) == int(row[coluna_contenido]) 
                    else "PROBLEMA",
        axis=1
    )

    df["ValidacionPrecio"] = validar_precio_por_categoria(df, coluna_preco, coluna_categoria)

    st.success("Processamento concluído!")

    # ----------------------------
    # Exibição da tabela com cores
    # ----------------------------
    st.dataframe(
        df.style.applymap(colorir_valores, subset=["ValidacaoContenido", "ValidacionPrecio"]),
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