import pandas as pd
from unidecode import unidecode

# =============================
# CONFIGURAÇÃO DOS ARQUIVOS E ABAS
# =============================
BASE_INTERNA = r"C:\Users\e_andressasam\Downloads\comparar\BASE_Interna_Conferida_2025.xlsx"
ABA_BASE = "Repasses Vitrine"

CREDENCIADOS = r"C:\Users\e_andressasam\Downloads\comparar\Planlilha_Consulta_Repasses_Credenciados_2025.xlsx"
ABA_CRED = "Repasses_Vitrine_Credenciados"

SAIDA = r"C:\Users\e_andressasam\Downloads\comparar\Relatorio_Correcao.xlsx"

COL_COD = "Código"
COL_SOLUCAO = "Solução"
COL_CARGA = "Carga horária"

# =============================
# FUNÇÃO DE NORMALIZAÇÃO
# =============================
def normalizar(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).strip().lower()
    return unidecode(texto)

# =============================
# CARREGAR PLANILHAS
# =============================
def carregar_planilha(caminho, aba):
    df = pd.read_excel(caminho, sheet_name=aba)
    df = df[df[COL_COD].notna()]
    df["cod_norm"] = df[COL_COD].apply(normalizar)
    return df

df_base = carregar_planilha(BASE_INTERNA, ABA_BASE)
df_cred = carregar_planilha(CREDENCIADOS, ABA_CRED)

# =============================
# COMPARAÇÃO LADO A LADO
# =============================
resultado = []

for _, row_base in df_base.iterrows():
    codigo = row_base["cod_norm"]
    alvo = df_cred[df_cred["cod_norm"] == codigo]

    if alvo.empty:
        resultado.append([
            row_base[COL_COD],
            row_base[COL_SOLUCAO],
            row_base[COL_CARGA],
            "", "",
            "Código ausente na planilha credenciada"
        ])
        continue

    alvo = alvo.iloc[0]

    dif_solucao = normalizar(row_base[COL_SOLUCAO]) != normalizar(alvo[COL_SOLUCAO])
    dif_carga = normalizar(row_base[COL_CARGA]) != normalizar(alvo[COL_CARGA])

    if dif_solucao or dif_carga:
        observacao = []
        if dif_solucao:
            observacao.append(f"Solução diferente: Base='{row_base[COL_SOLUCAO]}', Credenciado='{alvo[COL_SOLUCAO]}'")
        if dif_carga:
            observacao.append(f"Carga horária diferente: Base='{row_base[COL_CARGA]}', Credenciado='{alvo[COL_CARGA]}'")

        resultado.append([
            row_base[COL_COD],
            row_base[COL_SOLUCAO],
            row_base[COL_CARGA],
            alvo[COL_SOLUCAO],
            alvo[COL_CARGA],
            "; ".join(observacao)
        ])

# =============================
# DETECTAR CÓDIGOS NOVOS NA PLANILHA DE CREDENCIADOS
# =============================
codigos_base = set(df_base["cod_norm"])
codigos_cred = set(df_cred["cod_norm"])
novos = codigos_cred - codigos_base

for codigo in novos:
    linha = df_cred[df_cred["cod_norm"] == codigo].iloc[0]
    resultado.append([
        linha[COL_COD],
        linha[COL_SOLUCAO],
        linha[COL_CARGA],
        "", "",
        "Novo código na planilha credenciada"
    ])

# =============================
# GERAR PLANILHA DE CORREÇÃO
# =============================
df_saida = pd.DataFrame(resultado, columns=[
    "Código",
    "Solução Base",
    "Carga horária Base",
    "Solução Credenciado",
    "Carga horária Credenciado",
    "Diferença / Observação"
])

df_saida.to_excel(SAIDA, index=False)
print(f"✔ Relatório gerado com sucesso! {SAIDA}")
