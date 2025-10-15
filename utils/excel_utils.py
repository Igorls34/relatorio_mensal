import pandas as pd
import warnings


def ler_planilha(path_excel):
    """Lê a planilha Excel e normaliza as colunas.

    Observações:
    - Suprime um UserWarning conhecido do openpyxl quando a planilha não contém um estilo padrão.
    - Levanta RuntimeError com mensagem clara em caso de falha na leitura.
    """
    try:
        with warnings.catch_warnings():
            # openpyxl às vezes emite: "Workbook contains no default style, apply openpyxl's default"
            warnings.filterwarnings("ignore", message="Workbook contains no default style")
            df = pd.read_excel(path_excel)
    except Exception as e:
        raise RuntimeError(f"Erro ao ler o Excel '{path_excel}': {e}")

    # Normaliza nomes de colunas para facilitar buscas (caso-insensível e sem espaços)
    df.columns = [c.strip().upper() for c in df.columns]

    # Valida se a coluna obrigatória "RECADOS" existe
    if "RECADOS" not in df.columns:
        raise KeyError(f"A planilha não possui a coluna 'RECADOS'. Colunas encontradas: {list(df.columns)}")

    # Converte a coluna de data e remove linhas sem data válida
    df["FIM CONTRATO"] = pd.to_datetime(df["FIM CONTRATO"], errors="coerce")
    df = df.dropna(subset=["FIM CONTRATO"])
    df["MES_NUM"] = df["FIM CONTRATO"].dt.month
    df["ANO"] = df["FIM CONTRATO"].dt.year
    df["MES_NOME"] = df["FIM CONTRATO"].dt.strftime("%B").str.capitalize()
    # Normaliza o campo RECADOS para facilitar comparação de strings
    df["RECADOS"] = df["RECADOS"].astype(str).str.upper().fillna("")

    return df


def aplicar_filtros(df, ano_filtrado, mes_filtrado, tipo_filtro):
    """Aplica os filtros de mês, ano e tipo de recado."""
    meses = {
        "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4,
        "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8,
        "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
    }

    # Filtro pelo tipo de recado
    if tipo_filtro == "Compra de Crédito":
        df = df[df["RECADOS"].str.startswith("COMPRA DE CREDITO")]
    elif tipo_filtro == "Outros":
        df = df[~df["RECADOS"].str.startswith("COMPRA DE CREDITO")]

    # Filtro por ano e mês
    if ano_filtrado != "Todos":
        df = df[df["ANO"] == int(ano_filtrado)]
    if mes_filtrado != "Todos":
        df = df[df["MES_NUM"] == meses[mes_filtrado]]

    return df.sort_values(by=["ANO", "MES_NUM"])
