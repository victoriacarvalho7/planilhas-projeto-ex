import io
from pathlib import Path

import pandas as pd
import streamlit as st


def desmesclar():
    """Gera ``data/resultado.xlsx`` a partir do modelo presente em ``data``.

    A rotina não depende das demais funções ou constantes da aplicação.
    """
    from pathlib import Path

    from openpyxl import load_workbook

    pasta_data = Path(__file__).resolve().parent / "data"
    arquivo_modelo = next(
        (
            arquivo
            for arquivo in pasta_data.glob("*.xlsx")
            if arquivo.name != "resultado.xlsx" and not arquivo.name.startswith("~$")
        ),
        None,
    )

    if arquivo_modelo is None:
        raise FileNotFoundError("Nenhum arquivo .xlsx de modelo foi encontrado em data.")

    planilha = load_workbook(arquivo_modelo)
    primeira_aba = planilha.worksheets[0]

    # O openpyxl desmescla intervalos inteiros; por isso, desmesclamos todos
    # os intervalos que ocupam alguma célula das colunas A ou B. Antes de
    # desmesclar, guardamos o valor da célula principal para repeti-lo em
    # todas as células que pertenciam ao intervalo.
    for intervalo in list(primeira_aba.merged_cells.ranges):
        if intervalo.min_col <= 2:
            valor = primeira_aba.cell(intervalo.min_row, intervalo.min_col).value
            primeira_aba.unmerge_cells(str(intervalo))

            for linha in primeira_aba.iter_rows(
                min_row=intervalo.min_row,
                max_row=intervalo.max_row,
                min_col=intervalo.min_col,
                max_col=intervalo.max_col,
            ):
                for celula in linha:
                    celula.value = valor

    primeira_aba.delete_rows(1, 2)

    arquivo_resultado = pasta_data / "resultado.xlsx"
    planilha.save(arquivo_resultado)

    return arquivo_resultado


desmesclar()


# -----------------------------------------
# CONFIGURAÇÃO DA PÁGINA
# -----------------------------------------
st.set_page_config(
    page_title="Organizador de Projetos",
    page_icon="📊",
    layout="wide"
)


# -----------------------------------------
# CONSTANTES
# -----------------------------------------
# Arquivo de exemplo, usado só quando existe ao lado do código.
# No stlite (navegador) não há __file__ nem disco local, então cai no upload.
try:
    ARQUIVO_PADRAO = Path(__file__).parent / "dados.csv"

except NameError:
    ARQUIVO_PADRAO = Path("dados.csv")

# Colunas esperadas na planilha do projeto.
# "Status_Etapa" é a coluna a mais em relação ao formato anterior.
COL_ETAPA = "Etapa"
COL_STATUS_ETAPA = "Status_Etapa"
COL_ATIVIDADE = "Atividade"
COL_STATUS_ATIVIDADE = "Status atividade"
COL_CORRECAO = "Correção"

# dados.csv vem com separador ";" e encoding cp1252,
# então não dá para confiar nos padrões do pandas.
ENCODINGS = ["utf-8-sig", "cp1252", "latin-1"]
SEPARADORES = [";", ",", "\t"]

# Valores que contam como atividade concluída.
STATUS_CONCLUIDO = {"closed", "concluído", "concluido", "done", "finalizado"}

# Status que a etapa deve ter em cada um dos dois casos.
STATUS_ETAPA_FECHADA = "Closed"
STATUS_ETAPA_ANDAMENTO = "Progress"

# A linha 1 do arquivo é o cabeçalho, então a primeira linha de dados é a 2.
OFFSET_CABECALHO = 2

# Uma cor por bloco de etapa, repetindo se houver mais blocos que cores.
CORES_BLOCOS = [
    "#DCE9F9",  # azul
    "#DDEFDF",  # verde
    "#FAEBD3",  # âmbar
    "#EADFF5",  # roxo
    "#D8EFF1",  # ciano
    "#F7E2E9",  # rosa
]

COR_ERRO = "#FFC1C1"        # fundo só da linha que causa a inconsistência
COR_BORDA_ERRO = "#C0392B"  # contorno da célula a corrigir
COR_TEXTO = "#111111"       # fixa, para o fundo claro valer nos dois temas


# -----------------------------------------
# LEITURA DOS ARQUIVOS
# -----------------------------------------
def ler_bytes(fonte):
    """Devolve o conteúdo bruto de um Path ou de um upload do Streamlit."""

    if isinstance(fonte, Path):
        return fonte.read_bytes()

    fonte.seek(0)

    return fonte.read()


def decodificar(conteudo):
    """Decodifica o conteúdo testando os encodings mais comuns."""

    for encoding in ENCODINGS:

        try:
            return conteudo.decode(encoding)

        except UnicodeDecodeError:
            continue

    # latin-1 aceita qualquer byte, serve de último recurso.
    return conteudo.decode("latin-1")


def ler_csv(texto):
    """Lê o CSV testando os separadores mais comuns."""

    for separador in SEPARADORES:

        df = pd.read_csv(
            io.StringIO(texto),
            sep=separador,
            dtype=str
        )

        # Mais de uma coluna significa que o separador estava certo.
        if len(df.columns) > 1:
            return df

    return pd.read_csv(io.StringIO(texto), dtype=str)


def ler_markdown(texto):
    """Extrai a primeira tabela Markdown encontrada no texto."""

    dados = []

    for linha in texto.splitlines():

        if "|" not in linha:
            continue

        partes = [
            item.strip()
            for item in linha.strip().strip("|").split("|")
        ]

        # Ignora a linha de separação (|---|:--:|)
        if all(set(item) <= {"-", ":"} for item in partes):
            continue

        dados.append(partes)

    if len(dados) < 2:
        return None

    return pd.DataFrame(dados[1:], columns=dados[0])


def carregar(fonte, nome):
    """Carrega qualquer formato suportado em um DataFrame já normalizado."""

    extensao = Path(nome).suffix.lower()
    conteudo = ler_bytes(fonte)

    if extensao in [".xlsx", ".xls"]:
        df = pd.read_excel(io.BytesIO(conteudo), dtype=str)

    elif extensao == ".md":
        df = ler_markdown(decodificar(conteudo))

    else:
        df = ler_csv(decodificar(conteudo))

    if df is None:
        return None

    # Nomes de coluna e células sem espaços sobrando.
    df.columns = [str(coluna).strip() for coluna in df.columns]

    return df.fillna("").apply(
        lambda coluna: coluna.astype(str).str.strip()
    )


def esta_concluido(valor):
    return str(valor).strip().lower() in STATUS_CONCLUIDO


# -----------------------------------------
# VALIDAÇÃO DE CONSISTÊNCIA DAS ETAPAS
# -----------------------------------------
def blocos_de_etapa(df):
    """Numera blocos de linhas consecutivas com a mesma Etapa.

    O número só muda quando o nome da etapa muda, então uma etapa que
    reapareça mais adiante no arquivo forma um bloco separado.
    """

    etapas = df[COL_ETAPA]

    return (etapas != etapas.shift()).cumsum()


def analisar(df):
    """Aponta, em cada bloco, apenas a linha que quebra a regra.

    Regra: se todas as atividades do bloco estiverem concluídas, o bloco
    deve ter Status_Etapa "Closed"; caso contrário, "Progress".

    Quando o bloco está errado por causa de uma atividade pendente, quem
    é marcada é essa atividade — não o bloco inteiro, que apenas repete
    o mesmo Status_Etapa em todas as suas linhas.
    """

    erro = pd.Series(False, index=df.index)
    correcao = pd.Series("", index=df.index, dtype=object)

    for _, bloco in df.groupby(blocos_de_etapa(df), sort=False):

        concluidas = bloco[COL_STATUS_ATIVIDADE].map(esta_concluido)
        etapa = bloco[COL_ETAPA].iloc[0]
        status_etapa = bloco[COL_STATUS_ETAPA].str.strip().str.lower()

        if concluidas.all():

            # Bloco inteiro concluído: quem estiver fora de "Closed" erra.
            fora = status_etapa != STATUS_ETAPA_FECHADA.lower()

            alvos = bloco.index[fora]

            for posicao in alvos:
                atual = bloco.loc[posicao, COL_STATUS_ETAPA] or "(vazio)"
                correcao[posicao] = (
                    f"Status_Etapa deve ser {STATUS_ETAPA_FECHADA} "
                    f"(está {atual}): todas as atividades de "
                    f"{etapa} estão concluídas"
                )

        else:

            # Etapa tem pendência: só é erro se o Status_Etapa disser
            # o contrário. Marca a atividade pendente que causa isso.
            if (status_etapa == STATUS_ETAPA_ANDAMENTO.lower()).all():
                continue

            alvos = bloco.index[~concluidas]

            for posicao in alvos:
                atual = bloco.loc[posicao, COL_STATUS_ETAPA] or "(vazio)"
                pendente = bloco.loc[posicao, COL_STATUS_ATIVIDADE]
                correcao[posicao] = (
                    f"Atividade em {pendente}: Status_Etapa de "
                    f"{etapa} deve ser {STATUS_ETAPA_ANDAMENTO}, "
                    f"não {atual}"
                )

        erro[alvos] = True

    return erro, correcao


# -----------------------------------------
# CORES DA PLANILHA
# -----------------------------------------
def cor_do_bloco(numero):
    return CORES_BLOCOS[(int(numero) - 1) % len(CORES_BLOCOS)]


def estilizar(visao, blocos, erro):
    """Cor suave por bloco; vermelho apenas na linha que quebra a regra."""

    estilos = pd.DataFrame("", index=visao.index, columns=visao.columns)

    for posicao in visao.index:

        errada = bool(erro.get(posicao, False))

        if errada:
            base = (
                f"background-color: {COR_ERRO}; "
                f"color: {COR_TEXTO}; font-weight: 600;"
            )

        else:
            base = (
                f"background-color: {cor_do_bloco(blocos[posicao])}; "
                f"color: {COR_TEXTO};"
            )

        estilos.loc[posicao, :] = base

        if errada and COL_STATUS_ETAPA in visao.columns:

            estilos.loc[posicao, COL_STATUS_ETAPA] = (
                base + f" border: 2px solid {COR_BORDA_ERRO};"
            )

    return estilos


# -----------------------------------------
# TÍTULO
# -----------------------------------------
st.title("📊 Organizador de Tabela de Projeto")
st.write(
    "Faça upload de um arquivo CSV, Excel ou Markdown "
    "para visualizar e conferir a consistência das etapas."
)


# -----------------------------------------
# ESCOLHA DA FONTE DE DADOS
# -----------------------------------------
arquivo = st.file_uploader(
    "Escolha um arquivo do seu computador",
    type=["csv", "xlsx", "xls", "md"],
    accept_multiple_files=False
)

if arquivo is not None:
    fonte, nome = arquivo, arquivo.name

elif ARQUIVO_PADRAO.exists():
    fonte, nome = ARQUIVO_PADRAO, ARQUIVO_PADRAO.name

else:
    st.info(
        "Envie um arquivo acima para começar. "
        "O processamento acontece no seu próprio navegador — "
        "nenhum dado é enviado para servidores."
    )
    st.stop()


# -----------------------------------------
# PROCESSAMENTO
# -----------------------------------------
try:
    df = carregar(fonte, nome)

except Exception as erro_leitura:
    st.error(f"Não foi possível processar o arquivo: {erro_leitura}")
    st.stop()

if df is None or df.empty:
    st.error("Não foi encontrada nenhuma tabela válida no arquivo.")
    st.stop()

pode_validar = all(
    coluna in df.columns
    for coluna in [COL_ETAPA, COL_STATUS_ETAPA, COL_STATUS_ATIVIDADE]
)


# -----------------------------------------
# PLANILHA
# -----------------------------------------
if not pode_validar:

    st.dataframe(df, width="stretch", hide_index=True)
    st.stop()

blocos = blocos_de_etapa(df)
erro, correcao = analisar(df)

if erro.any():

    linhas = ", ".join(
        str(posicao + OFFSET_CABECALHO) for posicao in df.index[erro]
    )
    st.error(f"Inconsistência na(s) linha(s) {linhas} do arquivo.")

else:

    st.success("Nenhuma inconsistência encontrada.")

visao = df.copy()
visao[COL_CORRECAO] = correcao

st.dataframe(
    visao.style.apply(estilizar, blocos=blocos, erro=erro, axis=None),
    width="stretch",
    hide_index=True,
    column_config={
        COL_CORRECAO: st.column_config.TextColumn(
            COL_CORRECAO,
            width="large"
        )
    }
)

st.download_button(
    label="📥 Baixar planilha conferida",
    data=visao.to_csv(index=False, sep=";").encode("utf-8-sig"),
    file_name="projeto_conferido.csv",
    mime="text/csv"
)
