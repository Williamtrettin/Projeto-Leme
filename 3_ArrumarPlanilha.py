import re
import unicodedata
from pathlib import Path

import pandas as pd

# ============================================================
# PROJETO LEME/EPT - LIMPEZA DA PLANILHA dadosLeme.xlsx
#
# Este script:
# - lê dadosLeme.xlsx;
# - padroniza Orientador em Dr./Dra.;
# - limpa textos longos, principalmente Resumo Sintetizado;
# - melhora títulos/produtos em caixa alta sem apagar informações;
# - salva dadosLeme_arrumado.xlsx;
# - cria uma aba "Alterações" para conferir o que mudou.
# ============================================================

ARQUIVO_ENTRADA = Path("dadosLeme.xlsx")
ARQUIVO_SAIDA = Path("dadosLeme_arrumado.xlsx")

GENERO_MANUAL = {
    # Masculinos
    "reginaldo": "M", "joao": "M", "jose": "M", "carlos": "M", "paulo": "M",
    "marcos": "M", "ricardo": "M", "roberto": "M", "fernando": "M", "antonio": "M",
    "luiz": "M", "luis": "M", "pedro": "M", "rafael": "M", "daniel": "M",
    "fabio": "M", "cassio": "M", "william": "M", "jorge": "M", "andre": "M",
    "alexandre": "M", "renato": "M", "marcelo": "M", "mauricio": "M", "eder": "M",
    "edson": "M", "sergio": "M", "flavio": "M", "adriano": "M", "gustavo": "M",
    "eduardo": "M", "leandro": "M", "gilberto": "M", "sidnei": "M", "bruno": "M",
    "diego": "M", "tiago": "M", "thiago": "M", "lucas": "M", "felipe": "M",
    "francisco": "M", "jairo": "M", "ivandro": "M", "julio": "M",

    # Femininos
    "maria": "F", "ana": "F", "julia": "F", "juliana": "F", "patricia": "F",
    "claudia": "F", "cristina": "F", "fernanda": "F", "marcia": "F", "luciana": "F",
    "carla": "F", "camila": "F", "aline": "F", "eliane": "F", "renata": "F",
    "sandra": "F", "silvia": "F", "simone": "F", "leticia": "F", "vanessa": "F",
    "fatima": "F", "priscila": "F", "juliane": "F", "denise": "F", "helena": "F",
    "amanda": "F", "daniela": "F", "monica": "F", "rose": "F", "rosangela": "F",
    "cintia": "F", "viviane": "F", "luciane": "F", "karina": "F", "cristiane": "F",
    "adriana": "F", "clarice": "F", "michele": "F", "angelita": "F", "karen": "F",
    "tatiana": "F", "solange": "F", "elisabeth": "F", "elizabeth": "F", "raquel": "F",
    "miriam": "F", "mirela": "F", "edilene": "F", "elaine": "F", "vera": "F",
}

PREPOSICOES = {"de", "da", "do", "dos", "das", "e", "di", "du", "del", "della", "la", "le", "no", "na", "nos", "nas", "ao", "aos", "à", "às", "em", "com", "como", "para", "por", "sobre", "um", "uma", "o", "a", "os", "as"}
SIGLAS_COMUNS = {"EPT", "IFC", "PPC", "PPCS", "EJA", "PROEJA", "EMI", "PCD", "PCDS"}


def texto_vazio(texto) -> bool:
    if pd.isna(texto):
        return True
    texto = str(texto).strip()
    return texto == "" or texto.lower() in {"nan", "none", "null"}


def limpar_espacos(texto) -> str:
    if texto_vazio(texto):
        return ""
    texto = str(texto)
    texto = re.sub(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]", " ", texto)
    texto = re.sub(r"[\n\r\t]+", " ", texto)
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()


def remover_acentos(texto) -> str:
    texto = str(texto)
    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


def detectar_titulo_genero(original) -> str:
    texto = limpar_espacos(original).lower()
    if re.search(r"\b(dra|drª|profa|profª|professora)\b\.?,?", texto):
        return "F"
    if re.search(r"\b(dr|prof|professor)\b\.?,?", texto):
        return "M"
    return ""


def remover_titulos(nome) -> str:
    if texto_vazio(nome):
        return ""
    nome = limpar_espacos(nome)
    nome = nome.replace(".:", " ").replace(":", " ")

    padrao = (
        r"^(?:dr|dra|drª|dr\s*\(a\)|prof|profa|profª|professor|professora|"
        r"me|ma|msc|ms|mestre|mestra|phd)\.?\s+"
    )

    while re.match(padrao, nome, flags=re.IGNORECASE):
        nome = re.sub(padrao, "", nome, flags=re.IGNORECASE).strip()

    nome = re.sub(r"^(?:dr|dra|drª|prof|profa|profª)\.?\s*", "", nome, flags=re.IGNORECASE).strip()
    return nome.strip(" .,-;:")


def formatar_nome(nome) -> str:
    if texto_vazio(nome):
        return ""

    nome = limpar_espacos(nome).lower()
    partes_formatadas = []

    for parte in nome.split():
        if parte in PREPOSICOES:
            partes_formatadas.append(parte)
        elif "-" in parte:
            partes_formatadas.append(
                "-".join(sub.capitalize() if sub not in PREPOSICOES else sub for sub in parte.split("-"))
            )
        else:
            partes_formatadas.append(parte.capitalize())

    return " ".join(partes_formatadas)


def identificar_genero(nome_base, genero_por_titulo="") -> str:
    if genero_por_titulo:
        return genero_por_titulo

    nome_sem_acento = remover_acentos(nome_base).lower().strip()
    if not nome_sem_acento:
        return "M"

    primeiro_nome = nome_sem_acento.split()[0]

    if primeiro_nome in GENERO_MANUAL:
        return GENERO_MANUAL[primeiro_nome]

    if primeiro_nome.endswith("a"):
        return "F"

    return "M"


def padronizar_orientador(nome) -> str:
    if texto_vazio(nome):
        return "---"

    original = limpar_espacos(nome)
    if original in {"---", "-", "—"}:
        return "---"

    genero_por_titulo = detectar_titulo_genero(original)
    nome_base = remover_titulos(original)

    if texto_vazio(nome_base):
        return "---"

    genero = identificar_genero(nome_base, genero_por_titulo)
    nome_formatado = formatar_nome(nome_base)

    return f"Dra. {nome_formatado}" if genero == "F" else f"Dr. {nome_formatado}"


def limpar_texto_longo(texto) -> str:
    texto = limpar_espacos(texto)
    texto = re.sub(r"\s+([,.;:!?])", r"\1", texto)
    texto = re.sub(r"([.!?]){3,}", r"\1", texto)
    return texto


def formatar_titulo_caixa_alta(texto) -> str:
    texto = limpar_texto_longo(texto)
    if not texto:
        return texto

    letras = [c for c in texto if c.isalpha()]
    if not letras:
        return texto

    percentual_maiusculo = sum(1 for c in letras if c.isupper()) / len(letras)
    if percentual_maiusculo < 0.85:
        return texto

    palavras = []
    for palavra in texto.lower().split():
        limpa = palavra.strip(".,;:!?()[]{}")
        if limpa in PREPOSICOES:
            palavras.append(palavra.lower())
        elif limpa.upper() in SIGLAS_COMUNS:
            palavras.append(palavra.upper())
        elif "-" in limpa and all(parte.upper() in SIGLAS_COMUNS for parte in limpa.split("-") if parte):
            palavras.append("-".join(parte.upper() for parte in limpa.split("-")))
        else:
            palavras.append(palavra.capitalize())

    texto = " ".join(palavras)
    if texto:
        texto = texto[0].upper() + texto[1:]
    texto = texto.replace("Ppc´s", "PPCs").replace("Ppc’s", "PPCs")
    texto = texto.replace("Ifc", "IFC").replace("Ept", "EPT").replace("Eja", "EJA")
    texto = texto.replace("Proeja", "PROEJA").replace("Emi", "EMI")
    return texto


def limpar_planilha(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    df_limpo = df.copy()
    alteracoes = []

    def registrar(indice, coluna, antes, depois):
        antes_l = limpar_espacos(antes)
        depois_l = limpar_espacos(depois)
        if antes_l != depois_l:
            alteracoes.append({
                "Linha original": indice + 2,
                "Coluna": coluna,
                "Valor antes": antes_l,
                "Valor depois": depois_l,
            })

    if "Orientador" in df_limpo.columns:
        for i, valor in df_limpo["Orientador"].items():
            novo = padronizar_orientador(valor)
            registrar(i, "Orientador", valor, novo)
            df_limpo.at[i, "Orientador"] = novo

    for coluna in [
        "Resumo Sintetizado",
        "Observações",
        "Justificativa Linha/Macro",
        "Trecho onde aparece Linha/Macroprojeto",
        "Finalidade do Produto",
        "Público-alvo",
        "Termos Detectados Linha/Macro",
        "Termos Detectados Base",
    ]:
        if coluna in df_limpo.columns:
            for i, valor in df_limpo[coluna].items():
                novo = limpar_texto_longo(valor)
                registrar(i, coluna, valor, novo)
                df_limpo.at[i, coluna] = novo

    for coluna in ["Título da Dissertação", "Produto Educacional"]:
        if coluna in df_limpo.columns:
            for i, valor in df_limpo[coluna].items():
                novo = formatar_titulo_caixa_alta(valor)
                registrar(i, coluna, valor, novo)
                df_limpo.at[i, coluna] = novo

    if "Autor" in df_limpo.columns:
        for i, valor in df_limpo["Autor"].items():
            novo = formatar_nome(remover_titulos(valor)) if not texto_vazio(valor) else ""
            registrar(i, "Autor", valor, novo)
            df_limpo.at[i, "Autor"] = novo

    df_alteracoes = pd.DataFrame(alteracoes)
    return df_limpo, df_alteracoes


def main() -> None:
    try:
        if not ARQUIVO_ENTRADA.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {ARQUIVO_ENTRADA}")

        print(f"Lendo arquivo: {ARQUIVO_ENTRADA}")
        df = pd.read_excel(ARQUIVO_ENTRADA)
        df_limpo, df_alteracoes = limpar_planilha(df)

        with pd.ExcelWriter(ARQUIVO_SAIDA, engine="openpyxl") as writer:
            df_limpo.to_excel(writer, index=False, sheet_name="Planilha1")
            df_alteracoes.to_excel(writer, index=False, sheet_name="Alterações")

        print("\nLimpeza concluída com sucesso.")
        print(f"Arquivo gerado: {ARQUIVO_SAIDA}")
        print(f"Registros: {len(df_limpo)}")
        print(f"Alterações registradas: {len(df_alteracoes)}")

    except PermissionError:
        print("\nErro: não foi possível salvar o arquivo.")
        print("Feche a planilha no Excel/LibreOffice e rode o script novamente.")

    except Exception as erro:
        print("\nOcorreu um erro durante a limpeza da planilha.")
        print(f"Detalhes: {erro}")


if __name__ == "__main__":
    main()
