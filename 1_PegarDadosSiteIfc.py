import re
import unicodedata
import urllib3
from pathlib import Path
from urllib.parse import urljoin, urlparse, unquote

import pandas as pd
import requests
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


# ============================================================
# CONFIGURAÇÕES PRINCIPAIS
# ============================================================
# Se o endereço do site mudar futuramente, altere apenas aqui.
URL_PROFEPT = "https://profept.ifc.edu.br/dissertacoes/"

# Nome da planilha que será gerada.
ARQUIVO_SAIDA = Path("1_DadosExtraidosSiteIfc.xlsx")

# Pasta onde os PDFs serão salvos/reaproveitados.
PASTA_PDFS = Path("pdfs")

# Tempo máximo de espera para carregar a página, em milissegundos.
TIMEOUT_MS = 60000

# Timeout do download de cada PDF, em segundos.
TIMEOUT_DOWNLOAD_PDF = 90

# SSL do site do ProfEPT/IFC pode falhar no Python em alguns ambientes.
# Mantemos verify=False só no download dos PDFs públicos do site.
VERIFICAR_SSL_PDF = False
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# ============================================================
# FUNÇÕES DE APOIO PARA LINKS
# ============================================================

def extrair_links(texto: str) -> tuple[str, list[str]]:
    """
    Extrai links marcados no formato [LINK: url] e devolve:
    - o texto sem os marcadores de link;
    - uma lista com os links encontrados.
    """
    if not texto:
        return "", []

    links = re.findall(r"\[LINK:\s*(.*?)\]", texto)

    texto_limpo = re.sub(r"\[LINK:\s*.*?\]", "", texto)
    texto_limpo = limpar_espacos(texto_limpo)

    return texto_limpo, links


def normalizar_link(link: str) -> str:
    """
    Transforma link relativo em absoluto quando necessário.
    """
    link = limpar_espacos(link)
    if not link:
        return ""
    return urljoin(URL_PROFEPT, link)


def atribuir_links(entrada: dict, links: list[str], campo_origem: str) -> None:
    """
    Distribui os links encontrados para as colunas corretas.

    Regras usadas:
    - Se o campo atual for Produto Educacional, tenta salvar em Link do Produto.
    - Se o campo atual for Dissertação, tenta salvar em Link do PDF.
    - Se o link terminar ou contiver .pdf, salva em Link do PDF.
    - Se não conseguir identificar, preenche primeiro o Link do PDF e depois o Link do Produto.
    """
    campo_origem = (campo_origem or "").lower()

    for link in links:
        link = normalizar_link(link)
        link_lower = link.lower()

        if not link:
            continue

        if "produto" in campo_origem and not entrada["Link do Produto"]:
            entrada["Link do Produto"] = link

        elif "disserta" in campo_origem and not entrada["Link do PDF"]:
            entrada["Link do PDF"] = link

        elif ".pdf" in link_lower and not entrada["Link do PDF"]:
            entrada["Link do PDF"] = link

        elif not entrada["Link do PDF"]:
            entrada["Link do PDF"] = link

        elif not entrada["Link do Produto"]:
            entrada["Link do Produto"] = link


# ============================================================
# FUNÇÕES DE LIMPEZA DE TEXTO
# ============================================================

def limpar_espacos(texto: str) -> str:
    """
    Remove espaços duplicados, quebras estranhas e espaços no começo/fim.
    """
    if not texto:
        return ""

    return re.sub(r"\s+", " ", str(texto)).strip()


def remover_marcadores_linha(linha: str) -> str:
    """
    Remove marcadores comuns de lista, como:
    - texto
    • texto
    * texto
    """
    return re.sub(r"^[-•*]\s*", "", linha).strip()


def remover_acentos(texto: str) -> str:
    """
    Remove acentos para gerar nomes de arquivos seguros.
    """
    texto = unicodedata.normalize("NFD", texto)
    return "".join(c for c in texto if unicodedata.category(c) != "Mn")


def slugificar(texto: str, limite: int = 70) -> str:
    """
    Gera um pedaço de nome seguro para arquivo.
    """
    texto = limpar_espacos(texto)
    texto = remover_acentos(texto).lower()
    texto = re.sub(r"[^a-z0-9]+", "_", texto)
    texto = re.sub(r"_+", "_", texto).strip("_")

    if not texto:
        texto = "trabalho"

    return texto[:limite].strip("_") or "trabalho"


def nome_base_do_link(link: str) -> str:
    """
    Pega um nome aproveitável a partir do final da URL do PDF.
    """
    try:
        caminho = urlparse(link).path
        nome = unquote(Path(caminho).stem)
        return limpar_espacos(nome)
    except Exception:
        return ""


def preparar_html_com_links_marcados(html: str) -> str:
    """
    Recebe o HTML da página e substitui cada link <a> por:
    texto_do_link [LINK: url]
    """
    soup = BeautifulSoup(html, "html.parser")

    for a in soup.find_all("a"):
        href = a.get("href", "").strip()

        if not href:
            continue

        href = normalizar_link(href)
        texto_link = a.get_text(strip=True)
        a.replace_with(f"{texto_link} [LINK: {href}]")

    return soup.get_text(separator="\n")


def forcar_quebras_por_rotulos(texto: str) -> str:
    """
    Insere quebras de linha antes dos rótulos principais.
    """
    padrao_rotulos = (
        r"(Acadêmic[oa]s?:|"
        r"Alunos?:|"
        r"Autor(?:es)?s?:|"
        r"Mestrandos?:|"
        r"Orientador[a]?s?:|"
        r"Dissertaç[ãa]o\s*:|"
        r"Produto Educacional\s*:)"
    )

    return re.sub(padrao_rotulos, r"\n\1", texto, flags=re.IGNORECASE)


# ============================================================
# FUNÇÕES DE IDENTIFICAÇÃO DE CAMPOS
# ============================================================

def criar_entrada_vazia(ano: str) -> dict:
    """
    Cria o modelo padrão de um registro.
    """
    return {
        "Ano": ano,
        "Autor": "",
        "Orientador": "",
        "Título da Dissertação": "",
        "Produto Educacional": "",
        "Link do PDF": "",
        "Link do Produto": "",
    }


def identificar_ano(linha: str) -> str | None:
    """
    Tenta identificar se a linha representa um ano de defesa.
    """
    linha_limpa = limpar_espacos(linha)

    match = re.search(r"\b(20[0-9]{2})\b", linha_limpa)

    if not match:
        return None

    if len(linha_limpa) > 50:
        return None

    if re.search(r"(acadêmic|autor|aluno|dissertaç|produto|orientador)", linha_limpa, re.IGNORECASE):
        return None

    return match.group(1)


def extrair_valor_rotulo(linha: str, padrao: str) -> str:
    """
    Remove o rótulo do começo da linha e devolve apenas o valor.
    """
    valor = re.sub(padrao, "", linha, flags=re.IGNORECASE)
    return limpar_espacos(valor)


def linha_eh_autor(linha: str) -> bool:
    return bool(re.match(r"^(Acadêmic[oa]s?|Alunos?|Autor(?:es)?|Mestrandos?)\s*:", linha, re.IGNORECASE))


def linha_eh_orientador(linha: str) -> bool:
    return bool(re.match(r"^Orientador[a]?s?\s*:", linha, re.IGNORECASE))


def linha_eh_dissertacao(linha: str) -> bool:
    return bool(re.match(r"^Dissertaç[ãa]o\s*:", linha, re.IGNORECASE))


def linha_eh_produto(linha: str) -> bool:
    return bool(re.match(r"^Produto Educacional\s*:", linha, re.IGNORECASE))


# ============================================================
# COLETA DO HTML
# ============================================================

def baixar_html_pagina(url: str) -> str:
    """
    Abre a página com Playwright e retorna o HTML carregado.
    """
    print("Abrindo navegador virtual com Playwright...")

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()

            print("Acessando página do ProfEPT...")
            page.goto(url, wait_until="networkidle", timeout=TIMEOUT_MS)

            html = page.content()
            browser.close()

            return html

    except PlaywrightTimeoutError:
        raise RuntimeError(
            "A página demorou demais para carregar. "
            "Tente novamente ou aumente o TIMEOUT_MS."
        )

    except Exception as erro:
        raise RuntimeError(f"Erro ao acessar a página: {erro}")


# ============================================================
# EXTRAÇÃO DOS REGISTROS
# ============================================================

def processar_linhas(texto: str) -> list[dict]:
    """
    Processa o texto da página linha por linha e monta uma lista de registros.
    """
    linhas = texto.split("\n")

    entradas = []
    entrada_atual = None
    ano_atual = "Desconhecido"
    campo_atual = None

    for linha in linhas:
        linha = limpar_espacos(linha)
        linha = remover_marcadores_linha(linha)

        if not linha:
            continue

        ano_detectado = identificar_ano(linha)
        if ano_detectado:
            ano_atual = ano_detectado
            campo_atual = None
            continue

        if linha_eh_autor(linha):
            if entrada_atual:
                entradas.append(entrada_atual)

            entrada_atual = criar_entrada_vazia(ano_atual)

            valor = extrair_valor_rotulo(
                linha,
                r"^(Acadêmic[oa]s?|Alunos?|Autor(?:es)?|Mestrandos?)\s*:\s*"
            )

            valor, links = extrair_links(valor)

            entrada_atual["Autor"] = valor
            campo_atual = "Autor"

            atribuir_links(entrada_atual, links, campo_atual)
            continue

        if not entrada_atual:
            continue

        if linha_eh_orientador(linha):
            valor = extrair_valor_rotulo(linha, r"^Orientador[a]?s?\s*:\s*")
            valor, links = extrair_links(valor)

            entrada_atual["Orientador"] = valor
            campo_atual = "Orientador"

            atribuir_links(entrada_atual, links, campo_atual)
            continue

        if linha_eh_dissertacao(linha):
            valor = extrair_valor_rotulo(linha, r"^Dissertaç[ãa]o\s*:\s*")
            valor, links = extrair_links(valor)

            entrada_atual["Título da Dissertação"] = valor
            campo_atual = "Título da Dissertação"

            atribuir_links(entrada_atual, links, campo_atual)
            continue

        if linha_eh_produto(linha):
            valor = extrair_valor_rotulo(linha, r"^Produto Educacional\s*:\s*")
            valor, links = extrair_links(valor)

            entrada_atual["Produto Educacional"] = valor
            campo_atual = "Produto Educacional"

            atribuir_links(entrada_atual, links, campo_atual)
            continue

        valor, links = extrair_links(linha)

        if links:
            atribuir_links(entrada_atual, links, campo_atual or "")

        if valor and campo_atual:
            valor_lower = valor.lower()
            textos_genericos = {"clique aqui", "link", "download", "pdf", "acessar"}

            if valor_lower not in textos_genericos:
                entrada_atual[campo_atual] = limpar_espacos(
                    entrada_atual[campo_atual] + " " + valor
                )

    if entrada_atual:
        entradas.append(entrada_atual)

    return entradas


# ============================================================
# DOWNLOAD / REAPROVEITAMENTO DOS PDFs
# ============================================================

def gerar_caminho_pdf_local(indice: int, row: pd.Series) -> Path:
    """
    Gera o nome local do PDF no padrão:
    pdfs/001_titulo_do_trabalho.pdf

    Para reaproveitar os arquivos que você já baixou, antes de baixar o script
    também procura qualquer PDF começando com o mesmo número, por exemplo:
    pdfs/001_*.pdf
    """
    numero = f"{indice + 1:03d}"

    titulo = limpar_espacos(row.get("Título da Dissertação", ""))
    produto = limpar_espacos(row.get("Produto Educacional", ""))
    link_pdf = limpar_espacos(row.get("Link do PDF", ""))

    base_nome = titulo or produto or nome_base_do_link(link_pdf) or "trabalho"
    nome_arquivo = f"{numero}_{slugificar(base_nome)}.pdf"

    return PASTA_PDFS / nome_arquivo


def localizar_pdf_existente(indice: int, caminho_preferido: Path) -> Path | None:
    """
    Procura se o PDF já existe.

    Primeiro testa o caminho preferido.
    Depois procura qualquer arquivo com o mesmo prefixo numérico:
    001_*.pdf, 002_*.pdf etc.
    """
    if caminho_preferido.exists() and caminho_preferido.stat().st_size > 1000:
        return caminho_preferido

    prefixo = f"{indice + 1:03d}_"
    candidatos = sorted(PASTA_PDFS.glob(f"{prefixo}*.pdf"))

    for candidato in candidatos:
        if candidato.exists() and candidato.stat().st_size > 1000:
            return candidato

    return None


def baixar_pdf_para_arquivo(link_pdf: str, caminho_destino: Path) -> None:
    """
    Baixa o PDF para a pasta pdfs/.

    Usa verify=False porque o certificado do site do ProfEPT/IFC pode falhar
    no Python em alguns ambientes, mesmo abrindo normal pelo navegador.
    """
    url = limpar_espacos(link_pdf)
    if not url:
        raise ValueError("Link do PDF vazio")

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (X11; Linux x86_64; rv:120.0) "
            "Gecko/20100101 Firefox/120.0"
        ),
        "Accept": "application/pdf,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8",
        "Referer": URL_PROFEPT,
        "Connection": "keep-alive",
    }

    caminho_temporario = caminho_destino.with_suffix(".tmp")
    caminho_temporario.unlink(missing_ok=True)

    with requests.Session() as session:
        resposta = session.get(
            url,
            headers=headers,
            timeout=TIMEOUT_DOWNLOAD_PDF,
            allow_redirects=True,
            stream=True,
            verify=VERIFICAR_SSL_PDF,
        )

        resposta.raise_for_status()

        total_bytes = 0
        with open(caminho_temporario, "wb") as arquivo:
            for parte in resposta.iter_content(chunk_size=1024 * 128):
                if parte:
                    arquivo.write(parte)
                    total_bytes += len(parte)

    if total_bytes < 1000:
        caminho_temporario.unlink(missing_ok=True)
        raise ValueError("Arquivo baixado está muito pequeno ou vazio")

    inicio = caminho_temporario.read_bytes()[:200]
    content_type = resposta.headers.get("Content-Type", "").lower()

    if b"%PDF" not in inicio and "pdf" not in content_type:
        caminho_temporario.unlink(missing_ok=True)
        raise ValueError(
            f"O link não retornou PDF válido. Content-Type={content_type}; Bytes={total_bytes}"
        )

    caminho_temporario.replace(caminho_destino)


def baixar_ou_reaproveitar_pdfs(df: pd.DataFrame) -> pd.DataFrame:
    """
    Para cada linha da planilha:
    - verifica se o PDF já existe pelo nome/prefixo numérico;
    - se existe, reaproveita;
    - se não existe, baixa;
    - registra o resultado em colunas da planilha.
    """
    if df.empty:
        return df

    PASTA_PDFS.mkdir(exist_ok=True)

    colunas = [
        "Arquivo PDF Local",
        "Status Download PDF",
        "Erro Download PDF",
    ]

    for coluna in colunas:
        if coluna not in df.columns:
            df[coluna] = ""

    total = len(df)
    print("\nBaixando/reaproveitando PDFs...")

    for index, row in df.iterrows():
        link_pdf = limpar_espacos(row.get("Link do PDF", ""))
        caminho_preferido = gerar_caminho_pdf_local(index, row)
        existente = localizar_pdf_existente(index, caminho_preferido)

        if existente:
            df.at[index, "Arquivo PDF Local"] = str(existente)
            df.at[index, "Status Download PDF"] = "Sucesso local"
            df.at[index, "Erro Download PDF"] = ""
            print(f"[{index + 1}/{total}] OK local | {existente}")
            continue

        if not link_pdf:
            df.at[index, "Arquivo PDF Local"] = ""
            df.at[index, "Status Download PDF"] = "Sem link"
            df.at[index, "Erro Download PDF"] = "Coluna Link do PDF vazia"
            print(f"[{index + 1}/{total}] Sem link de PDF")
            continue

        try:
            print(f"[{index + 1}/{total}] Baixando PDF: {link_pdf}")
            baixar_pdf_para_arquivo(link_pdf, caminho_preferido)

            df.at[index, "Arquivo PDF Local"] = str(caminho_preferido)
            df.at[index, "Status Download PDF"] = "Sucesso baixado"
            df.at[index, "Erro Download PDF"] = ""
            print(f"   OK baixado | {caminho_preferido}")

        except Exception as erro:
            df.at[index, "Arquivo PDF Local"] = ""
            df.at[index, "Status Download PDF"] = "Erro"
            df.at[index, "Erro Download PDF"] = str(erro)
            print(f"   FALHOU PDF | {erro}")

    return df


# ============================================================
# TRATAMENTO FINAL E EXPORTAÇÃO
# ============================================================

def criar_dataframe(entradas: list[dict]) -> pd.DataFrame:
    """
    Converte a lista de registros em DataFrame e aplica limpezas finais.
    """
    df = pd.DataFrame(entradas)

    if df.empty:
        return df

    for coluna in df.columns:
        if df[coluna].dtype == "object":
            df[coluna] = df[coluna].fillna("").astype(str).map(limpar_espacos)

    df = df.drop_duplicates(
        subset=["Autor", "Título da Dissertação"],
        keep="first"
    ).reset_index(drop=True)

    colunas_ordenadas = [
        "Ano",
        "Autor",
        "Orientador",
        "Título da Dissertação",
        "Produto Educacional",
        "Link do PDF",
        "Link do Produto",
        "Arquivo PDF Local",
        "Status Download PDF",
        "Erro Download PDF",
    ]

    for coluna in colunas_ordenadas:
        if coluna not in df.columns:
            df[coluna] = ""

    df = df[colunas_ordenadas]

    return df


def salvar_planilha(df: pd.DataFrame, caminho_saida: Path) -> None:
    """
    Salva a planilha em Excel.
    """
    df.to_excel(caminho_saida, index=False)


# ============================================================
# FUNÇÃO PRINCIPAL
# ============================================================

def extrair_dados_profept() -> pd.DataFrame:
    """
    Executa o processo completo:
    1. baixa o HTML;
    2. marca os links;
    3. força quebras nos rótulos;
    4. extrai os registros;
    5. monta a planilha;
    6. baixa ou reaproveita os PDFs locais;
    7. salva a planilha final.
    """
    html = baixar_html_pagina(URL_PROFEPT)

    print("Preparando HTML e marcando links...")
    texto_com_links = preparar_html_com_links_marcados(html)

    print("Separando campos principais...")
    texto_separado = forcar_quebras_por_rotulos(texto_com_links)

    print("Extraindo registros...")
    entradas = processar_linhas(texto_separado)

    print(f"{len(entradas)} registros encontrados antes da limpeza.")

    df = criar_dataframe(entradas)

    df = baixar_ou_reaproveitar_pdfs(df)

    salvar_planilha(df, ARQUIVO_SAIDA)

    print(f"\nPlanilha salva com sucesso em: {ARQUIVO_SAIDA}")
    print(f"Total final de registros: {len(df)}")

    if "Status Download PDF" in df.columns:
        print("\nResumo dos PDFs:")
        for status, qtd in df["Status Download PDF"].value_counts().items():
            print(f"- {status}: {qtd}")

    return df


if __name__ == "__main__":
    extrair_dados_profept()
