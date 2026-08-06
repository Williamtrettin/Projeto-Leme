from pathlib import Path
import json
import re
import unicodedata
import traceback

import pandas as pd

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

# OCR opcional
try:
    import pytesseract
    from pdf2image import convert_from_path
except ImportError:
    pytesseract = None
    convert_from_path = None


# ============================================================
# CONFIGURAÇÕES
# ============================================================

ARQUIVO_PLANILHA_SITE = Path("1_DadosExtraidosSiteIfc.xlsx")

PASTA_ENTRADA = Path("pdfs")
PASTA_SAIDA = Path("lotes_json")

TAMANHO_LOTE = 10

USAR_OCR_SE_NECESSARIO = True

MIN_CARACTERES_TEXTO_NORMAL = 2500

MAX_PAGINAS_OCR = 20

IDIOMA_OCR = "por+eng"

LIMITES = {
    "resumo": 8000,
    "abstract": 4000,
    "introducao": 4000,
    "objetivos": 4000,
    "metodologia": 5000,
    "produto_educacional": 4000,
    "publico_alvo": 2000,
    "finalidade": 2000,
    "base_teorica_ou_referencial": 4000,
    "trecho_relevante": 800,
    "texto_base_para_ia": 14000,
    "texto_limpo_para_classificacao": 12000,
}


# ============================================================
# COLUNAS DA PLANILHA DO SITE
# ============================================================

MAPEAMENTO_SITE = {
    "ano_oficial": [
        "Ano",
        "Ano Limpo",
        "Ano de Defesa",
        "Ano Defesa",
    ],
    "autor_oficial": [
        "Autor",
        "Autora",
        "Discente",
        "Mestrando",
        "Mestranda",
    ],
    "orientador_oficial": [
        "Orientador",
        "Orientadora",
        "Professor Orientador",
        "Professora Orientadora",
    ],
    "titulo_oficial_dissertacao": [
        "Título da Dissertação",
        "Titulo da Dissertação",
        "Título da Dissertacao",
        "Titulo da Dissertacao",
        "Título",
        "Titulo",
        "Nome do Trabalho",
    ],
    "produto_educacional_oficial_site": [
        "Produto Educacional",
        "Produto",
        "Título do Produto",
        "Titulo do Produto",
        "Nome do Produto",
    ],
    "link_pdf_oficial": [
        "Link do PDF",
        "Link PDF",
        "PDF",
        "Link da Dissertação",
        "Link Dissertação",
        "Link da Dissertacao",
    ],
    "link_produto_oficial": [
        "Link do Produto",
        "Link Produto",
        "Link do Produto Educacional",
        "Produto Link",
        "Link PE",
    ],
    "arquivo_pdf_local_site": [
        "Arquivo PDF Local",
        "Nome do Arquivo",
        "Arquivo PDF",
        "PDF Local",
    ],
    "status_download_pdf_site": [
        "Status Download PDF",
        "Status do Download",
        "Status PDF",
    ],
}


# ============================================================
# FUNÇÕES DE LIMPEZA
# ============================================================

def remover_acentos(texto: str) -> str:
    if not texto:
        return ""

    return "".join(
        c for c in unicodedata.normalize("NFD", str(texto))
        if unicodedata.category(c) != "Mn"
    )


def normalizar_para_busca(texto: str) -> str:
    texto = remover_acentos(texto or "")
    texto = texto.lower()
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()


def normalizar_nome_coluna(texto: str) -> str:
    texto = normalizar_para_busca(texto)
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def limpar_valor(valor, padrao=""):
    if pd.isna(valor):
        return padrao

    texto = str(valor).strip()

    if not texto or texto.lower() in {"nan", "none", "null", "nat"}:
        return padrao

    return texto


def limpar_texto(texto: str) -> str:
    if not texto:
        return ""

    texto = str(texto)

    texto = texto.replace("\x00", " ")
    texto = texto.replace("\uf0b7", " ")
    texto = texto.replace("ﬁ", "fi").replace("ﬂ", "fl")

    texto = re.sub(r"(\w)-\s*\n\s*(\w)", r"\1\2", texto)
    texto = re.sub(r"\s*\n\s*", " ", texto)
    texto = re.sub(r"[ \t\r\f\v]+", " ", texto)
    texto = re.sub(r"[•●▪■□◆◇]+", " ", texto)
    texto = re.sub(r"\s+([,.;:!?])", r"\1", texto)

    return texto.strip()


def limitar(texto: str, limite: int) -> str:
    texto = limpar_texto(texto)

    if len(texto) <= limite:
        return texto

    return texto[:limite].rsplit(" ", 1)[0].strip() + "..."


def remover_linhas_repetidas(texto: str) -> str:
    if not texto:
        return ""

    linhas = [linha.strip() for linha in texto.splitlines() if linha.strip()]
    contagem = {}

    for linha in linhas:
        chave = normalizar_para_busca(linha)

        if 5 <= len(chave) <= 140:
            contagem[chave] = contagem.get(chave, 0) + 1

    linhas_filtradas = []

    for linha in linhas:
        chave = normalizar_para_busca(linha)

        if contagem.get(chave, 0) >= 5:
            continue

        if re.fullmatch(r"\d{1,4}", linha):
            continue

        if "http://" in linha.lower() or "https://" in linha.lower():
            continue

        linhas_filtradas.append(linha)

    return "\n".join(linhas_filtradas)


def parece_sumario_ou_lixo(texto: str) -> bool:
    if not texto:
        return False

    texto_norm = normalizar_para_busca(texto)

    sinais_lixo = [
        "sumario",
        "lista de figuras",
        "lista de tabelas",
        "ficha catalografica",
        "comissao examinadora",
        "assinado digitalmente",
        "documentos comprobatorios",
        "documentos comprobatórios",
        "para verificar a autenticidade",
        "agradecimentos",
        "dedicatoria",
        "dedicatória",
        "folha de aprovacao",
        "folha de aprovação",
    ]

    if any(sinal in texto_norm for sinal in sinais_lixo):
        return True

    qtd_pontinhos = texto.count(".....") + texto.count("....")

    if qtd_pontinhos >= 3:
        return True

    return False


def qualidade_campo(texto: str, termos_fortes=None):
    texto = limpar_texto(texto)
    termos_fortes = termos_fortes or []

    if not texto:
        return "vazio", ["campo vazio"]

    motivos = []

    if parece_sumario_ou_lixo(texto):
        motivos.append("parece sumário, capa, banca, agradecimento ou documento administrativo")
        return "descartado", motivos

    if len(texto) < 60:
        motivos.append("trecho muito curto")
        return "baixa", motivos

    texto_norm = normalizar_para_busca(texto)

    achados = []

    for termo in termos_fortes:
        termo_norm = normalizar_para_busca(termo)

        if termo_norm and termo_norm in texto_norm:
            achados.append(termo)

    if achados:
        motivos.append("indícios encontrados: " + ", ".join(achados[:6]))
        return "alta", motivos

    if len(texto) >= 250:
        motivos.append("trecho aproveitável, mas sem marcador forte")
        return "media", motivos

    motivos.append("sem indícios suficientes")
    return "baixa", motivos


def limpar_campo_por_qualidade(texto: str, qualidade: str) -> str:
    if qualidade in {"alta", "media"}:
        return limpar_texto(texto)

    return ""


# ============================================================
# LIMPAR JSONS ANTIGOS
# ============================================================

def limpar_jsons_antigos():
    """
    Apaga os JSONs antigos da pasta lotes_json antes de gerar os novos.
    Só apaga arquivos .json.
    """
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)

    arquivos_json = list(PASTA_SAIDA.glob("*.json"))

    if not arquivos_json:
        print(f"Nenhum JSON antigo encontrado em: {PASTA_SAIDA.resolve()}")
        return

    print(f"Limpando JSONs antigos em: {PASTA_SAIDA.resolve()}")

    for arquivo in arquivos_json:
        try:
            arquivo.unlink()
            print(f"   Removido: {arquivo.name}")
        except Exception as erro:
            print(f"   Não foi possível remover {arquivo.name}: {erro}")


# ============================================================
# LEITURA DA PLANILHA DO SITE
# ============================================================

def encontrar_coluna(df: pd.DataFrame, possiveis_nomes: list[str]):
    indice = {
        normalizar_nome_coluna(coluna): coluna
        for coluna in df.columns
    }

    for nome in possiveis_nomes:
        chave = normalizar_nome_coluna(nome)

        if chave in indice:
            return indice[chave]

    return None


def extrair_numero_inicial(nome: str):
    if not nome:
        return None

    match = re.match(r"^\s*(\d{1,4})[_\-\s]", str(nome))

    if match:
        return int(match.group(1))

    return None


def carregar_planilha_site() -> dict:
    """
    Retorna:
    {
        1: dados_da_linha_1,
        2: dados_da_linha_2,
        ...
    }

    Associação principal:
    001_arquivo.pdf -> linha 1 da planilha
    002_arquivo.pdf -> linha 2 da planilha
    """
    if not ARQUIVO_PLANILHA_SITE.exists():
        print(f"[AVISO] Planilha do site não encontrada: {ARQUIVO_PLANILHA_SITE.resolve()}")
        print("        Os JSONs serão gerados sem dados oficiais do site.")
        return {}

    print(f"Lendo planilha do site: {ARQUIVO_PLANILHA_SITE}")

    df = pd.read_excel(ARQUIVO_PLANILHA_SITE)
    df = df.fillna("")

    colunas_reais = {}

    for campo_final, opcoes in MAPEAMENTO_SITE.items():
        coluna = encontrar_coluna(df, opcoes)
        colunas_reais[campo_final] = coluna

    print("Colunas encontradas na planilha:")

    for campo, coluna in colunas_reais.items():
        print(f"   {campo}: {coluna if coluna else 'NÃO ENCONTRADA'}")

    dados_por_id = {}

    for indice, row in df.iterrows():
        id_linha = indice + 1
        dados = {}

        for campo_final, coluna_real in colunas_reais.items():
            if coluna_real:
                dados[campo_final] = limpar_valor(row.get(coluna_real), "")
            else:
                dados[campo_final] = ""

        dados["id_planilha_site"] = id_linha
        dados["linha_planilha_site"] = id_linha
        dados["fonte_dados_cadastrais"] = "planilha_site"

        dados_por_id[id_linha] = dados

    print(f"Registros carregados da planilha do site: {len(dados_por_id)}")

    return dados_por_id


def obter_dados_site_para_pdf(caminho_pdf: Path, dados_site_por_id: dict) -> tuple[dict, str]:
    numero = extrair_numero_inicial(caminho_pdf.name)

    if numero and numero in dados_site_por_id:
        return dados_site_por_id[numero], "numero_inicial_pdf"

    return {}, "sem_vinculo_site"


# ============================================================
# EXTRAÇÃO DE TEXTO DO PDF
# ============================================================

def extrair_texto_pdf(caminho_pdf: Path) -> str:
    if fitz is None:
        raise ImportError("PyMuPDF/fitz não está instalado.")

    partes = []

    with fitz.open(caminho_pdf) as doc:
        for pagina in doc:
            texto = pagina.get_text("text")

            if texto:
                partes.append(texto)

    texto_final = "\n".join(partes)
    texto_final = remover_linhas_repetidas(texto_final)

    return limpar_texto(texto_final)


def extrair_texto_ocr(caminho_pdf: Path) -> str:
    if pytesseract is None or convert_from_path is None:
        raise ImportError("pytesseract e/ou pdf2image não estão instalados.")

    kwargs = {
        "dpi": 250,
        "first_page": 1,
    }

    if MAX_PAGINAS_OCR is not None:
        kwargs["last_page"] = MAX_PAGINAS_OCR

    imagens = convert_from_path(str(caminho_pdf), **kwargs)

    partes = []

    for indice, imagem in enumerate(imagens, start=1):
        print(f"      OCR página {indice}/{len(imagens)}...")
        texto = pytesseract.image_to_string(imagem, lang=IDIOMA_OCR)
        partes.append(texto)

    texto_final = "\n".join(partes)
    texto_final = remover_linhas_repetidas(texto_final)

    return limpar_texto(texto_final)


# ============================================================
# BUSCA DE SEÇÕES
# ============================================================

def encontrar_posicao_titulo(texto_norm: str, padroes: list[str], inicio: int = 0):
    melhor = None

    for padrao in padroes:
        padrao_norm = normalizar_para_busca(padrao)
        regex = r"(?:^|\s|[0-9]\.?\s*)" + re.escape(padrao_norm) + r"(?:\s|:|\.|$)"
        match = re.search(regex, texto_norm[inicio:])

        if match:
            pos = inicio + match.start()

            if melhor is None or pos < melhor:
                melhor = pos

    return melhor


def extrair_entre_marcadores(
    texto: str,
    inicio_padroes: list[str],
    fim_padroes: list[str],
    limite: int,
    posicao_minima: int = 0
) -> str:
    if not texto:
        return ""

    texto_limpo = limpar_texto(texto)
    texto_norm = normalizar_para_busca(texto_limpo)

    inicio = encontrar_posicao_titulo(texto_norm, inicio_padroes, posicao_minima)

    if inicio is None:
        return ""

    trecho_inicio_norm = texto_norm[inicio:inicio + 300]
    fim_titulo = None

    for padrao in inicio_padroes:
        padrao_norm = normalizar_para_busca(padrao)
        match = re.search(re.escape(padrao_norm), trecho_inicio_norm)

        if match:
            candidato = inicio + match.end()

            if fim_titulo is None or candidato < fim_titulo:
                fim_titulo = candidato

    if fim_titulo is None:
        fim_titulo = inicio

    inicio_recorte = fim_titulo

    while inicio_recorte < len(texto_limpo) and texto_limpo[inicio_recorte] in " :.-–—":
        inicio_recorte += 1

    fim = None

    for fim_padrao in fim_padroes:
        pos = encontrar_posicao_titulo(texto_norm, [fim_padrao], inicio_recorte + 50)

        if pos is not None:
            if fim is None or pos < fim:
                fim = pos

    if fim is None:
        fim = min(len(texto_limpo), inicio_recorte + limite * 2)

    trecho = texto_limpo[inicio_recorte:fim]

    return limitar(trecho, limite)


def extrair_trecho_por_termos(texto: str, termos: list[str], limite: int, janela: int = 700) -> str:
    if not texto:
        return ""

    texto_limpo = limpar_texto(texto)
    texto_norm = normalizar_para_busca(texto_limpo)

    melhor_pos = None

    for termo in termos:
        termo_norm = normalizar_para_busca(termo)
        pos = texto_norm.find(termo_norm)

        if pos != -1:
            if melhor_pos is None or pos < melhor_pos:
                melhor_pos = pos

    if melhor_pos is None:
        return ""

    inicio = max(0, melhor_pos - 250)
    fim = min(len(texto_limpo), melhor_pos + janela)

    return limitar(texto_limpo[inicio:fim], limite)


# ============================================================
# EXTRAÇÕES ESPECÍFICAS
# ============================================================

def extrair_titulo_provavel(texto: str, nome_arquivo: str) -> str:
    if not texto:
        return "Não identificado"

    inicio = texto[:5000]
    inicio = limpar_texto(inicio)

    padroes = [
        r"título\s*[:\-]\s*(.{20,250})",
        r"titulo\s*[:\-]\s*(.{20,250})",
    ]

    for padrao in padroes:
        match = re.search(padrao, inicio, flags=re.IGNORECASE)

        if match:
            titulo = limpar_texto(match.group(1))
            titulo = re.split(
                r"(autor|autora|dissertação|dissertacao|produto educacional|orientador)",
                titulo,
                flags=re.IGNORECASE
            )[0]
            titulo = titulo.strip(" .:-")

            if 20 <= len(titulo) <= 250:
                return titulo

    nome = Path(nome_arquivo).stem
    nome = re.sub(r"[_\-]+", " ", nome)
    nome = re.sub(r"\d+", " ", nome)
    nome = limpar_texto(nome).title()

    return nome if nome else "Não identificado"


def extrair_palavras_chave(texto: str) -> list[str]:
    if not texto:
        return []

    texto_limpo = limpar_texto(texto)

    padrao = re.compile(
        r"(palavras[\s\-]?chave|palavras chave|keywords)\s*[:\-–—]\s*(.{10,700})",
        flags=re.IGNORECASE
    )

    match = padrao.search(texto_limpo)

    if not match:
        return []

    trecho = match.group(2)

    trecho = re.split(
        r"\b(abstract|resumo|introdução|introducao|sumário|sumario|1\s+introdução|1\s+introducao)\b",
        trecho,
        flags=re.IGNORECASE
    )[0]

    trecho = limpar_texto(trecho)
    partes = re.split(r";|,|\.", trecho)

    palavras = []
    vistos = set()

    for parte in partes:
        palavra = limpar_texto(parte).strip(" .:-–—")

        if 2 <= len(palavra) <= 80:
            chave = normalizar_para_busca(palavra)

            if chave not in vistos:
                palavras.append(palavra)
                vistos.add(chave)

    return palavras[:12]


def montar_trechos_relevantes(dados: dict) -> list[str]:
    campos_prioritarios = [
        "resumo",
        "objetivos",
        "metodologia",
        "produto_educacional",
        "publico_alvo",
        "finalidade",
        "base_teorica_ou_referencial",
        "introducao",
    ]

    trechos = []
    vistos = set()

    for campo in campos_prioritarios:
        valor = dados.get(campo, "")

        if valor:
            trecho = limitar(valor, LIMITES["trecho_relevante"])
            chave = normalizar_para_busca(trecho[:300])

            if chave and chave not in vistos:
                trechos.append(trecho)
                vistos.add(chave)

    return trechos[:10]


def montar_texto_limpo_para_classificacao(dados: dict) -> str:
    partes = []

    titulo_oficial = dados.get("titulo_oficial_dissertacao", "")

    if titulo_oficial:
        partes.append(f"TÍTULO OFICIAL: {titulo_oficial}")
    else:
        partes.append(f"TÍTULO PROVÁVEL: {dados.get('titulo_provavel', '')}")

    if dados.get("produto_educacional_oficial_site"):
        partes.append(f"PRODUTO EDUCACIONAL OFICIAL DO SITE: {dados['produto_educacional_oficial_site']}")

    if dados.get("resumo_limpo"):
        partes.append(f"RESUMO: {dados['resumo_limpo']}")

    if dados.get("palavras_chave_texto"):
        partes.append(f"PALAVRAS-CHAVE: {dados['palavras_chave_texto']}")

    if dados.get("objetivos_limpo"):
        partes.append(f"OBJETIVOS: {dados['objetivos_limpo']}")

    if dados.get("metodologia_limpo"):
        partes.append(f"METODOLOGIA: {dados['metodologia_limpo']}")

    if dados.get("produto_educacional_limpo"):
        partes.append(f"PRODUTO EDUCACIONAL EXTRAÍDO DO PDF: {dados['produto_educacional_limpo']}")

    if dados.get("publico_alvo_limpo"):
        partes.append(f"PÚBLICO-ALVO EXTRAÍDO DO PDF: {dados['publico_alvo_limpo']}")

    if dados.get("finalidade_limpo"):
        partes.append(f"FINALIDADE EXTRAÍDA DO PDF: {dados['finalidade_limpo']}")

    if dados.get("base_teorica_ou_referencial_limpo"):
        partes.append(f"BASE TEÓRICA OU REFERENCIAL: {dados['base_teorica_ou_referencial_limpo']}")

    if dados.get("alertas_qualidade_campos"):
        partes.append("ALERTAS DE QUALIDADE: " + " | ".join(dados["alertas_qualidade_campos"]))

    texto = "\n\n".join(partes)

    return limitar(texto, LIMITES["texto_limpo_para_classificacao"])


def montar_texto_base_para_ia(dados: dict) -> str:
    partes = []

    partes.append("DADOS OFICIAIS DA PLANILHA DO SITE")
    partes.append(f"ID: {dados.get('id_geral', '')}")
    partes.append(f"ANO OFICIAL: {dados.get('ano_oficial', '')}")
    partes.append(f"AUTOR OFICIAL: {dados.get('autor_oficial', '')}")
    partes.append(f"ORIENTADOR OFICIAL: {dados.get('orientador_oficial', '')}")
    partes.append(f"TÍTULO OFICIAL DA DISSERTAÇÃO: {dados.get('titulo_oficial_dissertacao', '')}")
    partes.append(f"PRODUTO EDUCACIONAL OFICIAL DO SITE: {dados.get('produto_educacional_oficial_site', '')}")
    partes.append(f"LINK PDF OFICIAL: {dados.get('link_pdf_oficial', '')}")
    partes.append(f"LINK PRODUTO OFICIAL: {dados.get('link_produto_oficial', '')}")

    partes.append("\nDADOS EXTRAÍDOS DO PDF")

    if dados.get("titulo_provavel"):
        partes.append(f"TÍTULO PROVÁVEL DO PDF: {dados['titulo_provavel']}")

    if dados.get("resumo"):
        partes.append(f"RESUMO: {dados['resumo']}")

    if dados.get("palavras_chave"):
        partes.append("PALAVRAS-CHAVE: " + "; ".join(dados["palavras_chave"]))

    if dados.get("objetivos"):
        partes.append(f"OBJETIVOS: {dados['objetivos']}")

    if dados.get("metodologia"):
        partes.append(f"METODOLOGIA: {dados['metodologia']}")

    if dados.get("produto_educacional"):
        partes.append(f"PRODUTO EDUCACIONAL EXTRAÍDO DO PDF: {dados['produto_educacional']}")

    if dados.get("publico_alvo"):
        partes.append(f"PÚBLICO-ALVO EXTRAÍDO DO PDF: {dados['publico_alvo']}")

    if dados.get("finalidade"):
        partes.append(f"FINALIDADE EXTRAÍDA DO PDF: {dados['finalidade']}")

    if dados.get("base_teorica_ou_referencial"):
        partes.append(f"BASE TEÓRICA OU REFERENCIAL EXTRAÍDA DO PDF: {dados['base_teorica_ou_referencial']}")

    if dados.get("trechos_relevantes_para_classificacao"):
        partes.append(
            "TRECHOS RELEVANTES PARA CLASSIFICAÇÃO: "
            + " | ".join(dados["trechos_relevantes_para_classificacao"])
        )

    partes.append(
        "\nORIENTAÇÃO PARA IA: use os dados oficiais da planilha para título, autor, orientador, ano e links. "
        "Use os dados do PDF para resumo, palavras-chave, classificação, público-alvo, finalidade e base teórica. "
        "Se algum campo do PDF parecer capa, sumário, banca, agradecimento ou trecho quebrado, não use como valor final; marque revisão manual."
    )

    texto = "\n\n".join(partes)

    return limitar(texto, LIMITES["texto_base_para_ia"])


def extrair_dados_academicos(texto: str, nome_arquivo: str, dados_site: dict) -> dict:
    fim_resumo = [
        "Palavras-chave",
        "Palavras chave",
        "Palavras-Chave",
        "Keywords",
        "Abstract",
        "Introdução",
        "Introducao",
        "Sumário",
        "Sumario",
    ]

    fim_abstract = [
        "Keywords",
        "Palavras-chave",
        "Palavras chave",
        "Introdução",
        "Introducao",
        "Sumário",
        "Sumario",
    ]

    fim_capitulo = [
        "Resumo",
        "Abstract",
        "Introdução",
        "Introducao",
        "Objetivos",
        "Objetivo geral",
        "Objetivos específicos",
        "Objetivos especificos",
        "Objetivo",
        "Metodologia",
        "Procedimentos metodológicos",
        "Procedimentos metodologicos",
        "Percurso metodológico",
        "Percurso metodologico",
        "Referencial teórico",
        "Referencial teorico",
        "Fundamentação teórica",
        "Fundamentacao teorica",
        "Produto Educacional",
        "Resultados",
        "Considerações finais",
        "Consideracoes finais",
        "Conclusão",
        "Conclusoes",
        "Conclusões",
        "Referências",
        "Referencias",
        "Apêndice",
        "Apendice",
        "Anexo",
    ]

    dados = {
        "titulo_provavel": extrair_titulo_provavel(texto, nome_arquivo),

        "resumo": extrair_entre_marcadores(
            texto,
            ["Resumo"],
            fim_resumo,
            LIMITES["resumo"]
        ),

        "palavras_chave": extrair_palavras_chave(texto),

        "abstract": extrair_entre_marcadores(
            texto,
            ["Abstract"],
            fim_abstract,
            LIMITES["abstract"]
        ),

        "introducao": extrair_entre_marcadores(
            texto,
            ["Introdução", "Introducao"],
            fim_capitulo,
            LIMITES["introducao"],
            posicao_minima=1000
        ),

        "objetivos": extrair_entre_marcadores(
            texto,
            [
                "Objetivos",
                "Objetivo geral",
                "Objetivos específicos",
                "Objetivos especificos",
                "Objetivo",
            ],
            fim_capitulo,
            LIMITES["objetivos"],
            posicao_minima=1000
        ),

        "metodologia": extrair_entre_marcadores(
            texto,
            [
                "Metodologia",
                "Procedimentos metodológicos",
                "Procedimentos metodologicos",
                "Percurso metodológico",
                "Percurso metodologico",
                "Caminho metodológico",
                "Caminho metodologico",
                "Método",
                "Metodo",
                "Métodos",
                "Metodos",
            ],
            fim_capitulo,
            LIMITES["metodologia"],
            posicao_minima=1000
        ),

        "produto_educacional": extrair_entre_marcadores(
            texto,
            [
                "Produto Educacional",
                "O Produto Educacional",
                "Descrição do Produto Educacional",
                "Descricao do Produto Educacional",
                "Produto",
            ],
            fim_capitulo,
            LIMITES["produto_educacional"],
            posicao_minima=1000
        ),

        "publico_alvo": extrair_trecho_por_termos(
            texto,
            [
                "público-alvo",
                "publico-alvo",
                "público alvo",
                "publico alvo",
                "destina-se a",
                "destinado a",
                "destinada a",
                "voltado a",
                "voltada a",
                "voltado para",
                "voltada para",
                "participantes da pesquisa",
                "sujeitos da pesquisa",
                "estudantes",
                "alunos",
                "professores",
                "docentes",
                "discentes",
                "trabalhadores",
                "EJA",
                "PROEJA",
                "ensino médio integrado",
                "ensino medio integrado",
                "curso técnico",
                "curso tecnico",
                "graduação",
                "graduacao",
                "comunidade",
                "servidores",
                "gestores",
            ],
            LIMITES["publico_alvo"]
        ),

        "finalidade": extrair_trecho_por_termos(
            texto,
            [
                "finalidade",
                "objetivo do produto",
                "tem como objetivo",
                "teve como objetivo",
                "visa",
                "busca",
                "pretende contribuir",
                "foi desenvolvido para",
                "foi desenvolvida para",
                "elaborado para",
                "elaborada para",
                "criado para",
                "criada para",
                "auxiliar",
                "contribuir com",
            ],
            LIMITES["finalidade"]
        ),

        "base_teorica_ou_referencial": extrair_entre_marcadores(
            texto,
            [
                "Referencial teórico",
                "Referencial teorico",
                "Fundamentação teórica",
                "Fundamentacao teorica",
                "Bases teóricas",
                "Bases teoricas",
                "Fundamentação",
                "Fundamentacao",
                "Marco teórico",
                "Marco teorico",
            ],
            fim_capitulo,
            LIMITES["base_teorica_ou_referencial"],
            posicao_minima=1000
        ),
    }

    if not dados["produto_educacional"]:
        dados["produto_educacional"] = extrair_trecho_por_termos(
            texto,
            [
                "produto educacional",
                "material educativo",
                "guia",
                "cartilha",
                "sequência didática",
                "sequencia didatica",
                "oficina",
                "curso",
                "aplicativo",
                "site",
                "vídeo",
                "video",
                "podcast",
                "manual",
                "ebook",
                "e-book",
                "infográfico",
                "infografico",
                "folder",
                "jogo",
                "quiz",
                "hq",
                "história em quadrinhos",
                "historia em quadrinhos",
            ],
            LIMITES["produto_educacional"],
            janela=1200
        )

    campos_para_qualidade = {
        "resumo": [
            "resumo",
            "esta dissertação",
            "esta dissertacao",
            "esta pesquisa",
            "objetivo",
        ],
        "introducao": [
            "introdução",
            "introducao",
        ],
        "objetivos": [
            "objetivo",
            "objetivos",
            "analisar",
            "compreender",
            "investigar",
            "identificar",
        ],
        "metodologia": [
            "metodologia",
            "pesquisa qualitativa",
            "análise documental",
            "analise documental",
            "pesquisa-ação",
            "pesquisa acao",
            "entrevista",
            "questionário",
            "questionario",
            "análise de conteúdo",
            "analise de conteudo",
        ],
        "produto_educacional": [
            "produto educacional",
            "cartilha",
            "guia",
            "manual",
            "e-book",
            "ebook",
            "vídeo",
            "video",
            "podcast",
            "site",
            "infográfico",
            "infografico",
            "sequência didática",
            "sequencia didatica",
            "oficina",
            "curso",
        ],
        "publico_alvo": [
            "público-alvo",
            "publico-alvo",
            "público alvo",
            "publico alvo",
            "destina-se",
            "destinado",
            "destinada",
            "voltado",
            "voltada",
            "participantes",
            "sujeitos",
        ],
        "finalidade": [
            "objetivo",
            "finalidade",
            "visa",
            "contribuir",
            "auxiliar",
            "foi desenvolvido para",
            "foi desenvolvida para",
        ],
        "base_teorica_ou_referencial": [
            "referencial teórico",
            "referencial teorico",
            "fundamentação teórica",
            "fundamentacao teorica",
            "trabalho como princípio educativo",
            "trabalho como principio educativo",
            "formação humana integral",
            "formacao humana integral",
            "currículo integrado",
            "curriculo integrado",
            "pedagogia histórico-crítica",
            "pedagogia historico-critica",
            "paulo freire",
        ],
    }

    alertas = []

    for campo, termos in campos_para_qualidade.items():
        qualidade, motivos = qualidade_campo(dados.get(campo, ""), termos)

        dados[f"{campo}_qualidade"] = qualidade
        dados[f"{campo}_motivos"] = motivos
        dados[f"{campo}_limpo"] = limpar_campo_por_qualidade(dados.get(campo, ""), qualidade)

        if qualidade in {"baixa", "descartado", "vazio"}:
            alertas.append(f"{campo}: {qualidade} ({'; '.join(motivos)})")

    dados["palavras_chave_texto"] = "; ".join(dados.get("palavras_chave", []))
    dados["palavras_chave_qualidade"] = "alta" if dados.get("palavras_chave") else "vazio"

    dados["trechos_relevantes_para_classificacao"] = montar_trechos_relevantes(dados)

    for chave, valor in dados_site.items():
        dados[chave] = valor

    dados["alertas_qualidade_campos"] = alertas
    dados["texto_limpo_para_classificacao"] = montar_texto_limpo_para_classificacao(dados)
    dados["texto_base_para_ia"] = montar_texto_base_para_ia(dados)

    return dados


# ============================================================
# PROCESSAMENTO PRINCIPAL
# ============================================================

def criar_objeto_erro(
    id_geral,
    id_lote,
    ordem_no_lote,
    nome_arquivo,
    erro,
    dados_site=None,
    metodo_vinculo_site=""
):
    dados_site = dados_site or {}

    objeto = {
        "id_geral": id_geral,
        "id_lote": id_lote,
        "ordem_no_lote": ordem_no_lote,
        "nome_arquivo": nome_arquivo,
        "titulo_provavel": "Não identificado",
        "resumo": "",
        "palavras_chave": [],
        "abstract": "",
        "introducao": "",
        "objetivos": "",
        "metodologia": "",
        "produto_educacional": "",
        "publico_alvo": "",
        "finalidade": "",
        "base_teorica_ou_referencial": "",
        "trechos_relevantes_para_classificacao": [],
        "texto_limpo_para_classificacao": "",
        "texto_base_para_ia": "",
        "metodo_extracao": "falhou",
        "status_extracao": "Erro",
        "erro": erro,
        "metodo_vinculo_site": metodo_vinculo_site,
        "dados_site": dados_site,
    }

    for chave, valor in dados_site.items():
        objeto[chave] = valor

    return objeto


def processar_pdf(
    caminho_pdf: Path,
    id_geral: int,
    id_lote: int,
    ordem_no_lote: int,
    dados_site_por_id: dict
) -> dict:
    nome_arquivo = caminho_pdf.name
    erro = ""
    metodo = ""
    status = ""

    dados_site, metodo_vinculo_site = obter_dados_site_para_pdf(
        caminho_pdf,
        dados_site_por_id
    )

    try:
        texto = ""

        try:
            texto = extrair_texto_pdf(caminho_pdf)
            metodo = "texto_pdf"
        except Exception as e:
            erro = f"Erro ao extrair texto normal: {e}"
            texto = ""

        if len(texto) < MIN_CARACTERES_TEXTO_NORMAL:
            if USAR_OCR_SE_NECESSARIO:
                try:
                    print("   Texto normal insuficiente. Tentando OCR...")
                    texto_ocr = extrair_texto_ocr(caminho_pdf)

                    if texto and texto_ocr:
                        texto = texto + "\n\n" + texto_ocr
                        metodo = "texto_pdf_e_ocr"
                    elif texto_ocr:
                        texto = texto_ocr
                        metodo = "ocr"

                except Exception as e:
                    erro_ocr = f"Erro no OCR: {e}"
                    erro = f"{erro} | {erro_ocr}" if erro else erro_ocr

        if not texto or len(texto) < 500:
            status = "Texto insuficiente"
        else:
            status = "Sucesso"

        dados_extraidos = extrair_dados_academicos(
            texto=texto,
            nome_arquivo=nome_arquivo,
            dados_site=dados_site
        )

        resultado = {
            "id_geral": id_geral,
            "id_lote": id_lote,
            "ordem_no_lote": ordem_no_lote,
            "nome_arquivo": nome_arquivo,
            "metodo_vinculo_site": metodo_vinculo_site,
            "dados_site": dados_site,

            "titulo_provavel": dados_extraidos.get("titulo_provavel", ""),
            "resumo": dados_extraidos.get("resumo", ""),
            "palavras_chave": dados_extraidos.get("palavras_chave", []),
            "abstract": dados_extraidos.get("abstract", ""),
            "introducao": dados_extraidos.get("introducao", ""),
            "objetivos": dados_extraidos.get("objetivos", ""),
            "metodologia": dados_extraidos.get("metodologia", ""),
            "produto_educacional": dados_extraidos.get("produto_educacional", ""),
            "publico_alvo": dados_extraidos.get("publico_alvo", ""),
            "finalidade": dados_extraidos.get("finalidade", ""),
            "base_teorica_ou_referencial": dados_extraidos.get("base_teorica_ou_referencial", ""),

            "resumo_qualidade": dados_extraidos.get("resumo_qualidade", ""),
            "resumo_motivos": dados_extraidos.get("resumo_motivos", []),
            "resumo_limpo": dados_extraidos.get("resumo_limpo", ""),

            "introducao_qualidade": dados_extraidos.get("introducao_qualidade", ""),
            "introducao_motivos": dados_extraidos.get("introducao_motivos", []),
            "introducao_limpo": dados_extraidos.get("introducao_limpo", ""),

            "objetivos_qualidade": dados_extraidos.get("objetivos_qualidade", ""),
            "objetivos_motivos": dados_extraidos.get("objetivos_motivos", []),
            "objetivos_limpo": dados_extraidos.get("objetivos_limpo", ""),

            "metodologia_qualidade": dados_extraidos.get("metodologia_qualidade", ""),
            "metodologia_motivos": dados_extraidos.get("metodologia_motivos", []),
            "metodologia_limpo": dados_extraidos.get("metodologia_limpo", ""),

            "produto_educacional_qualidade": dados_extraidos.get("produto_educacional_qualidade", ""),
            "produto_educacional_motivos": dados_extraidos.get("produto_educacional_motivos", []),
            "produto_educacional_limpo": dados_extraidos.get("produto_educacional_limpo", ""),

            "publico_alvo_qualidade": dados_extraidos.get("publico_alvo_qualidade", ""),
            "publico_alvo_motivos": dados_extraidos.get("publico_alvo_motivos", []),
            "publico_alvo_limpo": dados_extraidos.get("publico_alvo_limpo", ""),

            "finalidade_qualidade": dados_extraidos.get("finalidade_qualidade", ""),
            "finalidade_motivos": dados_extraidos.get("finalidade_motivos", []),
            "finalidade_limpo": dados_extraidos.get("finalidade_limpo", ""),

            "base_teorica_ou_referencial_qualidade": dados_extraidos.get("base_teorica_ou_referencial_qualidade", ""),
            "base_teorica_ou_referencial_motivos": dados_extraidos.get("base_teorica_ou_referencial_motivos", []),
            "base_teorica_ou_referencial_limpo": dados_extraidos.get("base_teorica_ou_referencial_limpo", ""),

            "palavras_chave_texto": dados_extraidos.get("palavras_chave_texto", ""),
            "palavras_chave_qualidade": dados_extraidos.get("palavras_chave_qualidade", ""),

            "alertas_qualidade_campos": dados_extraidos.get("alertas_qualidade_campos", []),
            "trechos_relevantes_para_classificacao": dados_extraidos.get("trechos_relevantes_para_classificacao", []),
            "texto_limpo_para_classificacao": dados_extraidos.get("texto_limpo_para_classificacao", ""),
            "texto_base_para_ia": dados_extraidos.get("texto_base_para_ia", ""),

            "metodo_extracao": metodo if metodo else "falhou",
            "status_extracao": status,
            "erro": erro,
        }

        for chave, valor in dados_site.items():
            resultado[chave] = valor

        return resultado

    except Exception:
        erro_completo = traceback.format_exc()

        return criar_objeto_erro(
            id_geral=id_geral,
            id_lote=id_lote,
            ordem_no_lote=ordem_no_lote,
            nome_arquivo=nome_arquivo,
            erro=erro_completo,
            dados_site=dados_site,
            metodo_vinculo_site=metodo_vinculo_site,
        )


def salvar_lote(lote: list[dict], numero_lote: int):
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)

    caminho_saida = PASTA_SAIDA / f"lote_{numero_lote:03d}.json"

    with open(caminho_saida, "w", encoding="utf-8") as arquivo:
        json.dump(lote, arquivo, ensure_ascii=False, indent=2)

    print(f"[LOTE SALVO] {caminho_saida}")


def main():
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)

    limpar_jsons_antigos()

    dados_site_por_id = carregar_planilha_site()

    if not PASTA_ENTRADA.exists():
        print(f"Erro: a pasta de entrada não existe: {PASTA_ENTRADA.resolve()}")
        print("Crie a pasta 'pdfs/' e coloque os PDFs dentro dela.")
        return

    pdfs = sorted(PASTA_ENTRADA.glob("*.pdf"), key=lambda p: p.name.lower())

    total = len(pdfs)

    if total == 0:
        print(f"Nenhum PDF encontrado em: {PASTA_ENTRADA.resolve()}")
        return

    print(f"\nEncontrados {total} PDFs.")

    lote_atual = []
    numero_lote = 1

    for indice, caminho_pdf in enumerate(pdfs, start=1):
        id_geral = indice
        id_lote = ((indice - 1) // TAMANHO_LOTE) + 1
        ordem_no_lote = ((indice - 1) % TAMANHO_LOTE) + 1

        print(f"\n[{indice}/{total}] Processando: {caminho_pdf.name}")

        resultado = processar_pdf(
            caminho_pdf=caminho_pdf,
            id_geral=id_geral,
            id_lote=id_lote,
            ordem_no_lote=ordem_no_lote,
            dados_site_por_id=dados_site_por_id,
        )

        print(f"   Método extração: {resultado.get('metodo_extracao')}")
        print(f"   Status extração: {resultado.get('status_extracao')}")
        print(f"   Vínculo site: {resultado.get('metodo_vinculo_site')}")

        titulo_site = resultado.get("titulo_oficial_dissertacao", "")

        if titulo_site:
            print(f"   Título oficial: {titulo_site[:90]}")

        if resultado.get("erro"):
            print(f"   Erro registrado: {resultado.get('erro')[:250]}")

        lote_atual.append(resultado)

        if len(lote_atual) == TAMANHO_LOTE:
            print(f"\n[{indice}/{total}] Salvando lote_{numero_lote:03d}.json")
            salvar_lote(lote_atual, numero_lote)
            lote_atual = []
            numero_lote += 1

    if lote_atual:
        print(f"\nSalvando lote_{numero_lote:03d}.json")
        salvar_lote(lote_atual, numero_lote)

    print("\nProcessamento concluído.")
    print(f"JSONs novos salvos em: {PASTA_SAIDA.resolve()}")


if __name__ == "__main__":
    main()