"""Etapa 2: lê a planilha oficial, executa OCR e prepara o envio manual ao GPT.

Execução normal:
    python 2_PdfParaJson.py

Saídas:
    dados/GPT/para_gpt.json
    dados/GPT/PromptClassificacaoGPT.md
"""

from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor
from functools import lru_cache
import json
from pathlib import Path
import re

import pandas as pd

try:
    import fitz
except ImportError:  # pragma: no cover - mensagem tratada em tempo de execução
    fitz = None

try:
    import pytesseract
    from PIL import Image
except ImportError:  # pragma: no cover - OCR é fallback opcional
    pytesseract = None
    Image = None

from util import (
    agora_utc,
    ler_json,
    limpar_texto,
    normalizar_busca,
    salvar_json_atomico,
    salvar_texto_atomico,
    sha256_arquivo,
)


ARQUIVO_ENTRADA = Path("dados/1_PegarDadosSiteIfc.xlsx")
ARQUIVO_IDENTIDADES = Path("dados/referencia/identidades.json")
ARQUIVO_CLASSIFICACOES = Path("dados/referencia/classificacoes_historicas.json")
ARQUIVO_VOCABULARIOS = Path("dados/referencia/vocabularios.json")
ARQUIVO_SAIDA = Path("dados/GPT/para_gpt.json")
ARQUIVO_PROMPT = Path("dados/GPT/PromptClassificacaoGPT.md")
ARQUIVO_RESPOSTA_GPT = Path("dados/GPT/gpt.json")
PASTA_PDFS = Path("pdf")
PASTA_CACHE_OCR = Path("dados/cache/ocr")
ARQUIVO_RELATORIO_OCR = Path("dados/relatorio_ocr.txt")
SCHEMA_ENTRADA_GPT = "leme-para-classificacao-gpt-v3"
SCHEMA_RESPOSTA_GPT = "leme-classificacoes-gpt-v2"
SCHEMA_VOCABULARIOS = "leme-vocabularios-v2"
SCHEMA_CLASSIFICACOES = "leme-classificacoes-historicas-v1"
VERSAO_ALIASES_ORIENTADORES = "confirmados_fontes_ifc_v1"
VERSAO_TIPOS_PRODUTO = "tipos_controlados_semanticos_v3"
MAX_PAGINAS_OCR = 15
MAX_PAGINAS_OCR_HISTORICO = 3
INSTITUICAO_CORPUS = "Instituto Federal Catarinense"
CAMPUS_CORPUS = "Blumenau"
PROGRAMA_CORPUS = "ProfEPT"

# Equivalências conferidas em páginas institucionais e/ou documentos assinados
# do IFC. Elas são deliberadamente explícitas: nomes parecidos que não estejam
# nesta lista continuam separados.
ORIENTADORES_CONFIRMADOS = {
    "Eduardo Augusto Werneck Ribeiro": {
        "aliases": ["Eduardo Werneck Ribeiro"],
        "evidencias": [
            "https://profept.ifc.edu.br/docentes/",
            "https://profept.ifc.edu.br/dissertacoes/",
        ],
    },
    "Fátima Peres Zago de Oliveira": {
        "aliases": ["Fátima Perez Zago de Oliveira"],
        "evidencias": [
            "https://profept.ifc.edu.br/docentes/",
            "https://profept.ifc.edu.br/dissertacoes/",
        ],
    },
    "Inge Renate Fröse Suhr": {
        "aliases": ["Inge Renate Frose Suhr", "Inge R. F. Suhr"],
        "evidencias": [
            "https://pedagogia.blumenau.ifc.edu.br/corpo-docente/",
            "https://sig.ifc.edu.br/sigaa/public/programa/defesas.jsf?id=1010",
        ],
    },
    "Leandro Marcos Salgado Alves": {
        "aliases": ["Leandro Marcos Alves Salgado"],
        "evidencias": [
            "https://profept.ifc.edu.br/agenda-de-eventos/",
            "https://profept.ifc.edu.br/dissertacoes/",
        ],
    },
    "Sonia Regina de Souza Fernandes": {
        "aliases": ["Sônia Regina Fernandes"],
        "evidencias": [
            "https://profept.ifc.edu.br/wp-content/uploads/sites/54/2023/09/DISSERTACAO-LUANA-TILLMANN-PDF-A.pdf",
            "https://profept.ifc.edu.br/dissertacoes/",
        ],
    },
}

LINHAS = {
    "Práticas Educativas em Educação Profissional e Tecnológica (EPT)": [
        "Práticas Educativas em EPT",
        "Práticas Educativas em Educação Profissional e Tecnológica (EPT)",
    ],
    "Organização e Memórias de Espaços Pedagógicos na Educação Profissional e Tecnológica (EPT)": [
        "Organização e Memórias de Espaços Pedagógicos na EPT",
        "Organização e Memórias de Espaços Pedagógicos na Educação Profissional e Tecnológica (EPT)",
    ],
}

MACROPROJETOS = {
    "MP1": "Propostas metodológicas e recursos didáticos em espaços formais e não formais de ensino na EPT",
    "MP2": "Inclusão e diversidade em espaços formais e não formais de ensino na EPT",
    "MP3": "Práticas Educativas no Currículo Integrado",
    "MP4": "História e memórias no contexto da EPT",
    "MP5": "Organização do currículo integrado na EPT",
    "MP6": "Organização de espaços pedagógicos na EPT",
}

CAMPOS_GPT = {
    "linha_pesquisa": "Linha de Pesquisa",
    "macroprojeto": "Macroprojeto",
    "base_epistemologica": "Base Epistemológica",
    "nivel_aplicacao": "Nível de Aplicação",
    "tipo_produto": "Tipo de Produto Educacional",
    "area_tematica": "Área Temática",
    "publico_alvo": "Público-alvo",
}

PUBLICOS_CONTROLADOS = [
    "Comunidade escolar",
    "Comunidade externa",
    "Docentes",
    "Egressos",
    "Estudantes",
    "Famílias e responsáveis",
    "Gestão e equipes institucionais",
    "Pesquisadores",
    "Servidores e profissionais",
]

REGRAS_PUBLICO = {
    "Comunidade escolar": ["comunidade escolar", "comunidade academica", "comunidade interna"],
    "Comunidade externa": ["comunidade externa", "sociedade", "populacao local", "publico interessado"],
    "Docentes": ["docente", "professor", "educador"],
    "Egressos": ["egresso"],
    "Estudantes": ["estudante", "discente", "aluno", "educando"],
    "Famílias e responsáveis": ["familia", "responsavel"],
    "Gestão e equipes institucionais": ["gestor", "gestao", "coordenador", "equipe", "comissao", "nucleo"],
    "Pesquisadores": ["pesquisador"],
    "Servidores e profissionais": ["servidor", "tecnico administrativo", "profissional", "enfermeiro", "tradutor"],
}

NIVEIS_CONTROLADOS = [
    "Aprendizagem Profissional",
    "Comunidade externa",
    "Curso Técnico",
    "Curso Técnico Subsequente",
    "Educação Profissional e Tecnológica",
    "EJA-EPT / PROEJA",
    "Ensino Médio",
    "Ensino Médio Integrado",
    "Formação de professores",
    "Institucional / Educação Profissional e Tecnológica",
    "Transição para o Ensino Médio Integrado",
]

# Tipos funcionais amplos usados nos filtros e gráficos. O valor histórico
# continua preservado separadamente; aqui são reunidos sinônimos e descrições
# compostas que indicam o mesmo formato principal.
TIPOS_PRODUTO_CONTROLADOS = {
    "Aplicativo / software": ["aplicativo móvel", "software/aplicativo / observatório"],
    "Apresentação": ["apresentação / palestra"],
    "Caderno": ["Caderno de orientações", "caderno didático-pedagógico"],
    "Cartilha": ["cartilha", "Cartilha de orientações"],
    "Curso / formação": [
        "Capacitação / curso formativo", "curso", "Curso de formação + vídeos",
        "curso / produto audiovisual", "curso/portfólio formativo", "curso/videoaulas",
    ],
    "E-book": [
        "e-book", "E-book", "ebook", "Ebook", "eBook", "livro digital", "Livro digital",
        "Livro digital / E-book", "E-book / curso formativo",
        "E-book / sequência didática gamificada", "livro digital / oficina",
        "livro digital/guia", "Livro digital / história em quadrinhos",
    ],
    "Exposição": ["exposição fotográfica"],
    "Folder / material informativo": [
        "folder", "folder educativo", "panfleto digital / material didático instrucional",
    ],
    "Guia": [
        "guia", "Guia", "guia / apresentação", "guia digital / e-book",
        "Guia digital bilíngue", "Guia formativo",
        "guia formativo / sequência didática", "guia/cartilha informativa",
    ],
    "História em quadrinhos": ["história em quadrinhos", "história em quadrinhos (HQ)"],
    "Infográfico": ["infográfico", "infográfico / hipertexto virtual"],
    "Instrumento de gestão": ["prontuário unificado e calendário de ações"],
    "Jogo educacional": ["quiz / jogo educacional"],
    "Manual": [
        "manual", "manual / e-book", "Manual / metodologia reflexiva",
        "Manual técnico / matriz de indicadores",
    ],
    "Material didático": [
        "Material de apoio / apresentação/tutorial", "material didático",
        "material histórico / produto educacional",
    ],
    "Oficina": ["oficina", "Oficina pedagógica", "oficina / relatório técnico"],
    "Podcast": ["podcast", "Podcast", "podcast educativo"],
    "Proposta normativa": ["proposta normativa"],
    "Proposta pedagógica": ["proposta pedagógica"],
    "Sequência didática": [
        "sequência didática", "Sequência didática",
        "sequência didática / e-book / cartilha eletrônica",
    ],
    "Site / portal": [
        "memorial virtual", "Página de internet / storytelling",
        "portal eletrônico / tour 360°", "Portal / site memorial", "site", "Site",
        "Site / ferramenta de avaliação diagnóstica", "Site / memorial digital",
        "site/portal", "Website / repositório digital",
    ],
    "Tutorial": ["tutorial", "Tutorial pedagógico / videoaula"],
    "Vídeo / audiovisual": [
        "audiovisual / animação", "produto audiovisual / lives",
        "projeto de extensão / narrativa audiovisual", "vídeo",
        "Vídeo em animação gráfica traduzido em Libras", "Vídeo explicativo",
        "vídeo / material audiovisual",
    ],
}

REGRAS_NIVEL = [
    ("Transição para o Ensino Médio Integrado", ["transicao", "9 ano", "nono ano", "ingresso no ensino medio integrado"]),
    ("EJA-EPT / PROEJA", ["eja ept", "proeja", "educacao de jovens e adultos"]),
    ("Aprendizagem Profissional", ["aprendizagem profissional", "jovem aprendiz"]),
    ("Curso Técnico Subsequente", ["subsequente"]),
    ("Ensino Médio Integrado", ["ensino medio integrado", "medio integrado", "emiept"]),
    ("Formação de professores", ["formacao de professor", "formacao docente", "licenciatura"]),
    ("Curso Técnico", ["curso tecnico", "cursos tecnicos", "educacao profissional tecnica"]),
    ("Ensino Médio", ["ensino medio"]),
    ("Comunidade externa", ["comunidade externa", "sociedade"]),
    ("Institucional / Educação Profissional e Tecnológica", ["institucional", "instituto federal", "rede federal", "ifc"]),
    ("Educação Profissional e Tecnológica", ["educacao profissional e tecnologica", "ept"]),
]


def classificar_nivel(valor) -> str:
    texto = normalizar_busca(valor)
    for canonico, termos in REGRAS_NIVEL:
        if any(normalizar_busca(termo) in texto for termo in termos):
            return canonico
    return "A revisar"

def argumentos() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--limite", type=int, default=0, help="limita registros para diagnóstico")
    parser.add_argument(
        "--max-paginas-ocr",
        type=int,
        default=MAX_PAGINAS_OCR,
        help="máximo de páginas OCR por PDF; 0 processa todas as páginas",
    )
    parser.add_argument("--forcar-ocr", action="store_true", help="ignora o cache OCR")
    parser.add_argument(
        "--workers-ocr",
        type=int,
        default=4,
        help="páginas OCR processadas em paralelo por PDF (padrão: 4)",
    )
    return parser.parse_args()


def limpar_texto_pdf(texto: str) -> str:
    texto = (texto or "").replace("\x00", " ").replace("ﬁ", "fi").replace("ﬂ", "fl")
    texto = re.sub(r"(\w)-\s*\n\s*(\w)", r"\1\2", texto)
    texto = re.sub(r"[ \t]+", " ", texto)
    texto = re.sub(r"\n{3,}", "\n\n", texto)
    return texto.strip()


@lru_cache(maxsize=1)
def idioma_ocr() -> str:
    if pytesseract is None:
        return ""
    instalados = set(pytesseract.get_languages(config=""))
    if {"por", "eng"}.issubset(instalados):
        return "por+eng"
    if "por" in instalados:
        return "por"
    if "eng" in instalados:
        return "eng"
    return ""


def extrair_texto_ocr(
    caminho: Path,
    max_paginas: int = MAX_PAGINAS_OCR,
    workers: int = 4,
) -> dict:
    """Obtém conteúdo exclusivamente por OCR e registra falhas página a página."""
    if fitz is None or pytesseract is None or Image is None:
        raise RuntimeError("PyMuPDF, pytesseract e Pillow são obrigatórios para o OCR")
    idioma = idioma_ocr()
    if not idioma:
        raise RuntimeError("Tesseract não possui idioma por ou eng instalado")
    textos: list[str] = []
    erros: list[dict] = []
    with fitz.open(caminho) as documento:
        total = documento.page_count
        quantidade = total if max_paginas <= 0 else min(total, max_paginas)
    def processar_pagina(numero: int) -> tuple[int, str, str]:
        try:
            with fitz.open(caminho) as documento:
                pagina = documento.load_page(numero)
                pixmap = pagina.get_pixmap(matrix=fitz.Matrix(1.7, 1.7), alpha=False)
                imagem = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
            texto = limpar_texto_pdf(
                pytesseract.image_to_string(imagem, lang=idioma, config="--psm 3")
            )
            aviso = "OCR retornou pouco ou nenhum texto" if len(texto) < 20 else ""
            return numero, texto, aviso
        except Exception as erro:  # uma página ruim não invalida todo o trabalho
            return numero, "", f"{type(erro).__name__}: {erro}"

    with ThreadPoolExecutor(max_workers=max(1, workers)) as executor:
        resultados = list(executor.map(processar_pagina, range(quantidade)))
    for numero, texto, erro in sorted(resultados):
        textos.append(texto)
        if erro:
            erros.append({"pagina": numero + 1, "erro": erro})
    return {
        "texto": limpar_texto_pdf("\n\n".join(textos)),
        "paginas_total_pdf": total,
        "paginas_processadas": quantidade,
        "paginas_nao_processadas": max(0, total - quantidade),
        "erros_paginas": erros,
        "metodo": f"ocr_{idioma}",
    }


def extrair_ocr_com_cache(
    identificador: str,
    caminho: Path,
    max_paginas: int,
    workers: int = 4,
    forcar: bool = False,
) -> dict:
    hash_pdf = sha256_arquivo(caminho)
    sufixo = "todas" if max_paginas <= 0 else str(max_paginas)
    cache = PASTA_CACHE_OCR / f"{identificador}_{hash_pdf[:16]}_{sufixo}.json"
    if cache.is_file() and not forcar:
        resultado = ler_json(cache)
        if isinstance(resultado, dict) and resultado.get("sha256") == hash_pdf:
            resultado["origem"] = "cache"
            return resultado
    resultado = extrair_texto_ocr(caminho, max_paginas, workers)
    resultado.update(
        {
            "id": identificador,
            "sha256": hash_pdf,
            "gerado_em": agora_utc(),
            "origem": "ocr_novo",
        }
    )
    salvar_json_atomico(resultado, cache)
    return resultado


def trecho_entre(texto: str, inicio: str, finais: list[str], limite: int) -> str:
    match = re.search(inicio, texto, flags=re.IGNORECASE)
    if not match:
        return ""
    fim = min(
        (m.start() for padrao in finais if (m := re.search(padrao, texto[match.end() :], flags=re.IGNORECASE))),
        default=min(len(texto) - match.end(), limite),
    )
    return limpar_texto(texto[match.end() : match.end() + min(fim, limite)])


def trecho_por_termos(texto: str, termos: list[str], limite: int = 1800) -> str:
    for termo in termos:
        # A posição precisa ser obtida no próprio texto que será recortado.
        # Usar o índice de uma cópia sem acentos desloca o recorte quando há
        # caracteres Unicode e pode associar conteúdo irrelevante ao campo.
        encontrado = re.search(re.escape(termo), texto, flags=re.IGNORECASE)
        if encontrado:
            inicio = max(0, encontrado.start() - 250)
            return limpar_texto(texto[inicio : inicio + limite])
    return ""


def extrair_campos(texto: str) -> dict:
    resumo = trecho_entre(
        texto,
        r"(?:^|\n)\s*RESUMO\s*[:\-]?",
        [r"\bPalavras[- ]chave\s*:", r"(?:^|\n)\s*ABSTRACT\b"],
        8000,
    )
    abstract = trecho_entre(
        texto,
        r"(?:^|\n)\s*ABSTRACT\s*[:\-]?",
        [r"\bKeywords?\s*:", r"(?:^|\n)\s*(?:INTRODUÇÃO|1\s+INTRODUÇÃO)\b"],
        5000,
    )
    match_palavras = re.search(r"Palavras[- ]chave\s*:\s*([^\n]{3,500})", texto, flags=re.IGNORECASE)
    palavras = []
    if match_palavras:
        palavras = [limpar_texto(item).strip(". ,;") for item in re.split(r"[;.]", match_palavras.group(1))]
        palavras = [item for item in palavras if item][:12]
    return {
        "resumo": resumo or None,
        "palavras_chave": palavras,
        "abstract": abstract or None,
        "objetivos": trecho_por_termos(texto, ["objetivo geral", "objetiva analisar", "tem como objetivo"], 1800) or None,
        "metodologia": trecho_por_termos(texto, ["procedimentos metodológicos", "metodologia", "pesquisa qualitativa"], 2200) or None,
        "produto_educacional_extraido": trecho_por_termos(texto, ["produto educacional"], 1800) or None,
        "publico_alvo_extraido": trecho_por_termos(texto, ["público-alvo", "público alvo"], 1200) or None,
        "base_teorica_extraida": trecho_por_termos(texto, ["referencial teórico", "fundamentação teórica"], 2000) or None,
        "linha_pesquisa_encontrada": trecho_por_termos(
            texto, ["linha de pesquisa", "práticas educativas em", "organização e memórias"], 900
        ) or None,
        "macroprojeto_encontrado": trecho_por_termos(
            texto, ["macroprojeto", "mp1", "mp2", "mp3", "mp4", "mp5", "mp6"], 900
        ) or None,
    }


def classificacao_padrao() -> dict:
    return {
        "Linha de Pesquisa": "A revisar",
        "Macroprojeto": "A revisar",
        "Código do Macroprojeto": "A revisar",
        "Base Epistemológica": "A revisar",
        "Nível de Aplicação": "A revisar",
        "Tipo de Produto Educacional": "A revisar",
        "Área Temática": "A revisar",
        "Revisão Manual": "Sim",
        "Observações": "Registro novo sem classificação histórica; nenhuma categoria foi inferida automaticamente.",
    }


def bool_planilha(valor) -> bool:
    return normalizar_busca(valor) in {"true", "sim", "1", "yes"}


def remover_titulo_pessoa(valor) -> str:
    texto = limpar_texto(valor)
    texto = re.sub(
        r"^\s*(?:(?:prof(?:essor|essora)?|prof[ªa]?|dr[ªa]?|dra)\.?\s*[:.-]?\s*)+",
        "",
        texto,
        flags=re.IGNORECASE,
    )
    return texto.strip(" .:;,-")


def escolher_canonico(valores: list[str]) -> str:
    contagem = Counter(limpar_texto(v) for v in valores if limpar_texto(v))
    if not contagem:
        return "A revisar"
    return sorted(
        contagem,
        key=lambda v: (-contagem[v], v.isupper(), len(v), normalizar_busca(v)),
    )[0]


def carregar_entradas() -> tuple[list[dict], dict[str, dict], dict[str, dict]]:
    if not ARQUIVO_ENTRADA.is_file():
        raise FileNotFoundError(
            f"Planilha oficial ausente: {ARQUIVO_ENTRADA}. Execute primeiro 1_PegarDadosSiteIfc.py"
        )
    quadro = pd.read_excel(ARQUIVO_ENTRADA, dtype=str).fillna("")
    obrigatorias = {
        "ID", "Ano", "Título da Dissertação", "Autor", "Orientador",
        "Instituição", "Campus", "Programa", "Produto Educacional",
        "Link do PDF", "Link do Produto", "Nome do PDF", "SHA-256",
        "Vínculo Confirmado",
    }
    faltantes = sorted(obrigatorias - set(quadro.columns))
    if faltantes:
        raise ValueError(f"Colunas ausentes em {ARQUIVO_ENTRADA}: {', '.join(faltantes)}")
    oficiais = quadro.to_dict(orient="records")
    ids = [limpar_texto(item["ID"]) for item in oficiais]
    duplicados = sorted({item for item in ids if item and ids.count(item) > 1})
    if duplicados or any(not re.fullmatch(r"EPT-\d{4,}", item) for item in ids):
        raise ValueError(f"IDs inválidos ou duplicados na planilha oficial: {duplicados}")

    identidades_doc = ler_json(ARQUIVO_IDENTIDADES)
    classificacoes_doc = ler_json(ARQUIVO_CLASSIFICACOES)
    if not identidades_doc or identidades_doc.get("schema_version") != "leme-identidades-v1":
        raise ValueError(f"Identidades inválidas: {ARQUIVO_IDENTIDADES}")
    if not classificacoes_doc or classificacoes_doc.get("schema_version") != SCHEMA_CLASSIFICACOES:
        raise ValueError(f"Classificações históricas inválidas: {ARQUIVO_CLASSIFICACOES}")
    identidades = {item["id"]: item for item in identidades_doc["identidades"]}
    classificacoes = classificacoes_doc["registros"]
    desconhecidos = sorted(set(ids) - set(identidades))
    if desconhecidos:
        raise ValueError(f"IDs da planilha ausentes nas identidades: {desconhecidos}")
    return oficiais, identidades, classificacoes


def criar_vocabulario(oficiais: list[dict], classificacoes: dict[str, dict]) -> dict:
    indice_confirmados = {
        normalizar_busca(alias): canonico
        for canonico, configuracao in ORIENTADORES_CONFIRMADOS.items()
        for alias in [canonico, *configuracao["aliases"]]
    }
    grupos_orientadores: defaultdict[str, list[str]] = defaultdict(list)
    for item in oficiais:
        original = limpar_texto(item.get("Orientador"))
        sem_titulo = remover_titulo_pessoa(original)
        if sem_titulo:
            canonico_confirmado = indice_confirmados.get(normalizar_busca(sem_titulo))
            chave = normalizar_busca(canonico_confirmado or sem_titulo)
            grupos_orientadores[chave].extend([original, sem_titulo])
    orientadores = {}
    for chave, aliases in grupos_orientadores.items():
        limpos = [remover_titulo_pessoa(item) for item in aliases]
        canonico = next(
            (
                nome
                for nome in ORIENTADORES_CONFIRMADOS
                if normalizar_busca(nome) == chave
            ),
            escolher_canonico(limpos),
        )
        confirmacao = ORIENTADORES_CONFIRMADOS.get(canonico)
        aliases_declarados = confirmacao.get("aliases", []) if confirmacao else []
        orientadores[canonico] = {
            "aliases": sorted(
                set(aliases + limpos + aliases_declarados + [canonico]),
                key=normalizar_busca,
            ),
            "confianca": "alta",
            "regra": (
                "equivalência nominal confirmada em fonte institucional do IFC"
                if confirmacao
                else "mesmo nome após remoção de título, caixa e pontuação"
            ),
            **({"evidencias": confirmacao["evidencias"]} if confirmacao else {}),
        }

    categorias = {}
    for campo in [
        "Base Epistemológica", "Nível de Aplicação",
        "Tipo de Produto Educacional", "Área Temática",
    ]:
        grupos: defaultdict[str, list[str]] = defaultdict(list)
        for registro in classificacoes.values():
            valor = limpar_texto(registro.get("dados", {}).get(campo))
            if valor and valor != "A revisar":
                grupos[normalizar_busca(valor)].append(valor)
        if campo == "Nível de Aplicação":
            aliases_por_nivel = {item: [] for item in NIVEIS_CONTROLADOS}
            for aliases in grupos.values():
                for alias in aliases:
                    canonico = classificar_nivel(alias)
                    if canonico != "A revisar":
                        aliases_por_nivel[canonico].append(alias)
            canonicos = {
                canonico: {
                    "aliases": sorted(set(aliases + [canonico]), key=normalizar_busca),
                    "confianca": "alta",
                    "regra": "nível controlado por termos educacionais explícitos",
                }
                for canonico, aliases in aliases_por_nivel.items()
            }
            categorias[campo] = {
                "modo": "niveis_controlados_v1",
                "preservar_nao_mapeados": True,
                "canonicos": canonicos,
            }
        elif campo == "Tipo de Produto Educacional":
            valores_historicos = {
                alias
                for aliases in grupos.values()
                for alias in aliases
            }
            aliases_mapeados = {
                normalizar_busca(alias)
                for canonico, aliases in TIPOS_PRODUTO_CONTROLADOS.items()
                for alias in [canonico, *aliases]
            }
            nao_mapeados = sorted(
                (
                    valor
                    for valor in valores_historicos
                    if normalizar_busca(valor) not in aliases_mapeados
                ),
                key=normalizar_busca,
            )
            categorias[campo] = {
                "modo": VERSAO_TIPOS_PRODUTO,
                "preservar_nao_mapeados": False,
                "canonicos": {
                    canonico: {
                        "aliases": sorted(set([canonico, *aliases]), key=normalizar_busca),
                        "confianca": "alta",
                        "regra": "equivalência semântica confirmada de formato de produto",
                    }
                    for canonico, aliases in TIPOS_PRODUTO_CONTROLADOS.items()
                },
                "pendentes_revisao": [
                    {
                        "valor": valor,
                        "motivo": "tipo histórico ainda sem equivalência controlada",
                    }
                    for valor in nao_mapeados
                ],
            }
        else:
            canonicos = {}
            for aliases in grupos.values():
                canonico = escolher_canonico(aliases)
                canonicos[canonico] = {
                    "aliases": sorted(set(aliases), key=normalizar_busca),
                    "confianca": "alta",
                    "regra": "variação determinística de caixa, acento ou pontuação",
                }
            categorias[campo] = {"preservar_nao_mapeados": True, "canonicos": canonicos}

    return {
        "schema_version": SCHEMA_VOCABULARIOS,
        "criado_em": agora_utc(),
        "politica": "Somente equivalências comprovadas são canônicas; dúvida permanece A revisar.",
        "linhas_de_pesquisa": LINHAS,
        "macroprojetos": MACROPROJETOS,
        "relacao_linha_macro": {
            list(LINHAS)[0]: ["MP1", "MP2", "MP3"],
            list(LINHAS)[1]: ["MP4", "MP5", "MP6"],
        },
        "orientadores": {
            "modo": VERSAO_ALIASES_ORIENTADORES,
            "canonicos": orientadores,
            "possiveis_equivalencias_nao_aplicadas": [],
        },
        "categorias": categorias,
        "publico_alvo": {
            "modo": "classes_amplas_v2",
            "ordem": PUBLICOS_CONTROLADOS,
            "regras_palavras_chave": [
                {
                    "canonico": canonico,
                    "termos": termos,
                    "regra": "palavra-chave controlada no público-alvo histórico",
                }
                for canonico, termos in REGRAS_PUBLICO.items()
            ],
        },
    }


def carregar_ou_criar_vocabulario(oficiais: list[dict], classificacoes: dict[str, dict]) -> dict:
    vocabulario = ler_json(ARQUIVO_VOCABULARIOS)
    if (
        vocabulario
        and vocabulario.get("schema_version") == SCHEMA_VOCABULARIOS
        and vocabulario.get("publico_alvo", {}).get("modo") == "classes_amplas_v2"
        and vocabulario.get("orientadores", {}).get("modo")
        == VERSAO_ALIASES_ORIENTADORES
        and vocabulario.get("categorias", {}).get("Nível de Aplicação", {}).get("modo")
        == "niveis_controlados_v1"
        and vocabulario.get("categorias", {}).get(
            "Tipo de Produto Educacional", {}
        ).get("modo") == VERSAO_TIPOS_PRODUTO
    ):
        return vocabulario
    vocabulario = criar_vocabulario(oficiais, classificacoes)
    salvar_json_atomico(vocabulario, ARQUIVO_VOCABULARIOS)
    return vocabulario


def indice_aliases(configuracoes: dict) -> dict[str, str]:
    indice = {}
    for canonico, configuracao in configuracoes.items():
        aliases = configuracao.get("aliases", []) if isinstance(configuracao, dict) else configuracao
        for alias in [canonico, *aliases]:
            indice[normalizar_busca(alias)] = canonico
    return indice


def padronizar_classificacao(campo: str, valor, vocabulario: dict) -> str:
    texto = limpar_texto(valor) or "A revisar"
    if normalizar_busca(texto) == "a revisar":
        return "A revisar"
    if campo == "Linha de Pesquisa":
        for canonico, aliases in vocabulario["linhas_de_pesquisa"].items():
            if normalizar_busca(texto) in {normalizar_busca(v) for v in [canonico, *aliases]}:
                return canonico
        return "A revisar"
    if campo == "Macroprojeto":
        match = re.search(r"\bMP\s*([1-6])\b", texto, flags=re.IGNORECASE)
        if not match:
            return "A revisar"
        codigo = f"MP{match.group(1)}"
        return f"{codigo} – {vocabulario['macroprojetos'][codigo]}"
    if campo == "Público-alvo":
        normalizado = normalizar_busca(texto)
        encontrados = [
            regra["canonico"]
            for regra in vocabulario["publico_alvo"]["regras_palavras_chave"]
            if any(
                normalizar_busca(termo) in normalizado
                for termo in regra.get("termos", [])
            )
        ]
        return "; ".join(encontrados) if encontrados else "A revisar"
    configuracoes = vocabulario.get("categorias", {}).get(campo, {}).get("canonicos", {})
    return indice_aliases(configuracoes).get(normalizar_busca(texto), "A revisar")


def orientador_canonico(valor, vocabulario: dict) -> str:
    chave = normalizar_busca(remover_titulo_pessoa(valor))
    for canonico, configuracao in vocabulario.get("orientadores", {}).get("canonicos", {}).items():
        aliases = [canonico, *configuracao.get("aliases", [])]
        if chave in {normalizar_busca(remover_titulo_pessoa(alias)) for alias in aliases}:
            return canonico
    return remover_titulo_pessoa(valor) or "A revisar"


def valores_permitidos(vocabulario: dict) -> dict:
    resultado = {
        "linha_pesquisa": list(vocabulario["linhas_de_pesquisa"]),
        "macroprojeto": [
            f"{codigo} – {descricao}"
            for codigo, descricao in vocabulario["macroprojetos"].items()
        ],
        "publico_alvo": list(vocabulario.get("publico_alvo", {}).get("ordem", [])),
    }
    for chave, campo in CAMPOS_GPT.items():
        if chave in resultado:
            continue
        resultado[chave] = list(
            vocabulario.get("categorias", {}).get(campo, {}).get("canonicos", {})
        )
    for chave in resultado:
        resultado[chave] = sorted(set(resultado[chave]), key=normalizar_busca) + ["A revisar"]
    return resultado


def esquema_resposta_exemplo(quantidade: int) -> dict:
    detalhe = {
        "valor_atual": "valor padronizado recebido",
        "valor": "valor permitido ou A revisar",
        "confianca": "alta|media|baixa",
        "evidencia": "trecho curto do contexto recebido",
        "motivo": "justificativa curta",
        "revisao_manual": True,
    }
    return {
        "schema_version": SCHEMA_RESPOSTA_GPT,
        "resumo": {"total_registros": quantidade, "registros_revisao": 0},
        "vocabularios_sugeridos": {},
        "trabalhos": [
            {
                "id": "EPT-0001",
                "orientador_padronizado": "Nome sem título acadêmico",
                "classificacoes": {chave: dict(detalhe) for chave in CAMPOS_GPT},
            }
        ],
    }


def montar_prompt(contexto: dict, quantidade: int) -> str:
    permitidos = json.dumps(contexto["valores_permitidos"], ensure_ascii=False, indent=2)
    esquema = json.dumps(esquema_resposta_exemplo(quantidade), ensure_ascii=False, indent=2)
    return f"""# Classificação do PAINEL EPT / LEME

## Arquivos recebidos

Você recebeu este prompt e `para_gpt.json`, contendo {quantidade} trabalhos do ProfEPT no Instituto Federal Catarinense. O campus é informado individualmente em cada registro.

## Tarefa

1. Leia primeiro o conjunto inteiro, as distribuições e os vocabulários.
2. Só depois classifique cada trabalho.
3. Devolva UM ÚNICO JSON com os {quantidade} trabalhos.
4. Não divida em lotes, não crie vários arquivos e não omita registros.

Padronize orientador, Linha de Pesquisa, Macroprojeto, Base Epistemológica, Nível de Aplicação, Tipo de Produto, Área Temática e Público-alvo. O objetivo é eliminar apenas duplicações reais de grafia ou conceito, sem reduzir categorias diferentes à força.

## Dados imutáveis

Não altere nem tente corrigir ID, título, autor, ano, instituição, campus, programa ou links. `ProfEPT` é o programa e a instituição é `Instituto Federal Catarinense`. O campus recebido foi obtido do eduCAPES quando havia informação explícita; caso contrário, foi mantido o contexto da unidade associada ao programa.

O orientador pode ser devolvido sem Dr., Dra. ou Prof., mas deve continuar sendo comprovadamente a mesma pessoa. Não una homônimos ou nomes apenas parecidos.

## Regras

- Use somente evidências contidas em `para_gpt.json`.
- Preserve uma classificação histórica quando ela continuar coerente.
- Use exatamente a mesma grafia canônica para trabalhos equivalentes.
- Não invente classificação para preencher campo.
- Na dúvida, use `A revisar`, confiança `baixa` e `revisao_manual: true`.
- Confianças aceitas: `alta`, `media`, `baixa`.
- Confiança média ou baixa exige revisão manual.
- Todos os sete campos de `classificacoes` são obrigatórios em todos os trabalhos.
- `valor_atual` deve copiar exatamente o `valor_padronizado` recebido.
- Linha e Macroprojeto precisam ser compatíveis conforme `contexto_global`.
- Em Tipo de Produto, trate `livro digital`, `ebook`, `eBook` e `e-book` como `E-book`.
- Não crie uma categoria nova só porque o produto também menciona oficina, vídeo, guia ou outro suporte; use um valor canônico permitido e explique a escolha.
- Uma sugestão nova de taxonomia deve ficar apenas em `vocabularios_sugeridos`, para revisão humana.

## Valores permitidos

```json
{permitidos}
```

## Verificação antes de responder

- existem exatamente {quantidade} trabalhos;
- cada ID da entrada aparece exatamente uma vez;
- não existe ID novo;
- nenhuma classificação obrigatória está ausente;
- a resposta inteira é JSON válido.

## Formato obrigatório

Responda SOMENTE com JSON válido, sem bloco Markdown e sem texto antes ou depois. Use esta estrutura:

```json
{esquema}
```

Em `resumo.total_registros`, use o número {quantidade}. O arquivo final deve ser salvo pelo usuário como `dados/GPT/gpt.json`.
"""


def trabalho_para_gpt(
    oficial: dict,
    campos: dict,
    historico: dict,
    vocabulario: dict,
) -> dict:
    valores = historico.get("dados", {}) if historico else classificacao_padrao()
    classificacao_atual = {}
    for chave, campo in CAMPOS_GPT.items():
        original = limpar_texto(valores.get(campo)) or "A revisar"
        valor_para_padronizar = original
        if chave == "macroprojeto" and original != "A revisar":
            codigo = limpar_texto(valores.get("Código do Macroprojeto"))
            valor_para_padronizar = f"{codigo} {original}" if codigo else original
        classificacao_atual[chave] = {
            "valor_original": original,
            "valor_padronizado": padronizar_classificacao(
                campo, valor_para_padronizar, vocabulario
            ),
        }
    return {
        "id": limpar_texto(oficial.get("ID")),
        "ano": limpar_texto(oficial.get("Ano")) or None,
        "titulo": limpar_texto(oficial.get("Título da Dissertação")) or None,
        "autor": limpar_texto(oficial.get("Autor")) or None,
        "orientador_original": limpar_texto(oficial.get("Orientador")) or None,
        "orientador_padronizado_atual": orientador_canonico(oficial.get("Orientador"), vocabulario),
        "instituicao": limpar_texto(oficial.get("Instituição")) or INSTITUICAO_CORPUS,
        "campus": limpar_texto(oficial.get("Campus")) or CAMPUS_CORPUS,
        "programa": limpar_texto(oficial.get("Programa")) or PROGRAMA_CORPUS,
        "resumo": campos.get("resumo") or limpar_texto(valores.get("Resumo Sintetizado")) or None,
        "palavras_chave": campos.get("palavras_chave", []),
        "objetivo_geral": campos.get("objetivos"),
        "objetivos_especificos": [],
        "metodologia": campos.get("metodologia"),
        "produto_educacional": limpar_texto(oficial.get("Produto Educacional")) or None,
        "descricao_produto": campos.get("produto_educacional_extraido")
        or limpar_texto(valores.get("Finalidade do Produto"))
        or None,
        "linha_pesquisa_encontrada": campos.get("linha_pesquisa_encontrada"),
        "macroprojeto_encontrado": campos.get("macroprojeto_encontrado"),
        "publico_alvo_extraido": campos.get("publico_alvo_extraido"),
        "base_teorica_extraida": campos.get("base_teorica_extraida"),
        "classificacao_atual": classificacao_atual,
    }


def executar(args: argparse.Namespace) -> dict:
    oficiais, identidades, classificacoes = carregar_entradas()
    vocabulario = carregar_ou_criar_vocabulario(oficiais, classificacoes)
    if args.limite:
        oficiais = oficiais[: args.limite]
    trabalhos = []
    relatorio = []
    for numero, oficial in enumerate(oficiais, 1):
        identificador = limpar_texto(oficial.get("ID"))
        print(f"[{numero}/{len(oficiais)}] OCR {identificador}")
        identidade = identidades[identificador]
        nome_pdf = limpar_texto(oficial.get("Nome do PDF"))
        caminho = PASTA_PDFS / nome_pdf
        status = "erro"
        erro_geral = ""
        campos = extrair_campos("")
        resultado_ocr = None
        if not bool_planilha(oficial.get("Vínculo Confirmado")):
            erro_geral = "vínculo entre trabalho e PDF não confirmado"
        elif not caminho.is_file():
            erro_geral = f"PDF ausente: {caminho}"
        else:
            hash_atual = sha256_arquivo(caminho)
            hash_esperado = limpar_texto(identidade.get("sha256"))
            if not hash_esperado or hash_atual != hash_esperado:
                erro_geral = "SHA-256 do PDF diverge da identidade permanente"
            else:
                try:
                    max_paginas_registro = (
                        min(args.max_paginas_ocr, MAX_PAGINAS_OCR_HISTORICO)
                        if identificador in classificacoes and args.max_paginas_ocr > 0
                        else args.max_paginas_ocr
                    )
                    resultado_ocr = extrair_ocr_com_cache(
                        identificador,
                        caminho,
                        max_paginas_registro,
                        args.workers_ocr,
                        args.forcar_ocr,
                    )
                    campos = extrair_campos(resultado_ocr["texto"])
                    status = "sucesso" if resultado_ocr["texto"] else "erro"
                    if not resultado_ocr["texto"]:
                        erro_geral = "OCR não produziu texto"
                except Exception as erro:
                    erro_geral = f"{type(erro).__name__}: {erro}"
        trabalhos.append(
            trabalho_para_gpt(
                oficial, campos, classificacoes.get(identificador, {}), vocabulario
            )
        )
        relatorio.append(
            {
                "id": identificador,
                "status": status,
                "paginas_processadas": (resultado_ocr or {}).get("paginas_processadas", 0),
                "paginas_total_pdf": (resultado_ocr or {}).get("paginas_total_pdf", 0),
                "erros_paginas": (resultado_ocr or {}).get("erros_paginas", []),
                "erro": erro_geral,
            }
        )

    ids = [item["id"] for item in trabalhos]
    duplicados = sorted({item for item in ids if ids.count(item) > 1})
    if duplicados:
        raise ValueError(f"IDs duplicados no JSON para GPT: {duplicados}")
    campi = sorted(
        {
            limpar_texto(oficial.get("Campus")) or CAMPUS_CORPUS
            for oficial in oficiais
        },
        key=normalizar_busca,
    )
    contexto = {
        "corpus": {
            "instituicao": INSTITUICAO_CORPUS,
            "campi": campi,
            "programa": PROGRAMA_CORPUS,
        },
        "orientadores_canonicos": sorted(
            vocabulario["orientadores"]["canonicos"], key=normalizar_busca
        ),
        "linhas_de_pesquisa_validas": list(vocabulario["linhas_de_pesquisa"]),
        "macroprojetos_validos": {
            codigo: f"{codigo} – {descricao}"
            for codigo, descricao in vocabulario["macroprojetos"].items()
        },
        "relacao_linha_macro": vocabulario["relacao_linha_macro"],
        "valores_permitidos": valores_permitidos(vocabulario),
        "distribuicoes_atuais": {
            chave: dict(Counter(
                item["classificacao_atual"][chave]["valor_padronizado"]
                for item in trabalhos
            ).most_common())
            for chave in CAMPOS_GPT
        },
    }
    documento = {
        "schema_version": SCHEMA_ENTRADA_GPT,
        "gerado_em": agora_utc(),
        "contexto_global": contexto,
        "quantidade": len(trabalhos),
        "trabalhos": trabalhos,
        "empacotamento": {
            "modo_resposta": "arquivo_unico",
            "arquivo_resposta_esperado": str(ARQUIVO_RESPOSTA_GPT),
            "schema_resposta": SCHEMA_RESPOSTA_GPT,
        },
    }
    texto_documento = json.dumps(documento, ensure_ascii=False, indent=2)
    documento["metricas"] = {
        "bytes_utf8_aproximados": len(texto_documento.encode("utf-8")),
        "megabytes_aproximados": round(len(texto_documento.encode("utf-8")) / 1048576, 3),
        "caracteres_aproximados": len(texto_documento),
    }
    salvar_json_atomico(documento, ARQUIVO_SAIDA)
    salvar_texto_atomico(montar_prompt(contexto, len(trabalhos)), ARQUIVO_PROMPT)
    linhas_relatorio = [
        "RELATÓRIO OCR — PAINEL EPT / LEME",
        f"Gerado em: {agora_utc()}",
        f"Registros: {len(relatorio)}",
        f"OCR com sucesso: {sum(item['status'] == 'sucesso' for item in relatorio)}",
        f"OCR com erro: {sum(item['status'] != 'sucesso' for item in relatorio)}",
        "",
    ]
    for item in relatorio:
        if item["status"] != "sucesso" or item["erros_paginas"]:
            linhas_relatorio.append(
                f"{item['id']}: {item['status']} | páginas {item['paginas_processadas']}/{item['paginas_total_pdf']}"
            )
            if item["erro"]:
                linhas_relatorio.append(f"  erro: {item['erro']}")
            for erro in item["erros_paginas"]:
                linhas_relatorio.append(f"  página {erro['pagina']}: {erro['erro']}")
    salvar_texto_atomico("\n".join(linhas_relatorio), ARQUIVO_RELATORIO_OCR)
    documento["resumo_processamento"] = {
        "quantidade": len(trabalhos),
        "ocr_sucesso": sum(item["status"] == "sucesso" for item in relatorio),
        "ocr_com_erro": sum(item["status"] != "sucesso" for item in relatorio),
        "ids_unicos": len(set(ids)),
        "json_bytes": ARQUIVO_SAIDA.stat().st_size,
    }
    return documento


def main() -> int:
    args = argumentos()
    resultado = executar(args)
    print(json.dumps(resultado["resumo_processamento"], ensure_ascii=False, indent=2))
    print(f"JSON para GPT: {ARQUIVO_SAIDA}")
    print(f"Prompt: {ARQUIVO_PROMPT}")
    return 0 if resultado["resumo_processamento"]["ocr_com_erro"] == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
