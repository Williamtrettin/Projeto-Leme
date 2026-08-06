import pandas as pd
import json
import re
import unicodedata
from datetime import datetime
from pathlib import Path

# ============================================================
# PROJETO LEME/EPT - ETAPA 05
# Geração do painel HTML para colar no Elementor
#
# Este script foi adaptado para a planilha:
# dadosLeme_arrumado.xlsx
#
# Ele:
# - lê a planilha arrumada;
# - aceita variações nos nomes das colunas;
# - NÃO recalcula Linha de Pesquisa, Macroprojeto ou Base Epistemológica;
# - NÃO altera nem salva planilha;
# - gera somente o arquivo HTML final do painel.
# ============================================================

ARQUIVO_ENTRADA = "dadosLeme_arrumado.xlsx"
ARQUIVO_HTML_SAIDA = "5_PainelLemeEpt_dadosLeme_v8_responsivo.html"

# Campos usados pelo JSON do painel e possíveis nomes na planilha.
# A busca das colunas ignora maiúsculas/minúsculas, acentos e espaços duplicados.
MAPEAMENTO_COLUNAS = {
    "Ano": ["Ano", "Ano Limpo", "Ano de Defesa", "Ano Defesa"],
    "Autor": ["Autor", "Autora", "Discente", "Mestrando", "Mestranda"],
    "Orientador": ["Orientador", "Orientadora", "Professor Orientador", "Professora Orientadora"],
    "Título da Dissertação": ["Título da Dissertação", "Titulo da Dissertação", "Título Provável", "Titulo Provavel", "Titulo", "Título", "Nome do Trabalho"],
    "Produto Educacional": ["Produto Educacional", "Produto", "Título do Produto", "Titulo do Produto", "Nome do Produto", "Produto Educacional Provável"],
    "Link do PDF": ["Link do PDF", "Link PDF", "PDF", "Link da Dissertação", "Link Dissertação", "Link da Dissertacao"],
    "Link do Produto": ["Link do Produto", "Link do Produto Educacional", "Link Produto", "Produto Link", "Link PE"],
    "Palavras-chave": ["Palavras-chave", "Palavras-Chave", "Palavras chave", "Palavras Chave", "Keywords"],
    "Tipo de Produto Original": ["Tipo de Produto", "Tipo de Produto Original", "Tipo Original"],
    "Tipo de Produto": ["Tipo de Produto V2", "Tipo de Produto", "Tipo de Produto Educacional", "Tipo Produto", "Tipo"],
    "Resumo": ["Resumo Sintetizado", "Resumo", "Resumo do Trabalho", "Resumo Limpo"],
    "Resumo Original": ["Resumo", "Resumo Sintetizado", "Resumo do Trabalho", "Resumo Limpo"],
    "Nível de Aplicação": ["Nível de Aplicação V2", "Nivel de Aplicação V2", "Nível de Aplicação", "Nivel de Aplicacao", "Nível", "Nivel"],
    "Base Epistemológica": ["Base Epistemológica V2", "Base Epistemologica V2", "Base Epistemológica", "Base Epistemologica"],
    "Confiança da Base": ["Confiança da Base", "Confianca da Base", "Confiança Base Epistemológica", "Confianca Base Epistemologica"],
    "Termos Detectados da Base": ["Termos Detectados da Base", "Termos Detectados Base", "Termos Base", "Termos Detectados Base Epistemológica"],
    "Área Temática": ["Área Temática", "Area Temática", "Área Tematica", "Area Tematica", "Área", "Area"],
    "Finalidade do Produto": ["Finalidade do Produto", "Finalidade", "Objetivo do Produto", "Uso do Produto"],
    "Público-alvo": ["Público-alvo", "Publico-alvo", "Publico Alvo", "Público Alvo", "Público", "Publico"],
    "Instituição/Campus": ["Instituição/Campus", "Instituicao/Campus", "Instituição", "Instituicao", "Campus", "IFC Campus"],
    "Linha de Pesquisa": ["Linha de Pesquisa", "Linha Pesquisa", "Linha"],
    "Macroprojeto": ["Macroprojeto", "Macro Projeto", "Macro-projeto"],
    "Código do Macroprojeto": ["Código do Macroprojeto", "Codigo do Macroprojeto", "Código Macroprojeto", "Codigo Macroprojeto", "Código", "Codigo"],
    "Confiança Linha/Macro": ["Confiança Linha/Macro", "Confianca Linha/Macro", "Confiança Linha Macro", "Confianca Linha Macro"],
    "Termos Detectados Linha/Macro": ["Termos Detectados Linha/Macro", "Termos Detectados Linha Macro", "Termos Linha/Macro", "Termos Linha Macro"],
    "Justificativa Linha/Macro": ["Justificativa Linha/Macro", "Justificativa Linha Macro", "Justificativa Macro", "Justificativa"],
    "Revisão Manual": ["Revisão Manual", "Revisao Manual", "Revisão", "Revisao"],
}


def limpar_valor(valor, padrao="---"):
    """Converte valores vazios/nulos para um padrão seguro."""
    if pd.isna(valor):
        return padrao
    valor = str(valor).strip()
    if not valor or valor.lower() in ["nan", "none", "null", "nat"]:
        return padrao
    return valor


def normalizar_nome_coluna(texto):
    """Normaliza nomes de colunas para comparação flexível."""
    texto = limpar_valor(texto, "")
    texto = texto.lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def normalizar_texto(texto):
    texto = limpar_valor(texto, "")
    texto = texto.lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def montar_indice_colunas(colunas):
    """Cria um índice {coluna_normalizada: coluna_original}."""
    indice = {}
    for coluna in colunas:
        chave = normalizar_nome_coluna(coluna)
        if chave and chave not in indice:
            indice[chave] = coluna
    return indice


def encontrar_coluna(colunas_disponiveis, *possiveis_nomes):
    """Encontra uma coluna ignorando acentos, caixa e espaços."""
    indice = montar_indice_colunas(colunas_disponiveis)
    for nome in possiveis_nomes:
        chave = normalizar_nome_coluna(nome)
        if chave in indice:
            return indice[chave]
    return None


def primeiro_valor(row, *colunas, padrao="---"):
    """Pega o primeiro valor válido entre várias possibilidades de coluna."""
    indice = montar_indice_colunas(row.keys())

    for coluna in colunas:
        chave = normalizar_nome_coluna(coluna)
        coluna_real = indice.get(chave)
        if not coluna_real:
            continue

        valor = limpar_valor(row.get(coluna_real), "")
        if valor and valor != "---":
            return valor

    return padrao


def valor_campo(row, nome_campo, padrao="---"):
    """Busca um campo pelo mapeamento oficial do script."""
    opcoes = MAPEAMENTO_COLUNAS.get(nome_campo, [nome_campo])
    return primeiro_valor(row, *opcoes, padrao=padrao)


def formatar_nome_orientador(valor):
    """Mantém o nome do orientador como está na planilha, só removendo sujeiras leves."""
    nome = limpar_valor(valor, "")
    if not nome or nome == "---":
        return "---"
    nome = re.sub(r"[:;]+", "", nome)
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome if nome else "---"


def chave_orientador(valor):
    """Cria chave normalizada para contar orientadores sem duplicidade."""
    nome = formatar_nome_orientador(valor)
    if nome == "---":
        return ""

    nome = unicodedata.normalize("NFD", nome)
    nome = "".join(c for c in nome if unicodedata.category(c) != "Mn")
    nome = nome.lower()
    nome = re.sub(r"\bdr[aª]?\b\.?", " ", nome, flags=re.IGNORECASE)
    nome = re.sub(r"\bdra\b\.?", " ", nome, flags=re.IGNORECASE)
    nome = re.sub(r"\bprof[aª]?\b\.?", " ", nome, flags=re.IGNORECASE)
    nome = re.sub(r"\bprofa\b\.?", " ", nome, flags=re.IGNORECASE)
    nome = re.sub(r"[^a-z0-9\s]", " ", nome)
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome


def ano_limpo(valor):
    texto = limpar_valor(valor, "")
    m = re.search(r"(19\d{2}|20\d{2})", texto)
    return m.group(1) if m else "---"


def dividir_tags(texto):
    texto = limpar_valor(texto, "")
    if not texto:
        return []

    partes = re.split(r"[;,|\n]+", texto)
    tags = []
    vistos = set()

    for parte in partes:
        tag = re.sub(r"\s+", " ", parte).strip(" .;,-")
        if not tag or tag == "---":
            continue

        chave = normalizar_texto(tag)
        if chave not in vistos:
            vistos.add(chave)
            tags.append(tag)

    return tags


def limitar_texto(texto, limite=280):
    texto = limpar_valor(texto, "")
    if not texto:
        return "---"
    if len(texto) <= limite:
        return texto
    return texto[:limite].rsplit(" ", 1)[0] + "..."


def link_valido(link):
    link = limpar_valor(link, "")
    return link.startswith("http://") or link.startswith("https://")


def avisar_colunas_ausentes(df):
    """Mostra avisos no terminal, mas não interrompe a geração do painel."""
    print("\nConferindo colunas da planilha...")
    ausentes = []

    for campo, alternativas in MAPEAMENTO_COLUNAS.items():
        if encontrar_coluna(df.columns, *alternativas) is None:
            ausentes.append(campo)

    if ausentes:
        print("ATENÇÃO: estes campos não foram encontrados e serão preenchidos com '---':")
        for campo in ausentes:
            print(f"- {campo}")
    else:
        print("Todas as colunas principais foram localizadas ou mapeadas.")


def carregar_planilha(arquivo=ARQUIVO_ENTRADA):
    caminho = Path(arquivo)
    if not caminho.exists():
        raise FileNotFoundError(
            f"Planilha de entrada não encontrada: {arquivo}\n"
            "Deixe o arquivo dadosLeme_arrumado.xlsx na mesma pasta deste script."
        )

    df = pd.read_excel(caminho)
    avisar_colunas_ausentes(df)
    return df.fillna("---")


def preparar_dados_para_html(df):
    registros = []

    for _, row_series in df.iterrows():
        row = row_series.to_dict()

        ano = valor_campo(row, "Ano")
        orientador = formatar_nome_orientador(valor_campo(row, "Orientador", ""))
        orientador_chave = chave_orientador(orientador)
        link_pdf = valor_campo(row, "Link do PDF", "")
        link_produto = valor_campo(row, "Link do Produto", "")
        resumo = valor_campo(row, "Resumo", "---")

        registro = {
            "ano": ano_limpo(ano),
            "autor": valor_campo(row, "Autor"),
            "orientador": orientador,
            "orientadorChave": orientador_chave,
            "titulo": valor_campo(row, "Título da Dissertação"),
            "produto": valor_campo(row, "Produto Educacional"),
            "linkPdf": link_pdf if link_valido(link_pdf) else "",
            "linkProduto": link_produto if link_valido(link_produto) else "",
            "palavrasChave": dividir_tags(valor_campo(row, "Palavras-chave", "")),
            "tipoProdutoOriginal": valor_campo(row, "Tipo de Produto Original"),
            "tipoProduto": valor_campo(row, "Tipo de Produto"),
            "resumo": resumo,
            "resumoCurto": limitar_texto(resumo, 280),
            "nivel": valor_campo(row, "Nível de Aplicação"),
            "base": valor_campo(row, "Base Epistemológica"),
            "confiancaBase": valor_campo(row, "Confiança da Base"),
            "termosBase": valor_campo(row, "Termos Detectados da Base"),
            "area": valor_campo(row, "Área Temática"),
            "finalidade": valor_campo(row, "Finalidade do Produto"),
            "publicoAlvo": valor_campo(row, "Público-alvo"),
            "instituicaoCampus": valor_campo(row, "Instituição/Campus"),
            "linhaPesquisa": valor_campo(row, "Linha de Pesquisa"),
            "macroprojeto": valor_campo(row, "Macroprojeto"),
            "codigoMacroprojeto": valor_campo(row, "Código do Macroprojeto"),
            "confiancaLinhaMacro": valor_campo(row, "Confiança Linha/Macro"),
            "termosLinhaMacro": valor_campo(row, "Termos Detectados Linha/Macro"),
            "justificativaLinhaMacro": valor_campo(row, "Justificativa Linha/Macro"),
            "revisaoManual": valor_campo(row, "Revisão Manual"),
        }

        registros.append(registro)

    return registros


def gerar_html(registros):
    dados_json = json.dumps(registros, ensure_ascii=False).replace("</", "<\\/")

    total = len(registros)
    total_orientadores = len({
        r.get("orientadorChave", "")
        for r in registros
        if r.get("orientadorChave", "")
    })
    total_bases = len({r["base"] for r in registros if r.get("base") and r["base"] != "---"})
    total_produtos = sum(1 for r in registros if r.get("linkProduto"))

    anos_validos = sorted([r["ano"] for r in registros if r.get("ano") and r["ano"] != "---"])
    periodo = f"{anos_validos[0]}–{anos_validos[-1]}" if anos_validos else "---"
    atualizado_em = datetime.now().strftime("%d/%m/%Y")

    html = r'''<!-- =========================================================
PAINEL LEME EPT V8
Gerado automaticamente por Python
Cole TODO este bloco no Elementor > Editor de Texto ou HTML
Melhorias: filtros dinâmicos, gráficos em HTML, modal de detalhes, navegação interna e acessibilidade.
========================================================= -->

<section id="leme-painel-ept" class="leme-painel-ept">
  <style>
    #leme-painel-ept {
      --leme-primary: #3730a3;
      --leme-primary-2: #4f46e5;
      --leme-primary-soft: #eef2ff;
      --leme-cyan: #06b6d4;
      --leme-green: #16a34a;
      --leme-green-soft: #dcfce7;
      --leme-red: #e11d48;
      --leme-red-soft: #fff1f2;
      --leme-bg: #f8fafc;
      --leme-card: rgba(255,255,255,.96);
      --leme-border: #e5e7eb;
      --leme-border-strong: #cfd4ff;
      --leme-text: #111827;
      --leme-muted: #64748b;
      --leme-shadow: 0 18px 50px rgba(15, 23, 42, .10);
      --leme-shadow-soft: 0 10px 28px rgba(15, 23, 42, .07);
      font-family: "Inter", "Segoe UI", Arial, sans-serif;
      color: var(--leme-text);
      background:
        radial-gradient(circle at top left, rgba(79,70,229,.16), transparent 34%),
        radial-gradient(circle at top right, rgba(6,182,212,.12), transparent 30%),
        linear-gradient(180deg, #ffffff, var(--leme-bg));
      border: 1px solid var(--leme-border);
      border-radius: 28px;
      padding: 26px;
      box-sizing: border-box;
      overflow: hidden;
      scroll-behavior: smooth;
    }

    #leme-painel-ept * { box-sizing: border-box; }
    #leme-painel-ept :focus-visible { outline: 4px solid rgba(79,70,229,.23); outline-offset: 3px; border-radius: 12px; }

    #leme-painel-ept .leme-skip {
      position: absolute;
      left: -999px;
      top: auto;
      width: 1px;
      height: 1px;
      overflow: hidden;
    }
    #leme-painel-ept .leme-skip:focus {
      position: static;
      width: auto;
      height: auto;
      display: inline-flex;
      margin-bottom: 12px;
      padding: 10px 14px;
      background: var(--leme-primary);
      color: #fff;
      border-radius: 999px;
      font-weight: 850;
      text-decoration: none;
    }

    #leme-painel-ept .leme-header {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 20px;
      align-items: start;
      margin-bottom: 18px;
    }
    #leme-painel-ept .leme-eyebrow {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 7px 12px;
      border-radius: 999px;
      background: var(--leme-primary-soft);
      color: var(--leme-primary);
      font-size: 12px;
      font-weight: 850;
      letter-spacing: .02em;
      margin-bottom: 10px;
    }
    #leme-painel-ept h2 {
      margin: 0;
      color: var(--leme-primary);
      font-size: clamp(25px, 3vw, 38px);
      line-height: 1.08;
      font-weight: 900;
      letter-spacing: -.04em;
    }
    #leme-painel-ept .leme-subtitle {
      margin: 10px 0 0;
      color: var(--leme-muted);
      font-size: 15px;
      line-height: 1.6;
      max-width: 940px;
    }
    #leme-painel-ept .leme-update {
      background: rgba(255,255,255,.86);
      border: 1px solid var(--leme-border-strong);
      border-radius: 18px;
      padding: 13px 15px;
      font-size: 12px;
      color: var(--leme-muted);
      white-space: nowrap;
      box-shadow: var(--leme-shadow-soft);
      backdrop-filter: blur(10px);
    }

    #leme-painel-ept .leme-nav {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin: 0 0 18px;
      padding: 10px;
      border-radius: 22px;
      background: rgba(255,255,255,.72);
      border: 1px solid var(--leme-border);
      box-shadow: var(--leme-shadow-soft);
      position: sticky;
      top: 8px;
      z-index: 15;
      backdrop-filter: blur(12px);
    }
    #leme-painel-ept .leme-nav button,
    #leme-painel-ept .leme-mini-nav button,
    #leme-painel-ept .leme-head-btn,
    #leme-painel-ept .leme-bottom-nav button,
    #leme-painel-ept .leme-export-btn,
    #leme-painel-ept .leme-view button,
    #leme-painel-ept .leme-page-btn,
    #leme-painel-ept .leme-card-detail-btn,
    #leme-painel-ept .leme-chip button,
    #leme-painel-ept .leme-modal-close {
      border: 1px solid var(--leme-border-strong);
      border-radius: 999px;
      background: #fff;
      color: var(--leme-primary);
      padding: 10px 14px;
      cursor: pointer;
      font-size: 12px;
      font-weight: 850;
      min-height: 40px;
      transition: transform .16s ease, background .16s ease, border-color .16s ease, color .16s ease, box-shadow .16s ease;
    }
    #leme-painel-ept .leme-nav button:hover,
    #leme-painel-ept .leme-mini-nav button:hover,
    #leme-painel-ept .leme-head-btn:hover,
    #leme-painel-ept .leme-bottom-nav button:hover,
    #leme-painel-ept .leme-export-btn:hover,
    #leme-painel-ept .leme-view button:hover,
    #leme-painel-ept .leme-page-btn:hover,
    #leme-painel-ept .leme-card-detail-btn:hover,
    #leme-painel-ept .leme-modal-close:hover {
      background: var(--leme-primary-soft);
      transform: translateY(-1px);
      box-shadow: 0 8px 18px rgba(79,70,229,.10);
    }
    #leme-painel-ept .leme-nav button.is-primary,
    #leme-painel-ept .leme-card-detail-btn,
    #leme-painel-ept .leme-view button.is-active,
    #leme-painel-ept .leme-page-btn.is-active {
      background: var(--leme-primary-2);
      color: #fff;
      border-color: var(--leme-primary-2);
    }

    #leme-painel-ept .leme-stats {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 14px;
      margin-bottom: 18px;
    }
    #leme-painel-ept .leme-stat,
    #leme-painel-ept .leme-insight-card {
      background: var(--leme-card);
      border: 1px solid var(--leme-border);
      border-radius: 22px;
      padding: 18px;
      box-shadow: var(--leme-shadow-soft);
      position: relative;
      overflow: hidden;
    }
    #leme-painel-ept .leme-stat::after,
    #leme-painel-ept .leme-insight-card::after {
      content: "";
      position: absolute;
      inset: auto -30px -42px auto;
      width: 105px;
      height: 105px;
      border-radius: 999px;
      background: rgba(79,70,229,.09);
      pointer-events: none;
    }
    #leme-painel-ept .leme-stat strong,
    #leme-painel-ept .leme-insight-card strong {
      display: block;
      color: var(--leme-primary);
      font-size: 29px;
      line-height: 1;
      margin-bottom: 7px;
      letter-spacing: -.03em;
    }
    #leme-painel-ept .leme-stat span,
    #leme-painel-ept .leme-insight-card span {
      color: var(--leme-muted);
      font-size: 12px;
      font-weight: 750;
      line-height: 1.35;
    }

    #leme-painel-ept .leme-panel {
      margin: 0 0 20px;
      background: rgba(255,255,255,.95);
      border: 1px solid var(--leme-border);
      border-radius: 26px;
      padding: 20px;
      box-shadow: var(--leme-shadow);
      backdrop-filter: blur(12px);
      scroll-margin-top: 88px;
    }
    #leme-painel-ept .leme-panel-head {
      display: flex;
      justify-content: space-between;
      gap: 14px;
      align-items: flex-start;
      margin-bottom: 16px;
      flex-wrap: wrap;
    }
    #leme-painel-ept .leme-panel-head h3,
    #leme-painel-ept .leme-section-label {
      margin: 0;
      color: var(--leme-primary);
      font-size: 22px;
      line-height: 1.2;
      font-weight: 900;
      letter-spacing: -.025em;
    }
    #leme-painel-ept .leme-panel-head p {
      margin: 6px 0 0;
      color: var(--leme-muted);
      font-size: 13px;
      line-height: 1.55;
      max-width: 800px;
    }
    #leme-painel-ept .leme-note {
      background: var(--leme-primary-soft);
      color: var(--leme-primary);
      border: 1px solid #d9ddff;
      border-radius: 999px;
      padding: 9px 12px;
      font-size: 12px;
      font-weight: 850;
      white-space: nowrap;
    }
    #leme-painel-ept .leme-panel-head-actions {
      display: flex;
      align-items: center;
      justify-content: flex-end;
      gap: 8px;
      flex-wrap: wrap;
    }
    #leme-painel-ept .leme-mini-nav { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 14px; }
    #leme-painel-ept .leme-bottom-nav {
      display: flex;
      justify-content: center;
      gap: 8px;
      flex-wrap: wrap;
      margin: 22px 0 0;
      padding: 14px;
      border: 1px solid var(--leme-border);
      border-radius: 22px;
      background: rgba(255,255,255,.82);
      box-shadow: var(--leme-shadow-soft);
    }

    #leme-painel-ept .leme-insights {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
      margin-bottom: 16px;
    }
    #leme-painel-ept .leme-insight-card strong { font-size: 21px; line-height: 1.18; word-break: break-word; }

    #leme-painel-ept .leme-charts-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 16px;
    }
    #leme-painel-ept .leme-chart-card {
      border: 1px solid var(--leme-border);
      border-radius: 22px;
      padding: 16px;
      background: linear-gradient(180deg, #ffffff, #fbfdff);
      min-height: 220px;
      box-shadow: 0 8px 20px rgba(15, 23, 42, .045);
    }
    #leme-painel-ept .leme-chart-card.is-wide { grid-column: 1 / -1; }
    #leme-painel-ept .leme-chart-title { margin: 0 0 4px; color: var(--leme-text); font-size: 15px; font-weight: 900; }
    #leme-painel-ept .leme-chart-subtitle { margin: 0 0 14px; color: var(--leme-muted); font-size: 12px; line-height: 1.45; }
    #leme-painel-ept .leme-bars { display: grid; gap: 10px; }
    #leme-painel-ept .leme-bar-row { display: grid; grid-template-columns: minmax(160px, 36%) 1fr auto; gap: 10px; align-items: center; }
    #leme-painel-ept .leme-bar-label { color: #334155; font-size: 12px; line-height: 1.25; font-weight: 800; overflow-wrap: anywhere; }
    #leme-painel-ept .leme-bar-track { height: 13px; border-radius: 999px; background: #eef2ff; overflow: hidden; border: 1px solid #e0e7ff; }
    #leme-painel-ept .leme-bar-fill { height: 100%; width: 0%; border-radius: 999px; background: linear-gradient(90deg, var(--leme-primary-2), var(--leme-cyan)); transition: width .3s ease; }
    #leme-painel-ept .leme-bar-value { color: var(--leme-primary); font-size: 12px; font-weight: 900; min-width: 24px; text-align: right; }
    #leme-painel-ept .leme-year-chart { display: flex; align-items: end; gap: 12px; height: var(--chart-height, 220px); padding: 12px 6px 0; border-bottom: 1px solid #e2e8f0; }
    #leme-painel-ept .leme-year-col { flex: 1; min-width: 42px; display: flex; flex-direction: column; align-items: center; justify-content: end; height: 100%; gap: 7px; }
    #leme-painel-ept .leme-year-value { color: var(--leme-primary); font-weight: 900; font-size: 12px; }
    #leme-painel-ept .leme-year-bar { width: min(52px, 78%); min-height: 6px; border-radius: 12px 12px 4px 4px; background: linear-gradient(180deg, var(--leme-primary-2), var(--leme-cyan)); box-shadow: 0 8px 18px rgba(79,70,229,.14); }
    #leme-painel-ept .leme-year-label { color: #475569; font-size: 12px; font-weight: 850; }
    #leme-painel-ept .leme-single-insight,
    #leme-painel-ept .leme-empty-chart {
      border: 1px dashed var(--leme-border-strong);
      background: linear-gradient(180deg, #ffffff, #fbfdff);
      border-radius: 18px;
      padding: 18px;
      color: var(--leme-muted);
      font-size: 13px;
      line-height: 1.55;
      min-height: 116px;
      display: grid;
      align-content: center;
    }
    #leme-painel-ept .leme-empty-chart::before {
      content: "Sem dados neste recorte";
      display: block;
      color: var(--leme-primary);
      font-weight: 900;
      font-size: 15px;
      margin-bottom: 5px;
    }
    #leme-painel-ept .leme-single-insight strong { color: var(--leme-primary); font-size: 18px; display:block; margin-bottom: 6px; }

    #leme-painel-ept .leme-search-row { display: grid; grid-template-columns: 1fr auto; gap: 12px; margin: 14px 0; }
    #leme-painel-ept .leme-search {
      width: 100%; border: 1px solid var(--leme-border-strong); border-radius: 18px; padding: 15px 16px;
      font-size: 15px; color: var(--leme-text); outline: none; background: #fff; min-height: 48px;
    }
    #leme-painel-ept .leme-search:focus { border-color: var(--leme-primary-2); box-shadow: 0 0 0 4px rgba(79,70,229,.12); }
    #leme-painel-ept .leme-clear {
      border: none; border-radius: 18px; padding: 0 17px; min-height: 48px;
      background: linear-gradient(135deg, var(--leme-primary-2), var(--leme-cyan)); color: #fff; font-weight: 850; cursor: pointer;
      box-shadow: 0 12px 22px rgba(79,70,229,.16);
    }
    #leme-painel-ept .leme-filters { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 11px; }
    #leme-painel-ept .leme-field { position: relative; }
    #leme-painel-ept .leme-field label { display:flex; align-items:center; justify-content:space-between; gap:8px; font-size: 11px; font-weight: 850; color: var(--leme-muted); margin: 0 0 6px; text-transform: uppercase; letter-spacing:.045em; }
    #leme-painel-ept .leme-field select { width:100%; border:1px solid var(--leme-border); border-radius:15px; padding: 11px 12px; background:#fff; color: var(--leme-text); font-size:13px; outline:none; min-height:44px; }
    #leme-painel-ept .leme-field select:disabled { background:#f8fafc; color:#94a3b8; cursor:not-allowed; }
    #leme-painel-ept .leme-field-count {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: max-content;
      max-width: 100%;
      margin-top: 6px;
      padding: 4px 8px;
      border-radius: 999px;
      background: #f8fafc;
      border: 1px solid #e2e8f0;
      color: #64748b;
      font-size: 11px;
      font-weight: 850;
      line-height: 1.2;
    }
    #leme-painel-ept .leme-field-count.is-selected {
      background: var(--leme-primary-soft);
      border-color: #d9ddff;
      color: var(--leme-primary);
    }
    #leme-painel-ept .leme-active-filters { display:flex; flex-wrap:wrap; gap:8px; margin: 12px 0 0; }
    #leme-painel-ept .leme-chip { display:inline-flex; align-items:center; gap:6px; padding: 6px 8px 6px 11px; border:1px solid #d9ddff; border-radius:999px; background:var(--leme-primary-soft); color:var(--leme-primary); font-size:12px; font-weight:850; }
    #leme-painel-ept .leme-chip button { min-height: 26px; padding:0 8px; border-color:transparent; background:#fff; }

    #leme-painel-ept .leme-toolbar { display:flex; justify-content:space-between; align-items:center; gap:12px; margin: 16px 0; flex-wrap:wrap; scroll-margin-top: 88px; }
    #leme-painel-ept .leme-count { font-size:14px; color:var(--leme-muted); font-weight:750; }
    #leme-painel-ept .leme-count strong { color:var(--leme-primary); }
    #leme-painel-ept .leme-actions-toolbar { display:flex; gap:8px; align-items:center; flex-wrap:wrap; }
    #leme-painel-ept .leme-view { display:inline-flex; gap:8px; align-items:center; }
    #leme-painel-ept .leme-page-btn:disabled { opacity:.45; cursor:not-allowed; transform:none; }

    #leme-painel-ept .leme-results { display:grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; align-items:stretch; grid-auto-rows: 1fr; }
    #leme-painel-ept .leme-results.is-table { display:block; overflow-x:auto; background:#fff; border:1px solid var(--leme-border); border-radius:22px; box-shadow:var(--leme-shadow-soft); }
    #leme-painel-ept .leme-card { min-height: 365px; height: 100%; background:var(--leme-card); border:1px solid var(--leme-border); border-radius:24px; padding:20px; box-shadow:var(--leme-shadow-soft); transition: transform .16s ease, box-shadow .16s ease, border-color .16s ease; display:flex; flex-direction:column; }
    #leme-painel-ept .leme-card:hover { transform:translateY(-3px); box-shadow:0 20px 44px rgba(15,23,42,.12); border-color:rgba(79,70,229,.34); }
    #leme-painel-ept .leme-card-title { margin:0 0 12px; color:var(--leme-primary); font-size:16px; line-height:1.35; font-weight:900; letter-spacing:-.01em; display:-webkit-box; -webkit-line-clamp:3; -webkit-box-orient:vertical; overflow:hidden; }
    #leme-painel-ept .leme-meta { display:grid; gap:7px; margin:0 0 12px; color:var(--leme-muted); font-size:13px; line-height:1.35; }
    #leme-painel-ept .leme-meta b { color:var(--leme-text); }
    #leme-painel-ept .leme-badges, #leme-painel-ept .leme-tags, #leme-painel-ept .leme-actions { display:flex; flex-wrap:wrap; gap:8px; }
    #leme-painel-ept .leme-badges { margin:12px 0; }
    #leme-painel-ept .leme-tags { margin-top:10px; }
    #leme-painel-ept .leme-actions { margin-top:15px; }
    #leme-painel-ept .leme-badge, #leme-painel-ept .leme-tag { display:inline-flex; align-items:center; border-radius:999px; padding:6px 10px; font-size:11px; font-weight:850; background:var(--leme-primary-soft); color:var(--leme-primary); border:1px solid #d9ddff; }
    #leme-painel-ept .leme-badge { max-width:100%; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
    #leme-painel-ept .leme-badge.green { background:var(--leme-green-soft); color:#166534; border-color:#bbf7d0; }
    #leme-painel-ept .leme-badge.red { background:var(--leme-red-soft); color:var(--leme-red); border-color:#fecdd3; }
    #leme-painel-ept .leme-tag { background:#f8fafc; border-color:#e2e8f0; color:#475569; font-weight:750; }
    #leme-painel-ept .leme-summary { color:#475569; font-size:13px; line-height:1.6; margin:13px 0; display:-webkit-box; -webkit-line-clamp:4; -webkit-box-orient:vertical; overflow:hidden; }
    #leme-painel-ept .leme-card-footer { display:flex; align-items:center; justify-content:space-between; gap:12px; flex-wrap:wrap; margin-top: auto; padding-top: 13px; border-top:1px solid var(--leme-border); }
    #leme-painel-ept .leme-link { display:inline-flex; align-items:center; justify-content:center; gap:6px; border-radius:999px; padding:9px 12px; text-decoration:none !important; font-size:12px; font-weight:900; border:1px solid transparent; min-height:40px; }
    #leme-painel-ept .leme-link.pdf { color:#fff !important; background:var(--leme-red); }
    #leme-painel-ept .leme-link.produto { color:#fff !important; background:var(--leme-green); }
    #leme-painel-ept .leme-empty { grid-column:1/-1; background:#fff; border:1px dashed var(--leme-border-strong); border-radius:22px; padding:32px; text-align:center; color:var(--leme-muted); font-weight:800; }

    #leme-painel-ept table.leme-table { width:100%; min-width:1150px; border-collapse:collapse; font-size:13px; }
    #leme-painel-ept .leme-table caption { text-align:left; padding: 14px; font-weight:900; color:var(--leme-primary); }
    #leme-painel-ept .leme-table th { background:var(--leme-primary-2); color:#fff; text-align:left; padding:14px; position:sticky; top:0; z-index:2; white-space:nowrap; }
    #leme-painel-ept .leme-table td { padding:14px; border-bottom:1px solid var(--leme-border); vertical-align:top; }
    #leme-painel-ept .leme-table-title { color:var(--leme-primary); font-weight:900; line-height:1.35; }
    #leme-painel-ept .leme-pagination { display:flex; justify-content:center; align-items:center; gap:7px; flex-wrap:wrap; margin:20px 0 4px; }
    #leme-painel-ept .leme-pagination-info { width:100%; text-align:center; color:var(--leme-muted); font-size:12px; font-weight:750; margin-top:4px; }
    #leme-painel-ept mark.leme-highlight { background:#fef08a; color:inherit; border-radius:5px; padding:0 2px; box-decoration-break:clone; -webkit-box-decoration-break:clone; }

    #leme-painel-ept .leme-modal-backdrop { position:fixed; inset:0; background:rgba(15,23,42,.55); z-index:9999; display:none; align-items:center; justify-content:center; padding:18px; }
    #leme-painel-ept .leme-modal-backdrop.is-open { display:flex; }
    #leme-painel-ept .leme-modal { width:min(920px, 100%); max-height:min(86vh, 900px); overflow:auto; background:#fff; border-radius:26px; border:1px solid var(--leme-border); box-shadow:0 25px 80px rgba(15,23,42,.28); padding:22px; }
    #leme-painel-ept .leme-modal-head { display:flex; justify-content:space-between; align-items:flex-start; gap:14px; border-bottom:1px solid var(--leme-border); padding-bottom:14px; margin-bottom:14px; }
    #leme-painel-ept .leme-modal-title { margin:0; color:var(--leme-primary); font-size:20px; line-height:1.25; font-weight:900; }
    #leme-painel-ept .leme-detail-grid { display:grid; grid-template-columns: repeat(2, minmax(0,1fr)); gap:10px 16px; }
    #leme-painel-ept .leme-detail-item { padding:12px; border:1px solid #eef2ff; border-radius:16px; background:#fbfdff; color:#374151; font-size:13px; line-height:1.55; }
    #leme-painel-ept .leme-detail-item.is-wide { grid-column:1/-1; }
    #leme-painel-ept .leme-detail-item b { color:#111827; }

    @media (max-width: 1180px) {
      #leme-painel-ept .leme-filters { grid-template-columns: repeat(3, minmax(0, 1fr)); }
      #leme-painel-ept .leme-insights { grid-template-columns: repeat(2, minmax(0, 1fr)); }
    }
    @media (max-width: 1024px) {
      #leme-painel-ept .leme-stats, #leme-painel-ept .leme-insights { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      #leme-painel-ept .leme-filters { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      #leme-painel-ept .leme-results { grid-template-columns: 1fr; }
      #leme-painel-ept .leme-charts-grid { grid-template-columns: 1fr; }
      #leme-painel-ept .leme-chart-card.is-wide { grid-column:auto; }
    }
    @media (max-width: 720px) {
      #leme-painel-ept { padding:16px; border-radius:20px; }
      #leme-painel-ept .leme-header { grid-template-columns:1fr; }
      #leme-painel-ept .leme-update { white-space:normal; }
      #leme-painel-ept .leme-stats, #leme-painel-ept .leme-insights { grid-template-columns:1fr; }
      #leme-painel-ept .leme-search-row { grid-template-columns:1fr; }
      #leme-painel-ept .leme-filters { grid-template-columns:1fr; }
      #leme-painel-ept .leme-view { display:none; }
      #leme-painel-ept .leme-actions-toolbar { width:100%; }
      #leme-painel-ept .leme-export-btn { flex:1; }
      #leme-painel-ept .leme-results { display:grid !important; grid-template-columns:1fr; overflow:visible !important; border:0 !important; background:transparent !important; grid-auto-rows:auto; }
      #leme-painel-ept .leme-card { min-height: 0; }
      #leme-painel-ept .leme-bar-row { grid-template-columns:1fr; gap:5px; }
      #leme-painel-ept .leme-bar-value { text-align:left; }
      #leme-painel-ept .leme-year-chart { overflow-x:auto; justify-content:flex-start; }
      #leme-painel-ept .leme-year-col { min-width:58px; }
      #leme-painel-ept .leme-detail-grid { grid-template-columns:1fr; }
      #leme-painel-ept .leme-nav { position:static; }
      #leme-painel-ept .leme-nav button, #leme-painel-ept .leme-bottom-nav button, #leme-painel-ept .leme-head-btn { flex: 1 1 150px; }
      #leme-painel-ept .leme-panel { padding: 15px; border-radius: 20px; }
      #leme-painel-ept .leme-panel-head-actions { width: 100%; justify-content: flex-start; }
      #leme-painel-ept .leme-toolbar { align-items: stretch; }
      #leme-painel-ept .leme-count { width: 100%; }
    }
    @media (max-width: 520px) {
      #leme-painel-ept { padding: 12px; }
      #leme-painel-ept .leme-actions-toolbar, #leme-painel-ept .leme-card-footer, #leme-painel-ept .leme-actions { width: 100%; }
      #leme-painel-ept .leme-export-btn, #leme-painel-ept .leme-link, #leme-painel-ept .leme-card-detail-btn { flex: 1 1 auto; }
      #leme-painel-ept .leme-modal { padding: 16px; border-radius: 20px; }
      #leme-painel-ept .leme-modal-head { flex-direction: column; }
      #leme-painel-ept .leme-modal-close { width: 100%; }
    }
    @media (prefers-reduced-motion: reduce) {
      #leme-painel-ept, #leme-painel-ept * { scroll-behavior:auto !important; transition:none !important; animation:none !important; }
    }
  </style>

  <a href="#leme-controles" class="leme-skip">Pular para pesquisa e filtros</a>

  <div class="leme-header">
    <div>
      <div class="leme-eyebrow" aria-hidden="true">📊 Painel bibliométrico · EPT</div>
      <h2>Repositório Digital da Produção ProfEPT</h2>
      <p class="leme-subtitle">
        Consulte trabalhos por tema, autor, orientador, instituição, base epistemológica, nível de aplicação, tipo de produto, linha de pesquisa e macroprojeto.
      </p>
    </div>
    <div class="leme-update">
      Atualizado em <strong>__ATUALIZADO_EM__</strong><br>
      Período analisado: <strong>__PERIODO__</strong>
    </div>
  </div>

  <div class="leme-nav" aria-label="Navegação rápida do painel">
    <button type="button" data-scroll-to="leme-graficos">Ver gráficos</button>
    <button type="button" class="is-primary" data-scroll-to="leme-controles">Ir para filtros</button>
  </div>

  <div class="leme-stats" aria-label="Indicadores gerais do painel">
    <div class="leme-stat"><strong>__TOTAL__</strong><span>trabalhos analisados</span></div>
    <div class="leme-stat"><strong>__TOTAL_PRODUTOS__</strong><span>produtos com link</span></div>
    <div class="leme-stat"><strong>__TOTAL_ORIENTADORES__</strong><span>orientadores</span></div>
    <div class="leme-stat"><strong>__TOTAL_BASES__</strong><span>bases identificadas</span></div>
  </div>

  <section id="leme-graficos" class="leme-panel" aria-label="Visualizações bibliométricas">
    <div class="leme-panel-head">
      <div>
        <h3>Gráficos e leitura rápida</h3>
        <p>Os gráficos mudam junto com a busca e os filtros. Quando há poucos dados, o painel troca gráficos poluídos por resumos mais claros.</p>
      </div>
      <div class="leme-panel-head-actions">
        <div id="leme-graficos-nota" class="leme-note">Todos os trabalhos</div>
        <button type="button" class="leme-head-btn" data-scroll-to="leme-controles">Ir para filtros</button>
      </div>
    </div>

    <div id="leme-insights" class="leme-insights" aria-label="Resumo dinâmico dos dados filtrados"></div>

    <div class="leme-charts-grid">
      <div class="leme-chart-card is-wide">
        <h4 class="leme-chart-title">Trabalhos por ano</h4>
        <p class="leme-chart-subtitle">Evolução temporal do conjunto filtrado.</p>
        <div id="leme-chart-ano" aria-live="polite"></div>
      </div>
      <div class="leme-chart-card">
        <h4 class="leme-chart-title">Orientadores</h4>
        <p class="leme-chart-subtitle">Ranking dos orientadores mais recorrentes.</p>
        <div id="leme-chart-orientador" aria-live="polite"></div>
      </div>
      <div class="leme-chart-card">
        <h4 class="leme-chart-title">Base epistemológica</h4>
        <p class="leme-chart-subtitle">Categorias principais, com agrupamento automático quando necessário.</p>
        <div id="leme-chart-base" aria-live="polite"></div>
      </div>
      <div class="leme-chart-card">
        <h4 class="leme-chart-title">Nível de aplicação</h4>
        <p class="leme-chart-subtitle">Onde os produtos e pesquisas se concentram.</p>
        <div id="leme-chart-nivel" aria-live="polite"></div>
      </div>
      <div class="leme-chart-card">
        <h4 class="leme-chart-title">Tipo de produto</h4>
        <p class="leme-chart-subtitle">Formatos de produtos educacionais mais encontrados.</p>
        <div id="leme-chart-tipo" aria-live="polite"></div>
      </div>
      <div class="leme-chart-card">
        <h4 class="leme-chart-title">Linha de pesquisa</h4>
        <p class="leme-chart-subtitle">Distribuição pelas linhas acadêmicas.</p>
        <div id="leme-chart-linha" aria-live="polite"></div>
      </div>
      <div class="leme-chart-card">
        <h4 class="leme-chart-title">Macroprojeto</h4>
        <p class="leme-chart-subtitle">Distribuição pelos macroprojetos.</p>
        <div id="leme-chart-macro" aria-live="polite"></div>
      </div>
    </div>

  </section>

  <section id="leme-controles" class="leme-panel" aria-label="Pesquisa e filtros">
    <div class="leme-panel-head">
      <div>
        <h3>Pesquisa e filtros</h3>
        <p>Os filtros são dependentes: depois de escolher uma opção, os demais mostram apenas valores que ainda existem dentro do resultado possível.</p>
      </div>
      <div class="leme-note">Filtros inteligentes</div>
    </div>

    <div class="leme-search-row">
      <input id="leme-busca" class="leme-search" type="search"
        placeholder="Busque por tema, título, autor, orientador, resumo, palavra-chave, campus..."
        aria-label="Busca no painel EPT">
      <button id="leme-limpar" class="leme-clear" type="button">Limpar filtros</button>
    </div>

    <div class="leme-filters">
      <div class="leme-field"><label for="leme-filtro-ano">Ano</label><select id="leme-filtro-ano" data-campo="ano" data-default="Todos"></select></div>
      <div class="leme-field"><label for="leme-filtro-nivel">Nível</label><select id="leme-filtro-nivel" data-campo="nivel" data-default="Todos"></select></div>
      <div class="leme-field"><label for="leme-filtro-linha">Linha de pesquisa</label><select id="leme-filtro-linha" data-campo="linhaPesquisa" data-default="Todas"></select></div>
      <div class="leme-field"><label for="leme-filtro-macro">Macroprojeto</label><select id="leme-filtro-macro" data-campo="macroprojeto" data-default="Todos"></select></div>
      <div class="leme-field"><label for="leme-filtro-area">Área temática</label><select id="leme-filtro-area" data-campo="area" data-default="Todas"></select></div>
      <div class="leme-field"><label for="leme-filtro-tipo">Tipo de produto</label><select id="leme-filtro-tipo" data-campo="tipoProduto" data-default="Todos"></select></div>
      <div class="leme-field"><label for="leme-filtro-orientador">Orientador</label><select id="leme-filtro-orientador" data-campo="orientador" data-default="Todos"></select></div>
      <div class="leme-field"><label for="leme-filtro-base">Base epistemológica</label><select id="leme-filtro-base" data-campo="base" data-default="Todas"></select></div>
    </div>

    <div id="leme-filtros-ativos" class="leme-active-filters" aria-live="polite"></div>

  </section>

  <div id="leme-resultados-area" class="leme-toolbar">
    <div id="leme-contador" class="leme-count" aria-live="polite">Carregando resultados...</div>
    <div class="leme-actions-toolbar">
      <button id="leme-exportar-csv" class="leme-export-btn" type="button">⬇️ Exportar filtrados</button>
      <button id="leme-copiar-resultados" class="leme-export-btn" type="button">📋 Copiar resultados</button>
      <div class="leme-view" aria-label="Alternar visualização">
        <button id="leme-view-cards" type="button" class="is-active" aria-pressed="true">Cards</button>
        <button id="leme-view-table" type="button" aria-pressed="false">Tabela</button>
      </div>
    </div>
  </div>

  <div id="leme-resultados" class="leme-results"></div>
  <nav id="leme-paginacao" class="leme-pagination" aria-label="Paginação de resultados"></nav>

  <div class="leme-bottom-nav" aria-label="Navegação final do painel">
    <button type="button" data-scroll-to="leme-controles">Voltar aos filtros</button>
    <button type="button" data-scroll-to="leme-graficos">Ver gráficos</button>
  </div>

  <div id="leme-modal-backdrop" class="leme-modal-backdrop" aria-hidden="true">
    <div class="leme-modal" role="dialog" aria-modal="true" aria-labelledby="leme-modal-title">
      <div class="leme-modal-head">
        <h3 id="leme-modal-title" class="leme-modal-title">Detalhes do trabalho</h3>
        <button type="button" id="leme-modal-close" class="leme-modal-close">Fechar</button>
      </div>
      <div id="leme-modal-content"></div>
    </div>
  </div>

  <script type="application/json" id="leme-dados-json">__DADOS_JSON__</script>

  <script>
    (function() {
      const raiz = document.getElementById("leme-painel-ept");
      if (!raiz) return;

      const dados = JSON.parse(raiz.querySelector("#leme-dados-json").textContent || "[]");
      const busca = raiz.querySelector("#leme-busca");
      const limpar = raiz.querySelector("#leme-limpar");
      const resultados = raiz.querySelector("#leme-resultados");
      const paginacao = raiz.querySelector("#leme-paginacao");
      const contador = raiz.querySelector("#leme-contador");
      const filtros = Array.from(raiz.querySelectorAll("select[data-campo]"));
      const filtrosAtivos = raiz.querySelector("#leme-filtros-ativos");
      const graficosNota = raiz.querySelector("#leme-graficos-nota");
      const insights = raiz.querySelector("#leme-insights");
      const exportarCsv = raiz.querySelector("#leme-exportar-csv");
      const copiarResultados = raiz.querySelector("#leme-copiar-resultados");
      const btnCards = raiz.querySelector("#leme-view-cards");
      const btnTable = raiz.querySelector("#leme-view-table");
      const modalBackdrop = raiz.querySelector("#leme-modal-backdrop");
      const modalContent = raiz.querySelector("#leme-modal-content");
      const modalTitle = raiz.querySelector("#leme-modal-title");
      const modalClose = raiz.querySelector("#leme-modal-close");

      let resultadosAtuais = [];
      let modo = "cards";
      let paginaAtual = 1;
      let debounceTimer = null;
      const itensPorPagina = 12;
      const camposBusca = ["ano", "autor", "orientador", "orientadorChave", "titulo", "produto", "resumo", "nivel", "base", "area", "finalidade", "publicoAlvo", "tipoProduto", "instituicaoCampus", "linhaPesquisa", "macroprojeto", "codigoMacroprojeto", "confiancaLinhaMacro", "termosLinhaMacro", "justificativaLinhaMacro", "revisaoManual"];

      function normalizarBusca(valor) {
        return String(valor || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/\s+/g, " ").trim();
      }
      function termosBusca() {
        const termo = normalizarBusca(busca.value || "");
        return termo ? termo.split(" ").filter(Boolean) : [];
      }
      function textoBusca(item) {
        const partes = camposBusca.map(campo => item[campo] || "");
        partes.push((item.palavrasChave || []).join(" "));
        return normalizarBusca(partes.join(" "));
      }
      function escapeHtml(valor) {
        return String(valor || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
      }
      function limparRotulo(valor) {
        const texto = String(valor || "").replace(/\s+/g, " ").trim();
        return texto && texto !== "---" ? texto : "Não identificado";
      }
      function labelCurto(valor, limite) {
        const texto = limparRotulo(valor);
        return texto.length > limite ? texto.slice(0, limite - 1).trim() + "…" : texto;
      }
      function highlightText(valor) {
        const textoOriginal = String(valor || "");
        const termos = termosBusca();
        if (!termos.length || !textoOriginal) return escapeHtml(textoOriginal);
        const textoNormalizado = normalizarBusca(textoOriginal);
        const ranges = [];
        termos.forEach(termo => {
          let inicio = 0;
          while (inicio < textoNormalizado.length) {
            const idx = textoNormalizado.indexOf(termo, inicio);
            if (idx === -1) break;
            ranges.push([idx, idx + termo.length]);
            inicio = idx + termo.length;
          }
        });
        if (!ranges.length) return escapeHtml(textoOriginal);
        ranges.sort((a, b) => a[0] - b[0] || b[1] - a[1]);
        const mesclados = [];
        ranges.forEach(range => {
          const ultimo = mesclados[mesclados.length - 1];
          if (!ultimo || range[0] > ultimo[1]) mesclados.push(range);
          else ultimo[1] = Math.max(ultimo[1], range[1]);
        });
        let saida = "";
        let cursor = 0;
        mesclados.forEach(([ini, fim]) => {
          saida += escapeHtml(textoOriginal.slice(cursor, ini));
          saida += '<mark class="leme-highlight">' + escapeHtml(textoOriginal.slice(ini, fim)) + '</mark>';
          cursor = fim;
        });
        saida += escapeHtml(textoOriginal.slice(cursor));
        return saida;
      }

      function filtrarDados(ignorarCampo) {
        const termos = termosBusca();
        return dados.filter(item => {
          if (termos.length && !termos.every(t => textoBusca(item).includes(t))) return false;
          for (const filtro of filtros) {
            const campo = filtro.dataset.campo;
            if (ignorarCampo && campo === ignorarCampo) continue;
            if (filtro.value && item[campo] !== filtro.value) return false;
          }
          return true;
        });
      }

      function contarOpcoes(lista, campo) {
        const mapa = new Map();
        lista.forEach(item => {
          const valor = item[campo];
          if (!valor || valor === "---") return;
          mapa.set(valor, (mapa.get(valor) || 0) + 1);
        });
        return Array.from(mapa.entries()).sort((a, b) => String(a[0]).localeCompare(String(b[0]), "pt-BR"));
      }

      function popularFiltrosDinamicos() {
        filtros.forEach(select => {
          const campo = select.dataset.campo;
          const atual = select.value;
          const base = filtrarDados(campo);
          const opcoes = contarOpcoes(base, campo);
          const padrao = select.dataset.default || "Todos";
          let html = '<option value="">' + escapeHtml(padrao) + '</option>';
          let existeAtual = !atual;
          let qtdAtual = 0;
          opcoes.forEach(([valor, qtd]) => {
            if (valor === atual) { existeAtual = true; qtdAtual = qtd; }
            html += '<option value="' + escapeHtml(valor) + '">' + escapeHtml(valor) + '</option>';
          });
          if (atual && !existeAtual) {
            html += '<option value="' + escapeHtml(atual) + '">' + escapeHtml(atual) + '</option>';
          }
          select.innerHTML = html;
          select.value = atual;
          select.disabled = opcoes.length === 0 && !atual;
          const field = select.closest(".leme-field");
          if (field) {
            let countEl = field.querySelector(".leme-field-count");
            if (!countEl) {
              countEl = document.createElement("div");
              countEl.className = "leme-field-count";
              field.appendChild(countEl);
            }
            countEl.classList.toggle("is-selected", Boolean(atual));
            if (atual) {
              countEl.textContent = (qtdAtual || 0) + " resultado(s) nessa opção";
            } else {
              countEl.textContent = opcoes.length + " opção(ões) disponível(is)";
            }
          }
        });
      }

      function renderFiltrosAtivos() {
        const chips = [];
        if (busca.value.trim()) chips.push('<span class="leme-chip">Busca: ' + escapeHtml(busca.value.trim()) + '<button type="button" data-clear-search="1" aria-label="Remover busca">×</button></span>');
        filtros.forEach(filtro => {
          if (filtro.value) {
            const label = raiz.querySelector('label[for="' + filtro.id + '"]');
            const qtd = filtrarDados().length;
            chips.push('<span class="leme-chip">' + escapeHtml(label ? label.textContent : filtro.dataset.campo) + ': ' + escapeHtml(labelCurto(filtro.value, 46)) + ' · ' + qtd + ' resultado(s)<button type="button" data-clear-filter="' + escapeHtml(filtro.dataset.campo) + '" aria-label="Remover filtro">×</button></span>');
          }
        });
        filtrosAtivos.innerHTML = chips.join("");
      }

      function contarPorCampo(lista, campo) {
        const mapa = new Map();
        lista.forEach(item => {
          const valor = limparRotulo(item[campo]);
          mapa.set(valor, (mapa.get(valor) || 0) + 1);
        });
        return Array.from(mapa.entries()).map(([label, valor]) => ({ label, valor })).sort((a, b) => b.valor - a.valor || String(a.label).localeCompare(String(b.label), "pt-BR"));
      }
      function contarOrientadores(lista, limite) {
        const mapa = new Map();
        lista.forEach(item => {
          const chave = item.orientadorChave || normalizarBusca(item.orientador);
          const nome = limparRotulo(item.orientador);
          if (!chave || nome === "Não identificado") return;
          if (!mapa.has(chave)) mapa.set(chave, { label: nome, valor: 0 });
          mapa.get(chave).valor += 1;
        });
        return Array.from(mapa.values()).sort((a,b) => b.valor - a.valor || String(a.label).localeCompare(String(b.label), "pt-BR")).slice(0, limite || 10);
      }
      function agruparTop(lista, limite) {
        if (lista.length <= limite) return lista;
        const top = lista.slice(0, limite);
        const outras = lista.slice(limite).reduce((soma, item) => soma + item.valor, 0);
        if (outras) top.push({ label: "Outras", valor: outras });
        return top;
      }
      function renderInsightCard(titulo, valor) {
        return '<div class="leme-insight-card"><strong>' + escapeHtml(valor || "---") + '</strong><span>' + escapeHtml(titulo) + '</span></div>';
      }
      function primeiroOuTraco(lista) { return lista && lista[0] ? lista[0].label : "---"; }
      function renderInsights(lista) {
        if (!lista.length) {
          insights.innerHTML = '<div class="leme-insight-card"><strong>0</strong><span>Nenhum resultado neste recorte. Limpe algum filtro ou use uma busca mais ampla.</span></div>';
          return;
        }
        insights.innerHTML = [
          renderInsightCard("resultados no conjunto atual", String(lista.length)),
          renderInsightCard("tipo de produto mais comum", primeiroOuTraco(contarPorCampo(lista, "tipoProduto"))),
          renderInsightCard("orientador em destaque", primeiroOuTraco(contarOrientadores(lista, 1))),
          renderInsightCard("macroprojeto mais recorrente", primeiroOuTraco(contarPorCampo(lista, "macroprojeto")))
        ].join("");
      }
      function renderHorizontalChart(el, dadosGrafico, limite) {
        const lista = agruparTop(dadosGrafico, limite || 8).filter(d => d.valor > 0);
        if (!lista.length) { el.innerHTML = '<div class="leme-empty-chart">Limpe algum filtro ou faça uma busca mais ampla para voltar a comparar categorias.</div>'; return; }
        if (lista.length === 1) {
          el.innerHTML = '<div class="leme-single-insight"><strong>' + escapeHtml(lista[0].label) + '</strong>Todos os ' + lista[0].valor + ' resultado(s) deste recorte estão nesta categoria.</div>';
          return;
        }
        const max = Math.max(...lista.map(d => d.valor), 1);
        el.innerHTML = '<div class="leme-bars">' + lista.map(d => {
          const pct = Math.max(4, Math.round((d.valor / max) * 100));
          return '<div class="leme-bar-row"><div class="leme-bar-label" title="' + escapeHtml(d.label) + '">' + escapeHtml(labelCurto(d.label, 56)) + '</div><div class="leme-bar-track"><div class="leme-bar-fill" style="width:' + pct + '%"></div></div><div class="leme-bar-value">' + d.valor + '</div></div>';
        }).join("") + '</div>';
      }
      function renderYearChart(el, lista) {
        const dadosAno = contarPorCampo(lista, "ano").sort((a,b) => String(a.label).localeCompare(String(b.label), "pt-BR"));
        if (!dadosAno.length) { el.innerHTML = '<div class="leme-empty-chart">Limpe algum filtro ou faça uma busca mais ampla para voltar a comparar os anos.</div>'; return; }
        if (dadosAno.length === 1) { el.innerHTML = '<div class="leme-single-insight"><strong>' + escapeHtml(dadosAno[0].label) + '</strong>Todos os ' + dadosAno[0].valor + ' resultado(s) filtrados estão neste ano.</div>'; return; }
        const max = Math.max(...dadosAno.map(d => d.valor), 1);
        const altura = Math.min(260, Math.max(150, 88 + dadosAno.length * 10));
        el.innerHTML = '<div class="leme-year-chart" style="--chart-height:' + altura + 'px">' + dadosAno.map(d => {
          const h = Math.max(8, Math.round((d.valor / max) * (altura - 62)));
          return '<div class="leme-year-col"><div class="leme-year-value">' + d.valor + '</div><div class="leme-year-bar" style="height:' + h + 'px"></div><div class="leme-year-label">' + escapeHtml(d.label) + '</div></div>';
        }).join("") + '</div>';
      }
      function atualizarGraficos(lista) {
        renderInsights(lista);
        renderYearChart(raiz.querySelector("#leme-chart-ano"), lista);
        renderHorizontalChart(raiz.querySelector("#leme-chart-orientador"), contarOrientadores(lista, 10), 10);
        renderHorizontalChart(raiz.querySelector("#leme-chart-base"), contarPorCampo(lista, "base"), 8);
        renderHorizontalChart(raiz.querySelector("#leme-chart-nivel"), contarPorCampo(lista, "nivel"), 8);
        renderHorizontalChart(raiz.querySelector("#leme-chart-tipo"), contarPorCampo(lista, "tipoProduto"), 8);
        renderHorizontalChart(raiz.querySelector("#leme-chart-linha"), contarPorCampo(lista, "linhaPesquisa"), 6);
        renderHorizontalChart(raiz.querySelector("#leme-chart-macro"), contarPorCampo(lista, "macroprojeto"), 6);
        graficosNota.textContent = lista.length === dados.length ? "Todos os trabalhos" : lista.length + " de " + dados.length + " trabalhos filtrados";
      }

      function renderBadges(item) {
        const badges = [];
        if (item.ano && item.ano !== "---") badges.push('<span class="leme-badge">' + escapeHtml(item.ano) + '</span>');
        if (item.nivel && item.nivel !== "---") badges.push('<span class="leme-badge green">' + escapeHtml(item.nivel) + '</span>');
        if (item.linhaPesquisa && item.linhaPesquisa !== "---") badges.push('<span class="leme-badge">' + escapeHtml(labelCurto(item.linhaPesquisa, 42)) + '</span>');
        if (item.tipoProduto && item.tipoProduto !== "---") badges.push('<span class="leme-badge red">' + escapeHtml(item.tipoProduto) + '</span>');
        return badges.join("");
      }
      function renderTags(tags) {
        if (!tags || !tags.length) return "";
        return '<div class="leme-tags">' + tags.slice(0, 12).map(t => '<span class="leme-tag">' + highlightText(t) + '</span>').join("") + '</div>';
      }
      function renderActions(item) {
        const links = [];
        if (item.linkPdf) links.push('<a class="leme-link pdf" href="' + escapeHtml(item.linkPdf) + '" target="_blank" rel="noopener" aria-label="Abrir PDF em nova aba">📄 PDF</a>');
        if (item.linkProduto) links.push('<a class="leme-link produto" href="' + escapeHtml(item.linkProduto) + '" target="_blank" rel="noopener" aria-label="Abrir produto educacional em nova aba">📦 Produto</a>');
        return links.length ? '<div class="leme-actions">' + links.join("") + '</div>' : "";
      }
      function renderCards(lista) {
        if (!lista.length) { resultados.innerHTML = '<div class="leme-empty">Nenhum trabalho encontrado com os filtros selecionados.</div>'; return; }
        resultados.innerHTML = lista.map(item => {
          const idx = dados.indexOf(item);
          return '<article class="leme-card">' +
            '<h3 class="leme-card-title">' + highlightText(item.titulo) + '</h3>' +
            '<div class="leme-meta"><div><b>Autor:</b> ' + highlightText(item.autor) + '</div><div><b>Orientador:</b> ' + highlightText(item.orientador) + '</div><div><b>Instituição/Campus:</b> ' + highlightText(item.instituicaoCampus) + '</div></div>' +
            '<div class="leme-badges">' + renderBadges(item) + '</div>' +
            '<p class="leme-summary">' + highlightText(item.resumoCurto || "Resumo não disponível.") + '</p>' +
            '<div class="leme-card-footer"><button type="button" class="leme-card-detail-btn" data-index="' + idx + '">Ver detalhes</button>' + renderActions(item) + '</div>' +
          '</article>';
        }).join("");
      }
      function renderTable(lista) {
        if (!lista.length) { resultados.innerHTML = '<div class="leme-empty">Nenhum trabalho encontrado com os filtros selecionados.</div>'; return; }
        resultados.innerHTML = '<table class="leme-table"><caption>Resultados filtrados do painel</caption><thead><tr><th scope="col">Título</th><th scope="col">Autor</th><th scope="col">Orientador</th><th scope="col">Ano</th><th scope="col">Nível</th><th scope="col">Base</th><th scope="col">Linha</th><th scope="col">Macroprojeto</th><th scope="col">Área</th><th scope="col">Links</th></tr></thead><tbody>' +
          lista.map(item => '<tr><td><div class="leme-table-title">' + highlightText(item.titulo) + '</div></td><td>' + highlightText(item.autor) + '</td><td>' + highlightText(item.orientador) + '</td><td>' + escapeHtml(item.ano) + '</td><td>' + escapeHtml(item.nivel) + '</td><td>' + escapeHtml(item.base) + '</td><td>' + escapeHtml(item.linhaPesquisa || "Não classificado") + '</td><td>' + escapeHtml(item.macroprojeto || "Não classificado") + '</td><td>' + escapeHtml(item.area) + '</td><td>' + renderActions(item) + '</td></tr>').join("") +
          '</tbody></table>';
      }
      function paginaLista(lista) {
        const totalPaginas = Math.max(1, Math.ceil(lista.length / itensPorPagina));
        if (paginaAtual > totalPaginas) paginaAtual = totalPaginas;
        const inicio = (paginaAtual - 1) * itensPorPagina;
        return lista.slice(inicio, inicio + itensPorPagina);
      }
      function paginasVisiveis(totalPaginas) {
        const paginas = new Set([1, totalPaginas, paginaAtual - 1, paginaAtual, paginaAtual + 1]);
        if (paginaAtual <= 3) [1,2,3,4].forEach(p => paginas.add(p));
        if (paginaAtual >= totalPaginas - 2) [totalPaginas - 3, totalPaginas - 2, totalPaginas - 1, totalPaginas].forEach(p => paginas.add(p));
        return [...paginas].filter(p => p >= 1 && p <= totalPaginas).sort((a, b) => a - b);
      }
      function botaoPagina(rotulo, pagina, classe, disabled) {
        const atual = classe && classe.includes("is-active");
        return '<button type="button" class="leme-page-btn ' + (classe || "") + '" data-pagina="' + pagina + '" ' + (disabled ? "disabled" : "") + (atual ? ' aria-current="page"' : '') + ' aria-label="Página ' + escapeHtml(rotulo) + '">' + escapeHtml(rotulo) + '</button>';
      }
      function renderPaginacao(totalItens) {
        const totalPaginas = Math.max(1, Math.ceil(totalItens / itensPorPagina));
        if (!totalItens || totalPaginas <= 1) { paginacao.innerHTML = ""; return; }
        let html = botaoPagina("Anterior", Math.max(1, paginaAtual - 1), "", paginaAtual === 1);
        const visiveis = paginasVisiveis(totalPaginas);
        let anterior = 0;
        visiveis.forEach(p => {
          if (anterior && p - anterior > 1) html += '<span class="leme-pagination-info" style="width:auto;margin:0;">...</span>';
          html += botaoPagina(String(p), p, p === paginaAtual ? "is-active" : "", false);
          anterior = p;
        });
        html += botaoPagina("Próximo", Math.min(totalPaginas, paginaAtual + 1), "", paginaAtual === totalPaginas);
        const inicio = (paginaAtual - 1) * itensPorPagina + 1;
        const fim = Math.min(paginaAtual * itensPorPagina, totalItens);
        html += '<div class="leme-pagination-info">Exibindo ' + inicio + '–' + fim + ' de ' + totalItens + ' resultados</div>';
        paginacao.innerHTML = html;
      }
      function atualizarContador(total) {
        contador.innerHTML = '<strong>' + total + '</strong> resultado(s) encontrado(s)';
      }
      function atualizarVisualizacao() {
        resultados.classList.toggle("is-table", modo === "table");
        btnCards.classList.toggle("is-active", modo === "cards");
        btnTable.classList.toggle("is-active", modo === "table");
        btnCards.setAttribute("aria-pressed", modo === "cards" ? "true" : "false");
        btnTable.setAttribute("aria-pressed", modo === "table" ? "true" : "false");
        const pagina = paginaLista(resultadosAtuais);
        if (modo === "cards") renderCards(pagina); else renderTable(pagina);
        renderPaginacao(resultadosAtuais.length);
      }
      function aplicarFiltros(resetPagina) {
        if (resetPagina) paginaAtual = 1;
        popularFiltrosDinamicos();
        resultadosAtuais = filtrarDados();
        atualizarContador(resultadosAtuais.length);
        renderFiltrosAtivos();
        atualizarGraficos(resultadosAtuais);
        atualizarVisualizacao();
      }
      function abrirModal(item) {
        if (!item) return;
        modalTitle.innerHTML = highlightText(item.titulo || "Detalhes do trabalho");
        modalContent.innerHTML = '<div class="leme-detail-grid">' +
          '<div class="leme-detail-item"><b>Autor:</b><br>' + highlightText(item.autor) + '</div>' +
          '<div class="leme-detail-item"><b>Orientador:</b><br>' + highlightText(item.orientador) + '</div>' +
          '<div class="leme-detail-item"><b>Ano:</b><br>' + escapeHtml(item.ano) + '</div>' +
          '<div class="leme-detail-item"><b>Instituição/Campus:</b><br>' + highlightText(item.instituicaoCampus) + '</div>' +
          '<div class="leme-detail-item is-wide"><b>Produto educacional:</b><br>' + highlightText(item.produto) + '</div>' +
          '<div class="leme-detail-item"><b>Área temática:</b><br>' + highlightText(item.area) + '</div>' +
          '<div class="leme-detail-item"><b>Tipo de produto:</b><br>' + highlightText(item.tipoProduto) + '</div>' +
          '<div class="leme-detail-item"><b>Nível de aplicação:</b><br>' + highlightText(item.nivel) + '</div>' +
          '<div class="leme-detail-item"><b>Base epistemológica:</b><br>' + highlightText(item.base) + '</div>' +
          '<div class="leme-detail-item is-wide"><b>Linha de pesquisa:</b><br>' + highlightText(item.linhaPesquisa || "Não classificado") + '</div>' +
          '<div class="leme-detail-item is-wide"><b>Macroprojeto:</b><br>' + highlightText(item.macroprojeto || "Não classificado") + '</div>' +
          '<div class="leme-detail-item"><b>Código do macroprojeto:</b><br>' + highlightText(item.codigoMacroprojeto || "Não classificado") + '</div>' +
          '<div class="leme-detail-item"><b>Confiança Linha/Macro:</b><br>' + highlightText(item.confiancaLinhaMacro || "Baixa") + ' · <b>Revisão manual:</b> ' + highlightText(item.revisaoManual || "Sim") + '</div>' +
          '<div class="leme-detail-item is-wide"><b>Finalidade:</b><br>' + highlightText(item.finalidade) + '</div>' +
          '<div class="leme-detail-item is-wide"><b>Público-alvo:</b><br>' + highlightText(item.publicoAlvo) + '</div>' +
          '<div class="leme-detail-item is-wide"><b>Justificativa Linha/Macro:</b><br>' + highlightText(item.justificativaLinhaMacro || "") + '</div>' +
          '<div class="leme-detail-item is-wide"><b>Resumo:</b><br>' + highlightText(item.resumo || "Resumo não disponível.") + renderTags(item.palavrasChave) + renderActions(item) + '</div>' +
        '</div>';
        modalBackdrop.classList.add("is-open");
        modalBackdrop.setAttribute("aria-hidden", "false");
        modalClose.focus();
      }
      function fecharModal() {
        modalBackdrop.classList.remove("is-open");
        modalBackdrop.setAttribute("aria-hidden", "true");
      }
      function csvEscape(valor) { return '"' + String(valor || "").replace(/"/g, '""') + '"'; }
      function baixarCsv() {
        const cols = [["ano","Ano"],["titulo","Título"],["autor","Autor"],["orientador","Orientador"],["nivel","Nível"],["base","Base"],["linhaPesquisa","Linha de Pesquisa"],["macroprojeto","Macroprojeto"],["area","Área"],["tipoProduto","Tipo de Produto"],["linkPdf","PDF"],["linkProduto","Produto"]];
        const linhas = [cols.map(c => csvEscape(c[1])).join(";")].concat(resultadosAtuais.map(item => cols.map(c => csvEscape(item[c[0]])).join(";")));
        const blob = new Blob(["\ufeff" + linhas.join("\n")], { type: "text/csv;charset=utf-8;" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "painel-leme-ept-filtrado.csv";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }
      function copiarTexto() {
        const texto = resultadosAtuais.map((item, i) => (i + 1) + ". " + item.titulo + " — " + item.autor + " (" + item.ano + ")").join("\n");
        navigator.clipboard.writeText(texto).then(() => { copiarResultados.textContent = "Copiado!"; setTimeout(() => copiarResultados.textContent = "📋 Copiar resultados", 1200); });
      }

      raiz.addEventListener("click", function(e) {
        const scrollBtn = e.target.closest("[data-scroll-to]");
        if (scrollBtn) {
          const alvo = raiz.querySelector("#" + scrollBtn.dataset.scrollTo);
          if (alvo) alvo.scrollIntoView({ behavior: "smooth", block: "start" });
        }
        const pageBtn = e.target.closest(".leme-page-btn[data-pagina]");
        if (pageBtn && !pageBtn.disabled) { paginaAtual = Number(pageBtn.dataset.pagina || 1); atualizarVisualizacao(); raiz.querySelector("#leme-resultados-area").scrollIntoView({ behavior: "smooth", block: "start" }); }
        const detailBtn = e.target.closest(".leme-card-detail-btn[data-index]");
        if (detailBtn) abrirModal(dados[Number(detailBtn.dataset.index)]);
        const clearFilter = e.target.closest("[data-clear-filter]");
        if (clearFilter) { const filtro = filtros.find(f => f.dataset.campo === clearFilter.dataset.clearFilter); if (filtro) filtro.value = ""; aplicarFiltros(true); }
        const clearSearch = e.target.closest("[data-clear-search]");
        if (clearSearch) { busca.value = ""; aplicarFiltros(true); busca.focus(); }
      });
      filtros.forEach(filtro => filtro.addEventListener("change", () => aplicarFiltros(true)));
      busca.addEventListener("input", () => { clearTimeout(debounceTimer); debounceTimer = setTimeout(() => aplicarFiltros(true), 160); });
      busca.addEventListener("keydown", e => { if (e.key === "Escape") { busca.value = ""; aplicarFiltros(true); } });
      limpar.addEventListener("click", () => { busca.value = ""; filtros.forEach(f => f.value = ""); aplicarFiltros(true); });
      btnCards.addEventListener("click", () => { modo = "cards"; atualizarVisualizacao(); });
      btnTable.addEventListener("click", () => { modo = "table"; atualizarVisualizacao(); });
      exportarCsv.addEventListener("click", baixarCsv);
      copiarResultados.addEventListener("click", copiarTexto);
      modalClose.addEventListener("click", fecharModal);
      modalBackdrop.addEventListener("click", e => { if (e.target === modalBackdrop) fecharModal(); });
      document.addEventListener("keydown", e => { if (e.key === "Escape" && modalBackdrop.classList.contains("is-open")) fecharModal(); });

      aplicarFiltros(true);
    })();
  </script>
</section>'''

    html = html.replace("__DADOS_JSON__", dados_json)
    html = html.replace("__TOTAL__", str(total))
    html = html.replace("__TOTAL_PRODUTOS__", str(total_produtos))
    html = html.replace("__TOTAL_ORIENTADORES__", str(total_orientadores))
    html = html.replace("__TOTAL_BASES__", str(total_bases))
    html = html.replace("__PERIODO__", periodo)
    html = html.replace("__ATUALIZADO_EM__", atualizado_em)
    return html

def salvar_arquivo_html(html, arquivo_saida=ARQUIVO_HTML_SAIDA):
    """
    Salva somente o HTML final do painel.

    Este script não cria nem altera planilhas.
    Ele apenas gera um arquivo .html ou .txt com o bloco completo
    para colar no Elementor.
    """
    caminho = Path(arquivo_saida)

    if caminho.suffix.lower() not in {".html", ".txt"}:
        print("Aviso: o arquivo de saída não termina com .html nem .txt.")
        print("O conteúdo será salvo mesmo assim, mas recomendo usar .html ou .txt.")

    caminho.write_text(html, encoding="utf-8")


def main():
    """
    Executa a geração do painel para Elementor.

    Entrada:
    - dados_finais_painel_ept_linhas_macroprojetos.xlsx

    Saída:
    - painel_leme_ept.html ou painel_leme_ept.txt

    Importante:
    Este script NÃO salva planilha.
    Este script NÃO recalcula classificações.
    Este script apenas gera o HTML final do painel.
    """
    try:
        print(f"Lendo planilha revisada: {ARQUIVO_ENTRADA}")
        df = carregar_planilha(ARQUIVO_ENTRADA)

        print("Preparando dados para o HTML sem recalcular classificações...")
        registros = preparar_dados_para_html(df)

        print("Gerando painel HTML...")
        html = gerar_html(registros)

        salvar_arquivo_html(html, ARQUIVO_HTML_SAIDA)

        print("\nPainel gerado com sucesso!")
        print(f"Arquivo gerado: {ARQUIVO_HTML_SAIDA}")
        print(f"Total de trabalhos no painel: {len(registros)}")
        print("\nAgora abra o arquivo gerado, copie todo o conteúdo e cole no Elementor.")

    except FileNotFoundError as erro:
        print("\nErro: planilha de entrada não encontrada.")
        print(erro)

    except PermissionError:
        print("\nErro: não foi possível salvar o arquivo HTML/TXT.")
        print("Feche o arquivo se ele estiver aberto e rode o script novamente.")

    except Exception as erro:
        print("\nOcorreu um erro ao gerar o painel.")
        print(f"Detalhes do erro: {erro}")


if __name__ == "__main__":
    main()
