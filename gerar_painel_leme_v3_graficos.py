import pandas as pd
import json
import re
import unicodedata
from datetime import datetime

# ============================================================
# PROJETO LEME - GERADOR DE PAINEL EPT V2
# WordPress + Elementor + HTML/CSS/JS puro
# ============================================================

ARQUIVO_ENTRADA = "dados_finais_painel_ept.xlsx"
ARQUIVO_HTML_SAIDA = "painel_leme_ept_v2.txt"
ARQUIVO_EXCEL_SAIDA = "dados_finais_painel_ept_classificados.xlsx"


def limpar_valor(valor, padrao="---"):
    if pd.isna(valor):
        return padrao
    valor = str(valor).strip()
    if not valor or valor.lower() in ["nan", "none", "null"]:
        return padrao
    return valor


def normalizar_texto(texto):
    texto = limpar_valor(texto, "")
    texto = texto.lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def ano_limpo(valor):
    texto = limpar_valor(valor, "")
    m = re.search(r"(20\d{2})", texto)
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
    if len(texto) <= limite:
        return texto
    return texto[:limite].rsplit(" ", 1)[0] + "..."


def link_valido(link):
    link = limpar_valor(link, "")
    return link.startswith("http://") or link.startswith("https://")


BASES_EPISTEMOLOGICAS = {
    "Formação Humana Integral / Omnilateralidade": {
        "fortes": ["formacao humana integral", "formacao integral", "omnilateral", "omnilateralidade", "politecnia", "educacao politecnica", "desenvolvimento integral", "formacao completa"],
        "medios": ["integral", "emancipatoria", "sujeito historico", "trabalho ciencia cultura tecnologia", "formacao dos sujeitos"],
    },
    "Trabalho como Princípio Educativo": {
        "fortes": ["trabalho como principio educativo", "trabalho como principio", "centralidade do trabalho", "mundo do trabalho", "ontologia do trabalho", "trabalho educativo"],
        "medios": ["juventude e trabalho", "relacao trabalho educacao", "formacao para o trabalho", "trabalho e educacao"],
    },
    "Currículo Integrado": {
        "fortes": ["curriculo integrado", "integracao curricular", "ensino medio integrado", "emi", "pratica pedagogica integradora", "praticas pedagogicas integradoras"],
        "medios": ["interdisciplinaridade", "transversal", "projeto integrador", "integracao", "curricularizacao", "componentes curriculares"],
    },
    "Materialismo Histórico-Dialético": {
        "fortes": ["materialismo historico dialetico", "historico dialetico", "materialismo", "dialetico", "dialetica", "marx", "marxista"],
        "medios": ["totalidade", "contradicao", "praxis", "modo de producao", "classes sociais", "consciencia de classe"],
    },
    "Pedagogia Histórico-Crítica": {
        "fortes": ["pedagogia historico critica", "historico critica", "saviani"],
        "medios": ["pratica social", "catarse", "instrumentalizacao", "conteudos classicos", "mediacao pedagogica"],
    },
    "Educação Crítica / Emancipatória": {
        "fortes": ["educacao critica", "emancipacao", "emancipatoria", "paulo freire", "freire", "conscientizacao", "dialogicidade"],
        "medios": ["autonomia", "critica", "acao dialogica", "educacao libertadora", "participacao", "protagonismo"],
    },
    "Educação Inclusiva / Direitos Humanos": {
        "fortes": ["inclusao", "educacao inclusiva", "educacao especial", "acessibilidade", "direitos humanos", "pessoa com deficiencia", "atendimento educacional especializado", "aee", "libras", "surdos", "estudantes trans", "lei de cotas"],
        "medios": ["diversidade", "permanencia", "acolhimento", "equidade", "desenho universal", "cotas", "ppi"],
    },
    "Tecnologias Educacionais na EPT": {
        "fortes": ["tecnologia educacional", "tecnologias digitais", "tdic", "moodle", "chatgpt", "inteligencia artificial", "telegram", "gamificacao", "jogo digital", "ambiente virtual"],
        "medios": ["audiovisual", "podcast", "video", "aplicativo", "portal", "site", "ferramentas digitais", "meios digitais"],
    },
    "História, Memória e Cultura Escolar": {
        "fortes": ["historia das instituicoes escolares", "memoria", "cultura escolar", "culturas escolares", "historia da educacao", "implantacao", "leme", "instituicao escolar"],
        "medios": ["arquivo", "fotografia", "jornais", "acervo", "historia", "identidade institucional"],
    },
    "Gestão, Permanência e Políticas Institucionais": {
        "fortes": ["permanencia e exito", "evasao", "assistencia estudantil", "programa de auxilios estudantis", "gestao", "politicas publicas", "conselho de classe", "processo seletivo", "ingresso"],
        "medios": ["regulamento", "comissoes", "gestao de conflitos", "identidade organizacional", "egressos", "acompanhamento"],
    },
}


AREAS_TEMATICAS = {
    "Currículo e Práticas Pedagógicas": ["curriculo", "ensino medio integrado", "pratica pedagogica", "sequencia didatica", "interdisciplinaridade", "componentes curriculares", "feira do conhecimento", "conselho de classe"],
    "Formação Docente": ["formacao docente", "formacao de professores", "professores", "docentes", "licenciatura", "pratica docente"],
    "Tecnologias e Recursos Digitais": ["tecnologia", "moodle", "chatgpt", "inteligencia artificial", "telegram", "gamificacao", "audiovisual", "podcast", "video", "aplicativo", "portal", "site"],
    "Inclusão, Diversidade e Acessibilidade": ["inclusao", "acessibilidade", "deficiencia", "aee", "libras", "surdos", "cotas", "ppi", "trans", "diversidade", "lei 10.639"],
    "História, Memória e Instituições": ["historia", "memoria", "cultura escolar", "instituicoes escolares", "implantacao", "fotografia", "cedup", "ifc", "leme"],
    "Permanência, Êxito e Gestão Educacional": ["permanencia", "exito", "evasao", "assistencia estudantil", "gestao", "conflitos", "ingresso", "processo seletivo", "egressos"],
    "Mundo do Trabalho e Formação Profissional": ["mundo do trabalho", "trabalho", "estagio", "jovem aprendiz", "empreendedorismo", "seguranca do trabalho", "curso tecnico"],
    "Educação Ambiental, Saúde e Sociedade": ["educacao ambiental", "saude", "alimentar", "nutricional", "violencias", "agroecologia", "agricultura familiar"],
}


FINALIDADES_PRODUTO = {
    "Apoio ao ensino": ["sequencia didatica", "material didatico", "cartilha", "e-book", "livro digital", "guia", "manual", "aula", "ensino", "aprendizagem"],
    "Formação de professores": ["formacao de professores", "formacao docente", "docentes", "professores", "pratica docente", "curso"],
    "Orientação institucional": ["guia", "manual", "orientativo", "regulamento", "protocolo", "gestao", "comissoes", "institucional"],
    "Divulgação científica e memória": ["historia", "memoria", "tour", "fotografia", "acervo", "cultura escolar", "leme", "identidade"],
    "Intervenção pedagógica": ["intervencao", "acao educativa", "oficina", "pratica pedagogica", "atividade", "proposta de ensino"],
    "Recurso digital/interativo": ["portal", "site", "jogo", "aplicativo", "podcast", "video", "moodle", "telegram", "audiovisual", "infografico"],
    "Diagnóstico e análise": ["analise", "percepcao", "diagnostico", "estudo de caso", "levantamento", "mapeamento", "estado do conhecimento"],
}


PUBLICOS_ALVO = {
    "Estudantes do Ensino Médio Integrado": ["ensino medio integrado", "emi", "estudantes do ensino medio", "alunos do ensino medio", "curso tecnico integrado"],
    "Estudantes da EJA/EPT": ["eja", "proeja", "jovens e adultos", "eja-ept"],
    "Professores/Educadores": ["professores", "docentes", "educadores", "formacao docente", "formacao de professores"],
    "Gestores e equipes pedagógicas": ["gestores", "equipe pedagogica", "conselho de classe", "comissoes", "gestao", "tae", "tecnico-administrativos"],
    "Comunidade externa": ["comunidade", "egressos", "familia", "9 ano", "ensino fundamental", "profissionais", "sociedade civil"],
    "Estudantes com deficiência ou público da inclusão": ["deficiencia", "aee", "libras", "surdos", "inclusao", "acessibilidade", "educacao especial"],
}


TIPOS_PRODUTO = {
    "Livro digital / E-book": ["livro digital", "e-book", "ebook"],
    "Guia / Manual / Cartilha": ["guia", "manual", "cartilha"],
    "Sequência didática": ["sequencia didatica"],
    "Curso / Oficina": ["curso", "oficina"],
    "Vídeo / Audiovisual": ["video", "audiovisual", "filme", "animacao"],
    "Podcast / Áudio": ["podcast", "audio"],
    "Portal / Site": ["portal", "site"],
    "Jogo / Gamificação": ["jogo", "gamificacao", "game"],
    "Infográfico / Apresentação": ["infografico", "apresentacao"],
    "Ferramenta / Aplicativo": ["aplicativo", "ferramenta", "telegram", "moodle"],
}


def montar_textos(row):
    titulo = normalizar_texto(row.get("Título da Dissertação", ""))
    produto = normalizar_texto(row.get("Produto Educacional", ""))
    palavras = normalizar_texto(row.get("Palavras-chave", ""))
    resumo = normalizar_texto(row.get("Resumo", ""))
    tipo = normalizar_texto(row.get("Tipo de Produto", ""))
    texto_total = " ".join([titulo, produto, palavras, resumo, tipo]).strip()
    return titulo, produto, palavras, resumo, tipo, texto_total


def pontuar_categoria(palavras, titulo, produto, resumo, config):
    pontos = 0
    termos = []

    for termo in config.get("fortes", []):
        termo_n = normalizar_texto(termo)
        if termo_n in palavras:
            pontos += 7
            termos.append(termo)
        if termo_n in titulo or termo_n in produto:
            pontos += 5
            termos.append(termo)
        if termo_n in resumo:
            pontos += 3
            termos.append(termo)

    for termo in config.get("medios", []):
        termo_n = normalizar_texto(termo)
        if termo_n in palavras:
            pontos += 4
            termos.append(termo)
        if termo_n in titulo or termo_n in produto:
            pontos += 3
            termos.append(termo)
        if termo_n in resumo:
            pontos += 1
            termos.append(termo)

    termos_unicos = []
    vistos = set()
    for t in termos:
        chave = normalizar_texto(t)
        if chave not in vistos:
            vistos.add(chave)
            termos_unicos.append(t)

    return pontos, termos_unicos


def classificar_base(row):
    titulo, produto, palavras, resumo, tipo, texto_total = montar_textos(row)

    placar = {}
    termos_por_base = {}

    for base, config in BASES_EPISTEMOLOGICAS.items():
        pontos, termos = pontuar_categoria(palavras, titulo, produto, resumo, config)
        placar[base] = pontos
        termos_por_base[base] = termos

    base_vencedora, pontos_vencedor = max(placar.items(), key=lambda x: x[1])

    if pontos_vencedor < 4:
        return "A definir (Análise Manual)", "Baixa", ""

    if pontos_vencedor >= 12:
        confianca = "Alta"
    elif pontos_vencedor >= 7:
        confianca = "Média"
    else:
        confianca = "Baixa"

    termos = "; ".join(termos_por_base.get(base_vencedora, [])[:8])
    return base_vencedora, confianca, termos


def classificar_por_mapa(row, mapa, padrao):
    titulo, produto, palavras, resumo, tipo, texto_total = montar_textos(row)

    melhor_categoria = padrao
    melhor_pontuacao = 0

    for categoria, termos in mapa.items():
        pontos = 0
        for termo in termos:
            termo_n = normalizar_texto(termo)
            if termo_n in palavras:
                pontos += 5
            if termo_n in titulo or termo_n in produto:
                pontos += 4
            if termo_n in tipo:
                pontos += 4
            if termo_n in resumo:
                pontos += 2

        if pontos > melhor_pontuacao:
            melhor_pontuacao = pontos
            melhor_categoria = categoria

    return melhor_categoria


def classificar_nivel(row):
    titulo, produto, palavras, resumo, tipo, texto_total = montar_textos(row)

    if any(t in texto_total for t in ["eja", "proeja", "jovens e adultos", "eja-ept"]):
        return "EJA/EPT"

    if any(t in texto_total for t in ["ensino medio integrado", "medio integrado", "emi", "curso tecnico integrado", "tecnico integrado"]):
        return "Ensino Médio Integrado"

    if any(t in texto_total for t in ["curso tecnico subsequente", "tecnico subsequente", "subsequente"]):
        return "Curso Técnico Subsequente"

    if any(t in texto_total for t in ["graduacao", "licenciatura", "curso superior", "superior"]):
        return "Graduação"

    if any(t in texto_total for t in ["pos-graduacao", "mestrado", "profept"]):
        return "Pós-graduação"

    if any(t in texto_total for t in ["professores", "docentes", "formacao docente"]):
        return "Formação continuada"

    return "Outros"


def extrair_instituicao_campus(row):
    texto_original = " ".join([
        limpar_valor(row.get("Título da Dissertação", ""), ""),
        limpar_valor(row.get("Resumo", ""), "")
    ])
    texto_n = normalizar_texto(texto_original)

    padroes = [
        ("IFC - Campus ", r"ifc\s*(?:-|–)?\s*campus\s+([a-zçãõáéíóúâêô\s]+)"),
        ("IFSC - Campus ", r"ifsc\s*(?:-|–)?\s*campus\s+([a-zçãõáéíóúâêô\s]+)"),
        ("CEDUP ", r"cedup\s*(?:-|–)?\s*([a-zçãõáéíóúâêô\s]+)"),
        ("SENAI ", r"senai\s+([a-zçãõáéíóúâêô\s]+)"),
    ]

    for prefixo, padrao in padroes:
        m = re.search(padrao, texto_n)
        if m:
            nome = m.group(1).strip()
            nome = re.split(r"\s+(sobre|no|na|em|do|da)\s+", nome)[0]
            nome = nome[:45].strip(" -–.,")
            if nome:
                return prefixo + nome.title()

    if "instituto federal catarinense" in texto_n or " ifc " in f" {texto_n} ":
        return "IFC"
    if "instituto federal de santa catarina" in texto_n or " ifsc " in f" {texto_n} ":
        return "IFSC"
    if "cedup" in texto_n:
        return "CEDUP"
    if "senai" in texto_n:
        return "SENAI"

    return "---"


def preparar_dados(df):
    registros = []

    for _, row in df.iterrows():
        row = row.to_dict()

        base, confianca_base, termos_base = classificar_base(row)
        nivel = classificar_nivel(row)

        area = classificar_por_mapa(row, AREAS_TEMATICAS, "Outras temáticas")
        finalidade = classificar_por_mapa(row, FINALIDADES_PRODUTO, "Consulta, estudo ou diagnóstico")
        publico = classificar_por_mapa(row, PUBLICOS_ALVO, "Público amplo da EPT")
        tipo_produto_detectado = classificar_por_mapa(row, TIPOS_PRODUTO, limpar_valor(row.get("Tipo de Produto", "Outro")))

        tags = dividir_tags(row.get("Palavras-chave", ""))

        registro = {
            "ano": ano_limpo(row.get("Ano", "")),
            "autor": limpar_valor(row.get("Autor", "")),
            "orientador": limpar_valor(row.get("Orientador", "")),
            "titulo": limpar_valor(row.get("Título da Dissertação", "")),
            "produto": limpar_valor(row.get("Produto Educacional", "")),
            "linkPdf": limpar_valor(row.get("Link do PDF", "")) if link_valido(row.get("Link do PDF", "")) else "",
            "linkProduto": limpar_valor(row.get("Link do Produto", "")) if link_valido(row.get("Link do Produto", "")) else "",
            "palavrasChave": tags,
            "tipoProdutoOriginal": limpar_valor(row.get("Tipo de Produto", "")),
            "tipoProduto": tipo_produto_detectado,
            "resumo": limpar_valor(row.get("Resumo", "")),
            "resumoCurto": limitar_texto(row.get("Resumo", ""), 280),
            "nivel": nivel,
            "base": base,
            "confiancaBase": confianca_base,
            "termosBase": termos_base,
            "area": area,
            "finalidade": finalidade,
            "publicoAlvo": publico,
            "instituicaoCampus": extrair_instituicao_campus(row),
        }

        registros.append(registro)

    return registros


def atualizar_excel_classificado(df, registros):
    df_saida = df.copy()

    df_saida["Ano Limpo"] = [r["ano"] for r in registros]
    df_saida["Base Epistemológica V2"] = [r["base"] for r in registros]
    df_saida["Confiança da Base"] = [r["confiancaBase"] for r in registros]
    df_saida["Termos Detectados da Base"] = [r["termosBase"] for r in registros]
    df_saida["Nível de Aplicação V2"] = [r["nivel"] for r in registros]
    df_saida["Área Temática"] = [r["area"] for r in registros]
    df_saida["Finalidade do Produto"] = [r["finalidade"] for r in registros]
    df_saida["Público-alvo"] = [r["publicoAlvo"] for r in registros]
    df_saida["Tipo de Produto V2"] = [r["tipoProduto"] for r in registros]
    df_saida["Instituição/Campus"] = [r["instituicaoCampus"] for r in registros]

    df_saida.to_excel(ARQUIVO_EXCEL_SAIDA, index=False)


def gerar_html(registros):
    dados_json = json.dumps(registros, ensure_ascii=False)

    total = len(registros)
    total_orientadores = len({r["orientador"] for r in registros if r["orientador"] != "---"})
    total_bases = len({r["base"] for r in registros if r["base"] != "---"})
    total_produtos = sum(1 for r in registros if r["linkProduto"])

    anos_validos = sorted([r["ano"] for r in registros if r["ano"] != "---"])
    periodo = f"{anos_validos[0]}–{anos_validos[-1]}" if anos_validos else "---"
    atualizado_em = datetime.now().strftime("%d/%m/%Y")

    html = r'''<!-- =========================================================
PAINEL LEME EPT V2
Gerado automaticamente por Python
Cole TODO este bloco no Elementor > Editor de Texto ou HTML
========================================================= -->

<section id="leme-painel-ept" class="leme-painel-ept">
  <style>
    #leme-painel-ept {
      --leme-primary: #4852f4;
      --leme-primary-dark: #3038c8;
      --leme-primary-soft: #eef0ff;
      --leme-green: #2d5a27;
      --leme-green-soft: #e8f5e9;
      --leme-red: #e63946;
      --leme-bg: #f7f8ff;
      --leme-card: #ffffff;
      --leme-border: #dfe3f5;
      --leme-text: #1f2937;
      --leme-muted: #6b7280;
      --leme-shadow: 0 16px 40px rgba(31, 41, 55, 0.10);
      font-family: "Segoe UI", Arial, sans-serif;
      color: var(--leme-text);
      background: linear-gradient(180deg, var(--leme-bg), #ffffff);
      border: 1px solid var(--leme-border);
      border-radius: 24px;
      padding: 24px;
      box-sizing: border-box;
      overflow: hidden;
    }

    #leme-painel-ept * { box-sizing: border-box; }

    #leme-painel-ept .leme-header {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 20px;
      align-items: start;
      margin-bottom: 22px;
    }

    #leme-painel-ept .leme-eyebrow {
      display: inline-flex;
      padding: 7px 11px;
      border-radius: 999px;
      background: var(--leme-primary-soft);
      color: var(--leme-primary-dark);
      font-size: 12px;
      font-weight: 700;
      letter-spacing: .02em;
      margin-bottom: 10px;
    }

    #leme-painel-ept h2 {
      margin: 0;
      color: var(--leme-primary-dark);
      font-size: clamp(24px, 3vw, 36px);
      line-height: 1.1;
      font-weight: 800;
    }

    #leme-painel-ept .leme-subtitle {
      margin: 10px 0 0;
      color: var(--leme-muted);
      font-size: 15px;
      line-height: 1.55;
      max-width: 920px;
    }

    #leme-painel-ept .leme-update {
      background: #fff;
      border: 1px solid var(--leme-border);
      border-radius: 16px;
      padding: 12px 14px;
      font-size: 12px;
      color: var(--leme-muted);
      white-space: nowrap;
      box-shadow: 0 8px 24px rgba(72, 82, 244, .08);
    }

    #leme-painel-ept .leme-stats {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
      margin-bottom: 18px;
    }

    #leme-painel-ept .leme-stat {
      background: var(--leme-card);
      border: 1px solid var(--leme-border);
      border-radius: 18px;
      padding: 16px;
      box-shadow: 0 10px 28px rgba(31, 41, 55, 0.06);
    }

    #leme-painel-ept .leme-stat strong {
      display: block;
      color: var(--leme-primary-dark);
      font-size: 26px;
      line-height: 1;
      margin-bottom: 6px;
    }

    #leme-painel-ept .leme-stat span {
      color: var(--leme-muted);
      font-size: 12px;
      font-weight: 600;
    }

    #leme-painel-ept .leme-controls {
      background: var(--leme-card);
      border: 1px solid var(--leme-border);
      border-radius: 22px;
      padding: 16px;
      margin-bottom: 18px;
      box-shadow: var(--leme-shadow);
    }

    #leme-painel-ept .leme-search-row {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 12px;
      margin-bottom: 14px;
    }

    #leme-painel-ept .leme-search {
      width: 100%;
      border: 2px solid var(--leme-border);
      border-radius: 16px;
      padding: 15px 16px;
      font-size: 15px;
      color: var(--leme-text);
      outline: none;
      background: #fff;
      transition: border-color .2s ease, box-shadow .2s ease;
    }

    #leme-painel-ept .leme-search:focus {
      border-color: var(--leme-primary);
      box-shadow: 0 0 0 4px rgba(72, 82, 244, .12);
    }

    #leme-painel-ept .leme-clear {
      border: none;
      border-radius: 16px;
      padding: 0 16px;
      background: var(--leme-primary);
      color: #fff;
      font-weight: 700;
      cursor: pointer;
    }

    #leme-painel-ept .leme-clear:hover { background: var(--leme-primary-dark); }

    #leme-painel-ept .leme-filters {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 10px;
    }

    #leme-painel-ept .leme-field label {
      display: block;
      font-size: 11px;
      font-weight: 800;
      color: var(--leme-muted);
      margin: 0 0 5px;
      text-transform: uppercase;
      letter-spacing: .04em;
    }

    #leme-painel-ept .leme-field select {
      width: 100%;
      border: 1px solid var(--leme-border);
      border-radius: 14px;
      padding: 11px 12px;
      background: #fff;
      color: var(--leme-text);
      font-size: 13px;
      outline: none;
    }

    #leme-painel-ept .leme-toolbar {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      margin: 16px 0;
      flex-wrap: wrap;
    }

    #leme-painel-ept .leme-count {
      font-size: 14px;
      color: var(--leme-muted);
      font-weight: 700;
    }

    #leme-painel-ept .leme-count strong { color: var(--leme-primary-dark); }

    #leme-painel-ept .leme-view {
      display: inline-flex;
      gap: 8px;
      align-items: center;
    }

    #leme-painel-ept .leme-view button {
      border: 1px solid var(--leme-border);
      border-radius: 999px;
      background: #fff;
      color: var(--leme-muted);
      padding: 8px 12px;
      cursor: pointer;
      font-size: 12px;
      font-weight: 800;
    }

    #leme-painel-ept .leme-view button.is-active {
      background: var(--leme-primary);
      border-color: var(--leme-primary);
      color: #fff;
    }

    #leme-painel-ept .leme-results {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 14px;
    }

    #leme-painel-ept .leme-results.is-table {
      display: block;
      overflow-x: auto;
      background: #fff;
      border: 1px solid var(--leme-border);
      border-radius: 18px;
    }

    #leme-painel-ept .leme-card {
      background: var(--leme-card);
      border: 1px solid var(--leme-border);
      border-radius: 20px;
      padding: 18px;
      box-shadow: 0 10px 26px rgba(31, 41, 55, .07);
      transition: transform .16s ease, box-shadow .16s ease, border-color .16s ease;
    }

    #leme-painel-ept .leme-card:hover {
      transform: translateY(-2px);
      box-shadow: 0 18px 42px rgba(31, 41, 55, .11);
      border-color: rgba(72, 82, 244, .35);
    }

    #leme-painel-ept .leme-card-title {
      margin: 0 0 10px;
      color: var(--leme-primary-dark);
      font-size: 16px;
      line-height: 1.32;
      font-weight: 850;
    }

    #leme-painel-ept .leme-meta {
      display: grid;
      gap: 6px;
      margin: 0 0 12px;
      color: var(--leme-muted);
      font-size: 13px;
      line-height: 1.35;
    }

    #leme-painel-ept .leme-meta b { color: var(--leme-text); }

    #leme-painel-ept .leme-badges {
      display: flex;
      flex-wrap: wrap;
      gap: 7px;
      margin: 12px 0;
    }

    #leme-painel-ept .leme-badge {
      display: inline-flex;
      align-items: center;
      border-radius: 999px;
      padding: 6px 9px;
      font-size: 11px;
      font-weight: 800;
      background: var(--leme-primary-soft);
      color: var(--leme-primary-dark);
      border: 1px solid #d9ddff;
    }

    #leme-painel-ept .leme-badge.green {
      background: var(--leme-green-soft);
      color: var(--leme-green);
      border-color: #c8e6c9;
    }

    #leme-painel-ept .leme-badge.red {
      background: #fff0f1;
      color: var(--leme-red);
      border-color: #ffd0d5;
    }

    #leme-painel-ept .leme-summary {
      color: #4b5563;
      font-size: 13px;
      line-height: 1.55;
      margin: 12px 0;
    }

    #leme-painel-ept details {
      margin-top: 10px;
      border-top: 1px solid var(--leme-border);
      padding-top: 10px;
    }

    #leme-painel-ept summary {
      cursor: pointer;
      color: var(--leme-primary-dark);
      font-weight: 800;
      font-size: 13px;
      list-style: none;
    }

    #leme-painel-ept summary::-webkit-details-marker { display: none; }

    #leme-painel-ept .leme-detail {
      margin-top: 12px;
      color: #374151;
      font-size: 13px;
      line-height: 1.55;
    }

    #leme-painel-ept .leme-tags {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      margin-top: 10px;
    }

    #leme-painel-ept .leme-tag {
      background: #f3f4f6;
      border: 1px solid #e5e7eb;
      color: #374151;
      border-radius: 999px;
      padding: 5px 8px;
      font-size: 11px;
      font-weight: 650;
    }

    #leme-painel-ept .leme-actions {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 14px;
    }

    #leme-painel-ept .leme-link {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 6px;
      border-radius: 999px;
      padding: 9px 12px;
      text-decoration: none !important;
      font-size: 12px;
      font-weight: 850;
      border: 1px solid transparent;
    }

    #leme-painel-ept .leme-link.pdf {
      color: #fff !important;
      background: var(--leme-red);
    }

    #leme-painel-ept .leme-link.produto {
      color: #fff !important;
      background: var(--leme-green);
    }

    #leme-painel-ept .leme-empty {
      grid-column: 1 / -1;
      background: #fff;
      border: 1px dashed var(--leme-border);
      border-radius: 18px;
      padding: 30px;
      text-align: center;
      color: var(--leme-muted);
      font-weight: 700;
    }

    #leme-painel-ept table.leme-table {
      width: 100%;
      min-width: 1150px;
      border-collapse: collapse;
      font-size: 13px;
    }

    #leme-painel-ept .leme-table th {
      background: var(--leme-primary);
      color: #fff;
      text-align: left;
      padding: 13px;
      position: sticky;
      top: 0;
      z-index: 2;
      white-space: nowrap;
    }

    #leme-painel-ept .leme-table td {
      padding: 13px;
      border-bottom: 1px solid var(--leme-border);
      vertical-align: top;
    }

    #leme-painel-ept .leme-table-title {
      color: var(--leme-primary-dark);
      font-weight: 850;
      line-height: 1.3;
    }


    #leme-painel-ept .leme-analytics {
      margin-top: 22px;
      background: #fff;
      border: 1px solid var(--leme-border);
      border-radius: 24px;
      padding: 20px;
      box-shadow: var(--leme-shadow);
    }

    #leme-painel-ept .leme-analytics-head {
      display: flex;
      justify-content: space-between;
      gap: 14px;
      align-items: flex-start;
      margin-bottom: 16px;
      flex-wrap: wrap;
    }

    #leme-painel-ept .leme-analytics-head h3 {
      margin: 0;
      color: var(--leme-primary-dark);
      font-size: 22px;
      line-height: 1.2;
      font-weight: 850;
    }

    #leme-painel-ept .leme-analytics-head p {
      margin: 6px 0 0;
      color: var(--leme-muted);
      font-size: 13px;
      line-height: 1.5;
      max-width: 780px;
    }

    #leme-painel-ept .leme-analytics-note {
      background: var(--leme-primary-soft);
      color: var(--leme-primary-dark);
      border: 1px solid #d9ddff;
      border-radius: 999px;
      padding: 8px 12px;
      font-size: 12px;
      font-weight: 800;
      white-space: nowrap;
    }

    #leme-painel-ept .leme-charts-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 14px;
    }

    #leme-painel-ept .leme-chart-card {
      border: 1px solid var(--leme-border);
      border-radius: 20px;
      padding: 16px;
      background: linear-gradient(180deg, #ffffff, #fbfbff);
      min-height: 290px;
    }

    #leme-painel-ept .leme-chart-card.is-wide {
      grid-column: 1 / -1;
    }

    #leme-painel-ept .leme-chart-title {
      margin: 0 0 4px;
      color: var(--leme-text);
      font-size: 15px;
      font-weight: 850;
    }

    #leme-painel-ept .leme-chart-subtitle {
      margin: 0 0 14px;
      color: var(--leme-muted);
      font-size: 12px;
      line-height: 1.45;
    }

    #leme-painel-ept .leme-year-chart {
      display: flex;
      align-items: end;
      gap: 10px;
      min-height: 210px;
      padding: 16px 4px 0;
      border-top: 1px solid #eef0ff;
    }

    #leme-painel-ept .leme-year-item {
      flex: 1;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: end;
      gap: 7px;
      min-width: 42px;
    }

    #leme-painel-ept .leme-year-value {
      font-size: 12px;
      font-weight: 850;
      color: var(--leme-primary-dark);
    }

    #leme-painel-ept .leme-year-bar {
      width: 100%;
      max-width: 46px;
      min-height: 8px;
      border-radius: 12px 12px 4px 4px;
      background: linear-gradient(180deg, var(--leme-primary), var(--leme-primary-dark));
      box-shadow: 0 10px 20px rgba(72, 82, 244, .18);
      transition: height .25s ease;
    }

    #leme-painel-ept .leme-year-label {
      color: var(--leme-muted);
      font-size: 11px;
      font-weight: 800;
    }

    #leme-painel-ept .leme-bars {
      display: grid;
      gap: 10px;
    }

    #leme-painel-ept .leme-bar-row {
      display: grid;
      grid-template-columns: minmax(120px, 220px) 1fr auto;
      gap: 10px;
      align-items: center;
    }

    #leme-painel-ept .leme-bar-label {
      color: var(--leme-text);
      font-size: 12px;
      font-weight: 750;
      line-height: 1.25;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    #leme-painel-ept .leme-bar-track {
      height: 11px;
      border-radius: 999px;
      background: #eef0ff;
      overflow: hidden;
      position: relative;
    }

    #leme-painel-ept .leme-bar-fill {
      height: 100%;
      min-width: 4px;
      border-radius: 999px;
      background: linear-gradient(90deg, var(--leme-primary), var(--leme-primary-dark));
    }

    #leme-painel-ept .leme-bar-value {
      color: var(--leme-primary-dark);
      font-size: 12px;
      font-weight: 850;
      min-width: 26px;
      text-align: right;
    }

    @media (max-width: 1024px) {
      #leme-painel-ept .leme-stats { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      #leme-painel-ept .leme-filters { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      #leme-painel-ept .leme-results { grid-template-columns: 1fr; }
      #leme-painel-ept .leme-charts-grid { grid-template-columns: 1fr; }
      #leme-painel-ept .leme-chart-card.is-wide { grid-column: auto; }
    }

    @media (max-width: 720px) {
      #leme-painel-ept {
        padding: 16px;
        border-radius: 18px;
      }

      #leme-painel-ept .leme-header { grid-template-columns: 1fr; }
      #leme-painel-ept .leme-update { white-space: normal; }
      #leme-painel-ept .leme-stats { grid-template-columns: 1fr; }
      #leme-painel-ept .leme-search-row { grid-template-columns: 1fr; }
      #leme-painel-ept .leme-clear { padding: 13px 16px; }
      #leme-painel-ept .leme-filters { grid-template-columns: 1fr; }
      #leme-painel-ept .leme-view { display: none; }

      #leme-painel-ept .leme-bar-row {
        grid-template-columns: 1fr;
        gap: 5px;
      }

      #leme-painel-ept .leme-bar-value {
        text-align: left;
      }

      #leme-painel-ept .leme-year-chart {
        overflow-x: auto;
        align-items: end;
      }

      #leme-painel-ept .leme-year-item {
        min-width: 52px;
      }

      #leme-painel-ept .leme-results {
        display: grid !important;
        grid-template-columns: 1fr;
        overflow: visible !important;
        border: 0 !important;
        background: transparent !important;
      }
    }
  </style>

  <div class="leme-header">
    <div>
      <div class="leme-eyebrow">Painel bibliométrico · EPT</div>
      <h2>Repositório Digital da Produção ProfEPT</h2>
      <p class="leme-subtitle">
        Consulte trabalhos por tema, autor, orientador, instituição, base epistemológica,
        nível de aplicação, finalidade do produto educacional e palavras-chave.
      </p>
    </div>
    <div class="leme-update">
      Atualizado em <strong>__ATUALIZADO_EM__</strong><br>
      Período analisado: <strong>__PERIODO__</strong>
    </div>
  </div>

  <div class="leme-stats" aria-label="Indicadores gerais do painel">
    <div class="leme-stat"><strong>__TOTAL__</strong><span>trabalhos analisados</span></div>
    <div class="leme-stat"><strong>__TOTAL_PRODUTOS__</strong><span>produtos com link</span></div>
    <div class="leme-stat"><strong>__TOTAL_ORIENTADORES__</strong><span>orientadores</span></div>
    <div class="leme-stat"><strong>__TOTAL_BASES__</strong><span>bases identificadas</span></div>
  </div>

  <div class="leme-controls">
    <div class="leme-search-row">
      <input id="leme-busca" class="leme-search" type="search"
        placeholder="Busque por tema, título, autor, orientador, resumo, palavra-chave, campus..."
        aria-label="Busca no painel EPT">
      <button id="leme-limpar" class="leme-clear" type="button">Limpar filtros</button>
    </div>

    <div class="leme-filters">
      <div class="leme-field">
        <label for="leme-filtro-ano">Ano</label>
        <select id="leme-filtro-ano" data-campo="ano"><option value="">Todos</option></select>
      </div>

      <div class="leme-field">
        <label for="leme-filtro-nivel">Nível</label>
        <select id="leme-filtro-nivel" data-campo="nivel"><option value="">Todos</option></select>
      </div>

      <div class="leme-field">
        <label for="leme-filtro-base">Base epistemológica</label>
        <select id="leme-filtro-base" data-campo="base"><option value="">Todas</option></select>
      </div>

      <div class="leme-field">
        <label for="leme-filtro-area">Área temática</label>
        <select id="leme-filtro-area" data-campo="area"><option value="">Todas</option></select>
      </div>

      <div class="leme-field">
        <label for="leme-filtro-orientador">Orientador</label>
        <select id="leme-filtro-orientador" data-campo="orientador"><option value="">Todos</option></select>
      </div>

      <div class="leme-field">
        <label for="leme-filtro-tipo">Tipo de produto</label>
        <select id="leme-filtro-tipo" data-campo="tipoProduto"><option value="">Todos</option></select>
      </div>

      <div class="leme-field">
        <label for="leme-filtro-finalidade">Finalidade</label>
        <select id="leme-filtro-finalidade" data-campo="finalidade"><option value="">Todas</option></select>
      </div>

      <div class="leme-field">
        <label for="leme-filtro-publico">Público-alvo</label>
        <select id="leme-filtro-publico" data-campo="publicoAlvo"><option value="">Todos</option></select>
      </div>
    </div>
  </div>

  <div class="leme-toolbar">
    <div id="leme-contador" class="leme-count">Carregando resultados...</div>
    <div class="leme-view" aria-label="Alternar visualização">
      <button id="leme-view-cards" type="button" class="is-active">Cards</button>
      <button id="leme-view-table" type="button">Tabela</button>
    </div>
  </div>

  <div id="leme-resultados" class="leme-results"></div>

  <section class="leme-analytics" aria-label="Visualizações bibliométricas">
    <div class="leme-analytics-head">
      <div>
        <h3>Visão geral da produção</h3>
        <p>
          Os gráficos resumem os trabalhos exibidos no painel. Ao usar a busca ou os filtros,
          as visualizações são atualizadas automaticamente.
        </p>
      </div>
      <div id="leme-graficos-nota" class="leme-analytics-note">Todos os trabalhos</div>
    </div>

    <div class="leme-charts-grid">
      <div class="leme-chart-card is-wide">
        <h4 class="leme-chart-title">Trabalhos por ano</h4>
        <p class="leme-chart-subtitle">Evolução temporal da produção cadastrada no repositório.</p>
        <div id="leme-chart-ano" class="leme-year-chart"></div>
      </div>

      <div class="leme-chart-card">
        <h4 class="leme-chart-title">Base epistemológica</h4>
        <p class="leme-chart-subtitle">Distribuição dos trabalhos pela base identificada automaticamente.</p>
        <div id="leme-chart-base" class="leme-bars"></div>
      </div>

      <div class="leme-chart-card">
        <h4 class="leme-chart-title">Área temática</h4>
        <p class="leme-chart-subtitle">Principais temas mapeados a partir de título, resumo e palavras-chave.</p>
        <div id="leme-chart-area" class="leme-bars"></div>
      </div>

      <div class="leme-chart-card">
        <h4 class="leme-chart-title">Nível de aplicação</h4>
        <p class="leme-chart-subtitle">Onde os produtos e pesquisas mais se concentram.</p>
        <div id="leme-chart-nivel" class="leme-bars"></div>
      </div>

      <div class="leme-chart-card">
        <h4 class="leme-chart-title">Tipo de produto</h4>
        <p class="leme-chart-subtitle">Formatos de produtos educacionais mais recorrentes.</p>
        <div id="leme-chart-produto" class="leme-bars"></div>
      </div>
    </div>
  </section>


  <script type="application/json" id="leme-dados-json">__DADOS_JSON__</script>

  <script>
    (function() {
      const raiz = document.getElementById("leme-painel-ept");
      if (!raiz) return;

      const dadosEl = raiz.querySelector("#leme-dados-json");
      const dados = JSON.parse(dadosEl.textContent || "[]");

      const busca = raiz.querySelector("#leme-busca");
      const limpar = raiz.querySelector("#leme-limpar");
      const resultados = raiz.querySelector("#leme-resultados");
      const contador = raiz.querySelector("#leme-contador");
      const filtros = Array.from(raiz.querySelectorAll("select[data-campo]"));

      const chartAno = raiz.querySelector("#leme-chart-ano");
      const chartBase = raiz.querySelector("#leme-chart-base");
      const chartArea = raiz.querySelector("#leme-chart-area");
      const chartNivel = raiz.querySelector("#leme-chart-nivel");
      const chartProduto = raiz.querySelector("#leme-chart-produto");
      const graficosNota = raiz.querySelector("#leme-graficos-nota");

      const btnCards = raiz.querySelector("#leme-view-cards");
      const btnTable = raiz.querySelector("#leme-view-table");

      let modo = "cards";

      function textoBusca(item) {
        return [
          item.ano,
          item.autor,
          item.orientador,
          item.titulo,
          item.produto,
          item.resumo,
          item.nivel,
          item.base,
          item.area,
          item.finalidade,
          item.publicoAlvo,
          item.tipoProduto,
          item.instituicaoCampus,
          (item.palavrasChave || []).join(" ")
        ].join(" ").toLowerCase();
      }

      function escapeHtml(valor) {
        return String(valor || "")
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;")
          .replace(/'/g, "&#039;");
      }

      function optionHtml(valor) {
        return '<option value="' + escapeHtml(valor) + '">' + escapeHtml(valor) + '</option>';
      }

      function popularFiltros() {
        filtros.forEach(select => {
          const campo = select.dataset.campo;
          const valores = [...new Set(dados.map(item => item[campo]).filter(v => v && v !== "---"))].sort();

          valores.forEach(valor => {
            select.insertAdjacentHTML("beforeend", optionHtml(valor));
          });
        });
      }

      function passaFiltros(item) {
        const termo = (busca.value || "").trim().toLowerCase();

        if (termo && !textoBusca(item).includes(termo)) {
          return false;
        }

        for (const filtro of filtros) {
          const campo = filtro.dataset.campo;
          const valor = filtro.value;

          if (valor && item[campo] !== valor) {
            return false;
          }
        }

        return true;
      }

      function renderBadges(item) {
        const badges = [
          '<span class="leme-badge">' + escapeHtml(item.ano) + '</span>',
          '<span class="leme-badge green">' + escapeHtml(item.nivel) + '</span>',
          '<span class="leme-badge">' + escapeHtml(item.base) + '</span>'
        ];

        if (item.tipoProduto && item.tipoProduto !== "---") {
          badges.push('<span class="leme-badge red">' + escapeHtml(item.tipoProduto) + '</span>');
        }

        return badges.join("");
      }

      function renderTags(tags) {
        if (!tags || !tags.length) return "";
        return '<div class="leme-tags">' + tags.slice(0, 10).map(t =>
          '<span class="leme-tag">' + escapeHtml(t) + '</span>'
        ).join("") + '</div>';
      }

      function renderActions(item) {
        const links = [];

        if (item.linkPdf) {
          links.push('<a class="leme-link pdf" href="' + escapeHtml(item.linkPdf) + '" target="_blank" rel="noopener">📄 PDF</a>');
        }

        if (item.linkProduto) {
          links.push('<a class="leme-link produto" href="' + escapeHtml(item.linkProduto) + '" target="_blank" rel="noopener">📦 Produto</a>');
        }

        return links.length ? '<div class="leme-actions">' + links.join("") + '</div>' : "";
      }

      function renderCards(lista) {
        if (!lista.length) {
          resultados.innerHTML = '<div class="leme-empty">Nenhum trabalho encontrado com os filtros selecionados.</div>';
          return;
        }

        resultados.innerHTML = lista.map(item => `
          <article class="leme-card">
            <h3 class="leme-card-title">${escapeHtml(item.titulo)}</h3>

            <div class="leme-meta">
              <div><b>Autor:</b> ${escapeHtml(item.autor)}</div>
              <div><b>Orientador:</b> ${escapeHtml(item.orientador)}</div>
              <div><b>Instituição/Campus:</b> ${escapeHtml(item.instituicaoCampus)}</div>
            </div>

            <div class="leme-badges">${renderBadges(item)}</div>

            <p class="leme-summary">${escapeHtml(item.resumoCurto || "Resumo não disponível.")}</p>

            <details>
              <summary>Ver detalhes da classificação e do trabalho</summary>
              <div class="leme-detail">
                <p><b>Produto educacional:</b> ${escapeHtml(item.produto)}</p>
                <p><b>Área temática:</b> ${escapeHtml(item.area)}</p>
                <p><b>Finalidade:</b> ${escapeHtml(item.finalidade)}</p>
                <p><b>Público-alvo:</b> ${escapeHtml(item.publicoAlvo)}</p>
                <p><b>Base epistemológica:</b> ${escapeHtml(item.base)}</p>
                <p><b>Resumo:</b> ${escapeHtml(item.resumo || "Resumo não disponível.")}</p>
                ${renderTags(item.palavrasChave)}
              </div>
            </details>

            ${renderActions(item)}
          </article>
        `).join("");
      }

      function renderTable(lista) {
        if (!lista.length) {
          resultados.innerHTML = '<div class="leme-empty">Nenhum trabalho encontrado com os filtros selecionados.</div>';
          return;
        }

        resultados.innerHTML = `
          <table class="leme-table">
            <thead>
              <tr>
                <th>Título</th>
                <th>Autor</th>
                <th>Orientador</th>
                <th>Ano</th>
                <th>Nível</th>
                <th>Base</th>
                <th>Área</th>
                <th>Links</th>
              </tr>
            </thead>
            <tbody>
              ${lista.map(item => `
                <tr>
                  <td><div class="leme-table-title">${escapeHtml(item.titulo)}</div></td>
                  <td>${escapeHtml(item.autor)}</td>
                  <td>${escapeHtml(item.orientador)}</td>
                  <td>${escapeHtml(item.ano)}</td>
                  <td>${escapeHtml(item.nivel)}</td>
                  <td>${escapeHtml(item.base)}</td>
                  <td>${escapeHtml(item.area)}</td>
                  <td>${renderActions(item)}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        `;
      }


      function contarPorCampo(lista, campo) {
        const mapa = new Map();

        lista.forEach(item => {
          const valor = item[campo] && item[campo] !== "---" ? item[campo] : "Não identificado";
          mapa.set(valor, (mapa.get(valor) || 0) + 1);
        });

        return Array.from(mapa.entries())
          .map(([label, valor]) => ({ label, valor }))
          .sort((a, b) => b.valor - a.valor || a.label.localeCompare(b.label));
      }

      function contarPorAno(lista) {
        const mapa = new Map();

        lista.forEach(item => {
          const valor = item.ano && item.ano !== "---" ? item.ano : "S/A";
          mapa.set(valor, (mapa.get(valor) || 0) + 1);
        });

        return Array.from(mapa.entries())
          .map(([label, valor]) => ({ label, valor }))
          .sort((a, b) => String(a.label).localeCompare(String(b.label)));
      }

      function renderYearChart(el, dadosGrafico) {
        if (!el) return;

        if (!dadosGrafico.length) {
          el.innerHTML = '<div class="leme-empty">Sem dados para exibir.</div>';
          return;
        }

        const max = Math.max(...dadosGrafico.map(d => d.valor), 1);

        el.innerHTML = dadosGrafico.map(d => {
          const altura = Math.max(8, Math.round((d.valor / max) * 170));
          return `
            <div class="leme-year-item" title="${escapeHtml(d.label)}: ${d.valor}">
              <div class="leme-year-value">${d.valor}</div>
              <div class="leme-year-bar" style="height:${altura}px"></div>
              <div class="leme-year-label">${escapeHtml(d.label)}</div>
            </div>
          `;
        }).join("");
      }

      function renderHorizontalBars(el, dadosGrafico, limite) {
        if (!el) return;

        const lista = dadosGrafico.slice(0, limite || 8);

        if (!lista.length) {
          el.innerHTML = '<div class="leme-empty">Sem dados para exibir.</div>';
          return;
        }

        const max = Math.max(...lista.map(d => d.valor), 1);

        el.innerHTML = lista.map(d => {
          const largura = Math.max(3, Math.round((d.valor / max) * 100));
          return `
            <div class="leme-bar-row" title="${escapeHtml(d.label)}: ${d.valor}">
              <div class="leme-bar-label">${escapeHtml(d.label)}</div>
              <div class="leme-bar-track">
                <div class="leme-bar-fill" style="width:${largura}%"></div>
              </div>
              <div class="leme-bar-value">${d.valor}</div>
            </div>
          `;
        }).join("");
      }

      function atualizarGraficos(lista) {
        renderYearChart(chartAno, contarPorAno(lista));
        renderHorizontalBars(chartBase, contarPorCampo(lista, "base"), 8);
        renderHorizontalBars(chartArea, contarPorCampo(lista, "area"), 8);
        renderHorizontalBars(chartNivel, contarPorCampo(lista, "nivel"), 8);
        renderHorizontalBars(chartProduto, contarPorCampo(lista, "tipoProduto"), 8);

        if (graficosNota) {
          if (lista.length === dados.length) {
            graficosNota.textContent = "Todos os trabalhos";
          } else {
            graficosNota.textContent = lista.length + " de " + dados.length + " trabalhos filtrados";
          }
        }
      }

      function aplicarFiltros() {
        const filtrados = dados.filter(passaFiltros);

        contador.innerHTML = '<strong>' + filtrados.length + '</strong> trabalhos encontrados de <strong>' + dados.length + '</strong> analisados';

        atualizarGraficos(filtrados);

        resultados.classList.toggle("is-table", modo === "table");

        if (modo === "table") {
          renderTable(filtrados);
        } else {
          renderCards(filtrados);
        }
      }

      function limparFiltros() {
        busca.value = "";
        filtros.forEach(f => f.value = "");
        aplicarFiltros();
        busca.focus();
      }

      busca.addEventListener("input", aplicarFiltros);
      limpar.addEventListener("click", limparFiltros);
      filtros.forEach(f => f.addEventListener("change", aplicarFiltros));

      btnCards.addEventListener("click", function() {
        modo = "cards";
        btnCards.classList.add("is-active");
        btnTable.classList.remove("is-active");
        aplicarFiltros();
      });

      btnTable.addEventListener("click", function() {
        modo = "table";
        btnTable.classList.add("is-active");
        btnCards.classList.remove("is-active");
        aplicarFiltros();
      });

      popularFiltros();
      aplicarFiltros();
    })();
  </script>
</section>
'''

    html = html.replace("__ATUALIZADO_EM__", atualizado_em)
    html = html.replace("__PERIODO__", periodo)
    html = html.replace("__TOTAL__", str(total))
    html = html.replace("__TOTAL_PRODUTOS__", str(total_produtos))
    html = html.replace("__TOTAL_ORIENTADORES__", str(total_orientadores))
    html = html.replace("__TOTAL_BASES__", str(total_bases))
    html = html.replace("__DADOS_JSON__", dados_json)

    return html


def gerar_painel_leme_ept_v2(arquivo_excel=ARQUIVO_ENTRADA):
    print(f"Lendo arquivo: {arquivo_excel}")

    df = pd.read_excel(arquivo_excel)
    df = df.fillna("---")

    registros = preparar_dados(df)

    atualizar_excel_classificado(df, registros)

    html = gerar_html(registros)

    with open(ARQUIVO_HTML_SAIDA, "w", encoding="utf-8") as f:
        f.write(html)

    print("Painel gerado com sucesso!")
    print(f"HTML para Elementor: {ARQUIVO_HTML_SAIDA}")
    print(f"Excel classificado: {ARQUIVO_EXCEL_SAIDA}")
    print(f"Total de trabalhos processados: {len(registros)}")


if __name__ == "__main__":
    gerar_painel_leme_ept_v2()
