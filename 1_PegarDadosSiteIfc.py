"""Coleta os dados e PDFs e, ao terminar, executa o OCR automaticamente.

Execução:
    python 1_PegarDadosSiteIfc.py
"""

from __future__ import annotations

import argparse
from collections import defaultdict
import json
from pathlib import Path
import re
import subprocess
import sys
from urllib.parse import urljoin, urlsplit

from bs4 import BeautifulSoup
import pandas as pd
import requests

from util import (
    agora_utc,
    criar_sessao_http,
    handle_educapes,
    ler_json,
    limpar_texto,
    normalizar_busca,
    salvar_json_atomico,
    sha256_arquivo,
    slugificar,
    url_canonica,
)


URL_PROFEPT = "https://profept.ifc.edu.br/dissertacoes/"
URL_API_PROFEPT = "https://profept.ifc.edu.br/wp-json/wp/v2/pages?slug=dissertacoes"
ARQUIVO_IDENTIDADES = Path("dados/referencia/identidades.json")
ARQUIVO_SAIDA = Path("dados/1_PegarDadosSiteIfc.xlsx")
PASTA_PDFS = Path("pdf")
PASTA_CACHE = Path("dados/cache")
SCHEMA_IDENTIDADES = "leme-identidades-v1"
INSTITUICAO_CORPUS = "Instituto Federal Catarinense"
CAMPUS_CORPUS = "Blumenau"
PROGRAMA_CORPUS = "ProfEPT"
ORIGEM_CONTEXTO_INSTITUCIONAL = (
    "Página oficial de dissertações do ProfEPT/IFC e metadados complementares "
    "do eduCAPES."
)
HOSTS_FALLBACK_TLS = {"profept.ifc.edu.br"}

ROTULOS = re.compile(
    r"(?P<autor>\b(?:Acad[eê]mic[oa]s?|Alun[oa]s?|Autor(?:es)?|Mestrand[oa]s?)\s*:)|"
    r"(?P<coorientador>\bCoorientador(?:a|es|as)?\s*:)|"
    r"(?P<orientador>\b(?:Orientador(?:a|es|as)?|Orietador(?:a|es|as)?|Orientatador)\s*:)|"
    r"(?P<dissertacao>(?<![-/_])\bDisserta[cç][aã]o(?:\s+Intitulada)?\s*:?\s*)|"
    r"(?P<produto>\bProdutos?\s+Educ(?:acional|ional)\s*:)",
    flags=re.IGNORECASE,
)


def argumentos() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--offline", action="store_true", help="usa somente caches já existentes")
    parser.add_argument("--sem-educapes", action="store_true", help="não consulta metadados do eduCAPES")
    parser.add_argument("--sem-download", action="store_true", help="não baixa PDFs")
    parser.add_argument(
        "--inseguro-tls",
        action="store_true",
        help="desativa verificação TLS explicitamente (somente para contornar o certificado IFC local)",
    )
    parser.add_argument("--limite", type=int, default=0, help="limita registros para diagnóstico")
    return parser.parse_args()


def requisicao_http(
    sessao: requests.Session,
    url: str,
    args: argparse.Namespace,
    **opcoes,
) -> requests.Response:
    """Tenta TLS normal e contorna apenas o certificado conhecido do IFC."""
    inseguro_solicitado = bool(getattr(args, "inseguro_tls", False))
    try:
        return sessao.get(url, verify=not inseguro_solicitado, **opcoes)
    except requests.exceptions.SSLError:
        host = (urlsplit(url).hostname or "").casefold()
        if inseguro_solicitado or host not in HOSTS_FALLBACK_TLS:
            raise
        requests.packages.urllib3.disable_warnings(
            requests.packages.urllib3.exceptions.InsecureRequestWarning
        )
        if not getattr(args, "_fallback_tls_avisado", False):
            print(
                "AVISO: o certificado do site do IFC não pôde ser validado. "
                "A conexão será repetida sem validar o certificado somente para esse domínio."
            )
            args._fallback_tls_avisado = True
        return sessao.get(url, verify=False, **opcoes)


def marcar_links_no_elemento(elemento) -> str:
    copia = BeautifulSoup(str(elemento), "html.parser")
    for ancora in copia.find_all("a"):
        href = urljoin(URL_PROFEPT, limpar_texto(ancora.get("href")))
        texto = limpar_texto(ancora.get_text(" ", strip=True))
        ancora.replace_with(f"{texto} [LINK:{href}]")
    return limpar_texto(copia.get_text(" ", strip=True))


def separar_links(texto: str) -> tuple[str, list[str]]:
    links = [url_canonica(url) for url in re.findall(r"\[LINK:(.*?)\]", texto)]
    return limpar_texto(re.sub(r"\[LINK:.*?\]", "", texto)), links


def fragmentos_rotulados(texto: str) -> list[tuple[str, str, list[str]]]:
    correspondencias = list(ROTULOS.finditer(texto))
    resultado: list[tuple[str, str, list[str]]] = []
    for posicao, match in enumerate(correspondencias):
        fim = correspondencias[posicao + 1].start() if posicao + 1 < len(correspondencias) else len(texto)
        valor, links = separar_links(texto[match.end() : fim])
        campo = next(nome for nome, conteudo in match.groupdict().items() if conteudo)
        resultado.append((campo, valor, links))
    return resultado


def extrair_trabalhos_html(html: str) -> list[dict]:
    """Extrai registros pela sequência de rótulos, tolerando mais de um trabalho por lista."""
    soup = BeautifulSoup(html, "html.parser")
    ano_atual = ""
    atual: dict | None = None
    trabalhos: list[dict] = []

    def concluir() -> None:
        nonlocal atual
        if atual and atual.get("autor"):
            trabalhos.append(atual)
        atual = None

    for elemento in soup.find_all(["h1", "h2", "h3", "h4", "h5", "h6", "p", "li"]):
        if elemento.name != "li" and elemento.find_parent("li") is not None:
            continue
        texto_simples = limpar_texto(elemento.get_text(" ", strip=True))
        if elemento.name != "li":
            if len(texto_simples) <= 80:
                anos = re.findall(r"\b20\d{2}\b", texto_simples)
                if anos and re.search(r"defesa|disserta|20\d{2}", texto_simples, re.IGNORECASE):
                    ano_atual = anos[-1]
            continue
        texto = marcar_links_no_elemento(elemento)
        for campo, valor, links in fragmentos_rotulados(texto):
            if campo == "autor":
                concluir()
                atual = {
                    "ano": ano_atual,
                    "autor": valor,
                    "orientador": "",
                    "coorientador": "",
                    "titulo": "",
                    "produto_educacional": "",
                    "url_pdf": "",
                    "url_produto": "",
                }
                continue
            if atual is None:
                continue
            if campo == "orientador":
                atual["orientador"] = valor
            elif campo == "coorientador":
                atual["coorientador"] = valor
            elif campo == "dissertacao":
                atual["titulo"] = valor
                if links:
                    atual["url_pdf"] = links[0]
            elif campo == "produto":
                atual["produto_educacional"] = valor
                if links:
                    atual["url_produto"] = links[0]
    concluir()

    for posicao, trabalho in enumerate(trabalhos, 1):
        trabalho["ordem_na_pagina"] = posicao
    return trabalhos


def obter_pagina_oficial(sessao: requests.Session, args: argparse.Namespace) -> tuple[str, dict]:
    cache = PASTA_CACHE / "profept_dissertacoes.json"
    if args.offline:
        resposta = ler_json(cache)
        if not resposta:
            raise FileNotFoundError(f"Cache oficial não encontrado: {cache}")
    else:
        requisicao = requisicao_http(sessao, URL_API_PROFEPT, args, timeout=90)
        requisicao.raise_for_status()
        resposta = requisicao.json()
        salvar_json_atomico(resposta, cache)
    if not isinstance(resposta, list) or not resposta:
        raise ValueError("A API oficial não retornou a página de dissertações")
    pagina = resposta[0]
    html = pagina.get("content", {}).get("rendered", "")
    if not html:
        raise ValueError("A página oficial não contém content.rendered")
    metadados = {
        "url": URL_PROFEPT,
        "api": URL_API_PROFEPT,
        "pagina_id": pagina.get("id"),
        "modificado_em": pagina.get("modified"),
        "coletado_em": agora_utc(),
    }
    return html, metadados


def carregar_identidades(caminho: Path = ARQUIVO_IDENTIDADES) -> dict:
    documento = ler_json(caminho)
    if not documento or documento.get("schema_version") != SCHEMA_IDENTIDADES:
        raise ValueError(
            f"Identidades ausentes ou inválidas em {caminho}. "
            "Restaure dados/referencia/identidades.json do Git."
        )
    return documento


def chaves_identidade(identidade: dict) -> dict[str, set]:
    aliases = identidade.get("aliases_conhecidos", {})
    urls = set(aliases.get("urls_pdf", [])) | {identidade.get("url_pdf", "")}
    titulos = set(aliases.get("titulos", [])) | {identidade.get("titulo", "")}
    autores = set(aliases.get("autores", [])) | {identidade.get("autor", "")}
    return {
        "urls": {url_canonica(v) for v in urls if v},
        "titulos": {normalizar_busca(v) for v in titulos if v},
        "autores": {normalizar_busca(v) for v in autores if v},
    }


def resolver_identidades(trabalhos: list[dict], documento: dict) -> list[dict]:
    identidades = documento["identidades"]
    por_url: defaultdict[str, list[dict]] = defaultdict(list)
    por_trinca: defaultdict[tuple[str, str, str], list[dict]] = defaultdict(list)
    for identidade in identidades:
        chaves = chaves_identidade(identidade)
        for url in chaves["urls"]:
            por_url[url].append(identidade)
        for autor in chaves["autores"]:
            for titulo in chaves["titulos"]:
                por_trinca[(autor, titulo, limpar_texto(identidade.get("ano")))].append(identidade)

    usados: set[str] = set()
    proximo = max(
        int(documento.get("proximo_numero", 1)),
        max(int(i["id"].split("-")[1]) for i in identidades) + 1,
    )
    for trabalho in trabalhos:
        url = url_canonica(trabalho.get("url_pdf"))
        trinca = (
            normalizar_busca(trabalho.get("autor")),
            normalizar_busca(trabalho.get("titulo")),
            limpar_texto(trabalho.get("ano")),
        )
        candidatos_url = [i for i in por_url.get(url, []) if i["id"] not in usados] if url else []
        candidatos_trinca = [i for i in por_trinca.get(trinca, []) if i["id"] not in usados]
        candidatos = candidatos_url or candidatos_trinca
        metodo = "url_pdf_exata" if candidatos_url else "autor_titulo_ano_exatos"
        if len(candidatos) == 1:
            identidade = candidatos[0]
        elif not candidatos and all(trinca) and url:
            identificador = f"EPT-{proximo:04d}"
            proximo += 1
            identidade = {
                "id": identificador,
                "titulo": trabalho["titulo"],
                "autor": trabalho["autor"],
                "ano": trabalho["ano"],
                "url_pdf": url,
                "novo_nome_pdf": f"{identificador}_{slugificar(trabalho['titulo'])}.pdf",
                "sha256": "",
                "aliases_conhecidos": {
                    "titulos": [trabalho["titulo"]],
                    "autores": [trabalho["autor"]],
                    "urls_pdf": [url],
                },
                "vinculo_confirmado": True,
            }
            identidades.append(identidade)
            por_url[url].append(identidade)
            por_trinca[trinca].append(identidade)
            metodo = "novo_registro_oficial_campos_fortes_unicos"
        else:
            trabalho.update(
                {
                    "id": None,
                    "vinculo_confirmado": False,
                    "metodo_vinculo": "ambiguo" if candidatos else "evidencia_insuficiente",
                    "candidatos_identidade": [item["id"] for item in candidatos],
                }
            )
            continue

        identificador = identidade["id"]
        usados.add(identificador)
        trabalho.update(
            {
                "id": identificador,
                "vinculo_confirmado": bool(identidade.get("vinculo_confirmado", True)),
                "metodo_vinculo": metodo,
                "candidatos_identidade": [identificador],
            }
        )
        identidade["titulo"] = trabalho["titulo"] or identidade.get("titulo", "")
        identidade["autor"] = trabalho["autor"] or identidade.get("autor", "")
        identidade["ano"] = trabalho["ano"] or identidade.get("ano", "")
        identidade["url_pdf"] = url or identidade.get("url_pdf", "")
        aliases = identidade["aliases_conhecidos"]
        for chave, valor in (("titulos", trabalho["titulo"]), ("autores", trabalho["autor"]), ("urls_pdf", url)):
            if valor and valor not in aliases[chave]:
                aliases[chave].append(valor)

    documento["identidades"].sort(key=lambda item: item["id"])
    documento["proximo_numero"] = proximo
    return trabalhos


def consultar_educapes(sessao: requests.Session, url: str, args: argparse.Namespace) -> dict:
    handle = handle_educapes(url)
    if not handle:
        return {"status": "sem_handle", "handle": "", "campos_dc": {}, "bitstreams": []}
    cache = PASTA_CACHE / "educapes" / f"{slugificar(handle)}.html"
    if cache.exists():
        html = cache.read_text(encoding="utf-8")
        origem = "cache"
    elif args.offline:
        return {"status": "cache_ausente", "handle": handle, "campos_dc": {}, "bitstreams": []}
    else:
        resposta = requisicao_http(
            sessao,
            f"https://educapes.capes.gov.br/handle/{handle}?mode=full",
            args,
            timeout=90,
        )
        resposta.raise_for_status()
        html = resposta.text
        cache.parent.mkdir(parents=True, exist_ok=True)
        temporario = cache.with_suffix(".html.tmp")
        temporario.write_text(html, encoding="utf-8")
        temporario.replace(cache)
        origem = "rede"

    soup = BeautifulSoup(html, "html.parser")
    campos: defaultdict[str, list[str]] = defaultdict(list)
    for rotulo in soup.select(".metadataFieldLabel"):
        valor = rotulo.find_next(class_="metadataFieldValue")
        if valor is None:
            continue
        match = re.search(r"\b(dc\.[a-z0-9_.-]+)", limpar_texto(rotulo.get_text(" ", strip=True)), re.I)
        texto = limpar_texto(valor.get_text(" ", strip=True))
        if match and texto and texto not in campos[match.group(1).casefold()]:
            campos[match.group(1).casefold()].append(texto)
    bitstreams = []
    for ancora in soup.select('a[href*="/bitstream/"]'):
        link = urljoin("https://educapes.capes.gov.br", ancora.get("href", ""))
        if link and link not in bitstreams:
            bitstreams.append(link)
    return {
        "status": "sucesso",
        "origem_consulta": origem,
        "handle": handle,
        "campos_dc": dict(sorted(campos.items())),
        "bitstreams": bitstreams,
    }


def nome_campus(valor: str) -> str:
    palavras = re.findall(r"[0-9A-Za-zÀ-ÖØ-öø-ÿ]+", limpar_texto(valor))
    conectivos = {"da", "das", "de", "do", "dos", "e"}
    return " ".join(
        palavra.casefold() if palavra.casefold() in conectivos else palavra.capitalize()
        for palavra in palavras
    )


def resolver_instituicao_campus(educapes: dict) -> tuple[str, str, str]:
    """Usa campus explícito do eduCAPES sem confundir local pesquisado com campus do programa."""
    contribuidores = educapes.get("campos_dc", {}).get("dc.contributor", [])
    for valor in contribuidores:
        texto = limpar_texto(valor)
        campus = ""
        encontrado = re.search(
            r"\bcampus\s*[-:/]?\s*([A-Za-zÀ-ÖØ-öø-ÿ][A-Za-zÀ-ÖØ-öø-ÿ\s'-]*)",
            texto,
            flags=re.IGNORECASE,
        )
        if encontrado:
            campus = nome_campus(encontrado.group(1))
        else:
            encontrado = re.fullmatch(
                r"\s*IFC\s*[-/]\s*([A-Za-zÀ-ÖØ-öø-ÿ][A-Za-zÀ-ÖØ-öø-ÿ\s'-]*)\s*\.?\s*",
                texto,
                flags=re.IGNORECASE,
            )
            if encontrado:
                campus = nome_campus(encontrado.group(1))
        if campus:
            return INSTITUICAO_CORPUS, campus, "eduCAPES: dc.contributor"
    return (
        INSTITUICAO_CORPUS,
        CAMPUS_CORPUS,
        "contexto do ProfEPT/IFC; eduCAPES sem campus explícito",
    )


def pdf_valido(caminho: Path) -> bool:
    if not caminho.is_file() or caminho.stat().st_size < 1024:
        return False
    with caminho.open("rb") as arquivo:
        return arquivo.read(5) == b"%PDF-"


def pessoa_equivalente(valor_oficial: str, candidatos: list[str]) -> bool:
    def assinatura(valor: str) -> set[str]:
        termos = normalizar_busca(valor).split()
        descartados = {"dr", "dra", "prof", "profa", "me", "mestre", "mestra"}
        return {termo for termo in termos if termo not in descartados and len(termo) > 1}

    oficial = assinatura(valor_oficial)
    for candidato in candidatos:
        termos = assinatura(candidato)
        comuns = oficial & termos
        if comuns and (
            len(comuns) / min(len(oficial), len(termos)) >= 0.75
            or len(comuns) / len(oficial) >= 0.65
        ):
            return True
    return False


def titulo_equivalente(valores_oficiais: list[str], candidatos: list[str]) -> bool:
    oficiais = [set(normalizar_busca(valor).split()) for valor in valores_oficiais if valor]
    for candidato in candidatos:
        termos = set(normalizar_busca(candidato).split())
        for oficial in oficiais:
            if termos and oficial and len(termos & oficial) / min(len(termos), len(oficial)) >= 0.55:
                return True
    return False


def obter_pdf(
    sessao: requests.Session,
    trabalho: dict,
    identidade: dict,
    args: argparse.Namespace,
) -> dict:
    destino = PASTA_PDFS / identidade["novo_nome_pdf"]
    esperado = identidade.get("sha256", "")
    if destino.exists():
        hash_atual = sha256_arquivo(destino) if pdf_valido(destino) else ""
        status = "validado" if hash_atual and (not esperado or hash_atual == esperado) else "conflito_local"
    elif args.sem_download:
        hash_atual = ""
        status = "não_processado_por_opção"
    elif args.offline:
        hash_atual = ""
        status = "pdf_ausente_offline"
    else:
        url = trabalho.get("url_pdf", "")
        if not url:
            return {"caminho": str(destino), "status": "url_pdf_ausente", "sha256": ""}
        destino.parent.mkdir(parents=True, exist_ok=True)
        temporario = destino.with_suffix(".pdf.tmp")
        try:
            with requisicao_http(sessao, url, args, timeout=180, stream=True) as resposta:
                resposta.raise_for_status()
                with temporario.open("wb") as arquivo:
                    for bloco in resposta.iter_content(1024 * 1024):
                        if bloco:
                            arquivo.write(bloco)
            if not pdf_valido(temporario):
                raise ValueError("conteúdo baixado não é um PDF válido")
            hash_atual = sha256_arquivo(temporario)
            temporario.replace(destino)
            status = "baixado"
        finally:
            temporario.unlink(missing_ok=True)

    if hash_atual and not identidade.get("sha256"):
        identidade["sha256"] = hash_atual
    return {"caminho": str(destino), "status": status, "sha256": hash_atual}


def registros_para_planilha(registros: list[dict]) -> pd.DataFrame:
    """Transforma a coleta estruturada na entrada pública e estável da etapa 2."""
    linhas = []
    for registro in registros:
        arquivo = registro.get("arquivo_pdf", {})
        validacao = registro.get("validacao", {})
        linhas.append(
            {
                "ID": registro.get("id"),
                "Ano": registro.get("ano"),
                "Título da Dissertação": registro.get("titulo"),
                "Autor": registro.get("autor"),
                "Orientador": registro.get("orientador"),
                "Coorientador": registro.get("coorientador"),
                "Instituição": registro.get("instituicao"),
                "Campus": registro.get("campus"),
                "Fonte do Campus": registro.get("fonte_campus"),
                "Programa": registro.get("programa"),
                "Produto Educacional": registro.get("produto_educacional"),
                "Link da Página": registro.get("url_oficial"),
                "Link do PDF": registro.get("url_pdf"),
                "Link do Produto": registro.get("url_produto"),
                "Handle eduCAPES": registro.get("handle_educapes"),
                "Nome do PDF": Path(arquivo.get("caminho", "")).name,
                "SHA-256": arquivo.get("sha256"),
                "Status do PDF": arquivo.get("status"),
                "Vínculo Confirmado": validacao.get("vinculo_confirmado", False),
                "Método do Vínculo": validacao.get("metodo_vinculo"),
                "Ordem Atual na Página": registro.get("ordem_na_pagina"),
                "Conflitos de Fonte": json.dumps(
                    validacao.get("conflitos", []), ensure_ascii=False
                ),
            }
        )
    return pd.DataFrame(linhas)


def salvar_planilha_atomica(registros: list[dict], caminho: Path = ARQUIVO_SAIDA) -> None:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    temporario = caminho.with_name(caminho.stem + ".tmp.xlsx")
    try:
        registros_para_planilha(registros).to_excel(temporario, index=False)
        temporario.replace(caminho)
    finally:
        temporario.unlink(missing_ok=True)


def executar(args: argparse.Namespace) -> dict:
    if args.inseguro_tls:
        requests.packages.urllib3.disable_warnings(requests.packages.urllib3.exceptions.InsecureRequestWarning)
        print("AVISO: verificação TLS desativada explicitamente para esta execução.")
    sessao = criar_sessao_http()
    html, fonte = obter_pagina_oficial(sessao, args)
    trabalhos = extrair_trabalhos_html(html)
    if not trabalhos:
        raise ValueError("Nenhum trabalho foi extraído da página oficial")
    if args.limite:
        trabalhos = trabalhos[: args.limite]

    documento_identidades = carregar_identidades()
    resolver_identidades(trabalhos, documento_identidades)
    por_id = {item["id"]: item for item in documento_identidades["identidades"]}
    registros = []
    for numero, trabalho in enumerate(trabalhos, 1):
        if numero == 1 or numero % 10 == 0:
            print(f"Processando registro {numero}/{len(trabalhos)}...")
        identificador = trabalho.get("id")
        identidade = por_id.get(identificador) if identificador else None
        educapes = (
            consultar_educapes(sessao, trabalho.get("url_produto", ""), args)
            if not args.sem_educapes
            else {
                "status": "não_consultado_por_opção",
                "handle": handle_educapes(trabalho.get("url_produto")),
                "campos_dc": {},
                "bitstreams": [],
            }
        )
        arquivo_pdf = (
            obter_pdf(sessao, trabalho, identidade, args)
            if identidade
            else {"caminho": "", "status": "identidade_não_resolvida", "sha256": ""}
        )
        instituicao, campus, fonte_campus = resolver_instituicao_campus(educapes)
        conflitos = []
        autores_educapes = educapes.get("campos_dc", {}).get("dc.contributor.author", [])
        titulos_educapes = educapes.get("campos_dc", {}).get("dc.title", [])
        if autores_educapes and not pessoa_equivalente(trabalho["autor"], autores_educapes):
            conflitos.append({"campo": "autor", "oficial_profept": trabalho["autor"], "educapes": autores_educapes})
        if (
            titulos_educapes
            and trabalho["produto_educacional"]
            and not titulo_equivalente([trabalho["produto_educacional"]], titulos_educapes)
        ):
            conflitos.append(
                {
                    "campo": "titulo_produto",
                    "oficial_profept": trabalho["produto_educacional"],
                    "educapes": titulos_educapes,
                }
            )
        registros.append(
            {
                "id": identificador,
                "ano": trabalho["ano"] or None,
                "autor": trabalho["autor"] or None,
                "orientador": trabalho["orientador"] or None,
                "coorientador": trabalho["coorientador"] or None,
                "titulo": trabalho["titulo"] or None,
                "produto_educacional": trabalho["produto_educacional"] or None,
                "instituicao": instituicao,
                "campus": campus,
                "programa": PROGRAMA_CORPUS,
                "instituicao_campus": f"IFC {campus}",
                "fonte_campus": fonte_campus,
                "url_oficial": URL_PROFEPT,
                "url_pdf": trabalho["url_pdf"] or None,
                "url_produto": trabalho["url_produto"] or None,
                "handle_educapes": educapes.get("handle") or None,
                "ordem_na_pagina": trabalho["ordem_na_pagina"],
                "arquivo_pdf": arquivo_pdf,
                "metadados_educapes": educapes,
                "proveniencia": {
                    "campos_oficiais": "Página oficial de dissertações do ProfEPT no IFC",
                    "contexto_institucional": ORIGEM_CONTEXTO_INSTITUCIONAL,
                    "metadados_complementares": "eduCAPES" if educapes.get("status") == "sucesso" else None,
                    "coletado_em": agora_utc(),
                },
                "validacao": {
                    "vinculo_confirmado": trabalho["vinculo_confirmado"],
                    "metodo_vinculo": trabalho["metodo_vinculo"],
                    "candidatos_identidade": trabalho["candidatos_identidade"],
                    "conflitos": conflitos,
                },
            }
        )

    ids = [registro["id"] for registro in registros if registro["id"]]
    pendentes = [
        registro.get("id") or f"ordem-{registro['ordem_na_pagina']}"
        for registro in registros
        if not registro["validacao"]["vinculo_confirmado"]
    ]
    duplicados = sorted({item for item in ids if ids.count(item) > 1})
    saida = {
        "schema_version": "leme-coleta-site-v2",
        "gerado_em": agora_utc(),
        "fonte": fonte,
        "resumo": {
            "quantidade": len(registros),
            "vinculos_confirmados": len(registros) - len(pendentes),
            "vinculos_para_revisao": pendentes,
            "ids_duplicados": duplicados,
        },
        "trabalhos": registros,
    }
    salvar_json_atomico(documento_identidades, ARQUIVO_IDENTIDADES)
    salvar_planilha_atomica(registros, ARQUIVO_SAIDA)
    if duplicados:
        raise ValueError(f"IDs duplicados após resolução: {duplicados}")
    return saida


def main() -> int:
    args = argumentos()
    resultado = executar(args)
    print(json.dumps(resultado["resumo"], ensure_ascii=False, indent=2))
    if resultado["resumo"]["ids_duplicados"]:
        return 2

    script_ocr = Path(__file__).with_name("2_PdfParaJson.py")
    comando = [sys.executable, str(script_ocr)]
    if args.limite:
        comando.extend(["--limite", str(args.limite)])
    print("\nColeta concluída. Iniciando o OCR...")
    return subprocess.run(comando, cwd=script_ocr.parent, check=False).returncode


if __name__ == "__main__":
    raise SystemExit(main())
