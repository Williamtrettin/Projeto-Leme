"""Funções pequenas compartilhadas pelos quatro scripts do PAINEL EPT / LEME."""

from __future__ import annotations

from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
import socket
import unicodedata
from urllib.parse import urlsplit, urlunsplit

import requests
from requests.adapters import HTTPAdapter
import urllib3.util.connection
from urllib3.util.retry import Retry


def agora_utc() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def limpar_texto(valor) -> str:
    if valor is None:
        return ""
    texto = str(valor).replace("\xa0", " ").strip()
    if texto.casefold() in {"nan", "none", "null", "nat"}:
        return ""
    return re.sub(r"\s+", " ", texto)


def normalizar_busca(valor) -> str:
    texto = unicodedata.normalize("NFKD", limpar_texto(valor).casefold())
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def slugificar(valor, limite: int = 110) -> str:
    slug = normalizar_busca(valor).replace(" ", "_")
    return (slug[:limite].strip("_") or "trabalho")


def url_canonica(valor) -> str:
    texto = limpar_texto(valor)
    if not texto:
        return ""
    partes = urlsplit(texto)
    if partes.scheme not in {"http", "https"} or not partes.netloc:
        return texto
    caminho = re.sub(r"/{2,}", "/", partes.path).rstrip("/")
    return urlunsplit((partes.scheme.casefold(), partes.netloc.casefold(), caminho, partes.query, ""))


def handle_educapes(valor) -> str:
    texto = limpar_texto(valor)
    match = re.search(r"(?:handle/)?(capes/\d+)", texto, flags=re.IGNORECASE)
    return match.group(1).casefold() if match else ""


def sha256_arquivo(caminho: Path) -> str:
    resumo = hashlib.sha256()
    with Path(caminho).open("rb") as arquivo:
        for bloco in iter(lambda: arquivo.read(1024 * 1024), b""):
            resumo.update(bloco)
    return resumo.hexdigest()


def ler_json(caminho: Path, padrao=None):
    caminho = Path(caminho)
    if not caminho.is_file():
        return padrao
    with caminho.open(encoding="utf-8") as arquivo:
        return json.load(arquivo)


def salvar_json_atomico(documento, caminho: Path) -> None:
    caminho = Path(caminho)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    temporario = caminho.with_name(caminho.name + ".tmp")
    with temporario.open("w", encoding="utf-8", newline="\n") as arquivo:
        json.dump(documento, arquivo, ensure_ascii=False, indent=2)
        arquivo.write("\n")
    os.replace(temporario, caminho)


def salvar_texto_atomico(texto: str, caminho: Path) -> None:
    caminho = Path(caminho)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    temporario = caminho.with_name(caminho.name + ".tmp")
    temporario.write_text(texto.rstrip() + "\n", encoding="utf-8")
    os.replace(temporario, caminho)


def criar_sessao_http() -> requests.Session:
    # O host do IFC anuncia uma rota IPv6 que não responde em alguns ambientes.
    # Forçar IPv4 evita esperas de vários minutos sem desativar TLS por padrão.
    urllib3.util.connection.allowed_gai_family = lambda: socket.AF_INET
    sessao = requests.Session()
    sessao.headers.update(
        {
            "User-Agent": "PAINEL-EPT-LEME/1.0 (coleta acadêmica pública)",
            "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.5",
        }
    )
    tentativas = Retry(
        total=3,
        connect=3,
        read=3,
        backoff_factor=1,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET"}),
    )
    sessao.mount("http://", HTTPAdapter(max_retries=tentativas))
    sessao.mount("https://", HTTPAdapter(max_retries=tentativas))
    return sessao
