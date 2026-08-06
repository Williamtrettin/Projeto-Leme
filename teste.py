import json
import re
import sys
import time
import unicodedata
from pathlib import Path

import requests


# ============================================================
# CONFIGURAÇÕES OBRIGATÓRIAS
# ============================================================

MODELO = "qwen2.5:7b-instruct"
OLLAMA_URL = "http://localhost:11434/api/generate"

ARQUIVO_ENTRADA = "lotes_json/lote_001.json"
ARQUIVO_SAIDA = "teste_json/classificacao_lote_001.json"

TIMEOUT_REQUISICAO = 900


# ============================================================
# LIMITES PARA REDUZIR O TEXTO ENVIADO À IA
# ============================================================

LIMITE_RESUMO = 2500
LIMITE_METODOLOGIA = 1200
LIMITE_OBJETIVOS = 1200
LIMITE_PRODUTO = 800
LIMITE_PALAVRAS_CHAVE = 500


# ============================================================
# PROMPT DA IA
# ============================================================

PROMPT_FIXO = """
Você é um classificador acadêmico de dissertações e produtos educacionais do ProfEPT.

Sua tarefa é analisar os dados do trabalho e retornar apenas um JSON válido.

Classifique o trabalho como um todo, considerando título, autor, orientador, ano, resumo, metodologia, objetivos, produto educacional e palavras-chave.

Regras obrigatórias:
1. Responda somente em JSON válido.
2. Não use Markdown.
3. Não escreva explicações fora do JSON.
4. Não invente informação sem evidência.
5. Nunca responda apenas "Sim", "Não", "Existe", "Presente", "Ausente", "Verdadeiro" ou "Falso" em campos de classificação.
6. Cada campo de classificação deve conter o valor específico encontrado ou inferido.
7. Se não houver evidência suficiente, use "Não identificado".
8. Se não tiver certeza, escolha a opção mais provável, use confiança "Baixa" e coloque revisao_manual como true.
9. A justificativa deve ser curta.
10. A evidência deve ser um pequeno trecho textual usado para decidir.

Campos esperados:
- linha_pesquisa: nome da linha de pesquisa, não responda "Sim".
- macroprojeto: nome ou código do macroprojeto, não responda "Sim".
- tipo_produto: tipo específico do produto educacional.
- nivel_aplicacao: nível ou contexto de aplicação.
- publico_alvo: público-alvo específico.
- finalidade: finalidade principal do trabalho/produto.
- base_epistemologica: base teórica/epistemológica identificada.

Formato obrigatório da resposta:

{
  "linha_pesquisa": "",
  "macroprojeto": "",
  "tipo_produto": "",
  "nivel_aplicacao": "",
  "publico_alvo": "",
  "finalidade": "",
  "base_epistemologica": "",
  "confianca": {
    "linha_pesquisa": "Alta/Média/Baixa",
    "macroprojeto": "Alta/Média/Baixa",
    "tipo_produto": "Alta/Média/Baixa",
    "nivel_aplicacao": "Alta/Média/Baixa",
    "publico_alvo": "Alta/Média/Baixa",
    "finalidade": "Alta/Média/Baixa",
    "base_epistemologica": "Alta/Média/Baixa"
  },
  "justificativas": {
    "linha_pesquisa": "",
    "macroprojeto": "",
    "tipo_produto": "",
    "nivel_aplicacao": "",
    "publico_alvo": "",
    "finalidade": "",
    "base_epistemologica": ""
  },
  "evidencias": {
    "linha_pesquisa": "",
    "macroprojeto": "",
    "tipo_produto": "",
    "nivel_aplicacao": "",
    "publico_alvo": "",
    "finalidade": "",
    "base_epistemologica": ""
  },
  "revisao_manual": false,
  "observacoes": ""
}
""".strip()


CAMPOS_CLASSIFICACAO = [
    "linha_pesquisa",
    "macroprojeto",
    "tipo_produto",
    "nivel_aplicacao",
    "publico_alvo",
    "finalidade",
    "base_epistemologica",
]

RESPOSTAS_GENERICAS_PROIBIDAS = {
    "sim",
    "não",
    "nao",
    "existe",
    "presente",
    "ausente",
    "verdadeiro",
    "falso",
    "yes",
    "no",
    "true",
    "false",
}


# ============================================================
# FUNÇÕES DE TEXTO E NORMALIZAÇÃO
# ============================================================

def normalizar_chave(texto):
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    texto = re.sub(r"[^a-z0-9]+", "_", texto)
    texto = re.sub(r"_+", "_", texto).strip("_")
    return texto


def cortar_texto(valor, limite):
    if valor is None:
        return ""

    texto = str(valor).strip()

    if len(texto) <= limite:
        return texto

    return texto[:limite].strip() + " ...[texto cortado para melhorar desempenho]"


def texto_curto(texto, limite=80):
    texto = str(texto or "").replace("\n", " ").strip()
    if len(texto) <= limite:
        return texto
    return texto[:limite].strip() + "..."


def formatar_tempo(segundos):
    segundos = float(segundos or 0)

    if segundos < 60:
        return f"{segundos:.2f}s"

    minutos = int(segundos // 60)
    resto = segundos % 60
    return f"{minutos}min {resto:.2f}s"


def nanos_para_segundos(valor_nanos):
    if not valor_nanos:
        return 0.0
    return valor_nanos / 1_000_000_000


# ============================================================
# LEITURA DO JSON
# ============================================================

def carregar_json(caminho):
    arquivo = Path(caminho)

    if not arquivo.exists():
        raise FileNotFoundError(f"Arquivo de entrada não encontrado: {arquivo.resolve()}")

    try:
        with arquivo.open("r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as erro:
        raise ValueError(f"JSON de entrada mal formatado: {erro}") from erro


def detectar_lista_trabalhos(dados):
    if isinstance(dados, list):
        return dados

    if isinstance(dados, dict):
        chaves_possiveis = [
            "trabalhos",
            "dados",
            "items",
            "itens",
            "registros",
            "resultados",
            "data",
        ]

        mapa = {normalizar_chave(k): k for k in dados.keys()}

        for chave in chaves_possiveis:
            chave_norm = normalizar_chave(chave)
            if chave_norm in mapa:
                valor = dados[mapa[chave_norm]]
                if isinstance(valor, list):
                    return valor

        # Caso tenha só uma lista dentro do objeto, usa ela.
        listas = [v for v in dados.values() if isinstance(v, list)]
        if len(listas) == 1:
            return listas[0]

    raise ValueError(
        "Não consegui identificar a lista de trabalhos. "
        "O JSON deve ser uma lista ou conter uma chave como "
        "'trabalhos', 'dados', 'items' ou 'registros'."
    )


# ============================================================
# EXTRAÇÃO ROBUSTA DE CAMPOS
# ============================================================

def buscar_valor_recursivo(objeto, candidatos):
    """
    Busca um campo dentro de dicionários/listas, mesmo se estiver aninhado.
    Compara as chaves normalizadas para evitar problema com acento, espaço etc.
    """
    candidatos_norm = {normalizar_chave(c) for c in candidatos}

    if isinstance(objeto, dict):
        # Primeiro tenta busca direta no nível atual
        for chave, valor in objeto.items():
            if normalizar_chave(chave) in candidatos_norm:
                if valor not in [None, ""]:
                    return valor

        # Depois busca recursiva
        for valor in objeto.values():
            encontrado = buscar_valor_recursivo(valor, candidatos)
            if encontrado not in [None, ""]:
                return encontrado

    elif isinstance(objeto, list):
        for item in objeto:
            encontrado = buscar_valor_recursivo(item, candidatos)
            if encontrado not in [None, ""]:
                return encontrado

    return ""


def montar_trabalho_reduzido(trabalho, indice):
    """
    Monta apenas os campos úteis para mandar para a IA.
    Também cria fallback para ID e título caso não encontre.
    """
    id_trabalho = buscar_valor_recursivo(
        trabalho,
        [
            "id",
            "ID",
            "codigo",
            "código",
            "identificador",
            "id_trabalho",
            "trabalho_id",
            "numero",
            "número",
        ],
    )

    titulo = buscar_valor_recursivo(
        trabalho,
        [
            "titulo",
            "título",
            "Titulo",
            "Título",
            "title",
            "nome_trabalho",
            "titulo_trabalho",
            "título do trabalho",
            "titulo do trabalho",
            "nome",
        ],
    )

    autor = buscar_valor_recursivo(
        trabalho,
        [
            "autor",
            "Autor",
            "autores",
            "Autores",
            "discente",
            "aluno",
            "estudante",
        ],
    )

    orientador = buscar_valor_recursivo(
        trabalho,
        [
            "orientador",
            "Orientador",
            "orientadora",
            "Orientadora",
            "professor_orientador",
        ],
    )

    ano = buscar_valor_recursivo(
        trabalho,
        [
            "ano",
            "Ano",
            "ano_defesa",
            "ano_publicacao",
            "data",
            "data_defesa",
        ],
    )

    resumo = buscar_valor_recursivo(
        trabalho,
        [
            "resumo",
            "Resumo",
            "abstract",
            "sinopse",
        ],
    )

    metodologia = buscar_valor_recursivo(
        trabalho,
        [
            "metodologia",
            "Metodologia",
            "método",
            "metodo",
            "procedimentos_metodologicos",
            "procedimentos metodológicos",
        ],
    )

    objetivos = buscar_valor_recursivo(
        trabalho,
        [
            "objetivos",
            "Objetivos",
            "objetivo",
            "Objetivo",
            "objetivo_geral",
            "objetivos_especificos",
        ],
    )

    produto = buscar_valor_recursivo(
        trabalho,
        [
            "produto",
            "Produto",
            "produto_educacional",
            "Produto Educacional",
            "tipo_produto",
            "Tipo de Produto",
            "produto educacional",
            "descricao_produto",
        ],
    )

    palavras_chave = buscar_valor_recursivo(
        trabalho,
        [
            "palavras_chave",
            "palavras-chave",
            "Palavras-chave",
            "Palavras-Chave",
            "keywords",
            "descritores",
        ],
    )

    if not id_trabalho:
        id_trabalho = f"item_{indice:03d}"

    if not titulo:
        titulo = "Título não identificado"

    return {
        "id": str(id_trabalho).strip(),
        "titulo": str(titulo).strip(),
        "autor": str(autor or "").strip(),
        "orientador": str(orientador or "").strip(),
        "ano": str(ano or "").strip(),
        "resumo": cortar_texto(resumo, LIMITE_RESUMO),
        "metodologia": cortar_texto(metodologia, LIMITE_METODOLOGIA),
        "objetivos": cortar_texto(objetivos, LIMITE_OBJETIVOS),
        "produto": cortar_texto(produto, LIMITE_PRODUTO),
        "palavras_chave": cortar_texto(palavras_chave, LIMITE_PALAVRAS_CHAVE),
    }


def montar_prompt(trabalho_reduzido):
    trabalho_json = json.dumps(
        trabalho_reduzido,
        ensure_ascii=False,
        indent=2
    )

    return f"{PROMPT_FIXO}\n\nDADOS DO TRABALHO:\n{trabalho_json}"


# ============================================================
# OLLAMA
# ============================================================

def chamar_ollama(prompt):
    payload = {
        "model": MODELO,
        "prompt": prompt,
        "stream": False,
        "format": "json",
        "options": {
            "temperature": 0.1,
            "num_predict": 700,
        },
    }

    inicio_python = time.perf_counter()

    try:
        resposta = requests.post(
            OLLAMA_URL,
            json=payload,
            timeout=TIMEOUT_REQUISICAO
        )
    except requests.exceptions.ConnectionError as erro:
        raise ConnectionError(
            "Não consegui conectar ao Ollama. Verifique se ele está rodando com: ollama serve"
        ) from erro
    except requests.exceptions.Timeout as erro:
        raise TimeoutError(
            f"Timeout após {TIMEOUT_REQUISICAO}s. O modelo demorou demais para responder."
        ) from erro

    tempo_python = time.perf_counter() - inicio_python

    if resposta.status_code == 404:
        raise RuntimeError(
            f"Modelo não encontrado: {MODELO}. Rode: ollama pull {MODELO}"
        )

    if resposta.status_code != 200:
        raise RuntimeError(
            f"Erro na API do Ollama. Status: {resposta.status_code}. Resposta: {resposta.text}"
        )

    try:
        dados = resposta.json()
    except json.JSONDecodeError as erro:
        raise ValueError(f"A API do Ollama não retornou JSON válido: {resposta.text}") from erro

    resposta_ia = dados.get("response", "")

    total_duration = nanos_para_segundos(dados.get("total_duration", 0))
    load_duration = nanos_para_segundos(dados.get("load_duration", 0))
    prompt_eval_duration = nanos_para_segundos(dados.get("prompt_eval_duration", 0))
    eval_duration = nanos_para_segundos(dados.get("eval_duration", 0))

    prompt_eval_count = dados.get("prompt_eval_count", 0) or 0
    eval_count = dados.get("eval_count", 0) or 0

    if eval_duration > 0 and eval_count > 0:
        tokens_por_segundo = eval_count / eval_duration
    elif tempo_python > 0 and eval_count > 0:
        tokens_por_segundo = eval_count / tempo_python
    else:
        tokens_por_segundo = 0.0

    metricas = {
        "tempo_python_segundos": round(tempo_python, 4),
        "tempo_total_ollama_segundos": round(total_duration, 4),
        "tempo_carregamento_modelo_segundos": round(load_duration, 4),
        "tempo_processamento_prompt_segundos": round(prompt_eval_duration, 4),
        "tempo_geracao_resposta_segundos": round(eval_duration, 4),
        "tokens_prompt": prompt_eval_count,
        "tokens_resposta": eval_count,
        "tokens_por_segundo": round(tokens_por_segundo, 2),
    }

    return resposta_ia, metricas


# ============================================================
# TRATAMENTO DA RESPOSTA DA IA
# ============================================================

def remover_markdown(texto):
    texto = str(texto or "").strip()
    texto = re.sub(r"^```json\s*", "", texto, flags=re.IGNORECASE)
    texto = re.sub(r"^```\s*", "", texto)
    texto = re.sub(r"\s*```$", "", texto)
    return texto.strip()


def extrair_json_da_resposta(texto):
    texto = remover_markdown(texto)

    try:
        return json.loads(texto)
    except json.JSONDecodeError:
        pass

    inicio = texto.find("{")
    fim = texto.rfind("}")

    if inicio != -1 and fim != -1 and fim > inicio:
        trecho = texto[inicio:fim + 1]
        try:
            return json.loads(trecho)
        except json.JSONDecodeError:
            pass

    raise ValueError("A IA não retornou JSON válido.")


def valor_generico(valor):
    valor_norm = normalizar_chave(str(valor or ""))
    return valor_norm in {normalizar_chave(v) for v in RESPOSTAS_GENERICAS_PROIBIDAS}


def garantir_estrutura_resposta(classificacao):
    if not isinstance(classificacao, dict):
        raise ValueError("A resposta da IA não é um objeto JSON.")

    for campo in CAMPOS_CLASSIFICACAO:
        classificacao.setdefault(campo, "Não identificado")

    classificacao.setdefault("confianca", {})
    classificacao.setdefault("justificativas", {})
    classificacao.setdefault("evidencias", {})

    for campo in CAMPOS_CLASSIFICACAO:
        classificacao["confianca"].setdefault(campo, "Baixa")
        classificacao["justificativas"].setdefault(campo, "")
        classificacao["evidencias"].setdefault(campo, "")

    classificacao.setdefault("revisao_manual", False)
    classificacao.setdefault("observacoes", "")

    return classificacao


def sanitizar_respostas_genericas(classificacao):
    """
    Se a IA responder 'Sim', 'Não', 'Existe' etc.,
    o código troca por 'Não identificado' e marca revisão manual.
    """
    observacoes = []

    for campo in CAMPOS_CLASSIFICACAO:
        valor = classificacao.get(campo, "")

        if valor_generico(valor):
            classificacao[campo] = "Não identificado"
            classificacao["confianca"][campo] = "Baixa"
            classificacao["justificativas"][campo] = (
                "Resposta genérica da IA removida automaticamente."
            )
            classificacao["evidencias"][campo] = ""
            classificacao["revisao_manual"] = True
            observacoes.append(f"Campo '{campo}' veio com resposta genérica e foi corrigido.")

    if observacoes:
        obs_antiga = classificacao.get("observacoes", "")
        nova_obs = " ".join(observacoes)

        if obs_antiga:
            classificacao["observacoes"] = obs_antiga + " " + nova_obs
        else:
            classificacao["observacoes"] = nova_obs

    return classificacao


def marcar_revisao_por_confianca(classificacao):
    confiancas = classificacao.get("confianca", {})

    for campo in CAMPOS_CLASSIFICACAO:
        conf = str(confiancas.get(campo, "")).strip().lower()

        if conf in ["baixa", "baixo"]:
            classificacao["revisao_manual"] = True
            return classificacao

    return classificacao


def tratar_resposta_ia(resposta_bruta):
    classificacao = extrair_json_da_resposta(resposta_bruta)
    classificacao = garantir_estrutura_resposta(classificacao)
    classificacao = sanitizar_respostas_genericas(classificacao)
    classificacao = marcar_revisao_por_confianca(classificacao)
    return classificacao


# ============================================================
# SAÍDA E CONTINUAÇÃO
# ============================================================

def criar_saida_nova():
    return {
        "arquivo_entrada": ARQUIVO_ENTRADA,
        "modelo": MODELO,
        "total_processados": 0,
        "total_sucesso": 0,
        "total_erros": 0,
        "desempenho_geral": {
            "tempo_total_segundos": 0,
            "tempo_total_formatado": "",
            "tempo_medio_por_trabalho_segundos": 0,
            "tempo_medio_por_trabalho_formatado": "",
            "total_tokens_prompt": 0,
            "total_tokens_resposta": 0,
            "media_tokens_por_segundo": 0,
        },
        "resultados": [],
        "erros": [],
    }


def salvar_saida(saida, caminho):
    caminho_saida = Path(caminho)
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)

    with caminho_saida.open("w", encoding="utf-8") as f:
        json.dump(saida, f, ensure_ascii=False, indent=2)


def carregar_saida_existente(caminho):
    arquivo = Path(caminho)

    if not arquivo.exists():
        return None

    try:
        with arquivo.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def perguntar_modo_continuacao():
    arquivo = Path(ARQUIVO_SAIDA)

    if not arquivo.exists():
        return "novo"

    print("\nJá existe um arquivo de saída:")
    print(arquivo.resolve())
    print("\nEscolha uma opção:")
    print("1 - Continuar de onde parou")
    print("2 - Reprocessar tudo")

    while True:
        escolha = input("Digite 1 ou 2: ").strip()

        if escolha == "1":
            return "continuar"

        if escolha == "2":
            return "novo"

        print("Opção inválida. Digite 1 ou 2.")


def ids_ja_processados(saida):
    ids = set()

    for item in saida.get("resultados", []):
        id_item = str(item.get("id", "")).strip()
        if id_item:
            ids.add(id_item)

    for item in saida.get("erros", []):
        id_item = str(item.get("id", "")).strip()
        if id_item:
            ids.add(id_item)

    return ids


def recalcular_totais(saida):
    saida["total_sucesso"] = len(saida.get("resultados", []))
    saida["total_erros"] = len(saida.get("erros", []))
    saida["total_processados"] = saida["total_sucesso"] + saida["total_erros"]


def atualizar_desempenho_geral(saida, tempo_total):
    metricas = []

    for item in saida.get("resultados", []):
        m = item.get("metricas_desempenho", {})
        if m:
            metricas.append(m)

    for item in saida.get("erros", []):
        m = item.get("metricas_desempenho", {})
        if m:
            metricas.append(m)

    total_tokens_prompt = sum(m.get("tokens_prompt", 0) for m in metricas)
    total_tokens_resposta = sum(m.get("tokens_resposta", 0) for m in metricas)

    velocidades = [
        m.get("tokens_por_segundo", 0)
        for m in metricas
        if m.get("tokens_por_segundo", 0) > 0
    ]

    if velocidades:
        media_tokens_por_segundo = sum(velocidades) / len(velocidades)
    else:
        media_tokens_por_segundo = 0

    total_processados = saida.get("total_processados", 0)

    if total_processados > 0:
        tempo_medio = tempo_total / total_processados
    else:
        tempo_medio = 0

    saida["desempenho_geral"] = {
        "tempo_total_segundos": round(tempo_total, 4),
        "tempo_total_formatado": formatar_tempo(tempo_total),
        "tempo_medio_por_trabalho_segundos": round(tempo_medio, 4),
        "tempo_medio_por_trabalho_formatado": formatar_tempo(tempo_medio),
        "total_tokens_prompt": total_tokens_prompt,
        "total_tokens_resposta": total_tokens_resposta,
        "media_tokens_por_segundo": round(media_tokens_por_segundo, 2),
    }


# ============================================================
# PROCESSAMENTO PRINCIPAL
# ============================================================

def processar_lote():
    dados = carregar_json(ARQUIVO_ENTRADA)
    trabalhos = detectar_lista_trabalhos(dados)

    modo = perguntar_modo_continuacao()

    if modo == "continuar":
        saida = carregar_saida_existente(ARQUIVO_SAIDA)
        if saida is None:
            print("Não consegui ler o arquivo anterior. Vou começar do zero.")
            saida = criar_saida_nova()
    else:
        saida = criar_saida_nova()

    recalcular_totais(saida)

    processados = ids_ja_processados(saida)
    total = len(trabalhos)

    print("\n" + "=" * 80)
    print("INICIANDO PROCESSAMENTO")
    print("=" * 80)
    print(f"Entrada: {ARQUIVO_ENTRADA}")
    print(f"Saída: {ARQUIVO_SAIDA}")
    print(f"Modelo: {MODELO}")
    print(f"Total no lote: {total}")
    print(f"Já processados no arquivo de saída: {len(processados)}")
    print("=" * 80)

    inicio_lote = time.perf_counter()

    for indice, trabalho in enumerate(trabalhos, start=1):
        trabalho_reduzido = montar_trabalho_reduzido(trabalho, indice)

        id_trabalho = trabalho_reduzido["id"]
        titulo_trabalho = trabalho_reduzido["titulo"]

        if modo == "continuar" and id_trabalho in processados:
            print(f"[{indice}/{total}] PULANDO já processado | ID: {id_trabalho} | {texto_curto(titulo_trabalho)}")
            continue

        print(f"\n[{indice}/{total}] Processando")
        print(f"ID: {id_trabalho}")
        print(f"Título: {texto_curto(titulo_trabalho, 100)}")

        resposta_bruta = ""
        metricas = {}

        try:
            prompt = montar_prompt(trabalho_reduzido)

            resposta_bruta, metricas = chamar_ollama(prompt)

            classificacao = tratar_resposta_ia(resposta_bruta)

            resultado_item = {
                "id": id_trabalho,
                "titulo": titulo_trabalho,
                "dados_entrada_usados": trabalho_reduzido,
                "classificacao": classificacao,
                "metricas_desempenho": metricas,
            }

            saida["resultados"].append(resultado_item)

            print("Status: OK")
            print(f"Tempo Python: {formatar_tempo(metricas.get('tempo_python_segundos', 0))}")
            print(f"Tempo Ollama: {formatar_tempo(metricas.get('tempo_total_ollama_segundos', 0))}")
            print(f"Tokens prompt: {metricas.get('tokens_prompt', 0)}")
            print(f"Tokens resposta: {metricas.get('tokens_resposta', 0)}")
            print(f"Tokens/s: {metricas.get('tokens_por_segundo', 0)}")

        except Exception as erro:
            erro_item = {
                "id": id_trabalho,
                "titulo": titulo_trabalho,
                "erro": str(erro),
                "resposta_bruta_ia": resposta_bruta,
                "dados_entrada_usados": trabalho_reduzido,
                "metricas_desempenho": metricas,
            }

            saida["erros"].append(erro_item)

            print("Status: ERRO")
            print(f"Erro: {erro}")

        recalcular_totais(saida)

        tempo_parcial = time.perf_counter() - inicio_lote
        atualizar_desempenho_geral(saida, tempo_parcial)

        salvar_saida(saida, ARQUIVO_SAIDA)

        print("-" * 80)

    tempo_total = time.perf_counter() - inicio_lote
    recalcular_totais(saida)
    atualizar_desempenho_geral(saida, tempo_total)
    salvar_saida(saida, ARQUIVO_SAIDA)

    return saida


# ============================================================
# MAIN
# ============================================================

def main():
    try:
        resultado = processar_lote()

        desempenho = resultado["desempenho_geral"]

        print("\n" + "=" * 80)
        print("PROCESSAMENTO FINALIZADO")
        print("=" * 80)
        print(f"Total processado: {resultado['total_processados']}")
        print(f"Total com sucesso: {resultado['total_sucesso']}")
        print(f"Total com erro: {resultado['total_erros']}")
        print(f"Tempo total: {desempenho['tempo_total_formatado']}")
        print(f"Tempo médio por trabalho: {desempenho['tempo_medio_por_trabalho_formatado']}")
        print(f"Total tokens do prompt: {desempenho['total_tokens_prompt']}")
        print(f"Total tokens da resposta: {desempenho['total_tokens_resposta']}")
        print(f"Média tokens/s: {desempenho['media_tokens_por_segundo']}")
        print(f"Arquivo gerado: {Path(ARQUIVO_SAIDA).resolve()}")

    except KeyboardInterrupt:
        print("\nProcessamento interrompido pelo usuário.")
        print("O progresso já salvo continua no arquivo de saída.")
        sys.exit(1)

    except Exception as erro:
        print("\nERRO GERAL")
        print(str(erro))
        sys.exit(1)


if __name__ == "__main__":
    main()