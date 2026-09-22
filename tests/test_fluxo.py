from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest
from unittest.mock import Mock, patch

import pandas as pd
from openpyxl import load_workbook


RAIZ = Path(__file__).resolve().parents[1]


def carregar_modulo(nome: str, arquivo: str):
    spec = importlib.util.spec_from_file_location(nome, RAIZ / arquivo)
    modulo = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(modulo)
    return modulo


etapa1 = carregar_modulo("etapa1", "1_PegarDadosSiteIfc.py")
etapa2 = carregar_modulo("etapa2", "2_PdfParaJson.py")
etapa3 = carregar_modulo("etapa3", "3_ArrumarPlanilha.py")
etapa4 = carregar_modulo("etapa4", "4_GerarPainelHtml.py")


def resposta_simulada(registros: list[dict], linhas: list[dict]) -> dict:
    trabalhos = []
    linhas_por_id = {item["ID"]: item for item in linhas}
    for registro in registros:
        linha = linhas_por_id[registro["id"]]
        classificacoes = {}
        for chave, coluna in etapa3.CAMPOS_CLASSIFICACAO_GPT.items():
            atual = linha[coluna]
            classificacoes[chave] = {
                "valor_atual": atual,
                "valor": atual,
                "confianca": "alta" if atual != "A revisar" else "baixa",
                "evidencia": "Fixture: valor recebido foi preservado sem inferência.",
                "motivo": "Teste estrutural; não é classificação de produção.",
                "revisao_manual": atual == "A revisar",
            }
        trabalhos.append(
            {
                "id": registro["id"],
                "orientador_padronizado": linha["Orientador"],
                "classificacoes": classificacoes,
            }
        )
    return {
        "schema_version": etapa3.SCHEMA_RESPOSTA_GPT,
        "resumo": {
            "total_registros": len(trabalhos),
            "registros_revisao": sum(
                any(d["revisao_manual"] for d in item["classificacoes"].values())
                for item in trabalhos
            ),
        },
        "vocabularios_sugeridos": {},
        "trabalhos": trabalhos,
    }


class TestEtapa1(unittest.TestCase):
    def test_certificado_ifc_faz_fallback_automatico(self):
        sessao = Mock()
        resposta = Mock()
        sessao.get.side_effect = [
            etapa1.requests.exceptions.SSLError("certificado inválido"),
            resposta,
        ]
        args = SimpleNamespace(inseguro_tls=False)

        recebido = etapa1.requisicao_http(
            sessao,
            etapa1.URL_API_PROFEPT,
            args,
            timeout=90,
        )

        self.assertIs(resposta, recebido)
        self.assertTrue(sessao.get.call_args_list[0].kwargs["verify"])
        self.assertFalse(sessao.get.call_args_list[1].kwargs["verify"])

    def test_campus_explicito_do_educapes_tem_prioridade(self):
        educapes = {
            "campos_dc": {
                "dc.contributor": [
                    "Instituto Federal Catarinense - IFC Campus Camboriú"
                ]
            }
        }
        instituicao, campus, fonte = etapa1.resolver_instituicao_campus(educapes)
        self.assertEqual("Instituto Federal Catarinense", instituicao)
        self.assertEqual("Camboriú", campus)
        self.assertEqual("eduCAPES: dc.contributor", fonte)

    def test_campus_usa_contexto_quando_educapes_nao_informa(self):
        educapes = {
            "campos_dc": {"dc.contributor": ["Instituto Federal Catarinense"]}
        }
        _, campus, fonte = etapa1.resolver_instituicao_campus(educapes)
        self.assertEqual("Blumenau", campus)
        self.assertIn("eduCAPES sem campus explícito", fonte)

    def test_script_1_chama_ocr_em_sequencia(self):
        args = SimpleNamespace(limite=2)
        concluido = SimpleNamespace(returncode=0)
        with (
            patch.object(etapa1, "argumentos", return_value=args),
            patch.object(
                etapa1,
                "executar",
                return_value={"resumo": {"ids_duplicados": []}},
            ),
            patch.object(etapa1.subprocess, "run", return_value=concluido) as executar_ocr,
        ):
            self.assertEqual(0, etapa1.main())

        comando = executar_ocr.call_args.args[0]
        self.assertEqual(etapa1.sys.executable, comando[0])
        self.assertTrue(comando[1].endswith("2_PdfParaJson.py"))
        self.assertEqual(["--limite", "2"], comando[2:])

    def test_parser_aceita_multiplos_trabalhos(self):
        html = """
        <h3>DEFESAS 2026</h3>
        <ul><li>
        Acadêmico: Ana Exemplo<br>
        Orientadora: Dra. Maria Souza<br>
        Dissertação: Trabalho A <a href="https://example.org/a.pdf">PDF</a><br>
        Produto Educacional: Produto A <a href="https://example.org/a">Produto</a><br>
        Acadêmico: Bruno Exemplo<br>
        Orientador: Dr. João Lima<br>
        Dissertação: Trabalho B <a href="https://example.org/b.pdf">PDF</a><br>
        Produto Educacional: Produto B <a href="https://example.org/b">Produto</a>
        </li></ul>
        """
        trabalhos = etapa1.extrair_trabalhos_html(html)
        self.assertEqual(2, len(trabalhos))
        self.assertEqual("Ana Exemplo", trabalhos[0]["autor"])
        self.assertEqual("https://example.org/b.pdf", trabalhos[1]["url_pdf"])

    def test_planilha_e_pdfs_ativos_sao_coerentes(self):
        quadro = pd.read_excel(RAIZ / etapa1.ARQUIVO_SAIDA, dtype=str).fillna("")
        self.assertEqual(len(quadro), quadro["ID"].nunique())
        self.assertTrue(quadro["ID"].str.fullmatch(r"EPT-\d{4,}").all())
        self.assertEqual(len(quadro), quadro["Nome do PDF"].nunique())
        for _, linha in quadro.iterrows():
            pdf = RAIZ / etapa1.PASTA_PDFS / linha["Nome do PDF"]
            self.assertTrue(pdf.is_file(), linha["ID"])
            resumo = etapa1.sha256_arquivo(pdf)
            self.assertEqual(linha["SHA-256"], resumo, linha["ID"])


class TestEtapa2(unittest.TestCase):
    def test_json_para_gpt_tem_um_registro_por_id(self):
        documento = json.loads((RAIZ / etapa2.ARQUIVO_SAIDA).read_text(encoding="utf-8"))
        ids = [item["id"] for item in documento["trabalhos"]]
        self.assertEqual(documento["quantidade"], len(ids))
        self.assertEqual(len(ids), len(set(ids)))
        proibidos = {"sha256", "caminho_pdf", "texto_ocr_completo", "cache"}
        self.assertFalse(proibidos & set(documento["trabalhos"][0]))

    def test_prompt_exige_resposta_unica(self):
        texto = (RAIZ / etapa2.ARQUIVO_PROMPT).read_text(encoding="utf-8")
        self.assertIn("UM ÚNICO JSON", texto)
        self.assertIn("não divida em lotes", texto.casefold())
        self.assertIn("dados/GPT/gpt.json", texto)

    def test_publico_alvo_usa_classes_controladas(self):
        vocabulario = json.loads(
            (RAIZ / etapa2.ARQUIVO_VOCABULARIOS).read_text(encoding="utf-8")
        )
        self.assertEqual("classes_amplas_v2", vocabulario["publico_alvo"]["modo"])
        self.assertEqual(9, len(vocabulario["publico_alvo"]["ordem"]))

    def test_livro_digital_e_ebook_sao_um_unico_tipo(self):
        vocabulario = json.loads(
            (RAIZ / etapa2.ARQUIVO_VOCABULARIOS).read_text(encoding="utf-8")
        )
        tipos = vocabulario["categorias"]["Tipo de Produto Educacional"]
        self.assertEqual(etapa2.VERSAO_TIPOS_PRODUTO, tipos["modo"])
        indice = etapa2.indice_aliases(tipos["canonicos"])
        for alias in ["livro digital / oficina", "Ebook", "e-book", "E-book"]:
            self.assertEqual("E-book", indice[etapa2.normalizar_busca(alias)])
        self.assertNotIn("livro digital / oficina", tipos["canonicos"])


class TestEtapa3E4(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        _, cls.registros = etapa3.carregar_registros_entrada()
        cls.vocabulario = json.loads(
            (RAIZ / etapa3.ARQUIVO_VOCABULARIOS).read_text(encoding="utf-8")
        )
        cls.linhas = [
            etapa3.construir_linha(registro, cls.vocabulario)[0]
            for registro in sorted(cls.registros, key=lambda item: item["id"])
        ]
        cls.documento = resposta_simulada(
            sorted(cls.registros, key=lambda item: item["id"]), cls.linhas
        )

    def validar(self, documento: dict):
        with tempfile.TemporaryDirectory() as pasta:
            caminho = Path(pasta) / "gpt_fixture.json"
            caminho.write_text(json.dumps(documento, ensure_ascii=False), encoding="utf-8")
            return etapa3.validar_respostas_gpt(
                caminho,
                sorted(self.registros, key=lambda item: item["id"]),
                self.linhas,
                self.vocabulario,
            )

    def test_fixture_gpt_completa_e_valida(self):
        resultado = self.validar(self.documento)
        self.assertEqual(len(self.registros), len(resultado["sugestoes"]))

    def test_rejeita_id_duplicado(self):
        documento = json.loads(json.dumps(self.documento))
        documento["trabalhos"][1]["id"] = documento["trabalhos"][0]["id"]
        with self.assertRaisesRegex(ValueError, "duplicado"):
            self.validar(documento)

    def test_rejeita_id_inventado(self):
        documento = json.loads(json.dumps(self.documento))
        documento["trabalhos"][0]["id"] = "EPT-9999"
        with self.assertRaisesRegex(ValueError, "inexistente"):
            self.validar(documento)

    def test_rejeita_resposta_incompleta(self):
        documento = json.loads(json.dumps(self.documento))
        documento["trabalhos"].pop()
        documento["resumo"]["total_registros"] -= 1
        with self.assertRaisesRegex(ValueError, "incompleta"):
            self.validar(documento)

    def test_aplica_classificacao_nova_quando_alta_e_sem_revisao(self):
        documento = json.loads(json.dumps(self.documento))
        indice = next(
            i
            for i, item in enumerate(documento["trabalhos"])
            if item["classificacoes"]["tipo_produto"]["valor_atual"] == "A revisar"
        )
        permitidos = etapa3.valores_permitidos_gpt(self.linhas, self.vocabulario)
        novo = next(item for item in permitidos["tipo_produto"] if item != "A revisar")
        detalhe = documento["trabalhos"][indice]["classificacoes"]["tipo_produto"]
        detalhe.update(
            {
                "valor": novo,
                "confianca": "alta",
                "evidencia": "Fixture com categoria permitida e evidência suficiente.",
                "revisao_manual": False,
            }
        )
        validado = self.validar(documento)
        identificador = documento["trabalhos"][indice]["id"]
        linha = dict(next(item for item in self.linhas if item["ID"] == identificador))
        auditoria, _ = etapa3.aplicar_sugestoes_gpt(
            linha, validado["sugestoes"][identificador]
        )
        self.assertEqual(novo, linha["Tipo de Produto Educacional"])
        acao = next(item for item in auditoria if item["Campo"] == "Tipo de Produto Educacional")
        self.assertEqual("Aplicado automaticamente", acao["Ação"])

    def test_padroniza_nomes_titulos_e_termos_especiais(self):
        self.assertEqual(
            "Silvana de Oliveira Ribeiro da Silva",
            etapa3.padronizar_nome_pessoa("SILVANA DE OLIVEIRA RIBEIRO DA SILVA"),
        )
        self.assertEqual(
            "Educação na EPT: uma Experiência no IFC com ChatGPT e Libras",
            etapa3.padronizar_titulo(
                "EDUCAÇÃO NA EPT:UMA EXPERIÊNCIA NO IFC COM CHAT GPT E LIBRAS"
            ),
        )

    def test_normaliza_componentes_multivalor_sem_unir_conceitos(self):
        valor, erro = etapa3.normalizar_valor_aberto(
            "Base Epistemológica",
            "formação humana integral / EDUCAÇÃO CRÍTICA / Educação crítica",
            self.vocabulario,
        )
        self.assertIsNone(erro)
        self.assertEqual(
            "Educação Crítica / Formação Humana Integral",
            valor,
        )

    def test_detecta_exemplos_ambiguos_sem_corrigi_los(self):
        registro = {
            "autor": "C ÁTIA MARIA ALVES MONTEIRO",
            "titulo": "(RE)CONSTRUINDOOLHARES",
            "produto_educacional": "(RE)CONSTRUINDO OLHARES",
        }
        problemas = etapa3.detectar_problemas_textuais(
            registro,
            etapa3.padronizar_nome_pessoa(registro["autor"]),
            "Orientador Exemplo",
            "---",
            etapa3.padronizar_titulo(registro["titulo"]),
            etapa3.padronizar_titulo(registro["produto_educacional"]),
            {"construindo", "olhares"},
        )
        self.assertTrue(any("espaço suspeito" in item for item in problemas))
        self.assertTrue(any("possivelmente colada" in item for item in problemas))
        self.assertIn("construindoolhares", etapa3.normalizar_busca(registro["titulo"]))

    def test_painel_separa_filtros_multivalorados(self):
        self.assertEqual(
            ["Educação Crítica", "Formação Humana Integral"],
            etapa4.dividir_valores_filtro(
                "Educação Crítica / Formação Humana Integral / Educação Crítica"
            ),
        )
        self.assertEqual(
            ["Docentes", "Estudantes"],
            etapa4.dividir_valores_filtro(
                "Docentes; Estudantes; Docentes", separador="ponto_e_virgula"
            ),
        )

    def test_painel_prefere_https_no_educapes(self):
        self.assertEqual(
            "https://educapes.capes.gov.br/handle/capes/123",
            etapa4.normalizar_link_publico(
                "http://educapes.capes.gov.br/handle/capes/123"
            ),
        )
        self.assertEqual(
            "http://example.org/recurso",
            etapa4.normalizar_link_publico("http://example.org/recurso"),
        )

    def test_painel_nao_publica_finalidade_contaminada_por_ocr(self):
        self.assertEqual(
            "Informação em revisão.",
            etapa4.preparar_finalidade_publica(
                "DOCUMENTOS COMPROBATÓRIOS Nº 123/2026 - Nº do Protocolo: 456"
            ),
        )
        self.assertEqual(
            "Apoiar docentes na aplicação da proposta.",
            etapa4.preparar_finalidade_publica(
                "Apoiar docentes na aplicação da proposta."
            ),
        )

    def test_planilha_e_painel_com_fixture_sem_producao(self):
        with tempfile.TemporaryDirectory() as pasta:
            pasta = Path(pasta)
            resposta = pasta / "gpt_fixture.json"
            planilha = pasta / "final_fixture.xlsx"
            relatorio = pasta / "relatorio_fixture.txt"
            html = pasta / "painel_fixture.html"
            resposta.write_text(
                json.dumps(self.documento, ensure_ascii=False), encoding="utf-8"
            )
            with (
                patch.object(etapa3, "ARQUIVO_SAIDA", planilha),
                patch.object(etapa3, "ARQUIVO_RELATORIO", relatorio),
            ):
                resumo = etapa3.executar(resposta)
            self.assertEqual(len(self.registros), resumo["quantidade"])
            self.assertEqual(0, resumo["erros_estruturais"])
            with pd.ExcelFile(planilha) as arquivo_excel:
                self.assertEqual(
                    [
                        "Planilha1", "Alterações", "Validação", "Revisão Manual",
                        "Aliases Orientador", "Revisão GPT",
                    ],
                    arquivo_excel.sheet_names,
                )
            quadro = etapa4.carregar_planilha(planilha)
            registros_html = etapa4.preparar_dados_para_html(quadro)
            etapa4.salvar_arquivo_html(etapa4.gerar_html(registros_html), html)
            conteudo = html.read_text(encoding="utf-8")
            self.assertEqual(len(self.registros), len(registros_html))
            self.assertIn("leme-painel-ept", conteudo)
            self.assertNotIn("https://cdn", conteudo)
            self.assertIn('"baseItens":', conteudo)
            self.assertIn("camposMultivalor", conteudo)
            self.assertIn('data-campo="publicoAlvo"', conteudo)
            pasta_trabalho = load_workbook(planilha)
            try:
                aba = pasta_trabalho["Planilha1"]
                self.assertEqual("A2", aba.freeze_panes)
                self.assertEqual("4D5AFF", aba["A1"].fill.fgColor.rgb[-6:])
                self.assertTrue(aba["A1"].font.bold)
                self.assertTrue(aba.auto_filter.ref)
            finally:
                pasta_trabalho.close()


if __name__ == "__main__":
    unittest.main()
