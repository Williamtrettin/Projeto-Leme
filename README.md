# PAINEL EPT / LEME

Coleta as dissertações do ProfEPT/IFC, executa OCR nos PDFs, prepara a classificação pelo GPT e gera o painel HTML.

A explicação completa para a continuidade do projeto está em [GUIA.md](GUIA.md).
Os diagramas de arquitetura, dados e execução estão em [DIAGRAMAS.md](DIAGRAMAS.md).

## Instalação

É necessário ter Python 3.11+ e Tesseract OCR com o idioma português.

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

## Uso

Execute o primeiro script. Ele faz a coleta e chama o OCR automaticamente:

```bash
python 1_PegarDadosSiteIfc.py
```

Ao final, envie estes arquivos ao GPT:

```text
dados/GPT/PromptClassificacaoGPT.md
dados/GPT/para_gpt.json
```

Salve a resposta como `dados/GPT/gpt.json` e conclua:

```bash
python 3_ArrumarPlanilha.py
python 4_GerarPainelHtml.py
```

O resultado final fica em `painel/painel_leme_ept.html`.

## Opções úteis

```bash
# usar os dados e PDFs já armazenados localmente
python 1_PegarDadosSiteIfc.py --offline

# pular somente a consulta complementar ao eduCAPES
python 1_PegarDadosSiteIfc.py --sem-educapes

# normalmente o fallback é automático; esta opção força o modo compatível
python 1_PegarDadosSiteIfc.py --inseguro-tls

# testar coleta + OCR com poucos registros
python 1_PegarDadosSiteIfc.py --limite 2
```

## Testes

```bash
python -m py_compile util.py 1_PegarDadosSiteIfc.py 2_PdfParaJson.py 3_ArrumarPlanilha.py 4_GerarPainelHtml.py
python -m unittest discover -s tests -v
```
