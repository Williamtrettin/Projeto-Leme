# Guia de continuidade do PAINEL EPT / LEME

Estou deixando este documento para explicar o projeto a quem assumir sua continuidade. Minha intenção é que seja possível entender o sistema, preparar o computador, executar o fluxo completo e identificar problemas sem depender de explicações externas.

Os diagramas de arquitetura, dados, sequência e decisões estão em [DIAGRAMAS.md](DIAGRAMAS.md).

## 1. Objetivo do projeto

Este projeto coleta as produções acadêmicas do ProfEPT vinculadas ao Instituto Federal Catarinense, organiza os documentos, extrai informações dos PDFs por OCR, auxilia a classificação dos trabalhos com o GPT e gera um painel HTML para publicação no Elementor.

O sistema foi dividido em quatro etapas:

```text
Site oficial do ProfEPT/IFC
        ↓
1_PegarDadosSiteIfc.py
        ↓
Planilha da coleta + PDFs
        ↓
2_PdfParaJson.py (OCR)
        ↓
Prompt + JSON para o GPT
        ↓
Resposta do GPT salva manualmente
        ↓
3_ArrumarPlanilha.py
        ↓
Planilha final + relatório de validação
        ↓
4_GerarPainelHtml.py
        ↓
Painel HTML para o Elementor
```

O primeiro script já chama o segundo automaticamente. Na execução normal, portanto, não é necessário iniciar o OCR separadamente.

## 2. O que é automático e o que continua manual

O sistema automatiza:

- a consulta ao site oficial do ProfEPT/IFC;
- a coleta dos dados dos trabalhos;
- o download e a validação dos PDFs;
- a consulta complementar ao eduCAPES;
- o OCR dos PDFs;
- a preparação dos arquivos que serão enviados ao GPT;
- a validação da resposta do GPT;
- a geração da planilha final;
- a geração do painel HTML.

Continuam manuais:

- enviar o prompt e o JSON ao GPT;
- salvar a resposta recebida como `dados/GPT/gpt.json`;
- revisar as abas de validação e revisão manual da planilha final;
- copiar o HTML final para o Elementor;
- conferir o painel antes de publicar.

O projeto não possui chave de API da OpenAI e não envia dados automaticamente ao GPT.

## 3. Arquivos principais

### `1_PegarDadosSiteIfc.py`

Este é o ponto inicial do sistema. Ele:

- consulta a página oficial de dissertações do ProfEPT/IFC;
- extrai ano, título, autor, orientador, coorientador, produto educacional e links;
- relaciona cada trabalho a um ID permanente no formato `EPT-0001`;
- consulta dados complementares no eduCAPES;
- baixa ou valida os PDFs;
- calcula o SHA-256 de cada PDF;
- cria `dados/1_PegarDadosSiteIfc.xlsx`;
- chama automaticamente `2_PdfParaJson.py` para iniciar o OCR.

O ID não é baseado apenas na posição do trabalho no site. Ele é mantido no arquivo de identidades para que um trabalho não seja renumerado quando a página mudar.

### `2_PdfParaJson.py`

Este script executa o OCR e prepara o material para classificação. Ele:

- lê a planilha criada pela etapa 1;
- confirma o ID e o hash de cada PDF;
- converte páginas dos PDFs em imagens;
- usa o Tesseract para reconhecer o texto;
- procura resumo, palavras-chave, objetivos, metodologia, público-alvo e outras evidências;
- reutiliza o cache quando o PDF não mudou;
- combina o conteúdo encontrado com as classificações já preservadas;
- cria `dados/GPT/para_gpt.json`;
- cria `dados/GPT/PromptClassificacaoGPT.md`;
- cria `dados/relatorio_ocr.txt`.

Por padrão, o OCR processa até 15 páginas por PDF. Para trabalhos que já possuem classificação histórica, a leitura normal é reduzida para as páginas iniciais, evitando processamento desnecessário.

O texto integral do OCR não é enviado ao GPT. O JSON preparado contém apenas o contexto necessário para a classificação.

### `3_ArrumarPlanilha.py`

Este script recebe a resposta do GPT e produz a planilha final. Ele:

- lê `dados/GPT/gpt.json`;
- verifica se o JSON é válido;
- rejeita IDs inventados, duplicados ou ausentes;
- verifica se todos os trabalhos foram devolvidos;
- valida categorias e combinações de Linha de Pesquisa e Macroprojeto;
- preserva os dados oficiais coletados do site;
- aplica apenas sugestões consideradas seguras;
- encaminha casos incertos para revisão manual;
- cria `dados/dados_finais_painel_ept.xlsx`;
- cria `dados/relatorio_validacao.txt`.

A planilha final possui abas para os dados do painel, alterações, validações, revisão manual, aliases de orientadores e auditoria das sugestões do GPT.

O GPT não tem autorização para alterar título, autor, ano, links, instituição, campus, programa, PDF ou ID.

### `4_GerarPainelHtml.py`

Este script lê a planilha final e gera:

```text
painel/painel_leme_ept.html
```

Ele não recalcula classificações e não altera a planilha. Apenas transforma os dados revisados em um painel HTML com pesquisa, filtros, gráficos, cards, tabela, modal, paginação e exportação em CSV.

O HTML é autocontido. Depois de revisar o resultado, deve-se copiar todo o conteúdo do arquivo e colá-lo em um componente HTML do Elementor.

### `util.py`

Contém funções pequenas compartilhadas pelos scripts, como:

- limpeza e normalização de texto;
- criação de nomes seguros para arquivos;
- leitura e gravação de JSON;
- cálculo de SHA-256;
- preparação das conexões HTTP.

## 4. OCR

OCR significa Reconhecimento Óptico de Caracteres. Ele é necessário porque o sistema trata as páginas dos PDFs como imagens e precisa convertê-las em texto pesquisável.

Neste projeto, o OCR utiliza:

- PyMuPDF para abrir e renderizar os PDFs;
- Pillow para trabalhar com as imagens;
- pytesseract para conversar com o Tesseract;
- Tesseract para reconhecer efetivamente os caracteres.

O Tesseract deve estar instalado no computador, preferencialmente com o idioma português. Ele não é instalado pelo `requirements.txt`, pois é um programa do sistema operacional, e não uma biblioteca Python.

No Ubuntu ou Debian, normalmente a instalação é feita com:

```bash
sudo apt install tesseract-ocr tesseract-ocr-por
```

O cache do OCR fica em:

```text
dados/cache/ocr/
```

Cada cache é ligado ao ID, ao hash do PDF e à quantidade de páginas processadas. Se o PDF mudar, o hash muda e o OCR é executado novamente.

Para ignorar o cache e refazer o OCR:

```bash
python 2_PdfParaJson.py --forcar-ocr
```

Para processar todas as páginas:

```bash
python 2_PdfParaJson.py --max-paginas-ocr 0
```

Essa execução pode levar bastante tempo.

## 5. eduCAPES

O eduCAPES é utilizado como fonte complementar. A fonte principal continua sendo a página oficial do ProfEPT/IFC.

A consulta ao eduCAPES procura:

- o identificador permanente do registro, chamado `handle`;
- metadados Dublin Core;
- autores e títulos cadastrados no repositório;
- links dos arquivos disponíveis, chamados `bitstreams`.

Essas informações ajudam a conferir os dados e a localizar o produto educacional, mas não substituem silenciosamente as informações oficiais do ProfEPT.

Para o campus, uso somente uma indicação institucional explícita no campo `dc.contributor` do eduCAPES. Um campus citado apenas no título ou no resumo pode ser o local pesquisado e não deve ser confundido com a unidade associada ao programa. Quando o eduCAPES não informa o campus, o sistema registra essa ausência e mantém Blumenau como contexto do ProfEPT/IFC.

O cache fica em:

```text
dados/cache/educapes/
```

Se o eduCAPES estiver lento ou indisponível, é possível pular apenas essa consulta:

```bash
python 1_PegarDadosSiteIfc.py --sem-educapes
```

A coleta do site oficial e o processamento dos PDFs continuam funcionando.

## 6. Arquivos de referência

Os arquivos em `dados/referencia/` precisam permanecer no Git. Eles não são simples resultados temporários.

### `identidades.json`

Mantém a identidade permanente de cada trabalho. Entre os dados importantes estão:

- ID `EPT-xxxx`;
- título e autor usados no reconhecimento;
- URL do PDF;
- nome padronizado do arquivo;
- SHA-256 esperado;
- aliases conhecidos;
- situação de confirmação do vínculo.

Esse arquivo não deve ser apagado nem recriado do zero, pois isso pode renumerar trabalhos e quebrar a continuidade dos dados.

### `classificacoes_historicas.json`

Preserva a curadoria já realizada nos trabalhos anteriores. Ela funciona como ponto de partida e como memória do projeto.

Uma classificação histórica não é tratada como verdade absoluta: o GPT pode apontar uma possível mudança, mas alterações conceituais continuam sujeitas à revisão humana.

### `vocabularios.json`

Define os valores controlados e as equivalências conhecidas, incluindo:

- linhas de pesquisa;
- macroprojetos;
- relação permitida entre linha e macroprojeto;
- nomes canônicos e aliases de orientadores;
- tipos de produto educacional;
- níveis de aplicação;
- áreas temáticas;
- classes de público-alvo.

Esse arquivo evita que a mesma categoria apareça escrita de várias formas.

## 7. Relatórios

### `dados/relatorio_ocr.txt`

Mostra o resultado da leitura dos PDFs:

- quantidade processada;
- quantidade com sucesso;
- quantidade com erro;
- páginas que falharam;
- PDFs ausentes;
- divergências de hash.

### `dados/relatorio_validacao.txt`

Mostra o resultado da validação e da padronização:

- problemas estruturais;
- classificações que precisam de revisão;
- alterações aplicadas;
- avisos sobre a resposta do GPT;
- comparação das categorias antes e depois.

Esses relatórios são gerados novamente a cada execução e não precisam ser enviados ao Git.

## 8. O que é o `requirements.txt`

O `requirements.txt` é a lista das bibliotecas Python necessárias para executar o projeto. Em vez de instalar cada biblioteca manualmente, uso o seguinte comando:

```bash
python -m pip install -r requirements.txt
```

O arquivo possui estas dependências:

- `beautifulsoup4`: interpreta o HTML do ProfEPT e do eduCAPES;
- `openpyxl`: cria, lê e formata arquivos Excel `.xlsx`;
- `pandas`: organiza os dados em tabelas e faz a leitura e gravação das planilhas;
- `Pillow`: trabalha com as imagens produzidas a partir dos PDFs;
- `PyMuPDF`: abre os PDFs e renderiza suas páginas para o OCR;
- `pytesseract`: permite que o Python utilize o Tesseract;
- `requests`: realiza as consultas HTTP ao ProfEPT, aos PDFs e ao eduCAPES.

Um exemplo de linha é:

```text
requests>=2.34,<3
```

Isso significa que o projeto aceita a versão 2.34 ou superior da biblioteca, mas ainda não aceita a versão 3. Esse limite reduz o risco de uma atualização incompatível quebrar o sistema.

O `requirements.txt` instala somente bibliotecas Python. Ele não instala:

- o próprio Python;
- o Tesseract OCR;
- o pacote de idioma português do Tesseract;
- navegador ou editor de código;
- Elementor ou WordPress.

## 9. Preparação do computador

Recomendo Python 3.11 ou mais recente.

Na pasta do projeto:

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

No Windows, a ativação normalmente é:

```powershell
.venv\Scripts\activate
```

Também é necessário instalar o Tesseract e confirmar que ele está disponível no `PATH` do sistema.

A pasta `.venv/` é local e não deve ser enviada ao Git.

## 10. Execução normal completa

### Passo 1 — ativar o ambiente

```bash
source .venv/bin/activate
```

### Passo 2 — coletar e executar OCR

```bash
python 1_PegarDadosSiteIfc.py
```

Esse único comando executa as etapas 1 e 2 em sequência.

Ao final, devem existir:

```text
dados/1_PegarDadosSiteIfc.xlsx
dados/GPT/PromptClassificacaoGPT.md
dados/GPT/para_gpt.json
dados/relatorio_ocr.txt
pdf/
```

### Passo 3 — usar o GPT

Enviar juntos:

```text
dados/GPT/PromptClassificacaoGPT.md
dados/GPT/para_gpt.json
```

Solicitar que o GPT siga integralmente o prompt e devolva um único JSON válido, sem texto antes ou depois.

Salvar a resposta exatamente como:

```text
dados/GPT/gpt.json
```

É possível conferir a sintaxe com:

```bash
python -m json.tool dados/GPT/gpt.json > /dev/null
```

### Passo 4 — validar e gerar a planilha

```bash
python 3_ArrumarPlanilha.py
```

Antes de publicar qualquer coisa, abrir `dados/dados_finais_painel_ept.xlsx` e conferir principalmente as abas de validação, revisão manual e revisão GPT.

### Passo 5 — gerar o painel

```bash
python 4_GerarPainelHtml.py
```

O resultado será:

```text
painel/painel_leme_ept.html
```

### Passo 6 — publicar

Abrir o HTML, copiar todo o conteúdo e colar em um componente HTML do Elementor. Conferir pesquisa, filtros, gráficos, cards, tabela, links, paginação, exportação CSV e visualização no celular antes de publicar.

## 11. Opções úteis

Executar uma amostra pequena da coleta e do OCR:

```bash
python 1_PegarDadosSiteIfc.py --limite 2
```

Usar somente os arquivos e caches já existentes:

```bash
python 1_PegarDadosSiteIfc.py --offline
```

Não consultar o eduCAPES:

```bash
python 1_PegarDadosSiteIfc.py --sem-educapes
```

Não baixar PDFs novos:

```bash
python 1_PegarDadosSiteIfc.py --sem-download
```

Normalmente o sistema tenta a conexão segura e faz automaticamente uma segunda tentativa limitada ao domínio do IFC quando a validação do certificado falha. Para forçar esse modo desde o início:

```bash
python 1_PegarDadosSiteIfc.py --inseguro-tls
```

Essa última opção desativa a verificação TLS e só deve ser usada quando realmente necessária.

## 12. Testes

O arquivo `tests/test_fluxo.py` verifica se as partes principais continuam compatíveis. Os testes cobrem, entre outros pontos:

- extração de vários trabalhos da página;
- ligação automática entre coleta e OCR;
- IDs únicos;
- correspondência entre PDF e SHA-256;
- estrutura do JSON enviado ao GPT;
- rejeição de resposta incompleta, ID duplicado ou ID inventado;
- padronização de nomes e categorias;
- compatibilidade dos filtros do painel;
- proteção contra conteúdo de OCR inadequado no painel;
- geração de planilha e painel com uma resposta simulada.

Para compilar os scripts e encontrar erros de sintaxe:

```bash
python -m py_compile util.py 1_PegarDadosSiteIfc.py 2_PdfParaJson.py 3_ArrumarPlanilha.py 4_GerarPainelHtml.py
```

Para executar os testes:

```bash
python -m unittest discover -s tests -v
```

Os testes não substituem uma execução real da rede e do OCR. Eles protegem principalmente a estrutura, as validações e a integração entre as etapas.

## 13. O que vai para o Git

Devem permanecer versionados:

```text
1_PegarDadosSiteIfc.py
2_PdfParaJson.py
3_ArrumarPlanilha.py
4_GerarPainelHtml.py
util.py
requirements.txt
README.md
GUIA.md
DIAGRAMAS.md
dados/referencia/*.json
tests/test_fluxo.py
.gitignore
```

Não devem ser enviados ao Git:

- PDFs;
- planilhas geradas;
- respostas do GPT;
- cache do OCR e do eduCAPES;
- relatórios;
- painel HTML gerado;
- ambiente virtual `.venv/`;
- arquivos temporários e `__pycache__/`.

Esses itens já estão tratados no `.gitignore`.

## 14. Cuidados importantes

- Não apagar `dados/referencia/identidades.json`.
- Não renumerar os IDs `EPT-xxxx` manualmente.
- Não substituir os JSONs de referência por arquivos gerados durante a execução.
- Não confiar automaticamente em toda sugestão do GPT.
- Não publicar o painel antes de revisar as pendências.
- Não enviar os PDFs, caches e planilhas geradas ao Git.
- Não alterar a relação entre Linha de Pesquisa e Macroprojeto sem revisar o vocabulário e os testes.
- Se a instituição ou o programa mudarem, revisar as constantes nos scripts, os arquivos de referência e os testes. Para mudar a regra de identificação do campus, revisar também `resolver_instituicao_campus` no primeiro script.
- Se uma execução com `--limite` substituir os arquivos de saída por uma amostra, executar novamente sem o limite antes de enviar os dados ao GPT.

## 15. Problemas comuns

### `ModuleNotFoundError`

O ambiente virtual provavelmente não está ativo ou as dependências não foram instaladas:

```bash
source .venv/bin/activate
python -m pip install -r requirements.txt
```

### Tesseract não encontrado

Instalar o Tesseract no sistema e verificar se o executável está disponível no `PATH`.

### Idioma português ausente

Instalar o pacote de idioma português do Tesseract. O sistema pode usar inglês como alternativa, mas a qualidade tende a ser inferior para documentos em português.

### Erro de certificado ao acessar o IFC

O sistema tenta primeiro validar o certificado. Se o servidor do IFC apresentar a falha conhecida, ele mostra um aviso e repete automaticamente a conexão sem validação apenas para `profept.ifc.edu.br`. Se ainda for necessário forçar esse comportamento:

```bash
python 1_PegarDadosSiteIfc.py --inseguro-tls
```

### eduCAPES indisponível

Executar:

```bash
python 1_PegarDadosSiteIfc.py --sem-educapes
```

### Resposta do GPT rejeitada

Conferir se:

- o arquivo é um JSON válido;
- não existem blocos Markdown ou explicações fora do JSON;
- todos os IDs aparecem exatamente uma vez;
- não existem IDs inventados;
- todas as classificações obrigatórias foram preenchidas;
- o arquivo foi salvo como `dados/GPT/gpt.json`.

### Planilha ou HTML não pode ser salvo

Fechar o arquivo no Excel, navegador ou editor e executar novamente. Um programa aberto pode impedir a substituição do arquivo, especialmente no Windows.

## 16. Resumo para quem assumir

Para uma atualização normal, o fluxo é:

```bash
source .venv/bin/activate
python 1_PegarDadosSiteIfc.py
# enviar PromptClassificacaoGPT.md + para_gpt.json ao GPT
# salvar a resposta como dados/GPT/gpt.json
python 3_ArrumarPlanilha.py
python 4_GerarPainelHtml.py
python -m unittest discover -s tests -v
```

O ponto mais importante é manter os arquivos de referência, revisar a resposta do GPT e nunca publicar automaticamente uma classificação incerta. O sistema ajuda na coleta, leitura, padronização e apresentação, mas a decisão final de curadoria continua sendo humana.
