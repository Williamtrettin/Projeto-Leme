# Diagramas do PAINEL EPT / LEME

Este documento complementa o guia de continuidade com uma visão visual do sistema.

O código atual é procedural: ele está organizado em scripts e funções, não em classes Python. Por isso, o diagrama de classes abaixo é conceitual. Ele representa as responsabilidades do sistema e ajuda a entender como uma futura reorganização em classes poderia ser feita.

## 1. Visão geral do fluxo

```mermaid
flowchart TD
    SITE["Site oficial ProfEPT/IFC"]
    EDU["eduCAPES"]
    E1["1_PegarDadosSiteIfc.py<br/>Coleta e PDFs"]
    REF["dados/referencia/<br/>identidades, classificações e vocabulários"]
    XLS1["dados/1_PegarDadosSiteIfc.xlsx"]
    PDF["pdf/*.pdf"]
    E2["2_PdfParaJson.py<br/>OCR e preparação do GPT"]
    CACHE["dados/cache/<br/>eduCAPES e OCR"]
    PROMPT["PromptClassificacaoGPT.md"]
    JSONIN["para_gpt.json"]
    GPT["GPT<br/>etapa manual"]
    JSONOUT["gpt.json"]
    E3["3_ArrumarPlanilha.py<br/>Validação e padronização"]
    XLS2["dados_finais_painel_ept.xlsx"]
    REL["Relatórios de OCR e validação"]
    E4["4_GerarPainelHtml.py"]
    HTML["painel_leme_ept.html"]
    WP["Elementor / WordPress<br/>publicação manual"]

    SITE --> E1
    EDU --> E1
    REF --> E1
    E1 --> XLS1
    E1 --> PDF
    E1 --> E2
    XLS1 --> E2
    PDF --> E2
    REF --> E2
    CACHE <--> E1
    CACHE <--> E2
    E2 --> PROMPT
    E2 --> JSONIN
    E2 --> REL
    PROMPT --> GPT
    JSONIN --> GPT
    GPT --> JSONOUT
    JSONOUT --> E3
    XLS1 --> E3
    REF --> E3
    E3 --> XLS2
    E3 --> REL
    XLS2 --> E4
    E4 --> HTML
    HTML --> WP
```

## 2. Diagrama de classes conceitual

```mermaid
classDiagram
    class ColetorProfEPT {
        +URL paginaOficial
        +URL apiOficial
        +obterPagina()
        +extrairTrabalhos()
        +resolverIdentidades()
        +salvarPlanilha()
    }

    class ClienteEduCAPES {
        +consultar(handle)
        +extrairMetadadosDC()
        +listarBitstreams()
        +resolverInstituicaoCampus()
    }

    class ClienteHTTP {
        +criarSessao()
        +requisicaoSegura()
        +fallbackTLSdoIFC()
    }

    class GerenciadorIdentidades {
        +carregarReferencias()
        +encontrarPorURL()
        +encontrarPorAutorTituloAno()
        +criarProximoID()
        +validarHash()
    }

    class GerenciadorPDF {
        +baixarPDF()
        +validarAssinaturaPDF()
        +calcularSHA256()
        +nomearPorID()
    }

    class ProcessadorOCR {
        +abrirPDF()
        +renderizarPaginas()
        +reconhecerTexto()
        +usarCache()
        +extrairCampos()
    }

    class PreparadorGPT {
        +aplicarVocabularios()
        +montarContextoGlobal()
        +gerarJSONEntrada()
        +gerarPrompt()
    }

    class ValidadorGPT {
        +validarJSON()
        +validarIDs()
        +validarCategorias()
        +aplicarSugestoesSeguras()
        +encaminharRevisaoManual()
    }

    class GeradorPlanilha {
        +gerarAbaPrincipal()
        +gerarAlteracoes()
        +gerarValidacoes()
        +gerarRevisaoManual()
        +gerarAuditoriaGPT()
    }

    class GeradorPainel {
        +lerPlanilhaFinal()
        +prepararDados()
        +gerarHTML()
        +salvarPainel()
    }

    class RepositorioReferencias {
        +identidadesJson
        +classificacoesHistoricasJson
        +vocabulariosJson
    }

    class RepositorioCache {
        +cacheEduCAPES
        +cacheOCR
    }

    ColetorProfEPT --> ClienteHTTP : usa
    ColetorProfEPT --> ClienteEduCAPES : consulta
    ColetorProfEPT --> GerenciadorIdentidades : resolve IDs
    ColetorProfEPT --> GerenciadorPDF : baixa e valida
    ClienteEduCAPES --> ClienteHTTP : usa
    GerenciadorIdentidades --> RepositorioReferencias : lê e atualiza
    ClienteEduCAPES --> RepositorioCache : reutiliza HTML
    ProcessadorOCR --> GerenciadorPDF : lê PDFs validados
    ProcessadorOCR --> RepositorioCache : reutiliza OCR
    PreparadorGPT --> ProcessadorOCR : recebe evidências
    PreparadorGPT --> RepositorioReferencias : usa classificações
    ValidadorGPT --> RepositorioReferencias : valida vocabulários
    ValidadorGPT --> GeradorPlanilha : entrega dados aprovados
    GeradorPlanilha --> GeradorPainel : fornece planilha final
```

## 3. Modelo conceitual dos dados

```mermaid
classDiagram
    class Trabalho {
        +string id
        +string ano
        +string titulo
        +string autor
        +string orientador
        +string coorientador
        +string produtoEducacional
        +string instituicao
        +string campus
        +string programa
        +string urlPDF
        +string urlProduto
    }

    class Identidade {
        +string idPermanente
        +string urlPDFCanonica
        +string nomePDF
        +string sha256
        +list aliasesTitulos
        +list aliasesAutores
        +boolean vinculoConfirmado
    }

    class DocumentoPDF {
        +string caminho
        +string sha256
        +string status
        +int totalPaginas
    }

    class MetadadosEduCAPES {
        +string handle
        +map camposDublinCore
        +list bitstreams
        +string campusExplicito
        +string origemConsulta
    }

    class ConteudoOCR {
        +string resumo
        +list palavrasChave
        +string objetivos
        +string metodologia
        +string descricaoProduto
        +string publicoAlvo
        +list errosPaginas
    }

    class Classificacao {
        +string linhaPesquisa
        +string macroprojeto
        +string baseEpistemologica
        +string nivelAplicacao
        +string tipoProduto
        +string areaTematica
        +string publicoAlvo
        +string confianca
        +boolean revisaoManual
    }

    class Vocabulario {
        +map valoresCanonicos
        +map aliases
        +map relacaoLinhaMacro
        +list orientadoresCanonicos
    }

    Trabalho "1" --> "1" Identidade : possui
    Trabalho "1" --> "1" DocumentoPDF : referencia
    Trabalho "1" --> "0..1" MetadadosEduCAPES : complementado por
    DocumentoPDF "1" --> "0..1" ConteudoOCR : produz
    Trabalho "1" --> "1" Classificacao : recebe
    Vocabulario "1" --> "0..*" Classificacao : restringe
```

## 4. Diagrama de sequência da execução normal

```mermaid
sequenceDiagram
    actor Pessoa as Responsável pelo projeto
    participant E1 as Python 1
    participant IFC as ProfEPT/IFC
    participant EDU as eduCAPES
    participant REF as Referências
    participant E2 as Python 2 / OCR
    participant GPT as GPT
    participant E3 as Python 3
    participant E4 as Python 4
    participant WP as Elementor

    Pessoa->>E1: python 1_PegarDadosSiteIfc.py
    E1->>IFC: Consulta API com TLS
    alt Certificado válido
        IFC-->>E1: Página oficial
    else Certificado do IFC rejeitado
        E1->>IFC: Repete somente para o IFC sem validar certificado
        IFC-->>E1: Página oficial
    end
    E1->>REF: Carrega identidades permanentes
    loop Para cada trabalho
        E1->>EDU: Consulta metadados e campus explícito
        EDU-->>E1: Dublin Core e bitstreams
        E1->>IFC: Baixa PDF ausente
        IFC-->>E1: PDF
        E1->>REF: Confere ID e SHA-256
    end
    E1->>E2: Inicia OCR automaticamente
    loop Para cada PDF validado
        E2->>E2: Renderiza páginas e executa Tesseract
        E2->>E2: Extrai evidências e atualiza cache
    end
    E2-->>Pessoa: PromptClassificacaoGPT.md + para_gpt.json
    Pessoa->>GPT: Envia os dois arquivos
    GPT-->>Pessoa: JSON único de resposta
    Pessoa->>E3: Salva gpt.json e executa Python 3
    E3->>E3: Valida IDs, categorias e confiança
    E3-->>Pessoa: Planilha final + relatório
    Pessoa->>Pessoa: Faz revisão humana
    Pessoa->>E4: python 4_GerarPainelHtml.py
    E4-->>Pessoa: painel_leme_ept.html
    Pessoa->>WP: Cola o HTML e publica
```

## 5. Decisão da instituição e do campus

```mermaid
flowchart TD
    A["Registro do trabalho"] --> B["Consultar eduCAPES"]
    B --> C{"dc.contributor informa<br/>campus explicitamente?"}
    C -- Sim --> D["Usar o campus informado<br/>pelo eduCAPES"]
    C -- Não --> E["Usar Blumenau como contexto<br/>do ProfEPT/IFC"]
    D --> F["Registrar Fonte do Campus:<br/>eduCAPES: dc.contributor"]
    E --> G["Registrar Fonte do Campus:<br/>eduCAPES sem campus explícito"]
    F --> H["Salvar na planilha da etapa 1"]
    G --> H

    I["Campus citado apenas no<br/>título ou resumo"] --> J["Não usar automaticamente"]
    J --> K["Pode ser somente o local pesquisado"]
```

## 6. Decisão de conexão TLS e download

```mermaid
flowchart TD
    A["Iniciar consulta HTTPS"] --> B["Validar certificado normalmente"]
    B --> C{"Conexão funcionou?"}
    C -- Sim --> D["Continuar download seguro"]
    C -- Não --> E{"Erro é de certificado e o host<br/>é profept.ifc.edu.br?"}
    E -- Não --> F["Interromper e mostrar o erro real"]
    E -- Sim --> G["Mostrar aviso"]
    G --> H["Repetir sem validar certificado<br/>somente para o domínio do IFC"]
    H --> I{"Segunda tentativa funcionou?"}
    I -- Sim --> J["Continuar coleta e download"]
    I -- Não --> F
```

## 7. Ciclo de vida dos arquivos

```mermaid
flowchart LR
    subgraph Git["Arquivos versionados no Git"]
        COD["Scripts Python"]
        REQ["requirements.txt"]
        DOC["Documentação"]
        TEST["Testes"]
        RJSON["JSONs de referência"]
    end

    subgraph Local["Arquivos locais gerados"]
        PDFs["pdf/"]
        PLAN["Planilhas"]
        GPTF["Arquivos de troca com GPT"]
        CACHES["Caches"]
        REPORTS["Relatórios"]
        PANEL["Painel HTML"]
    end

    COD --> Local
    REQ --> COD
    RJSON --> Local
    TEST --> COD
    Local -.->|ignorado por .gitignore| Git
```

## 8. Dependências técnicas

```mermaid
flowchart TD
    PY["Python 3.11+"] --> REQ["requirements.txt"]
    REQ --> BS["BeautifulSoup<br/>interpretação de HTML"]
    REQ --> REQUESTS["requests<br/>HTTP"]
    REQ --> PANDAS["pandas<br/>tabelas"]
    REQ --> OPENPYXL["openpyxl<br/>Excel"]
    REQ --> PYMUPDF["PyMuPDF<br/>leitura e imagem do PDF"]
    REQ --> PILLOW["Pillow<br/>imagens"]
    REQ --> PYTESS["pytesseract<br/>integração Python"]
    TESS["Tesseract instalado no sistema"] --> PYTESS
    LANG["Idioma português do Tesseract"] --> TESS
```

O `requirements.txt` instala as bibliotecas Python. O Tesseract e seu pacote de idioma são dependências externas e precisam ser instalados no sistema operacional.
