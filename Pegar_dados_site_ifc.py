import re
import pandas as pd
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

def extrair_links(texto):
    """Extrai todas as ocorrências de [LINK: url] de dentro do texto."""
    links = re.findall(r'\[LINK:\s*(.*?)\]', texto)
    # Remove os links do texto original para manter a string do título limpa
    texto_limpo = re.sub(r'\[LINK:\s*.*?\]', '', texto).strip()
    return texto_limpo, links

def atribuir_links(entrada, links, campo_origem):
    """Distribui os links encontrados na linha atual para as colunas corretas."""
    for link in links:
        link_lower = link.lower()
        if 'produto' in campo_origem.lower() and not entrada['Link do Produto']:
            entrada['Link do Produto'] = link
        elif 'disserta' in campo_origem.lower() and not entrada['Link do PDF']:
            entrada['Link do PDF'] = link
        elif '.pdf' in link_lower and not entrada['Link do PDF']:
            # Se for um PDF direto, priorizamos no link de PDF
            entrada['Link do PDF'] = link
        # Fallbacks: se o script não souber, preenche o primeiro buraco vazio
        elif not entrada['Link do PDF']:
            entrada['Link do PDF'] = link
        elif not entrada['Link do Produto']:
            entrada['Link do Produto'] = link

def extrair_dados_profept():
    url = "https://profept.ifc.edu.br/dissertacoes/"
    
    print("Iniciando navegador virtual (Playwright)...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        
        print("Acessando a página e aguardando o carregamento da rede...")
        # 'networkidle' resolve a espera dinâmica sem usar time.sleep()
        page.goto(url, wait_until="networkidle", timeout=60000)
        
        html_content = page.content()
        browser.close()
        
    print("Analisando e planificando a estrutura HTML...")
    soup = BeautifulSoup(html_content, "html.parser")
    
    # ESTRATÉGIA DE ANCORAGEM DE LINKS:
    # Como as tags HTML estão caóticas, substituímos cada hiperlink 
    # pelo seu texto seguido de um marcador especial.
    for a in soup.find_all("a"):
        href = a.get("href", "")
        if href:
            texto_link = a.get_text(strip=True)
            a.replace_with(f"{texto_link} [LINK: {href}]")
            
    # Extrai o texto da página inteira respeitando as quebras do BeautifulSoup
    texto_completo = soup.get_text(separator="\n")
    
    # FORÇAR QUEBRAS DE LINHA (Resolve o problema do texto "colado")
    padrao_rotulos = r'(Acadêmic[oa]s?:|Alunos?:|Autor(?:es)?s?:|Mestrandos?:|Orientador[a]?s?:|Dissertaç[ãa]o\s*:|Produto Educacional\s*:)'
    # Injeta um "\n" exatamente antes de qualquer rótulo para separá-los
    texto_separado = re.sub(padrao_rotulos, r'\n\1', texto_completo, flags=re.IGNORECASE)
    
    # -----------------------------------------------
    # MÁQUINA DE ESTADOS (Processamento Linha a Linha)
    # -----------------------------------------------
    linhas = texto_separado.split('\n')
    
    entradas = []
    entrada_atual = None
    ano_atual = "Desconhecido"
    campo_atual = None
    
    print("Mapeando e extraindo os campos de cada dissertação...")
    for linha in linhas:
        linha = linha.strip()
        if not linha:
            continue
            
        # 1. Detectar Ano (Ex: "DEFESAS 2024")
        # Garante que é um ano se tiver 4 dígitos, texto curto e sem nomes de campos.
        if re.search(r'20[1-9]\d', linha) and len(linha) < 40 and not re.search(r'(acadêmic|dissertaç|produto|orientador)', linha, re.IGNORECASE):
            ano_atual = linha
            campo_atual = None
            continue
            
        # Limpar tracejados e bullet points indesejados
        linha = re.sub(r'^[-•*]\s*', '', linha)
        
        # 2. Identificar novo Acadêmico (Isso "fecha" o registro anterior e abre um novo)
        if re.match(r'^(Acadêmic[oa]s?|Alunos?|Autor(?:es)?|Mestrandos?)\s*:', linha, re.IGNORECASE):
            if entrada_atual:
                entradas.append(entrada_atual)
            
            # Molde do Dicionário para evitar o erro de KeyError/NoneType
            entrada_atual = {
                'Ano': ano_atual,
                'Autor': '',
                'Orientador': '',
                'Título da Dissertação': '',
                'Produto Educacional': '',
                'Link do PDF': '',
                'Link do Produto': ''
            }
            
            valor = re.sub(r'^(Acadêmic[oa]s?|Alunos?|Autor(?:es)?|Mestrandos?)\s*:\s*', '', linha, flags=re.IGNORECASE)
            valor, links = extrair_links(valor)
            entrada_atual['Autor'] = valor
            campo_atual = 'Autor'
            atribuir_links(entrada_atual, links, campo_atual)
            continue
            
        # Se encontrou texto inútil antes do primeiro aluno, ignore
        if not entrada_atual:
            continue
            
        # 3. Orientador
        if re.match(r'^Orientador[a]?s?\s*:', linha, re.IGNORECASE):
            valor = re.sub(r'^Orientador[a]?s?\s*:\s*', '', linha, flags=re.IGNORECASE)
            valor, links = extrair_links(valor)
            entrada_atual['Orientador'] = valor
            campo_atual = 'Orientador'
            atribuir_links(entrada_atual, links, campo_atual)
            continue
            
        # 4. Dissertação
        if re.match(r'^Dissertaç[ãa]o\s*:', linha, re.IGNORECASE):
            valor = re.sub(r'^Dissertaç[ãa]o\s*:\s*', '', linha, flags=re.IGNORECASE)
            valor, links = extrair_links(valor)
            entrada_atual['Título da Dissertação'] = valor
            campo_atual = 'Título da Dissertação'
            atribuir_links(entrada_atual, links, campo_atual)
            continue
            
        # 5. Produto Educacional
        if re.match(r'^Produto Educacional\s*:', linha, re.IGNORECASE):
            valor = re.sub(r'^Produto Educacional\s*:\s*', '', linha, flags=re.IGNORECASE)
            valor, links = extrair_links(valor)
            entrada_atual['Produto Educacional'] = valor
            campo_atual = 'Produto Educacional'
            atribuir_links(entrada_atual, links, campo_atual)
            continue
            
        # 6. Linha "órfã" (Continuação de um título grande que quebrou linha ou Link isolado)
        valor, links = extrair_links(linha)
        if links:
            atribuir_links(entrada_atual, links, campo_atual or "")
            
        if valor and campo_atual:
            # Evita poluir o título se o texto for apenas uma âncora genérica
            if valor.lower() not in ['clique aqui', 'link', 'download', 'pdf', 'acessar']:
                # Concatena a continuação do texto
                entrada_atual[campo_atual] += ' ' + valor

    # Fechar o último registro mapeado
    if entrada_atual:
        entradas.append(entrada_atual)
        
    # -----------------------------------------------
    # PANDAS: EXPORTAÇÃO E TRATAMENTO DE DUPLICIDADES
    # -----------------------------------------------
    print(f"Extração concluída. {len(entradas)} registros encontrados. Limpando dados...")
    df = pd.DataFrame(entradas)
    
    # Remove espaços em branco excessivos concatenados acidentalmente
    df = df.apply(lambda col: col.str.strip() if col.dtype == 'object' else col)
    
    # Remove duplicatas exatas caso a página mostre o mesmo registro na home e em uma listagem
    df.drop_duplicates(subset=['Autor', 'Título da Dissertação'], inplace=True)
    
    arquivo_saida = 'dados.xlsx'
    df.to_excel(arquivo_saida, index=False)
    print(f"Sucesso! Tabela salva em: {arquivo_saida}")
    
    return df

if __name__ == "__main__":
    df_resultados = extrair_dados_profept()