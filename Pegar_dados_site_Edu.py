import pandas as pd

from bs4 import BeautifulSoup

from playwright.sync_api import sync_playwright



# ================= CONFIGURAÇÃO DE CONTROLO =================

MODO_TESTE = False   # Deixa em False para rodar as 100 linhas

LIMITE_TESTE = 5     # Apenas usado se MODO_TESTE for True

# ============================================================



def identificar_base_epistemologica(texto):

    """Analisa o texto e retorna a base teórica provável."""

    texto = texto.lower()

   

    if any(x in texto for x in ['omnilateral', 'politecnia', 'formação integral', 'pleno desenvolvimento']):

        return "Formação Humana Integral / Omnilateralidade"

   

    elif any(x in texto for x in ['materialismo', 'histórico-dialético', 'práxis', 'totalidade', 'marx']):

        return "Materialismo Histórico-Dialético"

   

    elif any(x in texto for x in ['trabalho como princípio', 'centralidade do trabalho', 'ontologia do trabalho']):

        return "Trabalho como Princípio Educativo"

   

    elif any(x in texto for x in ['currículo integrado', 'integração curricular', 'interdisciplinaridade']):

        return "Currículo Integrado"

   

    elif any(x in texto for x in ['emancipação', 'freire', 'autonomia', 'conscientização', 'crítica']):

        return "Educação Crítica / Emancipatória"

   

    return "A definir (Análise Manual)"



def enriquecer_dados_educapes(arquivo_entrada='dados.xlsx', arquivo_saida='dados_educapes.xlsx'):

    print(f"Lendo a planilha base: {arquivo_entrada}")

   

    try:

        df = pd.read_excel(arquivo_entrada)

    except Exception as e:

        print(f"Erro ao ler Excel: {e}")

        return



    if MODO_TESTE:

        print(f"⚠️ MODO TESTE: Processando apenas {LIMITE_TESTE} linhas.")

        df = df.head(LIMITE_TESTE)



    # Criação das colunas necessárias para o Painel

    colunas = ['Palavras-chave', 'Tipo de Produto', 'Resumo', 'Nível de Aplicação (Auto)', 'Base Epistemológica (Auto)']

    for col in colunas:

        df[col] = df.get(col, '')



    with sync_playwright() as p:

        browser = p.chromium.launch(headless=True)

        context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

        page = context.new_page()



        for index, row in df.iterrows():

            link = str(row.get('Link do Produto', ''))

            if 'educapes' not in link.lower():

                continue



            print(f"[{index+1}/{len(df)}] Processando: {link}")

           

            try:

                # Acesso ao link com espera para carregamento de metadados

                page.goto(link, wait_until="domcontentloaded", timeout=90000)

                page.wait_for_timeout(5000)

               

                soup = BeautifulSoup(page.content(), "html.parser")

               

                # 1. EXTRAÇÃO VIA METADADOS (DSpace 5)

                resumo = ""

                keywords = ""

                tipo = ""

               

                # Busca Resumo

                meta_resumo = soup.find("meta", {"name": "DCTERMS.abstract"}) or soup.find("meta", {"name": "description"})

                if meta_resumo: resumo = meta_resumo.get("content", "")

               

                # Busca Palavras-chave

                meta_key = soup.find_all("meta", {"name": "citation_keywords"}) or soup.find_all("meta", {"name": "DC.subject"})

                keywords = "; ".join([m.get("content", "") for m in meta_key])



                # Busca Tipo de Produto

                meta_tipo = soup.find("meta", {"name": "DC.type"})

                if meta_tipo: tipo = meta_tipo.get("content", "")



                # 2. FALLBACK VIA TABELA (Caso metadados falhem)

                if not resumo:

                    linhas = soup.find_all("tr")

                    for l in linhas:

                        if "resumo" in l.get_text().lower():

                            celulas = l.find_all("td")

                            if len(celulas) >= 2: resumo = celulas[1].get_text(strip=True)



                # 3. CLASSIFICAÇÃO AUTOMÁTICA (BI / Dados)

                texto_total = (keywords + " " + resumo).lower()

               

                # Nível de Ensino

                nivel = "Outros"

                if any(x in texto_total for x in ['eja', 'jovens e adultos']): nivel = 'EJA'

                elif any(x in texto_total for x in ['integrado', 'médio', 'emi']): nivel = 'Ensino Médio Integrado'

                elif any(x in texto_total for x in ['superior', 'graduação']): nivel = 'Graduação'

               

                # Base Epistemológica

                base = identificar_base_epistemologica(texto_total)



                # Gravação no DataFrame

                df.at[index, 'Palavras-chave'] = keywords

                df.at[index, 'Tipo de Produto'] = tipo

                df.at[index, 'Resumo'] = resumo

                df.at[index, 'Nível de Aplicação (Auto)'] = nivel

                df.at[index, 'Base Epistemológica (Auto)'] = base



                status = "✅ Sucesso" if resumo else "⚠️ Sem Resumo"

                print(f"   {status} | Base: {base}")



            except Exception as e:

                print(f"   ❌ Erro: {e}")



        browser.close()



    # Salva o arquivo final

    nome_arquivo = "dados_finais_painel_ept.xlsx"

    df.to_excel(nome_arquivo, index=False)

    print(f"\n🚀 TUDO PRONTO! Planilha salva como: {nome_arquivo}")



if __name__ == "__main__":

    enriquecer_dados_educapes()