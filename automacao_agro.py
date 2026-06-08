import os
import glob
import requests
import logging
import unicodedata
import re
from datetime import datetime
from urllib.parse import urljoin
from bs4 import BeautifulSoup
import pandas as pd

# ==========================================
# 1. CONFIGURAÇÃO DE CAMINHOS E LOGS
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s - [ %(levelname)s ] - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(os.path.join(BASE_DIR, 'automacao_agro.log'), mode='a', encoding='utf-8')
    ]
)

# ==========================================
# 2. MÓDULO DE EXTRAÇÃO (WEB SCRAPING)
# ==========================================
class AgroScraper:
    def __init__(self, url: str, dir_downloads: str):
        self.url = url
        self.dir_downloads = dir_downloads
        os.makedirs(self.dir_downloads, exist_ok=True)

    def extrair_planilhas(self) -> bool:
        logging.info("--- INICIANDO WEB SCRAPING ---")
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        
        try:
            response = requests.get(self.url, headers=headers, timeout=20)
            response.raise_for_status()
        except Exception as e:
            logging.error(f"Falha de conexão com o site: {e}")
            return False

        soup = BeautifulSoup(response.content, 'html.parser')
        links = soup.find_all('a', href=True)
        planilhas_baixadas = 0

        for arquivo_antigo in glob.glob(os.path.join(self.dir_downloads, "*.xls*")):
            os.remove(arquivo_antigo)

        for link in links:
            href = link['href']
            if href.endswith('.xlsx') or href.endswith('.xls'):
                url_completa = urljoin(self.url, href) 
                nome_arquivo = url_completa.split('/')[-1].split('?')[0].lower()
                
                # Foca EXCLUSIVAMENTE nos dados do VBP Completo / VBP Brasil
                if 'completo' not in nome_arquivo and 'vbpbrasil' not in nome_arquivo:
                    continue
                
                # Filtra apenas os arquivos do ano atual em diante (ignora anos anteriores)
                ano_atual = datetime.now().year
                match_ano = re.search(r'\d{4}', nome_arquivo)
                if match_ano:
                    ano_arquivo = int(match_ano.group())
                    if ano_arquivo < ano_atual:
                        continue

                caminho_salvar = os.path.join(self.dir_downloads, nome_arquivo)
                
                logging.info(f"Baixando: {nome_arquivo}...")
                try:
                    file_response = requests.get(url_completa, headers=headers, stream=True, timeout=30)
                    file_response.raise_for_status()
                    with open(caminho_salvar, 'wb') as arquivo_excel:
                        for chunk in file_response.iter_content(chunk_size=8192):
                            arquivo_excel.write(chunk)
                    planilhas_baixadas += 1
                except Exception as e:
                    logging.error(f"Erro ao baixar {nome_arquivo}: {e}")

        return planilhas_baixadas > 0

    def padronizar_nomes_arquivos(self):
        arquivos = glob.glob(os.path.join(self.dir_downloads, "*.xls*"))
        for arquivo in arquivos:
            nome_base = os.path.basename(arquivo).lower()
            if 'vbp' in nome_base or 'valor' in nome_base:
                novo_nome = os.path.join(self.dir_downloads, f"tratado_vbp_{nome_base}")
            elif 'producao' in nome_base or 'area' in nome_base or 'quantidade' in nome_base:
                novo_nome = os.path.join(self.dir_downloads, f"tratado_volume_{nome_base}")
            else:
                novo_nome = os.path.join(self.dir_downloads, f"tratado_vbp_{nome_base}")
                
            if arquivo != novo_nome:
                os.rename(arquivo, novo_nome)

# ==========================================
# 3. MÓDULO DE TRATAMENTO DE DADOS E HTML (ETL)
# ==========================================
class AgroETL:
    def __init__(self, dir_downloads: str, dir_consolidados: str, dir_relatorios: str):
        self.dir_downloads = dir_downloads
        self.dir_consolidados = dir_consolidados
        self.dir_relatorios = dir_relatorios
        for diretorio in [self.dir_downloads, self.dir_consolidados, self.dir_relatorios]:
            os.makedirs(diretorio, exist_ok=True)

    def _limpar_colunas(self, df: pd.DataFrame) -> pd.DataFrame:
        if df.empty: return df
        colunas_limpas = []
        for col in df.columns:
            col_str = str(col).strip().lower()
            col_str = re.sub(r'\.0$', '', col_str)
            col_str = ''.join(c for c in unicodedata.normalize('NFD', col_str) if unicodedata.category(c) != 'Mn')
            col_str = col_str.replace(' ', '_').replace('.', '').replace('\n', '')
            colunas_limpas.append(col_str)
        df.columns = colunas_limpas
        return df

    def extrair_e_empilhar(self, padrao_nome: str) -> pd.DataFrame:
        arquivos = sorted(glob.glob(os.path.join(self.dir_downloads, padrao_nome)))
        lista_dfs = []
        if not arquivos:
            logging.warning(f"Nenhum arquivo encontrado com o padrão: {padrao_nome}")
            
        for arquivo in arquivos:
            nome_arquivo = os.path.basename(arquivo)
            
            # Extrai ano e mês do nome do arquivo para formatar a versão (ex: 2026 - 01)
            match_data = re.search(r'(\d{4})(\d{2})', nome_arquivo)
            if match_data:
                versao_formatada = f"{match_data.group(1)} - {match_data.group(2)}"
            else:
                versao_formatada = nome_arquivo
                
            logging.info(f"Lendo e empilhando planilha: {nome_arquivo} (Versão: {versao_formatada})...")
            try:
                # Lê todas as abas ignorando o cabeçalho automático (header=None) para não perder o topo
                dicionario_abas = pd.read_excel(arquivo, sheet_name=None, header=None)
                df_maior = pd.DataFrame()
                
                for nome_aba, df_aba in dicionario_abas.items():
                    if len(df_aba) > len(df_maior):
                        df_maior = df_aba
                        
                if not df_maior.empty:
                    # Identifica a linha do cabeçalho verdadeiro (ex: que começa com os anos 1989, 1990...)
                    idx_cabecalho = 0
                    for idx, row in df_maior.iterrows():
                        valores_linha = row.astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                        if '1989' in valores_linha.values or valores_linha.str.match(r'^(19|20)\d{2}$').sum() > 3:
                            idx_cabecalho = idx
                            break
                            
                    # Define o cabeçalho real e descarta o texto solto das primeiras linhas
                    df_maior.columns = df_maior.iloc[idx_cabecalho]
                    df_maior.columns.name = None
                    df_maior = df_maior.iloc[idx_cabecalho + 1:].reset_index(drop=True)

                    # Remove linhas 100% vazias para organizar a base
                    df_maior = df_maior.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)

                    # A Tabela 2 (Preços Correntes) empilhada repete o 1º item da Tabela 1 (ex: "LAVOURAS").
                    # Cortamos o arquivo exatamente onde essa repetição acontece.
                    if not df_maior.empty:
                        col_zero = df_maior.columns[0]
                        valores_validos = df_maior[col_zero].dropna()
                        if not valores_validos.empty:
                            primeiro_item = valores_validos.iloc[0]
                            repeticoes = df_maior[df_maior[col_zero] == primeiro_item].index
                            if len(repeticoes) > 1:
                                df_maior = df_maior.loc[:repeticoes[1]-1]

                    df_limpo = self._limpar_colunas(df_maior)
                    
                    if not df_limpo.empty and len(df_limpo.columns) > 0:
                        primeira_col = df_limpo.columns[0]
                        
                        # Corta a tabela descartando o rodapé ao encontrar palavras-chave
                        termos_rodape = 'fonte|nota|elaboração|elaboracao|atualizado'
                        mascara_rodape = df_limpo[primeira_col].astype(str).str.lower().str.contains(termos_rodape, na=False)
                        if mascara_rodape.any():
                            idx_rodape = mascara_rodape.idxmax()
                            df_limpo = df_limpo.loc[:idx_rodape].iloc[:-1]

                        # Remove duplicatas normalizando o texto (sem acentos ou espaços soltos)
                        df_limpo['chave_temp'] = df_limpo[primeira_col].astype(str).str.strip().str.lower()
                        df_limpo['chave_temp'] = df_limpo['chave_temp'].apply(
                            lambda x: ''.join(c for c in unicodedata.normalize('NFD', x) if unicodedata.category(c) != 'Mn')
                        )
                        df_limpo = df_limpo.drop_duplicates(subset=['chave_temp'], keep='first').drop(columns=['chave_temp'])
                        
                        # Adiciona uma coluna na primeira posição para identificar a versão/origem do dado
                        df_limpo.insert(0, 'versao_arquivo', versao_formatada)
                        
                    lista_dfs.append(df_limpo)
            except Exception as e:
                logging.error(f"Erro ao ler planilha {arquivo}: {e}")
        return pd.concat(lista_dfs, ignore_index=True) if lista_dfs else pd.DataFrame()

    def cruzar_e_salvar_versao(self) -> bool:
        logging.info("--- INICIANDO PROCESSAMENTO DE DADOS ---")
        df_volume = self.extrair_e_empilhar("*_volume_*.xls*")
        df_valor = self.extrair_e_empilhar("*_vbp_*.xls*")
        
        if df_valor.empty and df_volume.empty:
            logging.error("ERRO: As tabelas estão vazias. Nenhum arquivo consolidado será gerado.")
            return False

        chaves = ['estado', 'cultura']
        if not df_valor.empty and not df_volume.empty:
            if all(ch in df_volume.columns for ch in chaves) and all(ch in df_valor.columns for ch in chaves):
                df_consolidado = pd.merge(df_volume, df_valor, on=chaves, how='inner')
            else:
                df_consolidado = pd.concat([df_volume, df_valor], ignore_index=True)
        else:
            df_consolidado = df_valor if not df_valor.empty else df_volume

        # Preenche valores vazios de forma segura (evita erro com texto no Pandas)
        for col in df_consolidado.columns:
            try:
                df_consolidado[col] = df_consolidado[col].fillna(0)
            except TypeError:
                df_consolidado[col] = df_consolidado[col].fillna("")
                
        nome_arquivo = f"base_consolidada_{datetime.now().strftime('%m-%Y')}.xlsx"
        caminho_saida = os.path.join(self.dir_consolidados, nome_arquivo)
        df_consolidado.to_excel(caminho_saida, index=False)
        logging.info(f"Consolidação salva: {nome_arquivo}")
        return True

    def gerar_relatorio_html(self):
        logging.info("--- GERANDO DASHBOARD VISUAL (HTML) ---")
        arquivos = sorted(glob.glob(os.path.join(self.dir_consolidados, "base_consolidada_*.xlsx")), key=os.path.getmtime)
        if not arquivos:
            logging.warning("Nenhum arquivo consolidado encontrado.")
            return False
        
        caminho_base = arquivos[-1]
        df = pd.read_excel(caminho_base)

        if 'versao_arquivo' not in df.columns:
            logging.error("A coluna 'versao_arquivo' não existe. Rode a consolidação novamente.")
            return False

        # Padroniza as colunas de anos para string, remove '.0' e asteriscos (*) que o governo usa para projeções
        novas_colunas = []
        for i, c in enumerate(df.columns):
            if i > 0: # Ignora a coluna 'versao_arquivo'
                novas_colunas.append(re.sub(r'\.0$', '', str(c)).replace('*', '').strip())
            else:
                novas_colunas.append(c)
        df.columns = novas_colunas

        versoes = sorted(df['versao_arquivo'].dropna().unique())
        if not versoes:
            return False
            
        versao_atual = versoes[-1]
        
        logging.info(f"Montando painel com o histórico de versões até: {versao_atual}...")

        df_atual = df[df['versao_arquivo'] == versao_atual].copy()

        # A primeira coluna é 'versao_arquivo', a segunda deve ser o Produto/Cultura
        col_produto_str = df_atual.columns[1]

        # Inicia a tabela de exibição com as colunas de Produtos, 2024 e 2025
        cols_base = [col_produto_str]
        for ano in ['2024', '2025']:
            if ano in df_atual.columns: cols_base.append(ano)
                
        df_exibicao = df_atual[cols_base].copy()

        # Loop dinâmico: Extrai a projeção de 2026 de TODAS as versões (Jan, Fev, Mar...) e as coloca lado a lado
        colunas_2026 = []
        for v in versoes:
            df_v = df[df['versao_arquivo'] == v].copy()
            if '2026' in df_v.columns:
                nome_coluna_mes = f'2026 ({v})'
                df_v_2026 = df_v[[col_produto_str, '2026']].rename(columns={'2026': nome_coluna_mes})
                df_exibicao = pd.merge(df_exibicao, df_v_2026, on=col_produto_str, how='left')
                colunas_2026.append(nome_coluna_mes)

        df_exibicao = df_exibicao.rename(columns={col_produto_str: 'Produto / Cultura'})

        # Calcula a variação de 2026 (Sempre comparando a última com a penúltima disponível)
        coluna_var = 'Variação vs Mês Anterior (%)'
        if len(colunas_2026) >= 2:
            col_atual = colunas_2026[-1]
            col_ant = colunas_2026[-2]
            
            df_exibicao[col_atual] = pd.to_numeric(df_exibicao[col_atual], errors='coerce').fillna(0)
            df_exibicao[col_ant] = pd.to_numeric(df_exibicao[col_ant], errors='coerce').fillna(0)
            df_exibicao[coluna_var] = ((df_exibicao[col_atual] - df_exibicao[col_ant]) / df_exibicao[col_ant].replace(0, pd.NA)) * 100
        else:
            df_exibicao[coluna_var] = pd.NA

        # Inserção da coluna de Insights de Maquinário (Motor Heurístico / IA)
        def gerar_insight(row):
            cultura = str(row['Produto / Cultura']).lower()
            var = row[coluna_var]
            
            if pd.isna(var): return "-"
            
            maquinas = ""
            if any(c in cultura for c in ['soja', 'milho', 'trigo', 'arroz', 'algodão', 'algodao', 'sorgo']):
                maquinas = "Tratores de Alta Potência (240-339HP e >340HP), Colheitadeiras, Pulverizadores e Plantadeiras"
            elif any(c in cultura for c in ['café', 'cafe']):
                maquinas = "Tratores Estreitos/Especiais (50-79HP e 80-119HP) e Colheitadeiras de Café"
            elif 'cana' in cultura:
                maquinas = "Tratores Pesados (170-239HP e 240-339HP) para transbordo e Colheitadeiras de Cana"
            elif any(c in cultura for c in ['laranja', 'uva', 'maçã', 'maca', 'banana', 'cacau']):
                maquinas = "Tratores Leves/Fruteiros (0-49HP e 50-79HP)"
            elif any(c in cultura for c in ['feijão', 'feijao', 'batata', 'cebola', 'tomate', 'mandioca', 'amendoim']):
                maquinas = "Tratores Médios (80-119HP e 120-169HP) e Implementos Menores"
            else:
                maquinas = "Tratores Multiuso (80-119HP e 120-169HP)"
                
            if var > 2:
                return f"📈 OPORTUNIDADE: Aumento forte no VBP sinaliza capitalização. Ótimo cenário para venda de {maquinas}."
            elif var > 0:
                return f"↗️ AQUECIMENTO: Alta moderada estimula a renovação natural da frota de {maquinas}."
            elif var < -2:
                return f"🔴 ALERTA: Queda severa retrai financiamentos. Venda de {maquinas} deve cair; focar em Pós-Venda."
            elif var < 0:
                return f"↘️ CAUTELA: Leve retração faz o produtor segurar caixa. Giro de {maquinas} pode ficar mais lento."
            else:
                return f"➡️ ESTABILIDADE: Fluxo normal de reposição para {maquinas}."

        df_exibicao['Impacto em Maquinário (IA)'] = df_exibicao.apply(gerar_insight, axis=1)

        # Limpeza fina: Garante que os números são numéricos e exclui culturas 100% zeradas para não poluir o painel
        cols_numericas = [c for c in df_exibicao.columns if c not in ['Produto / Cultura', coluna_var, 'Impacto em Maquinário (IA)']]
        for col in cols_numericas:
            df_exibicao[col] = pd.to_numeric(df_exibicao[col], errors='coerce').fillna(0)
        df_exibicao = df_exibicao[(df_exibicao[cols_numericas] != 0).any(axis=1)]

        def formatar_cores(val):
            if pd.isna(val) or isinstance(val, str): return ''
            try:
                v = float(val)
                if v > 0.05: return 'color: #107C41; font-weight: 700;' # Verde corporativo
                if v < -0.05: return 'color: #D83B01; font-weight: 700;' # Vermelho corporativo
            except: pass
            return 'color: #383d41; background-color: #e2e3e5;' 
            
        def formata_br(x):
            if pd.isna(x): return "-"
            try: 
                if float(x) == 0: return "-"
                return f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except: return x

        # Estilo profissional AGCO
        estilo = [
            {'selector': 'table', 'props': 'border-collapse: collapse; margin: 20px 0; font-family: "Segoe UI", sans-serif; font-size: 0.85em; width: 100%; border: 1px solid #e0e0e0;'},
            {'selector': 'thead tr', 'props': 'background-color: #2c2c2c; color: #ffffff; text-align: right;'}, # Cinza chumbo/preto
            {'selector': 'th', 'props': 'padding: 12px 10px; border: 1px solid #444; text-transform: uppercase; font-size: 0.9em; letter-spacing: 0.5px;'},
            {'selector': 'td', 'props': 'padding: 10px 10px; border: 1px solid #e0e0e0; text-align: right; color: #333;'},
            {'selector': 'th:first-child, td:first-child', 'props': 'text-align: left; font-weight: 600; background-color: #fafafa; border-right: 2px solid #ccc;'},
            {'selector': 'th:last-child, td:last-child', 'props': 'text-align: left; font-size: 0.85em; max-width: 330px; white-space: normal; line-height: 1.4; border-left: 2px solid #ccc; background-color: #fcfcfc;'},
            {'selector': 'tbody tr:hover', 'props': 'background-color: #f1f1f1;'},
            {'selector': 'tbody tr:nth-of-type(even)', 'props': 'background-color: #fafafa;'}
        ]

        html = (df_exibicao.style.map(formatar_cores, subset=[coluna_var] if coluna_var in df_exibicao.columns else [])
                .format("{:.2f}%", subset=[coluna_var], na_rep="-")
                .format(formata_br, subset=cols_numericas)
                .set_table_styles(estilo).to_html(index=False))
        
        caminho_html = os.path.join(self.dir_relatorios, 'index.html')
        
        with open(caminho_html, 'w', encoding='utf-8') as f:
            f.write(f'''
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
                <meta charset="utf-8">
                <title>Dashboard VBP - Estilo Executivo</title>
                <style>
                    body {{
                        background-color: #e9ecef;
                        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                        margin: 0;
                        padding: 30px;
                        color: #333;
                    }}
                    .dashboard-container {{
                        background-color: #ffffff;
                        border-radius: 8px;
                        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.08);
                        padding: 30px 40px;
                        border-top: 6px solid #BA0C2F; /* Vermelho AGCO */
                        max-width: 1500px;
                        margin: 0 auto;
                    }}
                    .header {{
                        display: flex;
                        justify-content: space-between;
                        align-items: center;
                        border-bottom: 2px solid #f0f0f0;
                        padding-bottom: 15px;
                        margin-bottom: 25px;
                    }}
                    .title-area h2 {{
                        color: #2c2c2c;
                        margin: 0 0 5px 0;
                        font-size: 1.8em;
                        font-weight: 700;
                    }}
                    .title-area p {{
                        color: #666;
                        margin: 0;
                        font-size: 0.95em;
                    }}
                    .logo-area {{
                        text-align: right;
                    }}
                    .logo-text {{
                        font-size: 26px;
                        font-weight: 900;
                        color: #2c2c2c;
                        letter-spacing: 1.5px;
                    }}
                    .logo-text span {{
                        color: #BA0C2F; /* Vermelho AGCO */
                    }}
                    .info-box {{
                        background-color: #f8f9fa;
                        border-left: 4px solid #4a4a4a;
                        padding: 12px 15px;
                        margin-bottom: 20px;
                        font-size: 0.9em;
                        color: #555;
                    }}
                </style>
            </head>
            <body>
                <div class="dashboard-container">
                    <div class="header">
                        <div class="title-area">
                            <h2>Painel Analítico de Acompanhamento VBP</h2>
                            <p>Evolução de Safra e Valor Bruto da Produção Nacional</p>
                        </div>
                        <div class="logo-area">
                            <div class="logo-text">AG<span>CO</span></div>
                        </div>
                    </div>
                    
                    <div class="info-box">
                        <strong>Última Extração de Dados:</strong> {datetime.now().strftime('%d/%m/%Y %H:%M')} <br>
                        <strong>Comparativo de Projeções:</strong> Exibindo o histórico de 2024 e 2025, junto com a evolução de todas as revisões do Governo de 2026.
                    </div>
                    
                    {html}
                </div>
            </body>
            </html>
            ''')
            
        logging.info(f"Sucesso! Relatório gerado em: {os.path.abspath(caminho_html)}")
        return caminho_html

# ==========================================
# 4. EXECUÇÃO PRINCIPAL
# ==========================================
if __name__ == "__main__":
    print(f"\n{'='*60}\nATENÇÃO: Os arquivos e pastas estão sendo criados EXATAMENTE aqui:\n-> {BASE_DIR} <-\n{'='*60}\n")

    PASTAS = {
        "down": os.path.join(BASE_DIR, "downloads"), 
        "hist": os.path.join(BASE_DIR, "bases_historicas"), 
        "rel": BASE_DIR
    }
    URL = "https://www.gov.br/agricultura/pt-br/assuntos/politica-agricola/valor-bruto-da-producao-agropecuaria-vbp"

    scraper = AgroScraper(URL, PASTAS["down"])
    
    if scraper.extrair_planilhas():
        scraper.padronizar_nomes_arquivos()
        etl = AgroETL(PASTAS["down"], PASTAS["hist"], PASTAS["rel"])
        
        if etl.cruzar_e_salvar_versao():
            etl.gerar_relatorio_html()
