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
                
                # Filtra apenas planilhas gerais (Brasil)
                if 'regional' in nome_arquivo:
                    continue
                if 'vbp' not in nome_arquivo:
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
            
            match_data = re.search(r'(\d{4})(\d{2})', nome_arquivo)
            if match_data:
                versao_formatada = f"{match_data.group(1)} - {match_data.group(2)}"
            else:
                versao_formatada = nome_arquivo
                
            logging.info(f"Lendo e empilhando planilha: {nome_arquivo} (Versão: {versao_formatada})...")
            try:
                dicionario_abas = pd.read_excel(arquivo, sheet_name=None, header=None)
                df_maior = pd.DataFrame()
                
                for nome_aba, df_aba in dicionario_abas.items():
                    if len(df_aba) > len(df_maior):
                        df_maior = df_aba
                        
                if not df_maior.empty:
                    idx_cabecalho = 0
                    for idx, row in df_maior.iterrows():
                        valores_linha = row.astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                        if '1989' in valores_linha.values or valores_linha.str.match(r'^(19|20)\d{2}$').sum() > 3:
                            idx_cabecalho = idx
                            break
                            
                    df_maior.columns = df_maior.iloc[idx_cabecalho]
                    df_maior.columns.name = None
                    df_maior = df_maior.iloc[idx_cabecalho + 1:].reset_index(drop=True)

                    df_maior = df_maior.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)

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
                        
                        termos_rodape = 'fonte|nota|elaboração|elaboracao|atualizado'
                        mascara_rodape = df_limpo[primeira_col].astype(str).str.lower().str.contains(termos_rodape, na=False)
                        if mascara_rodape.any():
                            idx_rodape = mascara_rodape.idxmax()
                            df_limpo = df_limpo.loc[:idx_rodape].iloc[:-1]

                        df_limpo['chave_temp'] = df_limpo[primeira_col].astype(str).str.strip().str.lower()
                        df_limpo['chave_temp'] = df_limpo['chave_temp'].apply(
                            lambda x: ''.join(c for c in unicodedata.normalize('NFD', x) if unicodedata.category(c) != 'Mn')
                        )
                        df_limpo = df_limpo.drop_duplicates(subset=['chave_temp'], keep='first').drop(columns=['chave_temp'])
                        
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

        novas_colunas = []
        for i, c in enumerate(df.columns):
            if i > 0: 
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
        col_produto_str = df_atual.columns[1]

        # Remove duplicatas para evitar explosão de memória (Cartesian Product)
        df_atual = df_atual.drop_duplicates(subset=[col_produto_str])

        # Identifica dinamicamente o último ano (ex: 2026 ou 2027) e os anteriores
        anos_encontrados = []
        for c in df_atual.columns:
            match = re.match(r'^(20\d{2})$', str(c).strip())
            if match: anos_encontrados.append(int(match.group(1)))
            
        if anos_encontrados:
            ano_maximo = str(max(anos_encontrados))
            ano_anterior = str(max(anos_encontrados) - 1)
            ano_retrasado = str(max(anos_encontrados) - 2)
        else:
            ano_maximo, ano_anterior, ano_retrasado = '2026', '2025', '2024'

        cols_base = [col_produto_str]
        for ano in [ano_retrasado, ano_anterior]:
            if ano in df_atual.columns: cols_base.append(ano)
                
        df_exibicao = df_atual[cols_base].copy()

        colunas_ano_maximo = []
        for v in versoes:
            df_v = df[df['versao_arquivo'] == v].copy()
            df_v = df_v.drop_duplicates(subset=[col_produto_str])
            if ano_maximo in df_v.columns:
                nome_coluna_mes = f'{ano_maximo} ({v})'
                df_v_ano = df_v[[col_produto_str, ano_maximo]].rename(columns={ano_maximo: nome_coluna_mes})
                df_exibicao = pd.merge(df_exibicao, df_v_ano, on=col_produto_str, how='left')
                colunas_ano_maximo.append(nome_coluna_mes)

        df_exibicao = df_exibicao.rename(columns={col_produto_str: 'Produto / Cultura'})

        coluna_var_mes = 'Variação vs Mês Anterior (%)'
        if len(colunas_ano_maximo) >= 2:
            col_atual = colunas_ano_maximo[-1]
            col_ant = colunas_ano_maximo[-2]
            
            df_exibicao[col_atual] = pd.to_numeric(df_exibicao[col_atual], errors='coerce').fillna(0)
            df_exibicao[col_ant] = pd.to_numeric(df_exibicao[col_ant], errors='coerce').fillna(0)
            df_exibicao[coluna_var_mes] = ((df_exibicao[col_atual] - df_exibicao[col_ant]) / df_exibicao[col_ant].replace(0, pd.NA)) * 100
        else:
            df_exibicao[coluna_var_mes] = pd.NA

        # Variação vs Ano Anterior
        coluna_var_ano = f'Variação {ano_maximo} vs {ano_anterior} (%)'
        if colunas_ano_maximo and ano_anterior in df_exibicao.columns:
            col_atual = colunas_ano_maximo[-1]
            df_exibicao[ano_anterior] = pd.to_numeric(df_exibicao[ano_anterior], errors='coerce').fillna(0)
            df_exibicao[coluna_var_ano] = ((df_exibicao[col_atual] - df_exibicao[ano_anterior]) / df_exibicao[ano_anterior].replace(0, pd.NA)) * 100
        else:
            df_exibicao[coluna_var_ano] = pd.NA

        # -------------------------------------------------------------------
        # AJUSTE: Textos concisos para o Impacto em Maquinário (IA)
        # -------------------------------------------------------------------
        def gerar_insight(row):
            cultura = str(row['Produto / Cultura']).lower()
            var = row[coluna_var_mes]
            
            if pd.isna(var): return "-"
            
            maquinas = ""
            if any(c in cultura for c in ['soja', 'milho', 'trigo', 'sorgo']):
                maquinas = "Impacto maior: Tratores 240-339cv e Colheitadeiras Classe 7>"
            elif any(c in cultura for c in ['algodão', 'algodao']):
                maquinas = "Impacto maior: Tratores 240-339cv e Colheitadeiras de Algodão"
            elif 'arroz' in cultura:
                maquinas = "Impacto maior: Tratores 100-130cv e Colheitadeiras Arrozeiras"
            elif any(c in cultura for c in ['café', 'cafe']):
                maquinas = "Impacto maior: Tratores Estreitos e Colhedoras de Café"
            elif 'cana' in cultura:
                maquinas = "Impacto maior: Tratores >300cv e Colhedoras de Cana"
            elif any(c in cultura for c in ['laranja', 'uva', 'maçã', 'maca', 'banana', 'cacau']):
                maquinas = "Impacto maior: Tratores Fruteiros (<80cv)"
            elif any(c in cultura for c in ['feijão', 'feijao', 'amendoim']):
                maquinas = "Impacto maior: Tratores 100-140cv e Colheitadeiras Classe 5-6"
            elif any(c in cultura for c in ['batata', 'cebola', 'tomate', 'mandioca']):
                maquinas = "Impacto maior: Tratores Médios 100-140cv"
            else:
                maquinas = "Impacto maior: Tratores Multiuso"
                
            if var > 2:
                return f"📈 {maquinas}"
            elif var > 0:
                return f"↗️ {maquinas}"
            elif var < -2:
                return f"🔴 {maquinas}"
            elif var < 0:
                return f"↘️ {maquinas}"
            else:
                return f"➡️ {maquinas}"

        df_exibicao['Impacto Máquinas Agrícolas'] = df_exibicao.apply(gerar_insight, axis=1)

        cols_numericas = [c for c in df_exibicao.columns if c not in ['Produto / Cultura', coluna_var_mes, coluna_var_ano, 'Impacto Máquinas Agrícolas']]
        for col in cols_numericas:
            df_exibicao[col] = pd.to_numeric(df_exibicao[col], errors='coerce').fillna(0)
        df_exibicao = df_exibicao[(df_exibicao[cols_numericas] != 0).any(axis=1)]

        # Ajuste de Cores Condicionais (AGCO Red & Green)
        def formatar_cores(val):
            if pd.isna(val) or isinstance(val, str): return ''
            try:
                v = float(val)
                if v > 0: return 'color: #107C41; font-weight: 700;' 
                if v < 0: return 'color: #BA0C2F; font-weight: 700;' 
            except: pass
            return ''
            
        def formata_br(x):
            if pd.isna(x): return "-"
            try: 
                if float(x) == 0: return "-"
                return f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except: return x

        try:
            total_culturas = len(df_exibicao)
            valid_vars = df_exibicao.dropna(subset=[coluna_var_mes])
            if not valid_vars.empty:
                max_idx = valid_vars[coluna_var_mes].idxmax()
                min_idx = valid_vars[coluna_var_mes].idxmin()
                
                maior_alta_prod = valid_vars.loc[max_idx, 'Produto / Cultura']
                maior_alta_val = valid_vars.loc[max_idx, coluna_var_mes]
                str_alta = f"{maior_alta_prod} (+{maior_alta_val:.1f}%)"
                
                maior_queda_prod = valid_vars.loc[min_idx, 'Produto / Cultura']
                maior_queda_val = valid_vars.loc[min_idx, coluna_var_mes]
                str_queda = f"{maior_queda_prod} ({maior_queda_val:.1f}%)"
            else:
                str_alta, str_queda = "-", "-"
        except:
            total_culturas = 0; str_alta = "-"; str_queda = "-"

        cols_formatar_cores = [c for c in [coluna_var_mes, coluna_var_ano] if c in df_exibicao.columns]
        html = (df_exibicao.style.hide(axis="index")
                .map(formatar_cores, subset=cols_formatar_cores)
                .format("{:.2f}%", subset=cols_formatar_cores, na_rep="-")
                .format(formata_br, subset=cols_numericas)
                .to_html())
        
        caminho_html = os.path.join(self.dir_relatorios, 'index.html')
        
        # -------------------------------------------------------------------
        # AJUSTE: Layout Polido (Cores AGCO, KPIs reduzidos, Tradução)
        # -------------------------------------------------------------------
        with open(caminho_html, 'w', encoding='utf-8') as f:
            f.write(f'''
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
                <meta charset="utf-8">
                <title>Painel de Renda Agrícola</title>
                <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
                <style>
                    :root {{
                        --primary-color: #BA0C2F; /* Vermelho AGCO */
                        --dark-color: #231F20;    /* Preto Corporativo AGCO */
                        --bg-light: #F4F5F7;
                        --border-color: #E2E6E9;
                    }}
                    body {{
                        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                        background-color: var(--bg-light);
                        margin: 0;
                        padding: 30px;
                        color: #333;
                        top: 0 !important;
                    }}
                    .skiptranslate {{ display: none !important; }} 
                    
                    .dashboard-container {{
                        background-color: #ffffff;
                        border-radius: 6px;
                        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
                        padding: 25px 30px;
                        border-top: 5px solid var(--primary-color);
                        max-width: 1400px;
                        margin: 0 auto;
                    }}
                    .header {{
                        display: flex;
                        justify-content: space-between;
                        align-items: flex-start;
                        border-bottom: 1px solid var(--border-color);
                        padding-bottom: 15px;
                        margin-bottom: 25px;
                    }}
                    .title-area h2 {{ color: var(--dark-color); margin: 0 0 5px 0; font-size: 1.6em; font-weight: 700; text-transform: uppercase; letter-spacing: -0.5px; }}
                    .title-area p {{ color: #666; margin: 0; font-size: 0.9em; }}
                    
                    /* KPIs - Menores e Polidos */
                    .kpi-grid {{ display: flex; gap: 15px; margin-bottom: 25px; }}
                    .kpi-card {{ 
                        flex: 1; 
                        background: #fff; 
                        padding: 8px 12px; 
                        border-radius: 3px; 
                        box-shadow: 0 1px 3px rgba(0,0,0,0.05); 
                        border: 1px solid var(--border-color);
                        border-left: 4px solid var(--dark-color);
                        display: flex;
                        flex-direction: column;
                        justify-content: center;
                    }}
                    .kpi-card.brand {{ border-left-color: var(--primary-color); }}
                    .kpi-card.positive {{ border-left-color: #107C41; }}
                    .kpi-card.negative {{ border-left-color: var(--primary-color); }}
                    
                    .kpi-title {{ font-size: 9px; color: #777; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px; margin-bottom: 2px; }}
                    .kpi-value {{ font-size: 15px; font-weight: 800; color: var(--dark-color); margin: 0; }}
                    
                    /* Botões e Ações */
                    .action-buttons {{ display: flex; gap: 8px; margin-top: 12px; justify-content: flex-end; }}
                    .btn {{ 
                        background: #fff; border: 1px solid #ccc; padding: 5px 10px; border-radius: 4px; 
                        cursor: pointer; font-size: 11px; font-weight: 600; color: var(--dark-color);
                        transition: all 0.2s;
                    }}
                    .btn:hover {{ background: var(--bg-light); border-color: var(--dark-color); }}
                    .btn-excel {{ background: #107C41; color: white; border-color: #107C41; }}
                    .btn-excel:hover {{ background: #0c6132; border-color: #0c6132; }}

                    /* Tabela */
                    table {{ border-collapse: collapse; width: 100%; font-size: 13px; }}
                    thead th {{ background-color: var(--dark-color); color: #fff; text-align: center; padding: 12px 10px; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; border: none; }}
                    thead th:first-child {{ text-align: left; }}
                    tbody tr {{ border-bottom: 1px solid var(--border-color); }}
                    tbody td {{ padding: 8px 10px; text-align: center; color: #444; border: none; }}
                    tbody td:first-child, tbody th:first-child {{ text-align: left; font-weight: 600; color: var(--dark-color); }}
                    /* Controle super restrito da largura da coluna de IA */
                    thead th:last-child, tbody td:last-child {{ text-align: left; max-width: 160px; line-height: 1.1; font-size: 10px; color: #ffffff; background-color: #2c3e50; white-space: normal; }}
                    tbody tr:hover {{ background-color: #fafafa; }}
                </style>
            </head>
            <body>
                <div id="google_translate_element" style="display:none;"></div>
                
                <div class="dashboard-container">
                    <div class="header">
                        <div class="title-area">
                            <h2>Painel de renda agrícola</h2>
                            <p>Acompanhamento de Valor Bruto da Produção com Inteligência de Mercado</p>
                            <p style="margin-top: 5px; font-size: 0.85em; color: var(--primary-color); font-weight: 600;">Desenvolvido pela área da AGCO Reporting & Analytics</p>
                        </div>
                        <div style="text-align: right;">
                            <p style="margin: 0; color: #666; font-size: 0.85em;">Atualizado: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
                            <div class="action-buttons">
                                <button class="btn" onclick="doGTranslate('pt')">🇧🇷 PT</button>
                                <button class="btn" onclick="doGTranslate('en')">🇺🇸 EN</button>
                                <button class="btn" onclick="doGTranslate('es')">🇪🇸 ES</button>
                                <button class="btn btn-excel" onclick="exportExcel()">📊 Exportar Excel</button>
                            </div>
                        </div>
                    </div>

                    <div class="kpi-grid">
                        <div class="kpi-card brand">
                            <div class="kpi-title">Culturas Monitoradas</div>
                            <div class="kpi-value">{total_culturas}</div>
                        </div>
                        <div class="kpi-card positive">
                            <div class="kpi-title">Maior Alta (vs Mês Anterior)</div>
                            <div class="kpi-value">{str_alta}</div>
                        </div>
                        <div class="kpi-card negative">
                            <div class="kpi-title">Alerta de Queda (vs Mês Anterior)</div>
                            <div class="kpi-value">{str_queda}</div>
                        </div>
                    </div>

                    {html}
                </div>

                <script type="text/javascript">
                    function googleTranslateElementInit() {{
                        new google.translate.TranslateElement({{pageLanguage: 'pt', autoDisplay: false}}, 'google_translate_element');
                    }}
                    
                    function doGTranslate(lang) {{
                        if (lang === 'pt') {{
                            document.cookie = 'googtrans=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/;';
                            location.reload();
                            return;
                        }}
                        var select = document.querySelector('select.goog-te-combo');
                        if (select) {{
                            select.value = lang;
                            select.dispatchEvent(new Event('change'));
                        }}
                    }}
                    
                    function exportExcel() {{
                        var table = document.querySelector("table");
                        var wb = XLSX.utils.table_to_book(table, {{sheet: "Painel VBP"}});
                        XLSX.writeFile(wb, "Painel_VBP_Agro.xlsx");
                    }}
                </script>
                <script type="text/javascript" src="//translate.google.com/translate_a/element.js?cb=googleTranslateElementInit"></script>
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

    # Remove o arquivo antigo (se existir) para evitar confusão de visualização
    arquivo_antigo = os.path.join(BASE_DIR, "outputs", "relatorio_agronegocio.html")
    if os.path.exists(arquivo_antigo):
        try: os.remove(arquivo_antigo)
        except: pass

    scraper = AgroScraper(URL, PASTAS["down"])
    
    if scraper.extrair_planilhas():
        scraper.padronizar_nomes_arquivos()
        etl = AgroETL(PASTAS["down"], PASTAS["hist"], PASTAS["rel"])
        
        if etl.cruzar_e_salvar_versao():
            etl.gerar_relatorio_html()
