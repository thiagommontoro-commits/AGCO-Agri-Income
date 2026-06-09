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

        # Identifica dinamicamente qual é o ano mais recente que o governo está publicando
        anos_nos_links = []
        for link in links:
            h = link.get('href', '').lower()
            if '.xls' in h:
                anos_nos_links.extend([int(x) for x in re.findall(r'(20[1-3]\d)', h + (link.text or ""))])
        maior_ano_gov = max(anos_nos_links) if anos_nos_links else datetime.now().year
        anos_antigos = [str(ano) for ano in range(2010, maior_ano_gov)]

        for arquivo_antigo in glob.glob(os.path.join(self.dir_downloads, "*.xls*")):
            os.remove(arquivo_antigo)

        for link in links:
            href = link.get('href', '')
            href_lower = href.lower()
            if '.xls' in href_lower:
                url_completa = urljoin(self.url, href) 
                nome_arquivo = url_completa.split('/')[-1].split('?')[0].lower()
                if not nome_arquivo.endswith(('.xls', '.xlsx')):
                    nome_arquivo = f"vbp_brasil_{planilhas_baixadas}.xlsx"
                
                # Filtro restrito: apenas planilhas gerais (Brasil)
                texto_link = link.text.lower() if link.text else ""
                if 'regional' in nome_arquivo or 'regional' in href_lower or 'regional' in texto_link:
                    continue
                if 'vbp' not in nome_arquivo and 'vbp' not in href_lower and 'brasil' not in href_lower:
                    continue

                # Bloqueia o download de arquivos de anos anteriores. O histórico anual já vem consolidado na base vigente!
                if any(ano in nome_arquivo or ano in href_lower or ano in texto_link for ano in anos_antigos):
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
                base_name = re.sub(r'\.0$', '', str(c)).replace('*', '').strip()
                col_name = base_name
                contador = 1
                while col_name in novas_colunas:
                    col_name = f"{base_name}_{contador}"
                    contador += 1
                novas_colunas.append(col_name)
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

        if len(colunas_ano_maximo) >= 2:
            col_atual = colunas_ano_maximo[-1]
            col_ant = colunas_ano_maximo[-2]
            v_atual = versoes[-1]
            v_ant = versoes[-2]
            coluna_var_mes = f'Variação Mês ({v_atual} vs {v_ant})'
            
            df_exibicao[col_atual] = pd.to_numeric(df_exibicao[col_atual], errors='coerce').fillna(0)
            df_exibicao[col_ant] = pd.to_numeric(df_exibicao[col_ant], errors='coerce').fillna(0)
            df_exibicao[coluna_var_mes] = ((df_exibicao[col_atual] - df_exibicao[col_ant]) / df_exibicao[col_ant].replace(0, pd.NA)) * 100
        else:
            coluna_var_mes = 'Variação Mês (%)'
            df_exibicao[coluna_var_mes] = pd.NA

        # Variação vs Ano Anterior
        if colunas_ano_maximo and ano_anterior in df_exibicao.columns:
            col_atual = colunas_ano_maximo[-1]
            v_atual = versoes[-1]
            coluna_var_ano = f'Variação Ano ({v_atual} vs {ano_anterior})'
            df_exibicao[ano_anterior] = pd.to_numeric(df_exibicao[ano_anterior], errors='coerce').fillna(0)
            df_exibicao[coluna_var_ano] = ((df_exibicao[col_atual] - df_exibicao[ano_anterior]) / df_exibicao[ano_anterior].replace(0, pd.NA)) * 100
        else:
            coluna_var_ano = f'Variação Ano ({ano_maximo} vs {ano_anterior})'
            df_exibicao[coluna_var_ano] = pd.NA
            

        # -------------------------------------------------------------------
        # AJUSTE: Textos concisos para o Impacto em Maquinário (IA)
        # -------------------------------------------------------------------
        def gerar_insight(row):
            cultura = str(row['Produto / Cultura']).lower()
            var_mes = row.get(coluna_var_mes, pd.NA)
            var_ano = row.get(coluna_var_ano, pd.NA)
            
            # Se o robô rodar do zero (sem mês anterior salvo), a IA usa a variação anual (Safra vs Safra)
            var = var_ano if pd.isna(var_mes) else var_mes

            if pd.isna(var): return "-"
            
            maquinas = ""
            if any(c in cultura for c in ['soja', 'milho', 'trigo', 'sorgo']):
                maquinas = "Impacto maior: Trator Alta Potência (240-339cv) e Colheitadeiras"
            elif any(c in cultura for c in ['algodão', 'algodao']):
                maquinas = "Impacto maior: Trator Alta Potência (240-339cv) e Colheitadeiras"
            elif 'arroz' in cultura:
                maquinas = "Impacto maior: Trator Média Potência (100-130cv) e Colheitadeiras"
            elif any(c in cultura for c in ['café', 'cafe']):
                maquinas = "Impacto maior: Trator Baixa Potência (Estreitos) e Colheitadeiras"
            elif 'cana' in cultura:
                maquinas = "Impacto maior: Trator Alta Potência (>300cv) e Colheitadeiras"
            elif any(c in cultura for c in ['laranja', 'uva', 'maçã', 'maca', 'banana', 'cacau']):
                maquinas = "Impacto maior: Trator Baixa Potência (Fruteiros <80cv)"
            elif any(c in cultura for c in ['feijão', 'feijao', 'amendoim']):
                maquinas = "Impacto maior: Trator Média Potência (100-140cv) e Colheitadeiras"
            elif any(c in cultura for c in ['batata', 'cebola', 'tomate', 'mandioca']):
                maquinas = "Impacto maior: Trator Média Potência (100-140cv)"
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
            col_referencia = coluna_var_mes if not df_exibicao[coluna_var_mes].isna().all() else coluna_var_ano
            valid_vars = df_exibicao.dropna(subset=[col_referencia])
            if not valid_vars.empty:
                max_idx = valid_vars[col_referencia].idxmax()
                min_idx = valid_vars[col_referencia].idxmin()
                
                maior_alta_prod = valid_vars.loc[max_idx, 'Produto / Cultura']
                maior_alta_val = valid_vars.loc[max_idx, col_referencia]
                str_alta = f"{maior_alta_prod} (+{maior_alta_val:.1f}%)"
                
                maior_queda_prod = valid_vars.loc[min_idx, 'Produto / Cultura']
                maior_queda_val = valid_vars.loc[min_idx, col_referencia]
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
                <title>Painel Renda Agrícola</title>
                <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
                <style>
                    :root {{
                        --agco-red: #BA0C2F;
                        --text-main: #2c3e50;
                        --text-muted: #6c757d;
                        --bg-page: #f4f7f6;
                        --bg-card: #ffffff;
                        --border-light: #e9ecef;
                        --positive: #107C41;
                        --negative: #D83B01;
                        --header-bg: #1e293b;
                    }}
                    body {{ background-color: var(--bg-page); font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; margin: 0; padding: 20px; color: var(--text-main); top: 0 !important; }}
                    .skiptranslate {{ display: none !important; }}
                    .dashboard-container {{ background-color: var(--bg-card); border-radius: 12px; box-shadow: 0 8px 30px rgba(0, 0, 0, 0.05); max-width: 1550px; margin: 0 auto; overflow: hidden; }}
                    .header {{ display: flex; justify-content: space-between; align-items: center; background-color: var(--header-bg); padding: 20px 30px; border-bottom: 4px solid var(--agco-red); }}
                    .title-area h2 {{ margin: 0 0 5px 0; font-size: 1.8em; font-weight: 800; color: #ffffff; letter-spacing: -0.5px; }}
                    .title-area p {{ color: #94a3b8; margin: 0; font-size: 1em; font-weight: 500; }}
                    .developer-info {{ text-align: right; color: #94a3b8; font-size: 0.9em; line-height: 1.4; }}
                    .developer-info strong {{ color: #ffffff; font-size: 1.15em; display: block; margin-top: 4px; font-weight: 600; }}
                    
                    /* Menu de Navegação */
                    .navbar {{ background-color: var(--text-main); padding: 0 30px; display: flex; align-items: center; border-bottom: 2px solid var(--agco-red); }}
                    .nav-link {{ color: #94a3b8; text-decoration: none; padding: 12px 20px; font-weight: 600; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; transition: 0.2s; border-bottom: 3px solid transparent; }}
                    .nav-link:hover {{ color: #ffffff; }}
                    .nav-link.active {{ color: #ffffff; border-bottom-color: var(--agco-red); }}
                    .btn-lang {{ background: #ffffff; border: 1px solid var(--border-light); padding: 4px 8px; border-radius: 4px; cursor: pointer; font-size: 11px; font-weight: bold; color: var(--text-main); margin-left: 5px; transition: 0.2s; }}
                    .btn-lang:hover {{ background: #e9ecef; }}
                    .content-area {{ padding: 20px 30px 30px 30px; }}
                    
                    .info-strip {{ display: flex; flex-wrap: wrap; justify-content: space-between; background-color: #f8f9fa; border-left: 5px solid var(--agco-red); padding: 15px 20px; border-radius: 0 8px 8px 0; margin-bottom: 30px; font-size: 0.95em; color: #495057; gap: 15px; }}
                    .info-item strong {{ color: var(--text-main); }}
                    
                    .kpi-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 40px; }}
                    .kpi-card {{ background: #fff; border: 1px solid var(--border-light); border-radius: 10px; padding: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); position: relative; overflow: hidden; transition: transform 0.2s ease; }}
                    .kpi-card:hover {{ transform: translateY(-3px); box-shadow: 0 6px 12px rgba(0,0,0,0.05); }}
                    .kpi-card::before {{ content: ''; position: absolute; top: 0; left: 0; width: 100%; height: 4px; }}
                    .kpi-card.total::before {{ background-color: var(--text-main); }}
                    .kpi-card.positive::before {{ background-color: var(--positive); }}
                    .kpi-card.negative::before {{ background-color: var(--negative); }}
                    .kpi-title {{ font-size: 0.85em; text-transform: uppercase; font-weight: 700; color: var(--text-muted); margin-bottom: 10px; }}
                    .kpi-value {{ font-size: 2em; font-weight: 800; color: var(--text-main); margin-bottom: 5px; }}
                    
                    .table-container {{ overflow-x: auto; border-radius: 8px; border: 1px solid var(--border-light); }}
                    table {{ width: 100%; border-collapse: collapse; font-size: 0.9em; background-color: #fff; }}
                    thead {{ background-color: var(--text-main); color: #ffffff; }}
                    th {{ padding: 8px 6px; text-align: center; font-weight: 600; letter-spacing: 0px; text-transform: uppercase; font-size: 10.5px; white-space: nowrap; border: none; }}
                    th:first-child {{ text-align: left; position: sticky; left: 0; background-color: var(--text-main); z-index: 2; }}
                    td {{ padding: 6px 6px; text-align: center; border-bottom: 1px solid var(--border-light); color: #495057; font-variant-numeric: tabular-nums; border-top: none; border-left: none; border-right: none; font-size: 11.5px; }}
                    td:first-child {{ text-align: left; font-weight: 600; color: var(--text-main); position: sticky; left: 0; background-color: #fff; border-right: 2px solid var(--border-light); z-index: 1; font-size: 11.5px; }}
                    tbody tr:hover td {{ background-color: #f8f9fa; }}
                    
                    /* Coluna IA - Layout Moderno com Letra Preta e Fundo Claro */
                    thead th:last-child {{ text-align: left; max-width: 145px; background-color: var(--text-main); }}
                    tbody td:last-child {{ text-align: left; max-width: 145px; line-height: 1.25; font-size: 10.5px; color: #000000; font-weight: 700; white-space: normal; background-color: #f4f7f6; border-left: 2px solid var(--border-light); }}
                    
                    .action-buttons {{ margin-top: 25px; text-align: right; }}
                    .btn-excel {{ background: #107C41; color: white; border: none; padding: 12px 20px; border-radius: 5px; cursor: pointer; font-weight: bold; font-size: 14px; transition: 0.2s; }}
                    .btn-excel:hover {{ background: #0c6132; box-shadow: 0 4px 8px rgba(0,0,0,0.1); }}
                </style>
            </head>
            <body>
                <div id="google_translate_element" style="display:none;"></div>
                <div class="dashboard-container">
                    <div class="navbar">
                        <a href="index.html" class="nav-link active">Painel VBP (Renda)</a>
                        <a href="precos.html" class="nav-link">Painel de Preços (CEPEA)</a>
                    </div>
                    <div class="header">
                        <div class="title-area">
                            <h2>Painel renda agrícola</h2>
                            <p>Valor Bruto da Produção Nacional (em R$ Bilhões)</p>
                        </div>
                        <div class="developer-info">
                            Desenvolvido por<br>
                            <strong>Reporting & Analytics AGCO</strong>
                            <div style="margin-top: 8px;">
                                <button class="btn-lang" onclick="doGTranslate('pt')">🇧🇷 PT</button>
                                <button class="btn-lang" onclick="doGTranslate('en')">🇺🇸 EN</button>
                                <button class="btn-lang" onclick="doGTranslate('es')">🇪🇸 ES</button>
                            </div>
                        </div>
                    </div>

                    <div class="content-area">
                        <div class="info-strip">
                            <div class="info-item"><strong>Fonte dos Dados:</strong> Ministério da Agricultura do Brasil</div>
                            <div class="info-item"><strong>Última Atualização:</strong> {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>
                            <div class="info-item"><strong>Cenário:</strong> Base de projeções consolidada ({ano_maximo})</div>
                        </div>

                        <div class="kpi-grid">
                            <div class="kpi-card total">
                                <div class="kpi-title">Culturas Monitoradas</div>
                                <div class="kpi-value">{total_culturas}</div>
                            </div>
                            <div class="kpi-card positive">
                                <div class="kpi-title">Maior Alta ({coluna_var_mes})</div>
                                <div class="kpi-value" style="font-size: 1.4em;">{str_alta}</div>
                            </div>
                            <div class="kpi-card negative">
                                <div class="kpi-title">Alerta de Queda ({coluna_var_mes})</div>
                                <div class="kpi-value" style="font-size: 1.4em;">{str_queda}</div>
                            </div>
                        </div>

                        <div class="table-container">
                            {html}
                        </div>
                        
                        <div class="action-buttons">
                            <button class="btn-excel" onclick="exportExcel()">📊 Exportar Base Excel</button>
                        </div>
                    </div>
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
# 4. MÓDULO CEPEA (PREÇOS) - ESTRUTURA INICIAL
# ==========================================
class CepeaETL:
    def __init__(self, dir_relatorios: str):
        self.dir_relatorios = dir_relatorios
        os.makedirs(self.dir_relatorios, exist_ok=True)
        self.precos = {}

    def extrair_cotacoes_reais(self):
        logging.info("--- COLETANDO COTAÇÕES REAIS (CEPEA E BOLSAS) ---")
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        
        # 1. Indicadores Oficiais CEPEA (Mercado Interno Base)
        try:
            r = requests.get("https://www.cepea.esalq.usp.br/br/widget.aspx", headers=headers, timeout=15)
            soup = BeautifulSoup(r.text, 'html.parser')
            for tr in soup.find_all('tr'):
                th = tr.find('th')
                tds = tr.find_all('td')
                if th and tds:
                    nome = th.text.strip().lower()
                    val = tds[0].text.strip()
                    if not val: continue
                    
                    if 'soja' in nome: self.precos['soja_pr'] = f"R$ {val}"
                    elif 'milho' in nome: self.precos['milho_sp'] = f"R$ {val}"
                    elif 'boi' in nome and 'bezerro' not in nome: self.precos['boi_sp'] = f"R$ {val}"
                    elif 'arábica' in nome: self.precos['cafe_sp'] = f"R$ {val}"
                    elif 'conilon' in nome: self.precos['cafe_es'] = f"R$ {val}"
                    elif 'algodão' in nome: 
                        self.precos['algodao_mt'] = f"R$ {val}"
                        self.precos['algodao_ba'] = f"R$ {val}"
                    elif 'trigo' in nome: self.precos['trigo_pr'] = f"R$ {val}"
        except Exception as e:
            logging.error(f"Erro ao acessar CEPEA: {e}")

        # 2. Bolsas Internacionais e Praças Regionais (Notícias Agrícolas)
        try:
            r = requests.get("https://www.noticiasagricolas.com.br/cotacoes/", headers=headers, timeout=15)
            soup = BeautifulSoup(r.text, 'html.parser')
            
            def buscar_na_tabela(nome_tabela, nome_linha, simbolo="R$"):
                for table in soup.find_all('table', class_='cotacao'):
                    thead = table.find('thead')
                    if thead and nome_tabela.lower() in thead.text.lower():
                        for tr in table.find_all('tr'):
                            tds = tr.find_all('td')
                            if len(tds) >= 2 and nome_linha.lower() in tds[0].text.lower():
                                return f"{simbolo} {tds[1].text.strip()}"
                return None

            def buscar_primeira_linha(nome_tabela, simbolo="R$"):
                for table in soup.find_all('table', class_='cotacao'):
                    thead = table.find('thead')
                    if thead and nome_tabela.lower() in thead.text.lower():
                        tbody = table.find('tbody')
                        if tbody:
                            tr = tbody.find('tr')
                            if tr:
                                tds = tr.find_all('td')
                                if len(tds) >= 2:
                                    return f"{simbolo} {tds[1].text.strip()}"
                return None

            # Bolsas Internacionais (Chicago / NY / Londres)
            if v := buscar_primeira_linha('soja - cbot', 'US$'): self.precos['soja_cbot'] = v
            if v := buscar_primeira_linha('soja - prêmios', 'US$'): self.precos['soja_premio'] = v
            if v := buscar_primeira_linha('milho - cbot', 'US$'): self.precos['milho_cbot'] = v
            if v := buscar_primeira_linha('café - ny', 'US$'): self.precos['cafe_ny'] = v
            if v := buscar_primeira_linha('café - londres', 'US$'): self.precos['cafe_lon'] = v
            if v := buscar_primeira_linha('algodão - ny', 'US$'): self.precos['algodao_ny'] = v
            if v := buscar_primeira_linha('trigo - cbot', 'US$'): self.precos['trigo_cbot'] = v

            # Praças Regionais Mercado Físico
            if v := buscar_na_tabela('soja - mercado', 'rio verde'): self.precos['soja_go'] = v
            if v := buscar_na_tabela('soja - mercado', 'rondonópolis'): self.precos['soja_mt'] = v
            if v := buscar_na_tabela('milho - mercado', 'cascavel'): self.precos['milho_pr'] = v
            if v := buscar_na_tabela('milho - mercado', 'sorriso'): self.precos['milho_mt'] = v
            if v := buscar_na_tabela('boi gordo - mercado', 'goiânia'): self.precos['boi_go'] = v
            if v := buscar_na_tabela('boi gordo - mercado', 'cuiabá'): self.precos['boi_mt'] = v
            if v := buscar_na_tabela('laranja - mercado', 'indústria'): self.precos['laranja_sp'] = v

        except Exception as e:
            logging.error(f"Erro ao ler Bolsas/Praças extra: {e}")

    def get_preco(self, chave, default="R$ --,--"):
        return self.precos.get(chave, default)

    def gerar_relatorio_precos(self):
        logging.info("--- GERANDO PÁGINA DE PREÇOS (CEPEA) ---")
        self.extrair_cotacoes_reais()
        caminho_html = os.path.join(self.dir_relatorios, 'precos.html')
        
        with open(caminho_html, 'w', encoding='utf-8') as f:
            f.write(f'''
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
                <meta charset="utf-8">
                <title>Painel Preços Agrícolas - CEPEA</title>
                <style>
                    :root {{ --agco-red: #BA0C2F; --text-main: #2c3e50; --bg-page: #f4f7f6; --bg-card: #ffffff; --header-bg: #1e293b; }}
                    body {{ background-color: var(--bg-page); font-family: 'Segoe UI', sans-serif; margin: 0; padding: 20px; color: var(--text-main); }}
                    .dashboard-container {{ background-color: var(--bg-card); border-radius: 12px; box-shadow: 0 8px 30px rgba(0, 0, 0, 0.05); max-width: 1550px; margin: 0 auto; overflow: hidden; }}
                    .navbar {{ background-color: var(--text-main); padding: 0 30px; display: flex; align-items: center; border-bottom: 2px solid var(--agco-red); }}
                    .nav-link {{ color: #94a3b8; text-decoration: none; padding: 12px 20px; font-weight: 600; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; transition: 0.2s; border-bottom: 3px solid transparent; }}
                    .nav-link:hover {{ color: #ffffff; }}
                    .nav-link.active {{ color: #ffffff; border-bottom-color: var(--agco-red); }}
                    .header {{ padding: 30px; background-color: var(--header-bg); border-bottom: 4px solid var(--agco-red); }}
                    .header h2 {{ margin: 0 0 5px 0; font-size: 1.8em; font-weight: 800; color: #ffffff; letter-spacing: -0.5px; }}
                    .header p {{ color: #94a3b8; margin: 0; font-size: 1em; font-weight: 500; }}
                    .content-area {{ padding: 30px 40px; text-align: left; display: block; }}
                    
                    /* Grid de Preços */
                    .commodity-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 25px; margin-top: 10px; width: 100%; }}
                    .commodity-card {{ background: #fff; border: 1px solid var(--border-light); border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.02); }}
                    .commodity-header {{ background: var(--text-main); color: #fff; padding: 15px 20px; font-weight: 700; font-size: 1.1em; display: flex; justify-content: space-between; align-items: center; border-bottom: 4px solid var(--agco-red); }}
                    .price-section {{ padding: 15px 20px; border-bottom: 1px solid var(--border-light); background-color: #fafafa; }}
                    .price-section:last-child {{ border-bottom: none; background-color: #fff; }}
                    .section-title {{ font-size: 0.8em; text-transform: uppercase; color: var(--text-muted); font-weight: 800; margin-bottom: 12px; letter-spacing: 0.5px; }}
                    .price-row {{ display: flex; justify-content: space-between; margin-bottom: 8px; font-size: 0.95em; border-bottom: 1px dashed #e2e8f0; padding-bottom: 5px; }}
                    .price-row:last-child {{ margin-bottom: 0; border-bottom: none; padding-bottom: 0; }}
                    .price-row span {{ color: #475569; }}
                    .price-row strong {{ color: var(--text-main); font-weight: 700; font-variant-numeric: tabular-nums; }}
                </style>
            </head>
            <body>
                <div class="dashboard-container">
                    <div class="navbar">
                        <a href="index.html" class="nav-link">Painel VBP (Renda)</a>
                        <a href="precos.html" class="nav-link active">Painel de Preços (CEPEA)</a>
                    </div>
                    <div class="header">
                        <h2>Painel de Preços Agrícolas</h2>
                        <p>Acompanhamento de Cotações - Médias Mensais (CEPEA-ESALQ)</p>
                    </div>
                    <div class="content-area">
                        <p style="color: #64748b; font-size: 1.05em; margin-top: 0; margin-bottom: 25px; border-left: 4px solid var(--agco-red); padding-left: 15px;">
                            Estrutura preparada para receber integração com as <strong>Médias Mensais do Mercado Interno (CEPEA)</strong> e <strong>Mercado Internacional (Bolsas)</strong>.
                        </p>
                        
                        <div class="commodity-grid">
                            <!-- Soja -->
                            <div class="commodity-card">
                                <div class="commodity-header"><span>🌱 Soja</span></div>
                                <div class="price-section">
                                    <div class="section-title">Mercado Interno (Média Mensal)</div>
                                    <div class="price-row"><span>Paranaguá (PR)</span> <strong>{self.get_preco('soja_pr')} / sc 60kg</strong></div>
                                    <div class="price-row"><span>Rio Verde (GO)</span> <strong>{self.get_preco('soja_go')} / sc 60kg</strong></div>
                                    <div class="price-row"><span>Rondonópolis (MT)</span> <strong>{self.get_preco('soja_mt')} / sc 60kg</strong></div>
                                </div>
                                <div class="price-section">
                                    <div class="section-title">Mercado Internacional (Média Mensal)</div>
                                    <div class="price-row"><span>CBOT (Chicago)</span> <strong>{self.get_preco('soja_cbot')} / bu</strong></div>
                                    <div class="price-row"><span>Prêmio de Exportação</span> <strong>{self.get_preco('soja_premio')} / c</strong></div>
                                </div>
                            </div>

                            <!-- Milho -->
                            <div class="commodity-card">
                                <div class="commodity-header"><span>🌽 Milho</span></div>
                                <div class="price-section">
                                    <div class="section-title">Mercado Interno (Média Mensal)</div>
                                    <div class="price-row"><span>Campinas (SP)</span> <strong>{self.get_preco('milho_sp')} / sc 60kg</strong></div>
                                    <div class="price-row"><span>Cascavel (PR)</span> <strong>{self.get_preco('milho_pr')} / sc 60kg</strong></div>
                                    <div class="price-row"><span>Sorriso (MT)</span> <strong>{self.get_preco('milho_mt')} / sc 60kg</strong></div>
                                </div>
                                <div class="price-section">
                                    <div class="section-title">Mercado Internacional (Média Mensal)</div>
                                    <div class="price-row"><span>CBOT (Chicago)</span> <strong>{self.get_preco('milho_cbot')} / bu</strong></div>
                                </div>
                            </div>

                            <!-- Café -->
                            <div class="commodity-card">
                                <div class="commodity-header"><span>☕ Café</span></div>
                                <div class="price-section">
                                    <div class="section-title">Mercado Interno (Média Mensal)</div>
                                    <div class="price-row"><span>Arábica - São Paulo</span> <strong>{self.get_preco('cafe_sp')} / sc 60kg</strong></div>
                                    <div class="price-row"><span>Conilon - Espírito Santo</span> <strong>{self.get_preco('cafe_es')} / sc 60kg</strong></div>
                                </div>
                                <div class="price-section">
                                    <div class="section-title">Mercado Internacional (Média Mensal)</div>
                                    <div class="price-row"><span>ICE (Arábica - NY)</span> <strong>{self.get_preco('cafe_ny')} / lb</strong></div>
                                    <div class="price-row"><span>ICE (Robusta - Londres)</span> <strong>{self.get_preco('cafe_lon')} / ton</strong></div>
                                </div>
                            </div>

                            <!-- Algodão -->
                            <div class="commodity-card">
                                <div class="commodity-header"><span>☁️ Algodão</span></div>
                                <div class="price-section">
                                    <div class="section-title">Mercado Interno (Média Mensal)</div>
                                    <div class="price-row"><span>Mato Grosso (MT)</span> <strong>{self.get_preco('algodao_mt')} / lp</strong></div>
                                    <div class="price-row"><span>Bahia (BA)</span> <strong>{self.get_preco('algodao_ba')} / lp</strong></div>
                                </div>
                                <div class="price-section">
                                    <div class="section-title">Mercado Internacional (Média Mensal)</div>
                                    <div class="price-row"><span>ICE (Nova York)</span> <strong>{self.get_preco('algodao_ny')} / lb</strong></div>
                                </div>
                            </div>

                            <!-- Boi Gordo -->
                            <div class="commodity-card">
                                <div class="commodity-header"><span>🐂 Boi Gordo</span></div>
                                <div class="price-section">
                                    <div class="section-title">Mercado Interno (Média Mensal)</div>
                                    <div class="price-row"><span>São Paulo (SP)</span> <strong>{self.get_preco('boi_sp')} / @</strong></div>
                                    <div class="price-row"><span>Campo Grande (GO)</span> <strong>{self.get_preco('boi_go')} / @</strong></div>
                                    <div class="price-row"><span>Cuiabá (MT)</span> <strong>{self.get_preco('boi_mt')} / @</strong></div>
                                </div>
                                <div class="price-section">
                                    <div class="section-title">Mercado Internacional (Média Mensal)</div>
                                    <div class="price-row"><span>CME (Live Cattle - EUA)</span> <strong>{self.get_preco('boi_cme', 'US$ --,--')} / lb</strong></div>
                                </div>
                            </div>

                            <!-- Trigo & Aveia -->
                            <div class="commodity-card">
                                <div class="commodity-header"><span>🌾 Trigo & Aveia</span></div>
                                <div class="price-section">
                                    <div class="section-title">Mercado Interno (Média Mensal)</div>
                                    <div class="price-row"><span>Trigo - Paraná (PR)</span> <strong>{self.get_preco('trigo_pr')} / ton</strong></div>
                                    <div class="price-row"><span>Aveia - Rio G. do Sul (RS)</span> <strong>{self.get_preco('aveia_rs')} / ton</strong></div>
                                </div>
                                <div class="price-section">
                                    <div class="section-title">Mercado Internacional (Média Mensal)</div>
                                    <div class="price-row"><span>CBOT (Trigo - EUA)</span> <strong>{self.get_preco('trigo_cbot')} / bu</strong></div>
                                </div>
                            </div>

                            <!-- Frutas -->
                            <div class="commodity-card">
                                <div class="commodity-header"><span>🍊 Frutas (Destaques)</span></div>
                                <div class="price-section">
                                    <div class="section-title">Mercado Interno (Média Mensal - Atacado)</div>
                                    <div class="price-row"><span>Laranja (Indústria - SP)</span> <strong>{self.get_preco('laranja_sp')} / cx 40.8kg</strong></div>
                                    <div class="price-row"><span>Banana (Nanica - SP)</span> <strong>{self.get_preco('banana_sp')} / cx 22kg</strong></div>
                                    <div class="price-row"><span>Maçã (Fuji - SC)</span> <strong>{self.get_preco('maca_sc')} / cx 18kg</strong></div>
                                </div>
                            </div>
                            
                        </div>
                    </div>
                </div>
            </body>
            </html>
            ''')
        logging.info(f"Sucesso! Relatório de preços gerado em: {{os.path.abspath(caminho_html)}}")
        return caminho_html

# ==========================================
# 5. EXECUÇÃO PRINCIPAL
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

    # Gera a aba de preços
    cepea = CepeaETL(PASTAS["rel"])
    cepea.gerar_relatorio_precos()
