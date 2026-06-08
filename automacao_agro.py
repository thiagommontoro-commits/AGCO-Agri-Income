            if pd.isna(val) or isinstance(val, str): return ''
            try:
                v = float(val)
                if v > 0: return 'color: green; font-weight: bold;'
                if v < 0: return 'color: red; font-weight: bold;'
            except: pass
            return ''
            
        def formata_br(x):
            if pd.isna(x): return "-"
            try: 
                if float(x) == 0: return "-"
                return f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except: return x

        html = (df_exibicao.style.map(formatar_cores, subset=[coluna_var] if coluna_var in df_exibicao.columns else [])
                .format("{:.2f}%", subset=[coluna_var], na_rep="-")
                .format(formata_br, subset=cols_numericas)
                .to_html(index=False))
        
        caminho_html = os.path.join(self.dir_relatorios, 'index.html')
        
        with open(caminho_html, 'w', encoding='utf-8') as f:
            f.write(f'''
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
                <meta charset="utf-8">
                <title>Relatório VBP</title>
                <style>
                    body {{ font-family: Arial, sans-serif; margin: 20px; color: #333; }}
                    h2 {{ color: #005a9c; }}
                    table {{ border-collapse: collapse; width: 100%; font-size: 14px; margin-top: 15px; }}
                    th, td {{ border: 1px solid #dddddd; padding: 8px; text-align: right; }}
                    th {{ background-color: #f2f2f2; text-align: center; font-weight: bold; }}
                    td:first-child, th:first-child {{ text-align: left; font-weight: bold; }}
                    /* Mantém o texto da IA organizado para não esticar a tela infinitamente */
                    td:last-child, th:last-child {{ text-align: left; max-width: 400px; white-space: normal; }}
                    tr:nth-child(even) {{ background-color: #f9f9f9; }}
                    tr:hover {{ background-color: #f1f1f1; }}
                </style>
            </head>
            <body>
                <h2>Relatório Consolidado - Valor Bruto da Produção (VBP)</h2>
                <p><strong>Atualizado em:</strong> {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
                <hr>
                {html}
            </body>
            </html>
            ''')
