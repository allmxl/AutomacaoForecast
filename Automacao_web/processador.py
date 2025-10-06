import pandas as pd
import os
from datetime import datetime
from dateutil.relativedelta import relativedelta
import locale
import shutil

# --- FUNÇÃO 1: MAPEAMENTO (Versão Final e Estável) ---
def processando_fornecimento_código(caminho_dre):
    print("1. Criando mapa de fornecimento a partir do DRE...")
    try:
        # Leitura robusta que funcionou no diagnóstico
        df_dre = pd.read_excel(caminho_dre, sheet_name='Receita', header=1, dtype=str)
        
        if 'Código' in df_dre.columns and 'Tipo' in df_dre.columns:
            mapeamento = df_dre[['Código', 'Tipo']].copy()
            mapeamento['Código'] = mapeamento['Código'].str.strip()
            mapeamento.rename(columns={'Tipo': 'Fornecimento'}, inplace=True)
            mapeamento.drop_duplicates(subset=['Código'], inplace=True)
            print(' -> Mapa de fornecimento criado com sucesso.')
            return mapeamento
        else:
            print(' -> ERRO: Colunas "Código" e/ou "Tipo" não encontradas no DRE.')
            return None
    except Exception as e:
        print(f" -> ERRO inesperado ao criar mapa de fornecimento: {e}")
        return None

def processar_dre_realizado(caminho_dre):
    print("2. Processando dados 'Realizado' do DRE...")
    try:
        # 1. Define o período alvo (mês anterior) no formato 'ANO-MÊS', ex: "2025-09"
        periodo_alvo = (datetime.now() - relativedelta(months=1)).strftime('%Y-%m')
        print(f" -> Filtrando DRE para o período alvo: {periodo_alvo}")
        
        df_dre = pd.read_excel(caminho_dre, sheet_name='Receita', header=1, dtype={'Código': str})
        
        df_dre.dropna(subset=['Base', 'Período', 'Mês'], inplace=True)

        # 2. "Tradutor" de Mês em Português para número
        mes_map = {
            'janeiro': '01', 'fevereiro': '02', 'março': '03', 'abril': '04', 'maio': '05', 'junho': '06',
            'julho': '07', 'agosto': '08', 'setembro': '09', 'outubro': '10', 'novembro': '11', 'dezembro': '12'
        }
        
        # Cria uma coluna com o número do mês, buscando no "tradutor"
        df_dre['Mes_Num'] = df_dre['Mês'].str.lower().str.strip().map(mes_map)
        
        # 3. Cria a coluna 'AnoMes' de forma segura, ex: "2025-09"
        df_dre['AnoMes'] = df_dre['Período'].astype(str) + '-' + df_dre['Mes_Num']

        # 4. Filtra usando a nova coluna 'AnoMes' e a 'Base'
        filtro_periodo = (df_dre['AnoMes'] == periodo_alvo)
        filtro_base = (df_dre['Base'].str.strip().str.lower() == 'real')
        
        df_realizado = df_dre[filtro_periodo & filtro_base].copy()

        if df_realizado.empty:
            print(f" -> Nenhuma linha com base 'real' para o período {periodo_alvo} foi encontrada no DRE.")
            return None 

        # O resto do seu código original continua igual...
        df_realizado['Código'] = df_realizado['Código'].str.strip()
        df_realizado.rename(columns={'Período': 'Ano', 'VOLUME': 'Soma de Qtd', 'Receita Bruta': 'Soma de Valor', 'Tipo': 'Fornecimento'}, inplace=True)
        
        datas = pd.to_datetime(df_realizado['AnoMes'])
        df_realizado['Ano'] = datas.dt.year

        # Dicionário que traduz o NÚMERO do mês para a abreviação
        mes_num_map = {
            1: 'jan', 2: 'fev', 3: 'mar', 4: 'abr', 5: 'mai', 6: 'jun',
            7: 'jul', 8: 'ago', 9: 'set', 10: 'out', 11: 'nov', 12: 'dez'
        }
        df_realizado['Mês'] = datas.dt.month.map(mes_num_map)

        df_realizado['Base'] = 'Realizado'
        df_realizado['Ciclo'] = 'Realizado'
        df_realizado['Classificação'] = 'Realizado'
        df_realizado['Produto'] = df_realizado['Código'] + ' | ' + df_realizado['Descrição'].astype(str)
        
        print(f" -> {len(df_realizado)} linhas de 'Realizado' para o período {periodo_alvo} processadas.")
        
        # Garante a ordem e a presença de todas as colunas necessárias
        ordem_colunas = ['Base', 'Ciclo', 'Classificação', 'Ano', 'Mês', 'Código', 'Descrição', 'Produto', 'Marca', 'Fornecimento', 'Categoria', 'Soma de Qtd', 'Soma de Valor']
        for col in ordem_colunas:
            if col not in df_realizado.columns:
                df_realizado[col] = ''
        return df_realizado[ordem_colunas]
        
    except Exception as e:
        print(f" -> ERRO ao processar DRE para 'Realizado': {e}")
        return None

# --- FUNÇÃO 3: DRE ORÇAMENTO (Versão estável) ---
def processar_dre_orcamento(caminho_dre):
    print("2.1 Processando dados 'Orçamento' do DRE...")
    try:
        df = pd.read_excel(caminho_dre, sheet_name= 'Receita', header=1, dtype=str)
        
        ano_alvo = datetime.now().year + 1
        print(f" -> Procurando por orçamento com base 'orç' para o ano de {ano_alvo}...")

        df.dropna(subset=['Base', 'Período', 'Código'], inplace=True)
        
        # Usamos errors='coerce' para transformar datas inválidas em NaT
        df['Período_dt'] = pd.to_datetime(df['Período'], errors='coerce')
        
        filtro_base = (df['Base'].str.lower() == 'orç')
        filtro_ano = (df['Período_dt'].dt.year == ano_alvo)
        
        df_orcamento = df[filtro_base & filtro_ano].copy()

        if df_orcamento.empty:
            print(f" -> AVISO: Nenhuma linha de orçamento para {ano_alvo} foi encontrada no DRE.")
            return None

        df_orcamento.rename(columns={'Mês': 'Mês', 'Volume': 'Soma de Qtd', 'Receita Bruta': 'Soma de Valor', 'Tipo': 'Fornecimento'}, inplace=True)
        
        df_orcamento['Ano'] = pd.to_datetime(df_orcamento['Período']).dt.year
        df_orcamento['Mês'] = pd.to_datetime(df_orcamento['Ano'].astype(str) + '-' + df_orcamento['Mês'].astype(str), errors='coerce').dt.strftime('%B').str.capitalize()
        
        df_orcamento['Base'] = 'Orçamento'
        df_orcamento['Ciclo'] = 'Orçado'
        df_orcamento['Classificação'] = 'Orçamento'
        df_orcamento['Produto'] = df_orcamento['Código'].str.strip() + ' | ' + df_orcamento['Descrição'].astype(str)
        
        print(f" -> {len(df_orcamento)} linhas de 'Orçamento' para {ano_alvo} processadas.")
        
        ordem_colunas = ['Base', 'Ciclo', 'Classificação', 'Ano', 'Mês', 'Código', 'Descrição', 'Produto', 'Marca', 'Fornecimento', 'Categoria', 'Soma de Qtd', 'Soma de Valor']
        for col in ordem_colunas:
            if col not in df_orcamento.columns:
                df_orcamento[col] = ''
        return df_orcamento[ordem_colunas]
        
    except Exception as e:
        print(f" -> ERRO ao processar DRE para 'Orçamento': {e}")
        return None

# --- FUNÇÃO 4: FORECAST (Sua versão com correções) ---
def processar_planilha_forecast(caminho_arquivo, tipo_base, mapa_fornecimento):
    print(f"3. Processando Forecast: {os.path.basename(caminho_arquivo)}...")
    try:
        df = pd.read_excel(caminho_arquivo, header=[2, 3], dtype={'Código': str, 'Material': str})
        
        novas_colunas = []
        for col in df.columns:
            parte1 = str(col[0]); parte2 = str(col[1])
            if 'Unnamed' in parte1: novas_colunas.append(parte2)
            else: novas_colunas.append(f"{parte1}_{parte2}")
        df.columns = novas_colunas

        if 'Classificação' in df.columns: df = df.drop(columns=['Classificação'])
        
        colunas_id = ['Material', 'Código', 'Marca', 'Submarca', 'CategoriaMkt', 'Lançamento']
        for col in colunas_id:
            if col in df.columns: df[col] = df[col].astype(str).str.strip()
        
        df_produtos_master = df[colunas_id].copy().dropna(subset=['Código']).drop_duplicates()
        
        colunas_valor = [col for col in df.columns if 'Prev.Vendas' in col and '.' in col]
        df_melted = pd.melt(df, id_vars=colunas_id, value_vars=colunas_valor, var_name='Objetivo_Ciclo', value_name='Valor')
        df_melted = df_melted[df_melted['Valor'].notna()]
        
        split_data = df_melted['Objetivo_Ciclo'].str.split('_', expand=True)
        df_melted['Objetivo'] = split_data[0].str.replace(' ', '').str.replace('Vendas', 'vendas')
        df_melted['Ciclo_Raw'] = split_data.get(1)

        df_processado = df_melted.pivot_table(index=colunas_id + ['Ciclo_Raw'], columns='Objetivo', values='Valor', fill_value=0).reset_index()
        df_processado.rename(columns={'Prev.vendas(Qtde)': 'Soma de Qtd', 'Prev.vendas(R$)': 'Soma de Valor'}, inplace=True)

        todos_os_meses = df_processado[['Ciclo_Raw']].copy().drop_duplicates()
        if todos_os_meses.empty: return None
        df_esqueleto = df_produtos_master.merge(todos_os_meses, how='cross')

        df_final = pd.merge(df_esqueleto, df_processado, on=colunas_id + ['Ciclo_Raw'], how='left')
        df_final[['Soma de Qtd', 'Soma de Valor']] = df_final[['Soma de Qtd', 'Soma de Valor']].fillna(0)
        
        df_final.rename(columns={'Material': 'Descrição', 'CategoriaMkt': 'Categoria'}, inplace=True)
        df_final = df_final[df_final['Ciclo_Raw'].str.contains(r'\.')]
        split_ciclo = df_final['Ciclo_Raw'].str.split('.', expand=True)
        df_final['Ano'] = split_ciclo[1].apply(lambda ano: f"20{ano}" if len(str(ano)) == 2 else str(ano))
        df_final['Mês'] = split_ciclo[0].str.slice(0, 3).str.lower()
        
        data_atual = datetime.now()
        mes_abreviado = data_atual.strftime('%b').capitalize()
        ano_curto = data_atual.strftime('%y')
        mes_atual_numero = data_atual.month

        df_final['Base'] = f'Forecast {mes_abreviado}. {ano_curto} {tipo_base}'
        df_final['Ciclo'] = f'Ciclo {mes_atual_numero} {tipo_base} ({mes_abreviado}. {ano_curto})'
        df_final['Classificação'] = tipo_base
        df_final['Produto'] = df_final['Código'].astype(str) + ' | ' + df_final['Descrição'].astype(str)

        if mapa_fornecimento is not None and not mapa_fornecimento.empty:
            df_final = pd.merge(df_final, mapa_fornecimento, on='Código', how='left')
        
        df_final['Fornecimento'] = df_final['Fornecimento'].fillna('Terceiro')

        print(f" -> {len(df_final)} linhas de '{tipo_base}' processadas.")
        return df_final
    except Exception as e:
        print(f" -> ERRO ao processar o arquivo '{os.path.basename(caminho_arquivo)}': {e}")
        return None

# --- FUNÇÃO PRINCIPAL (Sua versão original com lógica de substituição melhorada) ---
def executar_processamento(caminho_dre, caminho_irrestrito=None, caminho_restrito=None):
    try:
        locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
        PASTA_OUTPUT_WEB = "output/"
        os.makedirs(PASTA_OUTPUT_WEB, exist_ok=True)

        caminho_bd_mestre = "K:/Administrativo Financeiro/Controladoria/GMD/2025/Qlik/Forecast/BD_Forecast.xlsx"
        pasta_arquivados_final = "Q:/ForecastArquivados/"
        pasta_saida_final = "Q:/ForecastPronto/"
        os.makedirs(pasta_arquivados_final, exist_ok=True)
        os.makedirs(pasta_saida_final, exist_ok=True)
        
        nome_arquivo_saida = f"BD_Processado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        caminho_final_para_salvar = os.path.join(pasta_saida_final, nome_arquivo_saida)
        
        ordem_colunas = ['Base', 'Ciclo', 'Classificação', 'Ano', 'Mês', 'Código', 'Descrição', 'Produto', 'Marca', 'Fornecimento', 'Categoria', 'Soma de Qtd', 'Soma de Valor']
        
        mapa_fornecimento = processando_fornecimento_código(caminho_dre)
        lista_dfs_novos = []

        df_realizado = processar_dre_realizado(caminho_dre)
        if df_realizado is not None: lista_dfs_novos.append(df_realizado)
        
        df_orcamento = processar_dre_orcamento(caminho_dre)
        if df_orcamento is not None: lista_dfs_novos.append(df_orcamento)

        arquivos_forecast = {}
        if caminho_irrestrito: arquivos_forecast[caminho_irrestrito] = "Irrestrito"
        if caminho_restrito: arquivos_forecast[caminho_restrito] = "Restrito"
        
        for caminho, tipo in arquivos_forecast.items():
            df_forecast = processar_planilha_forecast(caminho, tipo, mapa_fornecimento)
            if df_forecast is not None:
                lista_dfs_novos.append(df_forecast)
                shutil.move(caminho, os.path.join(pasta_arquivados_final, os.path.basename(caminho)))

        if not lista_dfs_novos:
            return None, "Nenhum dado novo foi processado."

        df_novos_dados = pd.concat(lista_dfs_novos, ignore_index=True)
        
        df_bd_final = None
        try:
            df_bd_antigo = pd.read_excel(caminho_bd_mestre, sheet_name='BD_Forecast')
            df_bd_antigo_limpo = df_bd_antigo.copy()
            
            # Lógica de substituição segura
            bases_e_periodos_novos = df_novos_dados[['Base', 'Ano', 'Mês']].drop_duplicates()
            
            for index, row in bases_e_periodos_novos.iterrows():
                base, ano, mes = row['Base'], row['Ano'], row['Mês']
                
                # Para Realizado e Orçamento, apaga por Ano/Mês
                if base in ['Realizado', 'Orçamento']:
                    condicao = (df_bd_antigo_limpo['Base'] == base) & (df_bd_antigo_limpo['Ano'] == ano) & (df_bd_antigo_limpo['Mês'] == mes)
                    df_bd_antigo_limpo = df_bd_antigo_limpo[~condicao]
                # Para Forecasts, apaga pela Base inteira (Ex: 'Forecast Out. 25 Irrestrito')
                else:
                    condicao = (df_bd_antigo_limpo['Base'] == base)
                    df_bd_antigo_limpo = df_bd_antigo_limpo[~condicao]
            
            df_bd_final = pd.concat([df_bd_antigo_limpo, df_novos_dados], ignore_index=True)
        except FileNotFoundError:
            df_bd_final = df_novos_dados

        df_bd_final = df_bd_final.reindex(columns=ordem_colunas)
        
        df_bd_final.to_excel(caminho_final_para_salvar, index=False, sheet_name='BD_Forecast')
        print(f" -> SUCESSO! Arquivo atualizado foi salvo em: {caminho_final_para_salvar}")

        caminho_para_download = os.path.join(PASTA_OUTPUT_WEB, nome_arquivo_saida)
        shutil.copy(caminho_final_para_salvar, caminho_para_download)
        
        return nome_arquivo_saida, None 
    except Exception as e:
        print(f"ERRO GERAL: {e}")
        return None, f"Ocorreu um erro geral no processamento: {str(e)}"