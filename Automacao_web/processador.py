import pandas as pd                   # Utilizado para a manipulação dos dados 
import os                             # Utilizado interação com sistema operacional 
from datetime import datetime         # Utilizado para trabalhar com datas e horas 
import locale                         # Utilizado para trabalhar com localizador (país/Idioma)
import shutil                         # Utilizado para manipulação de arquivos e diretórios em um nível maior que o 'import os' 

#  ------ Função que puxa o campo 'fornecimento' pela chave 'código' --------
def processando_fornecimento_código(caminho_dre):   # parâmetro criado para o caminho do arquivo DRE 
   
    print("1. Criando mapa de fornecimento a partir do DRE...")
    try:
        df_dre = pd.read_excel(caminho_dre, header=1)  #começa a contar pela linha 2 
        if 'Código' in df_dre.columns and 'Tipo' in df_dre.columns:  
            mapeamento = df_dre[['Código', 'Tipo']].copy()   # selecionando apenas as colunas código e tipo 
             # padronização dos dados -- astype(str): passa os valores para string -- 
            # .str.replace(r'\.0$', '', regex=True): Retira os .0 no final dos números vindos do excel 
            # .str.strip(): remove os espaços em branco 
            mapeamento['Código'] = mapeamento['Código'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            # renomeia a coluna tipo -- inplace é usado para conseguir alterar direto do dataframe e não criar outro 
            mapeamento.rename(columns={'Tipo': 'Fornecimento'}, inplace=True)
            # remove códigos duplicados 
            mapeamento.drop_duplicates(subset=['Código'], inplace=True)
            print(' -> Mapa de fornecimento criado com sucesso.')
            return mapeamento
        else: 
            print(' -> ERRO: Colunas "Código" e/ou "Tipo" não encontradas no DRE.')
            return None
         # Exception é retornado para erros inesperados que o try não vai conseguir rodar    
    except Exception as e:
        print(f" -> ERRO inesperado ao criar mapa de fornecimento: {e}")
        return None

#  ------ Função que puxa o campo Base 'Real'  --------

def processar_dre_realizado(caminho_dre):
    print("2. Processando dados 'Realizado' do DRE...")
    #definindo e filtrando as datas atuais -- mês, ano e padronização do mês em 'jan' 
    try:
        data_atual = datetime.now()
        ano_atual = data_atual.year
        mes_atual_padrao = data_atual.strftime('%b').lower()
        print(f" -> Filtrando DRE para o período atual: {mes_atual_padrao}/{ano_atual}")
        
        #Começa a ler o arquivo 
        df_dre = pd.read_excel(caminho_dre, header=1)
        
        # criação de filtros para validar os meses, ano e base correto para a tabela 
        mes_dre_padrao = df_dre['Mês'].astype(str).str.slice(0, 3).str.lower()
        filtro_ano = (df_dre['Período'] == ano_atual)
        filtro_mes = (mes_dre_padrao == mes_atual_padrao)
        filtro_base = (df_dre['Base'].str.lower() == 'real')
        df_realizado = df_dre[filtro_ano & filtro_mes & filtro_base].copy()

    # Verifica se df_realizado (que por ultimo foi definido com os 3 filtros) existe 
        if df_realizado.empty:
            print(f" -> Nenhuma linha com base 'real' para {mes_atual_padrao}/{ano_atual} foi encontrada no DRE.")
            return None 

            # padronização dos dados -- astype(str): passa os valores para string -- 
            # .str.replace(r'\.0$', '', regex=True): Retira os .0 no final dos números vindos do excel 
            # .str.strip(): remove os espaços em branco 
        df_realizado['Código'] = df_realizado['Código'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        # Renomeia as colunas para o novo padrão 
        df_realizado.rename(columns={
            'Período': 'Ano', 'Mês': 'Mês', 'VOLUME': 'Soma de Qtd',
            'Receita Líquida': 'Soma de Valor', 'Tipo': 'Fornecimento' 
        }, inplace=True)
        # Passa o Mês dentro do DRE para o formato padrão ('jan')
        df_realizado['Mês'] = df_realizado['Mês'].astype(str).str.slice(0, 3).str.lower()
        
        # Criação dos novos campos 
        df_realizado['Base'] = 'Realizado'
        df_realizado['Ciclo'] = 'Realizado'
        df_realizado['Classificação'] = 'Realizado'
        df_realizado['Produto'] = df_realizado['Código'].astype(str) + ' | ' + df_realizado['Descrição'].astype(str)
        
        print(f" -> {len(df_realizado)} linhas de 'Realizado' para {mes_atual_padrao}/{ano_atual} processadas.")
        return df_realizado
    except Exception as e:
        print(f" -> ERRO ao processar DRE para 'Realizado': {e}")
        return None


#  ------ Função que puxa o campo Base 'Real'  --------

def processar_planilha_forecast(caminho_arquivo, tipo_base, mapa_fornecimento):
    print(f"3. Processando Forecast: {os.path.basename(caminho_arquivo)}...")
    try:
        df = pd.read_excel(caminho_arquivo, header=[2, 3])
        
        # 1. Limpeza de Cabeçalho
        novas_colunas = []
        for col in df.columns:
            parte1 = str(col[0]); parte2 = str(col[1])
            if 'Unnamed' in parte1: novas_colunas.append(parte2)
            else: novas_colunas.append(f"{parte1}_{parte2}")
        df.columns = novas_colunas
        df = df.loc[:, ~df.columns.duplicated(keep='first')]

        if 'Classificação' in df.columns:
            df = df.drop(columns=['Classificação'])
        
        # 2. Padronizar TODAS as colunas de identificação para TEXTO desde o início
        colunas_id = ['Material', 'Código', 'Marca', 'Submarca', 'CategoriaMkt', 'Lançamento']
        for col in colunas_id:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
        df['Código'] = df['Código'].str.replace(r'\.0$', '', regex=True)
        
        # 3. Guarda a "Lista de Presença"
        df_produtos_master = df[colunas_id].copy().dropna(subset=['Código']).drop_duplicates()
        print(f" -> Encontrados {len(df_produtos_master)} códigos únicos no arquivo original.")

        # 4. Transformação dos dados
        colunas_valor = [col for col in df.columns if 'Prev.Vendas' in col and '.' in col]
        df_melted = pd.melt(df, id_vars=colunas_id, value_vars=colunas_valor,
                            var_name='Objetivo_Ciclo', value_name='Valor')
        df_melted = df_melted[df_melted['Valor'].notna()]
        
        split_data = df_melted['Objetivo_Ciclo'].str.split('_', expand=True)
        df_melted['Objetivo'] = split_data[0].str.replace(' ', '').str.replace('Vendas', 'vendas')
        df_melted['Ciclo_Raw'] = split_data.get(1)

        df_processado = df_melted.pivot_table(index=colunas_id + ['Ciclo_Raw'], columns='Objetivo', 
                                              values='Valor', fill_value=0).reset_index()
        df_processado.rename(columns={'Prev.vendas(Qtde)': 'Soma de Qtd', 'Prev.vendas(R$)': 'Soma de Valor'}, inplace=True)

        # 5. Criar o "esqueleto" e juntar com os dados processados
        todos_os_meses = df_processado[['Ciclo_Raw']].copy().drop_duplicates()
        if todos_os_meses.empty:
             print(f" -> AVISO: Nenhum dado mensal com valores encontrado no arquivo '{os.path.basename(caminho_arquivo)}'.")
             return None
        df_esqueleto = df_produtos_master.merge(todos_os_meses, how='cross')

        df_final = pd.merge(df_esqueleto, df_processado, on=colunas_id + ['Ciclo_Raw'], how='left')
        df_final[['Soma de Qtd', 'Soma de Valor']] = df_final[['Soma de Qtd', 'Soma de Valor']].fillna(0)

        # 6. Enriquecimento e formatação final
        df_final.rename(columns={'Material': 'Descrição', 'CategoriaMkt': 'Categoria'}, inplace=True)
        df_final = df_final[df_final['Ciclo_Raw'].str.contains(r'\.')]
        split_ciclo = df_final['Ciclo_Raw'].str.split('.', expand=True)
        df_final['Ano'] = split_ciclo[1]
        df_final['Mês'] = split_ciclo[0].str.slice(0, 3).str.lower()
        
        # --- BLOCO DE CÓDIGO ALTERADO ---
        data_atual = datetime.now()
        # Pega o nome do mês abreviado com 3 letras (ex: 'Set') e capitaliza a primeira
        mes_abreviado = data_atual.strftime('%b').capitalize()
        # Pega os dois últimos dígitos do ano (ex: 25)
        ano_curto = data_atual.strftime('%y')
        # Mantém a lógica do ciclo (mês - 1)
        mes_atual_numero = data_atual.month - 1

        # Usa o novo formato para criar a Base
        df_final['Base'] = f'Forecast {mes_abreviado}. {ano_curto} {tipo_base}'
        # Atualiza o Ciclo para manter a consistência
        df_final['Ciclo'] = f'Ciclo {mes_atual_numero} {tipo_base} ({mes_abreviado}. {ano_curto})'
        # ------------------------------------

        df_final['Classificação'] = tipo_base
        df_final['Produto'] = df_final['Código'].astype(str) + ' | ' + df_final['Descrição'].astype(str)

        if mapa_fornecimento is not None:
            df_final = pd.merge(df_final, mapa_fornecimento, on='Código', how='left')

        print(f" -> {len(df_final)} linhas de '{tipo_base}' processadas.")
        return df_final
    except Exception as e:
        print(f" -> ERRO ao processar o arquivo '{os.path.basename(caminho_arquivo)}': {e}")
        return None


def executar_processamento(caminho_dre, caminho_irrestrito=None, caminho_restrito=None):
    """Função principal que orquestra o processo de automação de forma segura."""
    try:
        # --- 1. CONFIGURAÇÃO DOS CAMINHOS FINAIS ---
        locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
        
        # Caminhos de trabalho do PROJETO WEB (para uploads e download temporário)
        PASTA_OUTPUT_WEB = "output/"
        PASTA_UPLOADS_WEB = "uploads/"
        os.makedirs(PASTA_OUTPUT_WEB, exist_ok=True)

        # Caminhos FINAIS na sua rede, conforme solicitado
        caminho_bd_mestre = "Q:/ForecastAutomacao/BD_Forecast_Teste.xlsx"
        pasta_arquivados_final = "Q:/ForecastArquivados/"
        pasta_saida_final = "Q:/ForecastPronto/"
        
        os.makedirs(pasta_arquivados_final, exist_ok=True)
        os.makedirs(pasta_saida_final, exist_ok=True)
        
        # Cria um nome dinâmico para o arquivo de saída
        mes_ano_atual = datetime.now().strftime("%b-%Y").lower() # ex: set-2025
        nome_arquivo_saida = f"BD_Forecast_{mes_ano_atual}1.xlsx"
        caminho_final_para_salvar = os.path.join(pasta_saida_final, nome_arquivo_saida)
        
        print(f"Arquivo mestre (leitura): {caminho_bd_mestre}")
        print(f"Arquivo de saída (gravação): {caminho_final_para_salvar}")
        print(f"Pasta de arquivados: {pasta_arquivados_final}")

        ordem_colunas = [
            'Base', 'Ciclo', 'Classificação', 'Ano', 'Mês', 'Código', 'Descrição', 
            'Produto', 'Marca', 'Fornecimento', 'Categoria', 'Soma de Qtd', 'Soma de Valor'
        ]
        
        # --- 2. PROCESSAMENTO ---
        mapa_fornecimento = processando_fornecimento_código(caminho_dre)
        lista_dfs_novos = []

        df_realizado = processar_dre_realizado(caminho_dre)
        if df_realizado is not None:
            lista_dfs_novos.append(df_realizado)

        arquivos_forecast = {}
        if caminho_irrestrito: arquivos_forecast[caminho_irrestrito] = "Irrestrito"
        if caminho_restrito: arquivos_forecast[caminho_restrito] = "Restrito"
        
        if arquivos_forecast:
            for caminho, tipo in arquivos_forecast.items():
                df_forecast = processar_planilha_forecast(caminho, tipo, mapa_fornecimento)
                if df_forecast is not None:
                    lista_dfs_novos.append(df_forecast)
                    # Move o arquivo de UPLOAD para a pasta final de ARQUIVADOS na rede
                    shutil.move(caminho, os.path.join(pasta_arquivados_final, os.path.basename(caminho)))

        if not lista_dfs_novos:
            return None, "Nenhum dado novo foi processado."

        # --- 3. CONSOLIDAÇÃO (Lendo o mestre, sem alterá-lo) ---
        df_novos_dados = pd.concat(lista_dfs_novos, ignore_index=True)
        bases_para_atualizar = df_novos_dados['Base'].unique()

        try:
            # Lê o arquivo mestre original, que não será alterado
            df_bd_antigo = pd.read_excel(caminho_bd_mestre, sheet_name="BD_Forecast")
            # Remove do histórico as bases que serão substituídas pelos novos dados
            df_historico_limpo = df_bd_antigo[~df_bd_antigo['Base'].isin(bases_para_atualizar)]
            # Junta o histórico limpo com os novos dados
            df_bd_final = pd.concat([df_historico_limpo, df_novos_dados], ignore_index=True)
        except (FileNotFoundError, ValueError):
            # Se o mestre não existe, o BD final são apenas os novos dados
            df_bd_final = df_novos_dados

        df_bd_final = df_bd_final.reindex(columns=ordem_colunas)
        
        # --- 4. SALVAMENTO (Criando o novo arquivo em Q:\ForecastPronto) ---
        # Salva o novo arquivo ATUALIZADO no local de saída final
        df_bd_final.to_excel(caminho_final_para_salvar, sheet_name="BD_Forecast", index=False)
        print(f" -> SUCESSO! Arquivo atualizado foi salvo em: {caminho_final_para_salvar}")

        # --- 5. PREPARAÇÃO PARA DOWNLOAD ---
        # Copia o arquivo recém-criado para a pasta 'output' do projeto web
        caminho_para_download = os.path.join(PASTA_OUTPUT_WEB, nome_arquivo_saida)
        shutil.copy(caminho_final_para_salvar, caminho_para_download)
        print(f" -> Cópia para download criada em: {caminho_para_download}")
        
        return nome_arquivo_saida, None # Retorna o nome do arquivo para o link de download
    except Exception as e:

        return None, f"Ocorreu um erro geral no processamento: {str(e)}"
