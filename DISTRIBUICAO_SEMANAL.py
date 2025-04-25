import pandas as pd
from datetime import datetime
import os
import glob

def carregar_dados(caminho_arquivo):
    """Carrega o arquivo Excel em um DataFrame."""
    try:
        df = pd.read_excel(caminho_arquivo)
        return df
    except FileNotFoundError:
        print(f"Erro: Arquivo não encontrado em {caminho_arquivo}")
        return None
    except Exception as e:
        print(f"Erro inesperado ao ler o arquivo {caminho_arquivo}: {e}")
        return None

def filtrar_e_salvar(df, filtro_tarefa, filtro_login, nome_responsavel, caminho_saida, nome_arquivo_sufixo=""):
    """
    Filtra e salva os dados com base na tarefa e, se informado, no login.
    Garante que a coluna 'Responsavel' esteja presente para uniformizar os DataFrames.
    """
    if filtro_login:
        df_filtrado = df[(df['TAREFA PENDENTE'] == filtro_tarefa) &
                         (df['LOGIN RETORNO EXIGÊNCIA'].str.contains(filtro_login, case=False))]
    else:
        df_filtrado = df[df['TAREFA PENDENTE'] == filtro_tarefa]

    # Se a coluna 'Responsavel' não estiver presente, cria-a com valor padrão vazio.
    if 'Responsavel' not in df_filtrado.columns:
        df_filtrado['Responsavel'] = ""

    if not df_filtrado.empty:
        data_atual = datetime.now().strftime("%Y%m%d")
        nome_arquivo = f"{nome_responsavel}_{data_atual}{nome_arquivo_sufixo}.xlsx"
        caminho_completo = os.path.join(caminho_saida, nome_arquivo)
        df_filtrado.to_excel(caminho_completo, index=False)
        print(f"Arquivo {nome_arquivo} gerado com sucesso.")
    else:
        print(f"Nenhum dado encontrado para {nome_responsavel}.")

def distribuir_restante(df, responsaveis, caminho_saida):
    """
    Distribui os protocolos restantes (da Conferência Técnica Preliminar que não foram atribuidos pelo login)
    entre os responsáveis e os combina com os protocolos atribuídos diretamente.
    Aqui garantimos que ambos os conjuntos (direto e distribuído) tenham a coluna 'Responsavel'.
    """
    # Seleciona os protocolos da Conferência Técnica Preliminar
    df_conferencia = df[df['TAREFA PENDENTE'] == 'Conferência Técnica Preliminar'].copy()
    df_conferencia['LOGIN RETORNO EXIGÊNCIA'] = df_conferencia['LOGIN RETORNO EXIGÊNCIA'].fillna('')

    # Garante a coluna 'Responsavel' em df_conferencia
    if 'Responsavel' not in df_conferencia.columns:
        df_conferencia['Responsavel'] = ''

    # Filtra os protocolos que não foram atribuídos (não contêm nenhum dos logins)
    df_restante = df_conferencia[~df_conferencia['LOGIN RETORNO EXIGÊNCIA']
        .str.contains('EDUARSOUZA|ROSSILVA|VSILVEIRA|KAMARQUES|PFREGOLON', case=False)]

    if not df_restante.empty:
        num_responsaveis_15 = 3
        num_responsaveis_11 = 2
        proporcao_15 = 15 / (num_responsaveis_15 * 15 + num_responsaveis_11 * 11)
        proporcao_11 = 11 / (num_responsaveis_15 * 15 + num_responsaveis_11 * 11)

        grupos_15 = responsaveis[:num_responsaveis_15]

        # Cria a coluna 'Responsavel' para df_restante (caso não exista) e distribui os protocolos restantes
        df_restante['Responsavel'] = ''
        for responsavel in responsaveis:
            if responsavel in grupos_15:
                num_registros = int(len(df_restante) * proporcao_15)
            else:
                num_registros = int(len(df_restante) * proporcao_11)

            df_temp = df_restante[df_restante['Responsavel'] == ''].head(num_registros)
            df_restante.loc[df_temp.index, 'Responsavel'] = responsavel

        # Para cada responsável, junta os protocolos atribuídos diretamente com os distribuídos
        for responsavel in responsaveis:
            # Protocolos atribuídos diretamente a partir do login
            df_responsavel_direto = df_conferencia[df_conferencia['LOGIN RETORNO EXIGÊNCIA']
                .str.contains(responsavel.upper(), case=False)].copy()
            if 'Responsavel' not in df_responsavel_direto.columns:
                df_responsavel_direto['Responsavel'] = ''
            # Atribui explicitamente o nome do responsável
            df_responsavel_direto.loc[:, 'Responsavel'] = responsavel

            df_responsavel_restante = df_restante[df_restante['Responsavel'] == responsavel].copy()

            # Concatena os registros diretos e os distribuídos
            df_responsavel_combinado = pd.concat([df_responsavel_direto, df_responsavel_restante], ignore_index=True)
            if not df_responsavel_combinado.empty:
                filtrar_e_salvar(df_responsavel_combinado, 'Conferência Técnica Preliminar', None, responsavel, caminho_saida, "_combinado")
            else:
                print(f"Nenhum dado para {responsavel}.")
    else:
        print("Nenhum dado restante para distribuir.")

def combinar_planilhas_por_responsavel(responsaveis, caminho_saida):
    """
    Para cada responsável (ex: EDUARDO, ROSANA, etc.), procura todos os arquivos da pasta de saída cujo nome inicie com o responsável,
    independentemente do sufixo (seja o arquivo direto ou com "_combinado") e os concatena em um arquivo final.
    """
    data_atual = datetime.now().strftime("%Y%m%d")
    for responsavel in responsaveis:
        pattern = os.path.join(caminho_saida, f"{responsavel}_*.xlsx")
        arquivos = glob.glob(pattern)
        if arquivos:
            lista_dfs = []
            for arquivo in arquivos:
                try:
                    df_temp = pd.read_excel(arquivo)
                    # Se a coluna 'Responsavel' estiver ausente, adiciona-a, para garantir a uniformidade
                    if 'Responsavel' not in df_temp.columns:
                        df_temp['Responsavel'] = ''
                    lista_dfs.append(df_temp)
                except Exception as e:
                    print(f"Erro ao ler o arquivo {arquivo}: {e}")
            if lista_dfs:
                df_final = pd.concat(lista_dfs, ignore_index=True)
                # Opcional: reorganiza as colunas para padronizar (por exemplo, em ordem alfabética)
                df_final = df_final.reindex(sorted(df_final.columns), axis=1)
                nome_arquivo_final = f"{responsavel}_FINAL_{data_atual}.xlsx"
                caminho_final = os.path.join(caminho_saida, nome_arquivo_final)
                df_final.to_excel(caminho_final, index=False)
                print(f"Planilha final para {responsavel} gerada: {nome_arquivo_final}")
            else:
                print(f"Não houve dados para combinar para {responsavel}.")
        else:
            print(f"Nenhum arquivo encontrado para {responsavel}.")

def main():
    """Função principal para coordenar o processo."""
    caminho_arquivo = "/content/CONSULTA_ATENDIMENTO_APOSENTADORIA_31032025100810.xlsx"  # Arquivo de entrada
    caminho_saida = "/content/SAIDA"  # Pasta de saída
    os.makedirs(caminho_saida, exist_ok=True)

    df = carregar_dados(caminho_arquivo)
    if df is not None:
        mapeamento_nomes = {
            "EDUARSOUZA": "EDUARDO",
            "ROSSILVA": "ROSANA",
            "VSILVEIRA": "VINICIUS",
            "KAMARQUES": "KAUE",
            "PFREGOLON": "PEDRO",
        }

        # Processa dados vinculados a "Despacho / Análise Técnica (Decisão)"
        filtrar_e_salvar(df, "Despacho / Análise Técnica (Decisão)", None, "MMARIO", caminho_saida)

        # Processa dados da Conferência Técnica Preliminar para cada login definido
        tarefa_conferencia = "Conferência Técnica Preliminar"
        for login in ["EDUARSOUZA", "ROSSILVA", "VSILVEIRA", "KAMARQUES", "PFREGOLON"]:
            nome_completo = mapeamento_nomes.get(login)
            filtrar_e_salvar(df, tarefa_conferencia, login, nome_completo, caminho_saida)

        # Distribui os protocolos restantes e gera arquivos _combinado para cada responsável
        responsaveis = list(mapeamento_nomes.values())
        distribuir_restante(df, responsaveis, caminho_saida)

        # Combina todas as planilhas geradas (tanto as filtradas diretamente quanto as _combinado) para cada responsável
        combinar_planilhas_por_responsavel(responsaveis, caminho_saida)

if __name__ == "__main__":
    main()
