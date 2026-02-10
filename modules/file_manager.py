"""
Módulo para gerenciamento de importação de arquivos
"""
import os
import pandas as pd
import webview


class FileManager:
    def __init__(self):
        pass
    
    def import_asis_file(self):
        """
        Importa o arquivo com AS IS e TO BE scenarios
        
        Returns:
            Tupla (status_dict, dataframe) onde status_dict contém informações
            sobre o resultado da operação e dataframe contém os dados (ou None)
        """
        try:
            active_window = webview.windows[0]
            
            # Abre diálogo de arquivo (filtra Excel e CSV)
            file_types = ('Arquivos de Dados (*.xlsx;*.xls;*.csv)', 'Todos os arquivos (*.*)')
            
            result = active_window.create_file_dialog(
                webview.FileDialog.OPEN,
                file_types=file_types
            )
            
            print(f"File dialog result: {result}")
            
            if result and len(result) > 0:
                file_path = result[0]
                print(f"Selected file: {file_path}")
                
                try:
                    # Lê o arquivo com AS IS e TO BE
                    # O arquivo tem uma estrutura especial com multi-level headers:
                    # Row 0: PN, AS IS, , TO BE, 
                    # Row 1: , QME, MDR, QME, MDR
                    if file_path.endswith('.csv'):
                        df_raw = pd.read_csv(file_path)
                    else:
                        df_raw = pd.read_excel(file_path)
                    
                    
                    # As primeira linha contém os nomes das colunas reais
                    # Precisamos renomear as colunas de forma apropriada
                    if len(df_raw) > 0:
                        # Verifica se a primeira linha é o header secundário
                        first_row = df_raw.iloc[0]
                        
                        # Se a primeira linha contém "QME", "MDR", é o header
                        if "QME" in str(first_row.values) or "MDR" in str(first_row.values):
                            # Pula a primeira linha (é o header real)
                            df = df_raw.iloc[1:].reset_index(drop=True)
                            
                            # Renomeia as colunas baseado na estrutura conhecida
                            # Unnamed: 0 = PN
                            # AS IS = QME AS IS
                            # Unnamed: 2 = MDR AS IS
                            # TO BE = QME TO BE
                            # Unnamed: 4 = MDR TO BE
                            new_columns = ['PN', 'AS_IS_QME', 'AS_IS_MDR', 'TO_BE_QME', 'TO_BE_MDR']
                            df.columns = new_columns[:len(df.columns)]
                        else:
                            df = df_raw.copy()
                            # Se não tem o padrão esperado, tenta renomear baseado em posição
                            if len(df.columns) >= 5:
                                new_columns = ['PN', 'AS_IS_QME', 'AS_IS_MDR', 'TO_BE_QME', 'TO_BE_MDR']
                                df.columns = new_columns[:len(df.columns)]
                    else:
                        df = df_raw.copy()
                    
                    # Remove linhas vazias
                    df = df.dropna(subset=['PN'])
                    
                    qtd_linhas = len(df)
                    
                    # Informações detalhadas sobre o arquivo
                    pn_examples = df['PN'].head(5).tolist() if 'PN' in df.columns else []
                    
                    # Calcula estatísticas para QME (somas) e MDR (valores distintos)
                    stats = {}
                    
                    # AS IS QME - Total Sum
                    if 'AS_IS_QME' in df.columns:
                        try:
                            as_is_qme_sum = pd.to_numeric(df['AS_IS_QME'], errors='coerce').sum()
                            stats['AS_IS_QME_Total'] = int(as_is_qme_sum)
                        except:
                            stats['AS_IS_QME_Total'] = 0
                    
                    # AS IS MDR - Distinct Values
                    if 'AS_IS_MDR' in df.columns:
                        distinct_mdr = df['AS_IS_MDR'].dropna().unique().tolist()
                        stats['AS_IS_MDR_Distinct'] = distinct_mdr
                    
                    # TO BE QME - Total Sum
                    if 'TO_BE_QME' in df.columns:
                        try:
                            to_be_qme_sum = pd.to_numeric(df['TO_BE_QME'], errors='coerce').sum()
                            stats['TO_BE_QME_Total'] = int(to_be_qme_sum)
                        except:
                            stats['TO_BE_QME_Total'] = 0
                    
                    # TO BE MDR - Distinct Values
                    if 'TO_BE_MDR' in df.columns:
                        distinct_mdr_tobe = df['TO_BE_MDR'].dropna().unique().tolist()
                        stats['TO_BE_MDR_Distinct'] = distinct_mdr_tobe
                    
                    status = {
                        "status": "success",
                        "filename": os.path.basename(file_path),
                        "message": f"{qtd_linhas} PNs carregados.",
                        "details": {
                            "rows": qtd_linhas,
                            "columns": list(df.columns),
                            "sample_pns": pn_examples,
                            "stats": stats
                        }
                    }
                    
                    return status, df
                    
                except Exception as e:
                    print(f"Error reading file: {str(e)}")
                    status = {
                        "status": "error",
                        "message": f"Erro ao ler arquivo: {str(e)}"
                    }
                    return status, None
            
            status = {
                "status": "cancel",
                "message": "Nenhum arquivo selecionado"
            }
            return status, None
            
        except Exception as e:
            print(f"Error in import_asis_file: {str(e)}")
            status = {
                "status": "error",
                "message": f"Erro ao abrir diálogo: {str(e)}"
            }
            return status, None
    
    def select_folder(self, folder_type):
        """
        Usa o diálogo nativo do pywebview para selecionar pastas
        
        Args:
            folder_type: Tipo de pasta (não utilizado nesta função, apenas retorna o caminho)
            
        Returns:
            Tupla (folder_path, folder_name) ou (None, "Não Selecionado")
        """
        try:
            active_window = webview.windows[0]
            
            # Abre diálogo de pasta usando nova API
            result = active_window.create_file_dialog(webview.FileDialog.FOLDER)
            
            if result and len(result) > 0:
                folder_selected = result[0]
                print(f"Folder selected: {folder_selected}")
                return folder_selected, os.path.basename(folder_selected)
            
            return None, "Não Selecionado"
            
        except Exception as e:
            print(f"Error selecting folder: {str(e)}")
            return None, "Não Selecionado"
