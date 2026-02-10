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
                    if file_path.endswith('.csv'):
                        df = pd.read_csv(file_path)
                    else:
                        df = pd.read_excel(file_path)
                    
                    qtd_linhas = len(df)
                    status = {
                        "status": "success",
                        "filename": os.path.basename(file_path),
                        "message": f"{qtd_linhas} linhas carregadas (AS IS + TO BE)."
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
