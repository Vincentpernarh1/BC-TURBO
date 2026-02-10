import webview
import os
import pandas as pd

class Api:
    def __init__(self):
        self.db_folder = ""
        self.result_folder = ""
        self.asis_data = None  # Para guardar os dados carregados do AS IS
        self.last_results = None  # Guardar último resultado para export
        self.sap_data = {}  # Cache de dados SAP

    def select_folder(self, folder_type):
        """Usa o diálogo nativo do pywebview para selecionar pastas"""
        active_window = webview.windows[0]
        
        # Abre diálogo de pasta usando nova API
        result = active_window.create_file_dialog(webview.FileDialog.FOLDER)
        
        if result and len(result) > 0:
            folder_selected = result[0]
            
            if folder_type == 'db':
                self.db_folder = folder_selected
                print(f"Database set to: {self.db_folder}")
            elif folder_type == 'result':
                self.result_folder = folder_selected
                print(f"Result folder set to: {self.result_folder}")
            
            return os.path.basename(folder_selected)
        
        return "Não Selecionado"

    def import_asis_file(self):
        """Importa o arquivo com AS IS e TO BE scenarios"""
        try:
            active_window = webview.windows[0]
            
            # Abre diálogo de arquivo (filtra Excel e CSV)
            file_types = ('Arquivos de Dados (*.xlsx;*.xls;*.csv)', 'Todos os arquivos (*.*)')
            
            result = active_window.create_file_dialog(webview.FileDialog.OPEN, file_types=file_types)
            
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
                    
                    self.asis_data = df
                    
                    qtd_linhas = len(df)
                    return {
                        "status": "success",
                        "filename": os.path.basename(file_path),
                        "message": f"{qtd_linhas} linhas carregadas (AS IS + TO BE)."
                    }
                except Exception as e:
                    print(f"Error reading file: {str(e)}")
                    return {"status": "error", "message": f"Erro ao ler arquivo: {str(e)}"}
            
            return {"status": "cancel", "message": "Nenhum arquivo selecionado"}
        except Exception as e:
            print(f"Error in import_asis_file: {str(e)}")
            return {"status": "error", "message": f"Erro ao abrir diálogo: {str(e)}"}

    def lookup_sap_data(self, cod_sap, planta, cidade_origem, cidade_destino):
        """Busca dados complementares baseado no SAP e outros inputs"""
        try:
            # Aqui você carregaria do database folder
            # Por enquanto, retornamos dados mock
            
            # Simula busca em arquivo da database
            if self.db_folder:
                # TODO: Carregar arquivo real da database
                pass
            
            # Mock data
            return {
                "status": "success",
                "data": {
                    "fornecedor": "FORNECEDOR XYZ LTDA",
                    "transportadora": "DHL Supply Chain",
                    "veiculo": "Truck",
                    "uf": "MG"
                }
            }
        except Exception as e:
            return {"status": "error", "message": str(e)}

    def calculate_qme(self, data):
        print("Dados recebidos:", data)
        
        if self.asis_data is None:
            return {"status": "error", "message": "Carregue o arquivo AS IS/TO BE antes de simular!"}

        # AQUI ENTRARÁ O CÁLCULO REAL USANDO O self.asis_data E OS INPUTS
        # Por enquanto, retornamos mock data para teste visual
        
        results = []
        for idx, row in self.asis_data.iterrows():
            # Mock calculation - será substituído pela lógica real
            results.append({
                "row": idx + 1,
                "pn": row.get('PN', f'PN-{idx+1}'),
                "qme_asis": row.get('QME_ASIS', 100),
                "qme_tobe": row.get('QME_TOBE', data.get('qme_tobe', 150)),
                "vol_asis": row.get('VOL_ASIS', 10),
                "vol_tobe": 0,  # Será calculado
                "savings": 0,  # Será calculado
                "status": "OK"
            })
        
        response = {
            "status": "success",
            "message": f"Simulação concluída para {len(results)} linhas.",
            "results": results,
            "summary": {
                "total_rows": len(results),
                "total_savings": sum(r['savings'] for r in results)
            }
        }
        
        self.last_results = results
        return response
    
    def export_results(self, filename=None):
        """Exporta os resultados para Excel"""
        if self.last_results is None:
            return {"status": "error", "message": "Nenhum resultado para exportar!"}
        
        if not self.result_folder:
            return {"status": "error", "message": "Selecione a pasta de resultados primeiro!"}
        
        try:
            df_results = pd.DataFrame(self.last_results)
            
            if not filename:
                from datetime import datetime
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"BC_Turbo_Results_{timestamp}.xlsx"
            
            filepath = os.path.join(self.result_folder, filename)
            df_results.to_excel(filepath, index=False, sheet_name='Resultados')
            
            return {
                "status": "success",
                "message": f"Arquivo exportado: {filename}",
                "filepath": filepath
            }
        except Exception as e:
            return {"status": "error", "message": str(e)}
