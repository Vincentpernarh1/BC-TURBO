"""
Módulo para exportação de resultados
"""
import os
import pandas as pd
from datetime import datetime


class ExportManager:
    def __init__(self):
        self.result_folder = None
    
    def set_result_folder(self, folder_path):
        """Define a pasta de destino para exportação"""
        self.result_folder = folder_path
    
    def export_results(self, results, filename=None):
        """
        Exporta os resultados para Excel
        
        Args:
            results: Lista de dicionários com os resultados
            filename: Nome do arquivo (opcional, gera timestamp se não fornecido)
            
        Returns:
            Dicionário com status da exportação
        """
        if results is None or len(results) == 0:
            return {
                "status": "error",
                "message": "Nenhum resultado para exportar!"
            }
        
        if not self.result_folder:
            return {
                "status": "error",
                "message": "Selecione a pasta de resultados primeiro!"
            }
        
        try:
            df_results = pd.DataFrame(results)
            
            if not filename:
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
            return {
                "status": "error",
                "message": str(e)
            }
    
    def get_result_folder(self):
        """Retorna a pasta de resultados configurada"""
        return self.result_folder
