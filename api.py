import webview
import os
import pandas as pd
from modules import SAPLookup, QMECalculator, FileManager, ExportManager


class Api:
    def __init__(self):
        self.db_folder = ""
        self.result_folder = ""
        self.loading_status = ""
        self.is_loading = False
        
        # Inicializa os módulos de processamento
        self.sap_lookup = SAPLookup()
        self.qme_calculator = QMECalculator()
        self.file_manager = FileManager()
        self.export_manager = ExportManager()
    
    def _update_loading_status(self, message):
        """Atualiza o status de carregamento"""
        self.loading_status = message
        print(f"Status: {message}")

    def select_folder(self, folder_type):
        """Usa o diálogo nativo do pywebview para selecionar pastas"""
        folder_path, folder_name = self.file_manager.select_folder(folder_type)
        
        if folder_path:
            if folder_type == 'db':
                self.db_folder = folder_path
                self.is_loading = True
                print(f"Database set to: {self.db_folder}")
                
                # Carrega dados com callback de progresso
                self.sap_lookup.update_db_folder(folder_path, progress_callback=self._update_loading_status)
                
                self.is_loading = False
                self.loading_status = "Ready"
                
            elif folder_type == 'result':
                self.result_folder = folder_path
                self.export_manager.set_result_folder(folder_path)
                print(f"Result folder set to: {self.result_folder}")
            
            return folder_name
        
        return folder_name
    
    def get_loading_status(self):
        """Retorna o status atual de carregamento"""
        return {
            "is_loading": self.is_loading,
            "status": self.loading_status
        }

    def import_asis_file(self):
        """Importa o arquivo com AS IS e TO BE scenarios"""
        status, data = self.file_manager.import_asis_file()
        
        if data is not None:
            self.qme_calculator.set_asis_data(data)
        
        return status

    def lookup_sap_data(self, cod_sap, planta, cidade_origem, cidade_destino):
        """Busca dados complementares baseado no SAP e outros inputs"""
        return self.sap_lookup.lookup_data(cod_sap, planta, cidade_origem, cidade_destino)
    

    def calculate_qme(self, data):
        """Calcula QME usando o módulo QMECalculator"""
        return self.qme_calculator.calculate(data)
    
    def export_results(self, filename=None):
        """Exporta os resultados para Excel"""
        results = self.qme_calculator.get_last_results()
        return self.export_manager.export_results(results, filename)
