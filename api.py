import webview
import os
import pandas as pd
import math
from modules import SAPLookup, QMECalculator, FileManager, ExportManager


def clean_nan_values(obj):
    """Recursively replace NaN values with None for JSON serialization"""
    if isinstance(obj, dict):
        return {k: clean_nan_values(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [clean_nan_values(item) for item in obj]
    elif isinstance(obj, float) and math.isnan(obj):
        return None
    else:
        return obj


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
        # Obtém o DataFrame completo de PFEP para filtrar por PNs do Astobe
        pfep_data = self.sap_lookup.get_pfep_data()
        
        # Obtém o DataFrame FILTRADO de NPRC do último lookup SAP (não o completo)
        # Isso garante que usamos apenas os dados NPRC relevantes para o SAP selecionado
        nprc_data = self.sap_lookup.get_cached_nprc_data()
        
        # Se não houver cache, usa o completo como fallback
        if nprc_data is None:
            nprc_data = self.sap_lookup.get_nprc_data()
            print("WARNING: Using full NPRC database (no cached filter available)")
        else:
            print(f"Using cached NPRC data: {len(nprc_data)} rows filtered by SAP lookup")
        
        # Obtém o DataFrame completo de MDR para lookup de volumes
        mdr_data = self.sap_lookup.get_mdr_data()
        
        # Passa tanto os dados do formulário quanto os dados PFEP, NPRC e MDR completos para o calculador
        result = self.qme_calculator.calculate(data, pfep_data, nprc_data, mdr_data)
        
        # Clean NaN values before returning (NaN is not valid JSON)
        return clean_nan_values(result)
    
    def export_results(self, filename=None):
        """Exporta a tabela de breakdown detalhado para Excel"""
        results = self.qme_calculator.get_last_results()
        
        if not results or 'summary' not in results:
            return {
                "status": "error",
                "message": "Nenhum resultado disponível para exportar!"
            }
        
        if not self.result_folder:
            return {
                "status": "error",
                "message": "Selecione a pasta de resultados primeiro!"
            }
        
        try:
            summary = results['summary']
            veiculo = results.get('veiculo', 'VEÍCULO')
            
            # Extract monthly data
            months = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 
                     'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
            
            monthly_m3_asis = summary.get('monthly_m3_asis', {})
            monthly_m3_tobe = summary.get('monthly_m3_tobe', {})
            
            # Create breakdown table data
            breakdown_data = []
            
            # Volume m³ row
            vol_row = {'Métrica': 'Volume m³'}
            for month in months:
                vol_row[f'{month} AS IS'] = float(monthly_m3_asis.get(month, 0))
                vol_row[f'{month} TO BE'] = float(monthly_m3_tobe.get(month, 0))
            vol_row['Total AS IS'] = sum(monthly_m3_asis.values())
            vol_row['Total TO BE'] = sum(monthly_m3_tobe.values())
            breakdown_data.append(vol_row)
            
            # Qtde de viagens row (placeholder)
            viagens_row = {'Métrica': f'Qtde de Viagens Semanal'}
            for month in months:
                viagens_row[f'{month} AS IS'] = '-'
                viagens_row[f'{month} TO BE'] = '-'
            viagens_row['Total AS IS'] = '-'
            viagens_row['Total TO BE'] = '-'
            breakdown_data.append(viagens_row)
            
            # Custo de veículo semanal row (placeholder)
            custo_row = {'Métrica': f'Custo de {veiculo} Semanal'}
            for month in months:
                custo_row[f'{month} AS IS'] = '-'
                custo_row[f'{month} TO BE'] = '-'
            custo_row['Total AS IS'] = '-'
            custo_row['Total TO BE'] = '-'
            breakdown_data.append(custo_row)
            
            # Custo total row (placeholder)
            custo_total_row = {'Métrica': f'Custo total de {veiculo} Semanal'}
            for month in months:
                custo_total_row[f'{month} AS IS'] = '-'
                custo_total_row[f'{month} TO BE'] = '-'
            custo_total_row['Total AS IS'] = '-'
            custo_total_row['Total TO BE'] = '-'
            breakdown_data.append(custo_total_row)
            
            # Create DataFrame
            df = pd.DataFrame(breakdown_data)
            
            # Generate filename
            if not filename:
                from datetime import datetime
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"BC_Turbo_Breakdown_{timestamp}.xlsx"
            
            filepath = os.path.join(self.result_folder, filename)
            
            # Write to Excel with formatting
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Breakdown Detalhado')
                
                # Auto-adjust column widths
                worksheet = writer.sheets['Breakdown Detalhado']
                for idx, col in enumerate(df.columns):
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(str(col))
                    )
                    worksheet.column_dimensions[chr(65 + idx)].width = min(max_length + 2, 30)
            
            return {
                "status": "success",
                "message": f"Breakdown exportado: {filename}",
                "filepath": filepath
            }
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Export breakdown error details:\n{error_details}")
            return {
                "status": "error",
                "message": f"Erro ao exportar: {str(e)}"
            }
    
    def export_pn_table(self, filename=None):
        """Exporta a tabela de PNs detalhada para Excel"""
        results = self.qme_calculator.get_last_results()
        
        if not results or 'results' not in results:
            return {
                "status": "error",  
                "message": "Nenhum resultado disponível para exportar!"
            }
        
        if not self.result_folder:
            return {
                "status": "error",
                "message": "Selecione a pasta de resultados primeiro!"
            }
        
        try:
            # Extract the detailed PN data
            pn_data = results['results']
            
            # Convert to DataFrame directly from the results list
            # This avoids the "mixing dicts with non-Series" error
            df_data = []
            
            for idx, row in enumerate(pn_data, start=1):
                # Extract monthly volumes safely
                monthly_vols = row.get('monthly_volumes', {})
                
                # Convert all values to basic Python types to avoid pandas issues
                row_data = {
                    'Linha': int(idx),
                    'PN': str(row.get('pn', '')),
                    'Jan': float(monthly_vols.get('Jan', 0)) if monthly_vols.get('Jan') else 0,
                    'Fev': float(monthly_vols.get('Fev', 0)) if monthly_vols.get('Fev') else 0,
                    'Mar': float(monthly_vols.get('Mar', 0)) if monthly_vols.get('Mar') else 0,
                    'Abr': float(monthly_vols.get('Abr', 0)) if monthly_vols.get('Abr') else 0,
                    'Mai': float(monthly_vols.get('Mai', 0)) if monthly_vols.get('Mai') else 0,
                    'Jun': float(monthly_vols.get('Jun', 0)) if monthly_vols.get('Jun') else 0,
                    'Jul': float(monthly_vols.get('Jul', 0)) if monthly_vols.get('Jul') else 0,
                    'Ago': float(monthly_vols.get('Ago', 0)) if monthly_vols.get('Ago') else 0,
                    'Set': float(monthly_vols.get('Set', 0)) if monthly_vols.get('Set') else 0,
                    'Out': float(monthly_vols.get('Out', 0)) if monthly_vols.get('Out') else 0,
                    'Nov': float(monthly_vols.get('Nov', 0)) if monthly_vols.get('Nov') else 0,
                    'Dez': float(monthly_vols.get('Dez', 0)) if monthly_vols.get('Dez') else 0,
                    'QME AS IS': float(row.get('qme_asis', 0)) if row.get('qme_asis') else 0,
                    'MDR AS IS': str(row.get('mdr_asis', '')),
                    'Vol AS IS (m³)': float(row.get('vol_asis_m3', 0)) if row.get('vol_asis_m3') else 0,
                    'Peso AS IS (kg)': float(row.get('peso_asis_kg', 0)) if row.get('peso_asis_kg') else 0,
                    'QME TO BE': float(row.get('qme_tobe', 0)) if row.get('qme_tobe') else 0,
                    'MDR TO BE': str(row.get('mdr_tobe', '')),
                    'Vol TO BE (m³)': float(row.get('vol_tobe_m3', 0)) if row.get('vol_tobe_m3') else 0,
                    'Peso TO BE (kg)': float(row.get('peso_tobe_kg', 0)) if row.get('peso_tobe_kg') else 0,
                    'Status': str(row.get('status', ''))
                }
                df_data.append(row_data)
            
            # Create DataFrame with explicit column order
            columns = ['Linha', 'PN', 'Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 
                      'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez',
                      'QME AS IS', 'MDR AS IS', 'Vol AS IS (m³)', 'Peso AS IS (kg)',
                      'QME TO BE', 'MDR TO BE', 'Vol TO BE (m³)', 'Peso TO BE (kg)', 'Status']
            
            df = pd.DataFrame(df_data, columns=columns)
            
            # Generate filename if not provided
            if not filename:
                from datetime import datetime
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"BC_Turbo_PN_Table_{timestamp}.xlsx"
            
            filepath = os.path.join(self.result_folder, filename)
            
            # Write to Excel with formatting
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Detalhes PN')
                
                # Auto-adjust column widths
                worksheet = writer.sheets['Detalhes PN']
                for idx, col in enumerate(df.columns):
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(str(col))
                    )
                    worksheet.column_dimensions[chr(65 + idx)].width = min(max_length + 2, 50)
            
            return {
                "status": "success",
                "message": f"Tabela de PNs exportada: {filename}",
                "filepath": filepath
            }
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Export error details:\n{error_details}")
            return {
                "status": "error",
                "message": f"Erro ao exportar: {str(e)}"
            }
