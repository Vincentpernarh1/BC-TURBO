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
        self.viajante_results = None  # Store Viajante processing results
        
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
    
    def prepare_viajante_data(self, cod_sap, cidade_destino=None, veiculo=None):
        """
        Prepara dados para integração com Viajante
        Cria arquivo no formato: Mês, COD FORNECEDOR, DESENHO, QTDE
        
        Args:
            cod_sap: Código SAP do fornecedor
            cidade_destino: Código IMS Destino (opcional)
            veiculo: Tipo de veículo selecionado (opcional)
            
        Returns:
            Status dict com informações sobre o arquivo gerado
        """
        status, dataframe = self.sap_lookup.prepare_viajante_data(cod_sap, cidade_destino, veiculo)
        
        
        
        
        # Clean NaN values before returning (NaN is not valid JSON)
        return clean_nan_values(status)
    
    def run_viajante(self, cod_sap):
        """
        Executa o processamento Viajante em modo headless (sem GUI)
        Usa os dados e parâmetros preparados por prepare_viajante_data()
        
        Args:
            cod_sap: Código SAP do fornecedor
            
        Returns:
            Dict com resultados do Viajante (Volume_por_rota.xlsx)
        """
        try:
            # Get stored parameters and demanda DataFrame
            params = self.sap_lookup.get_viajante_parameters()
            
            demanda_df = params.get('demanda_data')
            cidade_destino = params.get('cidade_destino')
            veiculo = params.get('veiculo')
            
            # Validate parameters
            if demanda_df is None or demanda_df.empty:
                return {
                    "status": "error",
                    "message": "Nenhum dado de demanda disponível. Execute prepare_viajante_data primeiro."
                }
            
            if not cidade_destino:
                return {
                    "status": "error",
                    "message": "Cidade destino não informada."
                }
            
            if not veiculo:
                return {
                    "status": "error",
                    "message": "Veículo não informado."
                }
            
            # Import Viajante headless function
            import sys
            from pathlib import Path
            
            # Get Viajante folder path
            current_path = Path(__file__).parent
            viajante_path = current_path / "Viajante"
            
            if not viajante_path.exists():
                return {
                    "status": "error",
                    "message": f"Pasta Viajante não encontrada: {viajante_path}"
                }
            
            # Add to path if not already there
            viajante_str = str(viajante_path)
            if viajante_str not in sys.path:
                sys.path.insert(0, viajante_str)
            
            # Change to Viajante directory (needed for relative file paths in DB.py)
            original_cwd = os.getcwd()
            os.chdir(viajante_path)
            
            try:
                # Import and run headless function
                from DB import run_viajante_headless  # type: ignore
                
                results = run_viajante_headless(
                    demanda_df=demanda_df,
                    cod_sap=cod_sap,
                    cod_destino=cidade_destino,
                    veiculo=veiculo,
                    caminho_BD='BD'
                )
                
                # Store Viajante results for trip calculation
                if results.get('status') == 'success':
                    self.viajante_results = results
                    print(f"\n{'='*60}")
                    print("✅ VIAJANTE RESULTS STORED IN self.viajante_results")
                    print(f"{'='*60}")
                    print(f"  Results count: {len(results.get('results', []))} rows")
                    print(f"  Status: {results.get('status')}")
                    print(f"  self.viajante_results is now: {'SET' if self.viajante_results else 'None'}")
                    print(f"{'='*60}\n")
                else:
                    print(f"\n⚠️ Viajante status was NOT success: {results.get('status')}")
                
                # Clean NaN values for JSON
                return clean_nan_values(results)
                
            finally:
                # Restore original working directory
                os.chdir(original_cwd)
                
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Error in run_viajante:\n{error_details}")
            return {
                "status": "error",
                "message": f"Erro ao executar Viajante: {str(e)}"
            }
    
    def _count_tdc_activations(self, cod_sap, origem, destino, veiculo, fluxo, trip):
        """
        Conta ativações únicas por mês no TDC (para fluxos que não são Milk Run/Line Haul)
        
        Filters:
        - Codigo IMS - Origem contains COD SAP/IMS
        - Codigo IMS Destino = destino
        - Veiculo = veiculo (vehicle type, not CrossDock!)
        - Fluxo Viagem = fluxo
        - Trip = trip (if provided)
        
        Group by Mês, remove duplicate Ativacao, count rows per month
        
        Returns:
            Dict with monthly counts {'Jan': 5, 'Fev': 7, ...}
        """
        try:
            tdc_data = self.sap_lookup.tdc_data
            
            if tdc_data is None or tdc_data.empty:
                print("Warning: No TDC data available for activation counting")
                return None
            
            # print(f"\n{'='*60}")
            # print("TDC ACTIVATION COUNTING (AS IS)")
            # print(f"{'='*60}")
            # print(f"Initial TDC rows: {len(tdc_data)}")
            # print(f"\n🔍 Filter Parameters:")
            # print(f"  cod_sap (input): '{cod_sap}'")
            # print(f"  destino: '{destino}'")
            # print(f"  veiculo: '{veiculo}'")
            # print(f"  fluxo: '{fluxo}'")
            # print(f"  trip: '{trip}'" + (" (empty - will skip)" if not trip else ""))
            # print(f"\n📋 TDC Columns available: {list(tdc_data.columns)}")
            
            # Apply filters
            filtered = tdc_data.copy()
            
            # Resolve SAP code → IMS code (TDC stores IMS codes, not SAP codes)
            cod_sap_str = str(cod_sap).strip().replace('.0', '')
            ims_code = cod_sap_str  # default: assume already IMS
            
            if len(cod_sap_str) > 6:
                # It's a SAP code — look up its COD IMS from the PFEP cache
                cache_key = f"COD SAP_{cod_sap_str}"
                if cache_key in self.sap_lookup.sap_cache:
                    cached_ims = self.sap_lookup.sap_cache[cache_key].get('cod_ims_for_tdc')
                    if cached_ims:
                        ims_code = str(cached_ims).strip()
                        print(f"  Resolved SAP {cod_sap_str} → IMS {ims_code} for TDC filter")
                    else:
                        print(f"  ⚠️ SAP {cod_sap_str} in cache but no IMS code — using SAP as fallback")
                else:
                    print(f"  ⚠️ SAP {cod_sap_str} not in cache — using SAP code as fallback")
            
            print(f"\n🔍 Applying filters step by step...\n")
            
            # Filter 1: Codigo IMS - Origem = IMS code
            if 'Codigo IMS - Origem' in filtered.columns:
                print(f"Filter 1: Codigo IMS - Origem = '{ims_code}'")
                print(f"  Sample values in column: {filtered['Codigo IMS - Origem'].head(10).tolist()}")
                
                filtered = filtered[
                    filtered['Codigo IMS - Origem'].astype(str).str.strip() == ims_code
                ]
                print(f"  ✓ After filter: {len(filtered)} rows")
            else:
                print(f"  ⚠️ Column 'Codigo IMS - Origem' NOT found!")
            
            if filtered.empty:
                print(f"\n❌ No data after Codigo IMS - Origem filter!")
                months = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
                return {month: 0 for month in months}
            
            # Filter 2: Codigo IMS Destino = destino  
            if destino and 'Codigo IMS Destino' in filtered.columns:
                print(f"\nFilter 2: Codigo IMS Destino contains '{destino}'")
                print(f"  Sample values in column: {filtered['Codigo IMS Destino'].unique()[:10].tolist()}")
                
                filtered = filtered[
                    filtered['Codigo IMS Destino'].astype(str).str.contains(str(destino), case=False, na=False)
                ]
                print(f"  ✓ After filter: {len(filtered)} rows")
            elif destino:
                print(f"\n  ⚠️ Column 'Codigo IMS Destino' NOT found!")
            
            if filtered.empty:
                print(f"\n❌ No data after Codigo IMS Destino filter!")
                months = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
                return {month: 0 for month in months}
            
            # Filter 3: Veiculo = veiculo (CHANGED FROM CrossDock!)
            if veiculo and 'Veiculo' in filtered.columns:
                print(f"\nFilter 3: Veiculo = '{veiculo}'")
                print(f"  Sample values in column: {filtered['Veiculo'].unique()[:20].tolist()}")
                
                # Try exact match
                filtered_exact = filtered[
                    filtered['Veiculo'].astype(str).str.strip().str.upper() == str(veiculo).strip().upper()
                ]
                
                if len(filtered_exact) > 0:
                    filtered = filtered_exact
                    print(f"  ✓ After exact match: {len(filtered)} rows")
                else:
                    # Try contains match
                    print(f"  ⚠️ No exact matches, trying contains...")
                    filtered = filtered[
                        filtered['Veiculo'].astype(str).str.contains(str(veiculo), case=False, na=False)
                    ]
                    print(f"  ✓ After contains match: {len(filtered)} rows")
            elif veiculo:
                print(f"\n  ⚠️ Column 'Veiculo' NOT found!")
            
            if filtered.empty:
                print(f"\n❌ No data after Veiculo filter!")
                months = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
                return {month: 0 for month in months}
            
            # Filter 4: Fluxo Viagem = fluxo
            if fluxo and 'Fluxo Viagem' in filtered.columns:
                print(f"\nFilter 4: Fluxo Viagem = '{fluxo}'")
                print(f"  Sample values in column: {filtered['Fluxo Viagem'].unique()[:10].tolist()}")
                
                # Try exact match
                filtered_exact = filtered[
                    filtered['Fluxo Viagem'].astype(str).str.strip().str.upper() == str(fluxo).strip().upper()
                ]
                
                if len(filtered_exact) > 0:
                    filtered = filtered_exact
                    print(f"  ✓ After exact match: {len(filtered)} rows")
                else:
                    # Try contains match
                    print(f"  ⚠️ No exact matches, trying contains...")
                    filtered = filtered[
                        filtered['Fluxo Viagem'].astype(str).str.contains(str(fluxo), case=False, na=False)
                    ]
                    print(f"  ✓ After contains match: {len(filtered)} rows")
            elif fluxo:
                print(f"\n  ⚠️ Column 'Fluxo Viagem' NOT found!")
            
            if filtered.empty:
                print(f"\n❌ No data after Fluxo Viagem filter!")
                months = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
                return {month: 0 for month in months}
            
            # Filter 5: Trip = trip (optional - skip if empty)
            if trip and trip.strip() and 'Trip' in filtered.columns:
                print(f"\nFilter 5: Trip = '{trip}'")
                print(f"  Sample values in column: {filtered['Trip'].unique()[:10].tolist()}")
                
                # Try exact match
                filtered_exact = filtered[
                    filtered['Trip'].astype(str).str.strip().str.upper() == str(trip).strip().upper()
                ]
                
                if len(filtered_exact) > 0:
                    filtered = filtered_exact
                    print(f"  ✓ After exact match: {len(filtered)} rows")
                else:
                    # Try contains match
                    print(f"  ⚠️ No exact matches, trying contains...")
                    filtered = filtered[
                        filtered['Trip'].astype(str).str.contains(str(trip), case=False, na=False)
                    ]
                    print(f"  ✓ After contains match: {len(filtered)} rows")
            elif trip and trip.strip():
                print(f"\n  ⚠️ Column 'Trip' NOT found!")
            else:
                print(f"\n  ℹ️ Trip filter skipped (empty value)")
            
            if filtered.empty:
                print(f"\n❌ WARNING: No TDC data matches ALL filters combined!")
                print(f"   This could mean:")
                print(f"   - Some filter values don't match TDC data")
                print(f"   - Column names are different")
                print(f"   - Data for this combination doesn't exist in TDC")
                print(f"\n   Returning zero counts for all months")
                
                # Return zero counts instead of None
                months = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 
                         'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
                return {month: 0 for month in months}
            
            print(f"\n✅ Final filtered data: {len(filtered)} rows")
            
            # Group by Mês, remove duplicate Ativacao, count per month
            if 'Mês' not in filtered.columns or 'Ativacao' not in filtered.columns:
                print("Warning: Required columns (Mês, Ativacao) not found in TDC data")
                print(f"Available columns: {list(filtered.columns)}")
                return None
            
            # Map month names to Portuguese abbreviations
            month_mapping = {
                'JANEIRO': 'Jan', 'JAN': 'Jan',
                'FEVEREIRO': 'Fev', 'FEV': 'Fev',
                'MARÇO': 'Mar', 'MAR': 'Mar',
                'ABRIL': 'Abr', 'ABR': 'Abr',
                'MAIO': 'Mai', 'MAI': 'Mai',
                'JUNHO': 'Jun', 'JUN': 'Jun',
                'JULHO': 'Jul', 'JUL': 'Jul',
                'AGOSTO': 'Ago', 'AGO': 'Ago',
                'SETEMBRO': 'Set', 'SET': 'Set',
                'OUTUBRO': 'Out', 'OUT': 'Out',
                'NOVEMBRO': 'Nov', 'NOV': 'Nov',
                'DEZEMBRO': 'Dez', 'DEZ': 'Dez'
            }
            
            monthly_counts = {}
            months = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 
                     'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
            
            # Initialize all months with 0
            for month in months:
                monthly_counts[month] = 0
            
            # Group by month and count unique activations
            # print(f"\nProcessing TDC data by month:")
            for month_name, group in filtered.groupby('Mês'):
                # Normalize month name
                month_upper = str(month_name).strip().upper()
                month_abbr = month_mapping.get(month_upper, month_upper)
                
                # print(f"  Raw month name: '{month_name}' -> Upper: '{month_upper}' -> Abbr: '{month_abbr}'")
                
                # Count unique Ativacao values in this month
                unique_activations = group['Ativacao'].nunique()
                
                if month_abbr in monthly_counts:
                    monthly_counts[month_abbr] = int(unique_activations)
                    print(f"  ✓ Mapped {month_abbr}: {unique_activations} unique activations")
                else:
                    print(f"  ⚠️ Month '{month_abbr}' not in expected months list!")
            
            # print(f"\n📊 FINAL TDC ACTIVATION COUNTS:")
            for month in months:
                print(f"  {month}: {monthly_counts[month]} activations")
            
            print(f"{'='*60}\n")
            
            return monthly_counts
            
        except Exception as e:
            print(f"Error counting TDC activations: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def _normalize_veiculo(self, veiculo):
        """Normalizes vehicle names so TDC and Tarifa tables can be matched"""
        if not veiculo:
            return ''
        v = str(veiculo).strip().upper()
        if 'BITREM' in v:             return 'BITREM'
        if 'VANDERLEIA' in v:         return 'VANDERLEIA'
        if 'CARRETA' in v:            return 'CARRETA'
        if 'VAN' in v or 'DUCATO' in v: return 'VAN'
        if '3/4' in v or '0.75' in v: return '3/4'
        if 'TOCO' in v:               return 'TOCO'
        if 'TRUCK' in v:              return 'TRUCK'
        if 'FIORINO' in v:            return 'FIORINO'
        return v

    def _calculate_weekly_trips(self, qme_results, viajante_results, fluxo='', cod_sap='', origem='', destino='', veiculo='', trip='', km=None):
        """
        Calcula quantidade de viagens semanais (TO BE e AS IS) e frete
        
        TO BE Formula: Volume m³ TO BE / CAP. ÚTIL (m³)
        
        AS IS Logic:
        - If FLUXO contains "Milk Run" or "Line Haul": AS IS Volume / CAP. ÚTIL
        - Otherwise: Count unique TDC Ativacao per month
        
        Args:
            qme_results: Resultados do QME com monthly volumes
            viajante_results: Resultados do Viajante com CAP. ÚTIL (m³) por mês
            fluxo: Tipo de fluxo (para determinar método de cálculo AS IS)
            cod_sap: Código SAP/IMS (para filtrar TDC)
            origem: Cidade origem (para filtrar TDC)
            destino: Cidade destino (para filtrar TDC)
            veiculo: Veículo selecionado (para filtrar TDC)
            trip: Trip selecionado (para filtrar TDC)
            km: Distância em KM (extraída do TDC)
            
        Returns:
            Dict com quantidade de viagens por mês (AS IS e TO BE) e dados de frete
        """
        try:
            # Extract monthly volumes from QME
            summary = qme_results.get('summary', {})
            monthly_m3_tobe = summary.get('monthly_m3_tobe', {})
            monthly_m3_asis = summary.get('monthly_m3_asis', {})
            
            if not monthly_m3_tobe:
                print("Warning: No monthly TO BE data available for trip calculation")
                return None
            
            # Extract CAP. ÚTIL from Viajante results
            viajante_data = viajante_results.get('results', [])
            
            if not viajante_data:
                print("Warning: No Viajante results available for trip calculation")
                return None
            
            # Create a mapping of month to CAP. ÚTIL (m³)
            # Viajante returns data per route/month with English abbreviations
            month_capacity = {}
            
            # Map English month abbreviations (from Viajante) to Portuguese (from QME)
            english_to_portuguese = {
                'Jan': 'Jan',
                'Feb': 'Fev',
                'Mar': 'Mar',
                'Apr': 'Abr',
                'May': 'Mai',
                'Jun': 'Jun',
                'Jul': 'Jul',
                'Aug': 'Ago',
                'Sep': 'Set',
                'Oct': 'Out',
                'Nov': 'Nov',
                'Dec': 'Dez'
            }
            
            for row in viajante_data:
                mes_english = row.get('Mês', '')  # English abbreviation from Viajante
                mes_portuguese = english_to_portuguese.get(mes_english, mes_english)
                cap_util = row.get('CAP. ÚTIL (m³)', 0)
                
                # Store the first capacity found for each month (should be consistent)
                if mes_portuguese and mes_portuguese not in month_capacity and cap_util:
                    month_capacity[mes_portuguese] = cap_util
            
            # Calculate TO BE trips for each month
            monthly_trips_tobe = {}
            months = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 
                     'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
            
            for month in months:
                volume_tobe = monthly_m3_tobe.get(month, 0)
                capacity = month_capacity.get(month, 0)
                
                if capacity > 0 and volume_tobe > 0:
                    trips = round(volume_tobe / capacity, 0)
                    monthly_trips_tobe[month] = int(trips)
                else:
                    monthly_trips_tobe[month] = 0
            
            # Calculate AS IS trips based on FLUXO type
            monthly_trips_asis = {}
            fluxo_lower = str(fluxo).lower()
            
            if 'milk run' in fluxo_lower or 'line haul' in fluxo_lower:
                # Method 1: Calculate AS IS trips = AS IS Volume / CAP. ÚTIL
                print(f"\n{'='*60}")
                print(f"AS IS TRIPS CALCULATION - MILK RUN/LINE HAUL")
                print(f"{'='*60}")
                
                for month in months:
                    volume_asis = monthly_m3_asis.get(month, 0)
                    capacity = month_capacity.get(month, 0)
                    
                    if capacity > 0 and volume_asis > 0:
                        trips = round(volume_asis / capacity, 0)
                        monthly_trips_asis[month] = int(trips)
                    else:
                        monthly_trips_asis[month] = 0
                    
                    print(f"  {month}: {volume_asis:.2f} m³ / {capacity:.2f} m³ = {monthly_trips_asis[month]} viagens")
                
                print(f"{'='*60}\n")
            else:
                # Method 2: Count unique TDC activations and convert to weekly
                print(f"\n{'='*60}")
                print(f"AS IS TRIPS CALCULATION - TDC ACTIVATION COUNT")
                print(f"{'='*60}")
                
                tdc_counts = self._count_tdc_activations(cod_sap, origem, destino, veiculo, fluxo, trip)
                
                if tdc_counts:
                    # Convert monthly activations to weekly quantities by dividing by 4.4
                    print(f"\n📊 Converting monthly activations to weekly quantities (÷ 4.4):")
                    monthly_trips_asis = {}
                    for month in months:
                        monthly_count = tdc_counts.get(month, 0)
                        weekly_count = int(round(monthly_count / 4.4)) if monthly_count > 0 else 0
                        monthly_trips_asis[month] = weekly_count
                        print(f"  {month}: {monthly_count} monthly activations → {weekly_count} weekly trips")
                else:
                    # Fallback to 0 if no TDC data
                    monthly_trips_asis = {month: 0 for month in months}
                
                print(f"{'='*60}\n")
            
            # Print TO BE trips summary
            print(f"\n{'='*60}")
            print("TO BE TRIPS CALCULATION")
            print(f"{'='*60}")
            for month in months:
                volume = monthly_m3_tobe.get(month, 0)
                capacity = month_capacity.get(month, 0)
                trips = monthly_trips_tobe.get(month, 0)
                print(f"  {month}: {volume:.2f} m³ / {capacity:.2f} m³ = {trips} viagens")
            print(f"{'='*60}\n")
            
            # Print final summary of what's being returned
            print(f"\n{'='*60}")
            print("📦 FINAL TRIP DATA BEING RETURNED TO FRONTEND")
            print(f"{'='*60}")
            print(f"FLUXO Type: '{fluxo}' (Milk Run/Line Haul: {('milk run' in fluxo_lower or 'line haul' in fluxo_lower)})")
            print(f"\n📊 AS IS Trips by Month:")
            for month in months:
                print(f"  {month}: {monthly_trips_asis.get(month, 0)}")
            print(f"\n📊 TO BE Trips by Month:")
            for month in months:
                print(f"  {month}: {monthly_trips_tobe.get(month, 0)}")
            print(f"\n📊 Month Capacities (CAP. ÚTIL):")
            for month in months:
                print(f"  {month}: {month_capacity.get(month, 0):.2f} m³")
            print(f"{'='*60}\n")
            
            # =====================================================
            # FREIGHT (TARIFA) CALCULATION
            # =====================================================
            freight_result = None
            normalized_veiculo = self._normalize_veiculo(veiculo)
            
            print(f"\n{'='*60}")
            print("FREIGHT CALCULATION (TARIFA)")
            print(f"{'='*60}")
            print(f"  Fluxo (from form):           '{fluxo}'")
            print(f"  Veiculo (from form):          '{veiculo}'")
            print(f"  Veiculo (normalized):         '{normalized_veiculo}'")
            print(f"  Origem (from form):           '{origem}'")
            print(f"  Destino (from form):          '{destino}'")
            print(f"  KM (from TDC):                {km}")
            print(f"  Trip (from form):             '{trip}'")
            print(f"  Available fluxos in Tarifa:   {self.sap_lookup.get_available_fluxos()}")
            
            if km and km > 0:
                # Find the best matching fluxo name in Tarifa data
                available_fluxos = self.sap_lookup.get_available_fluxos()
                matched_fluxo = None
                
                # Try to match the TDC fluxo value against Tarifa folder names
                for tf in available_fluxos:
                    if str(fluxo).lower() in tf.lower() or tf.lower() in str(fluxo).lower():
                        matched_fluxo = tf
                        break
                
                print(f"  Matched Tarifa fluxo folder: '{matched_fluxo}'")
                
                if matched_fluxo:
                    # Determine trip type for Tarifa (RT/OW)
                    trip_upper = str(trip).upper()
                    if 'ROUND' in trip_upper or trip_upper == 'RT':
                        viagem_code = 'RT'
                    elif 'ONE WAY' in trip_upper or trip_upper == 'OW':
                        viagem_code = 'OW'
                    else:
                        viagem_code = None  # Let tarifa_manager match what it finds
                    
                    print(f"  Viagem code for Tarifa:      '{viagem_code}'")
                    
                    freight_result = self.sap_lookup.calculate_tariff(
                        fluxo_name=matched_fluxo,
                        origem=origem,
                        destino=destino,
                        veiculo=normalized_veiculo,
                        km_value=km,
                        viagem=viagem_code
                    )
                    
                    print(f"\n  Tarifa result status: '{freight_result.get('status')}'")
                    if freight_result.get('status') == 'success':
                        print(f"  Tarifa original:      R$ {freight_result.get('tarifa_original', 0):.2f}")
                        print(f"  Tarifa real (@ {km}km): R$ {freight_result.get('tarifa_real', 0):.2f}")
                        print(f"  Transportadora:       '{freight_result.get('transportadora')}'")
                        print(f"  Veiculo (Tarifa):     '{freight_result.get('veiculo')}'")
                        print(f"  Outras opções:        {len(freight_result.get('outras_opcoes', []))}")
                    else:
                        print(f"  Message: {freight_result.get('message', 'N/A')}")
                else:
                    print(f"  ⚠️  No Tarifa fluxo matched for '{fluxo}'")
                    print(f"       Available: {available_fluxos}")
                    freight_result = {'status': 'not_found', 'message': f"Fluxo '{fluxo}' not found in Tarifa data"}
            else:
                print(f"  ⚠️  KM not available — skipping freight calculation")
                freight_result = {'status': 'no_km', 'message': 'KM not available from TDC lookup'}
            
            print(f"{'='*60}\n")
            
            # Compute monthly freight costs (tarifa_real × trips per month)
            if freight_result and freight_result.get('status') == 'success':
                tarifa_real = freight_result.get('tarifa_real', 0)
                freight_result['monthly_freight_asis'] = {
                    m: round((monthly_trips_asis.get(m, 0) or 0) * tarifa_real, 2) for m in months
                }
                freight_result['monthly_freight_tobe'] = {
                    m: round((monthly_trips_tobe.get(m, 0) or 0) * tarifa_real, 2) for m in months
                }
                freight_result['monthly_freight_savings'] = {
                    m: round(freight_result['monthly_freight_asis'][m] - freight_result['monthly_freight_tobe'][m], 2)
                    for m in months
                }
            
            return {
                'monthly_trips_tobe': monthly_trips_tobe,
                'monthly_trips_asis': monthly_trips_asis,
                'month_capacity': month_capacity,
                'freight': freight_result
            }
            
        except Exception as e:
            print(f"Error calculating weekly trips: {e}")
            import traceback
            traceback.print_exc()
            return None

    def calculate_qme(self, data):
        """Calcula QME usando o módulo QMECalculator"""
        
        print(f"\n{'='*60}")
        print("🔵 CALCULATE_QME CALLED")
        print(f"{'='*60}")
        print(f"  self.viajante_results is: {'SET' if self.viajante_results else 'None'}")
        if self.viajante_results:
            print(f"  Viajante results count: {len(self.viajante_results.get('results', []))} rows")
        print(f"{'='*60}\n")
        
        # Obtém o DataFrame completo de PFEP para filtrar por PNs do Astobe 
        pfep_data = self.sap_lookup.get_pfep_data()
        
        # Extrai o cod_sap dos dados para usar no cache lookup
        cod_sap = data.get('cod_sap', '')
        
        # Obtém o DataFrame FILTRADO de NPRC para este SAP code específico
        # Isso garante que usamos apenas os dados NPRC relevantes para o SAP selecionado
        nprc_data = self.sap_lookup.get_cached_nprc_data(cod_sap)
        
        # Se não houver cache, usa o completo como fallback
        if nprc_data is None:
            nprc_data = self.sap_lookup.get_nprc_data()
            print(f"WARNING: Using full NPRC database (no cached filter available for {cod_sap})")
        else:
            print(f"Using cached NPRC data for {cod_sap}: {len(nprc_data)} rows filtered by SAP lookup")
        
        # Obtém o DataFrame completo de MDR para lookup de volumes
        mdr_data = self.sap_lookup.get_mdr_data()
        
        # Passa tanto os dados do formulário quanto os dados PFEP, NPRC e MDR completos para o calculador
        result = self.qme_calculator.calculate(data, pfep_data, nprc_data, mdr_data)
        
        # Calculate weekly trips if Viajante results are available
        if result.get('status') == 'success' and self.viajante_results:
            # Extract user input parameters for TDC filtering
            cod_sap = data.get('cod_sap', '')
            origem = data.get('origem', '')
            destino = data.get('destino', '')
            veiculo = data.get('veiculo', '')
            fluxo = data.get('fluxo', '')
            trip = data.get('trip', '')
            
            # Extract KM from last TDC lookup result (KM is a TDC column)
            last_lookup = self.sap_lookup.get_last_lookup_result() or {}
            km = last_lookup.get('KM', None)
            if km is not None:
                try:
                    km = float(km)
                except (ValueError, TypeError):
                    km = None
            
            print(f"\n{'='*60}")
            print("CALLING _calculate_weekly_trips WITH PARAMETERS:")
            print(f"{'='*60}")
            print(f"  cod_sap: '{cod_sap}'")
            print(f"  origem: '{origem}'")
            print(f"  destino: '{destino}'")
            print(f"  veiculo: '{veiculo}'")
            print(f"  fluxo: '{fluxo}'")
            print(f"  trip: '{trip}'")
            print(f"  km (from TDC): '{km}'")
            print(f"{'='*60}\n")
            
            trip_data = self._calculate_weekly_trips(
                result, 
                self.viajante_results,
                fluxo=fluxo,
                cod_sap=cod_sap,
                origem=origem,
                destino=destino,
                veiculo=veiculo,
                trip=trip,
                km=km
            )
            if trip_data:
                result['weekly_trips'] = trip_data
                print(f"\n✅ Trip data added to result: {list(trip_data.keys())}")
            else:
                print(f"\n⚠️ No trip data returned from calculation")
        
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
            
            # Get weekly trips data if available
            weekly_trips = results.get('weekly_trips', {})
            monthly_trips_tobe = weekly_trips.get('monthly_trips_tobe', {})
            monthly_trips_asis = weekly_trips.get('monthly_trips_asis', {})
            
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
            
            # Qtde de viagens row - now with calculated AS IS and TO BE values
            viagens_row = {'Métrica': f'Qtde de Viagens Semanal'}
            for month in months:
                # AS IS trips (from TDC or calculation)
                trips_asis = monthly_trips_asis.get(month, 0) if monthly_trips_asis else 0
                viagens_row[f'{month} AS IS'] = trips_asis if trips_asis > 0 else '-'
                
                # TO BE trips (calculated)
                trips_tobe = monthly_trips_tobe.get(month, 0) if monthly_trips_tobe else 0
                viagens_row[f'{month} TO BE'] = trips_tobe if trips_tobe > 0 else '-'
            
            # Calculate total trips (sum of all months)
            if monthly_trips_asis:
                total_trips_asis = sum(v for v in monthly_trips_asis.values() if isinstance(v, (int, float)))
                viagens_row['Total AS IS'] = int(total_trips_asis) if total_trips_asis > 0 else '-'
            else:
                viagens_row['Total AS IS'] = '-'
            
            if monthly_trips_tobe:
                total_trips_tobe = sum(v for v in monthly_trips_tobe.values() if isinstance(v, (int, float)))
                viagens_row['Total TO BE'] = int(total_trips_tobe) if total_trips_tobe > 0 else '-'
            else:
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
