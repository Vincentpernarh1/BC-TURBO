"""
Módulo para busca de dados SAP
"""
import pandas as pd
from pathlib import Path
import os

class SAPLookup:
    def __init__(self, db_folder=None):
        self.db_folder = db_folder
        self.sap_cache = {}
        self.pfep_data = None
        self.tdc_data = None
    
    def _needs_parquet_conversion(self, excel_path, parquet_path):
        """Verifica se o arquivo Excel precisa ser convertido para Parquet"""
        parquet_file = Path(parquet_path)
        excel_file = Path(excel_path)
        
        # Se Parquet não existe, precisa converter
        if not parquet_file.exists():
            return True
        
        # Se Excel foi modificado depois do Parquet, precisa reconverter
        excel_mtime = excel_file.stat().st_mtime
        parquet_mtime = parquet_file.stat().st_mtime
        
        return excel_mtime > parquet_mtime
    
    def _convert_to_parquet(self, excel_path, parquet_path, usecols, header_row, is_pfep=True):
        """Converte arquivo Excel para Parquet com limpeza de dados"""
        try:
            print(f"  Converting to Parquet for faster loading...")
            df = pd.read_excel(excel_path, usecols=usecols, engine='openpyxl', header=header_row)
            
            # Limpa os dados ANTES de salvar no Parquet
            df = self._clean_data(df, is_pfep=is_pfep)
            
            # Converte colunas restantes com tipos mistos para string
            for col in df.columns:
                if df[col].dtype == 'object':
                    try:
                        df[col] = df[col].astype(str).replace('nan', '')
                    except:
                        pass
            
            df.to_parquet(parquet_path, engine='pyarrow', compression='snappy')
            print(f"  ✓ Parquet cache created")
            return True
        except Exception as e:
            print(f"  Warning: Parquet conversion failed ({e}), will use Excel")
            return False
    
    def _clean_data(self, df, is_pfep=True):
        """Limpa e otimiza os dados do DataFrame"""
        if is_pfep:
            # Limpa e converte COD SAP e COD IMS (remove .0 e converte para string)
            if 'COD SAP' in df.columns:
                df['COD SAP'] = pd.to_numeric(df['COD SAP'], errors='coerce').fillna(0).astype('Int64').astype(str).replace('0', '').replace('<NA>', '')
            
            if 'COD IMS' in df.columns:
                df['COD IMS'] = pd.to_numeric(df['COD IMS'], errors='coerce').fillna(0).astype('Int64').astype(str).replace('0', '').replace('<NA>', '')
            
            # Converte QME para melhor tipo de dado
            if 'QME (Pecas/Embalagem)' in df.columns:
                df['QME (Pecas/Embalagem)'] = pd.to_numeric(df['QME (Pecas/Embalagem)'], errors='coerce').fillna(0).astype('Int64')
                
            if 'Estado Fornecedor' in df.columns:
                df['Estado Fornecedor'] = df['Estado Fornecedor'].astype(str).str.strip()
        else:
            # Limpa TDC - CodigoFornecedor é equivalente ao COD IMS
            if 'CodigoFornecedor' in df.columns:
                df['CodigoFornecedor'] = pd.to_numeric(df['CodigoFornecedor'], errors='coerce').fillna(0).astype('Int64').astype(str).replace('0', '').replace('<NA>', '')
                
        return df
    
    def _load_pfep_files(self):
        """Carrega arquivos PFEP FIASA ou BETIM com cache Parquet para performance"""
        if not self.db_folder:
            return None
        
        db_path = Path(self.db_folder)
        if not db_path.exists():
            return None
        
        # Define colunas específicas para ler (ajuste conforme necessário)
        pfep_columns = [
            "Part Number","Pecas por semana","COD IMS","COD SAP", "Nome Fornecedor", "Cidade Fornecedor", "Estado Fornecedor",
            "Metro Cúbico Semanal","COD Embalagem","QME (Pecas/Embalagem)"
        ]
        
        dataframes = []
        
        # Procura arquivos PFEP FIASA ou BETIM (.xlsx)
        for file in db_path.glob("*.xlsx"):
            filename = file.name.upper()
            if "PFEP" in filename and ("FIASA" in filename or "BETIM" in filename) and '~$' not in filename:
                print(f"Loading PFEP file: {file.name}")
                
                # Define caminho do arquivo Parquet
                parquet_path = file.with_suffix('.parquet')
                
                try:
                    # Verifica se precisa converter para Parquet
                    if self._needs_parquet_conversion(file, parquet_path):
                        self._convert_to_parquet(file, parquet_path, 
                                                lambda x: x in pfep_columns, header_row=9, is_pfep=True)
                    
                    # Tenta carregar do Parquet (100x mais rápido)
                    if parquet_path.exists():
                        print(f"  Loading from Parquet cache (fast mode)...")
                        df = pd.read_parquet(parquet_path, engine='pyarrow')
                    else:
                        # Fallback para Excel se Parquet falhou
                        print(f"  Loading from Excel (slow mode)...")
                        df = pd.read_excel(file, usecols=lambda x: x in pfep_columns, 
                                         engine='openpyxl', header=9)
                        df = self._clean_data(df, is_pfep=True)
                    
                    dataframes.append(df)
                    
                except Exception as e:
                    print(f"Error reading {file.name}: {e}")
        
        # Procura arquivos .xlsm
        for file in db_path.glob("*.xlsm"):
            if '~$' in file.name:
                continue
            filename = file.name.upper()
            if "PFEP" in filename and ("FIASA" in filename or "BETIM" in filename):
                print(f"Loading PFEP file: {file.name}")
                
                parquet_path = file.with_suffix('.parquet')
                
                try:
                    if self._needs_parquet_conversion(file, parquet_path):
                        self._convert_to_parquet(file, parquet_path, 
                                                lambda x: x in pfep_columns, header_row=10, is_pfep=True)
                    
                    if parquet_path.exists():
                        print(f"  Loading from Parquet cache (fast mode)...")
                        df = pd.read_parquet(parquet_path, engine='pyarrow')
                    else:
                        print(f"  Loading from Excel (slow mode)...")
                        df = pd.read_excel(file, usecols=lambda x: x in pfep_columns, 
                                         engine='openpyxl', header=10)
                        df = self._clean_data(df, is_pfep=True)
                    
                    dataframes.append(df)
                    
                except Exception as e:
                    print(f"Error reading {file.name}: {e}")
        
        if dataframes:
            self.pfep_data = pd.concat(dataframes, ignore_index=True)
            
            print(f"Loaded {len(self.pfep_data)} rows from PFEP files")
            return self.pfep_data
        
        return None
    
    def _load_tdc_files(self):
        """Carrega arquivos TDC com cache Parquet para performance"""
        if not self.db_folder:
            return None
        
        db_path = Path(self.db_folder)
        if not db_path.exists():
            return None
        
        # Define colunas específicas para TDC (ajuste conforme necessário)
        tdc_columns = [
            "CodigoFornecedor", "Transportadora", "CNPJ Origem", "Destino Materiais", 
            "Codigo de Rota", "Tipo de Fluxo", "Km Total","Veiculo a ser Utilizado"
        ]
        
        dataframes = []
        
        # Procura arquivos TDC
        for file in db_path.glob("*TDC*.xlsx"):
            if '~$' in file.name:
                continue
            print(f"Loading TDC file: {file.name}")
            
            parquet_path = file.with_suffix('.parquet')
            
            try:
                if self._needs_parquet_conversion(file, parquet_path):
                    self._convert_to_parquet(file, parquet_path, 
                                            lambda x: x in tdc_columns, header_row=0, is_pfep=False)
                
                if parquet_path.exists():
                    print(f"  Loading from Parquet cache (fast mode)...")
                    df = pd.read_parquet(parquet_path, engine='pyarrow')
                else:
                    print(f"  Loading from Excel (slow mode)...")
                    df = pd.read_excel(file, usecols=lambda x: x in tdc_columns)
                    df = self._clean_data(df, is_pfep=False)
                
                dataframes.append(df)
                
            except Exception as e:
                print(f"Error reading {file.name}: {e}")
        
        if dataframes:
            self.tdc_data = pd.concat(dataframes, ignore_index=True)
            print(f"Loaded {len(self.tdc_data)} rows from TDC files")
            return self.tdc_data
        
        return None
    
    def lookup_data(self, cod_sap, planta, cidade_origem, cidade_destino):
        """Busca dados complementares baseado no SAP e outros inputs"""
        try:
            # Limpa e converte o código de entrada (remove .0 se for float)
            if isinstance(cod_sap, (int, float)):
                cod_sap_str = str(int(cod_sap)).strip()
            else:
                cod_sap_str = str(cod_sap).strip().replace('.0', '')
            
            cod_length = len(cod_sap_str)
            
            # Determina qual coluna usar baseado no tamanho do código
            if cod_length < 7:
                # Usa COD IMS
                filter_column = "COD IMS"
            elif 6 < cod_length < 10:
                # Usa COD SAP
                filter_column = "COD SAP"
            else:
                # Código inválido
                return {
                    "status": "error",
                    "message": "Código SAP ou IMS inválido. Use IMS (menos de 7 dígitos) ou SAP (7-9 dígitos)"
                }
            
            # Verifica se os dados foram carregados (normalmente já carregados ao selecionar pasta)
            if self.pfep_data is None and self.tdc_data is None:
                return {
                    "status": "error",
                    "message": "Nenhuma base de dados carregada. Por favor, selecione uma pasta de database primeiro."
                }
            
            # Busca nos dados PFEP
            pfep_result = None
            cod_ims_for_tdc = None  # IMS code to use for TDC lookup
            
            if self.pfep_data is not None:
                # Filtra dados baseado no código correto (IMS ou SAP)
                mask = (self.pfep_data[filter_column] == cod_sap_str)
                pfep_match = self.pfep_data[mask]
                
                print(f"PFEP lookup for {filter_column}={cod_sap_str}: found {len(pfep_match)} matches")
                
                if not pfep_match.empty:
                    pfep_result = pfep_match.iloc[0].to_dict()
                    
                    # Extrai COD IMS para busca no TDC
                    if filter_column == "COD IMS":
                        # User já forneceu IMS, usa direto
                        cod_ims_for_tdc = cod_sap_str
                    elif filter_column == "COD SAP" and 'COD IMS' in pfep_result:
                        # User forneceu SAP, pega IMS do resultado PFEP
                        cod_ims_for_tdc = str(pfep_result['COD IMS']).strip()
            
            # Busca nos dados TDC usando COD IMS
            tdc_result = None
            if self.tdc_data is not None and cod_ims_for_tdc:
                # TDC usa CodigoFornecedor que corresponde ao COD IMS
                mask = (self.tdc_data['CodigoFornecedor'].astype(str).str.strip() == cod_ims_for_tdc)
                tdc_match = self.tdc_data[mask]
                
                print(f"TDC lookup for CodigoFornecedor={cod_ims_for_tdc}: found {len(tdc_match)} matches")
                
                if not tdc_match.empty:
                    tdc_result = tdc_match.iloc[0].to_dict()
            
            # Combina resultados PFEP e TDC
            if pfep_result or tdc_result:
                combined_data = {}
                if pfep_result:
                    combined_data.update(pfep_result)
                if tdc_result:
                    combined_data.update(tdc_result)
                
                return {
                    "status": "success",
                    "data": combined_data,
                    "filter_used": filter_column
                }
            
            # Nenhum dado encontrado
            return {
                "status": "not_found",
                "message": f"Nenhum dado encontrado para {filter_column}: {cod_sap_str}",
                "filter_used": filter_column
            }
        except Exception as e:
            return {"status": "error", "message": str(e)}
    
    def update_db_folder(self, db_folder, progress_callback=None):
        """Atualiza o caminho da pasta de database e carrega os dados imediatamente"""
        self.db_folder = db_folder
        # Limpa dados antigos
        self.pfep_data = None
        self.tdc_data = None
        
        # Notifica início do carregamento
        if progress_callback:
            progress_callback("Preparing to load database files...")
        
        # Carrega dados imediatamente (inclui conversão para Parquet se necessário)
        print("\n" + "="*60)
        print("📂 Loading and preparing database files...")
        print("="*60)
        
        if progress_callback:
            progress_callback("Loading PFEP files...")
        self._load_pfep_files()
        
        if progress_callback:
            progress_callback("Loading TDC files...")
        self._load_tdc_files()
        
        print("="*60)
        print("✓ Database ready! You can now perform searches.")
        print("="*60 + "\n")
        
        if progress_callback:
            progress_callback("Database ready!")
    
    def reload_data(self):
        """Recarrega os dados dos arquivos"""
        self.pfep_data = None
        self.tdc_data = None
        self._load_pfep_files()
        self._load_tdc_files()
    
    def clear_cache(self):
        """Limpa o cache de dados SAP"""
        self.sap_cache.clear()
        self.pfep_data = None
        self.tdc_data = None
