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
        self.mdr_data = None
        self.nprc_data = None
        self.last_lookup_result = None  # Store last lookup result to reuse in calculations
    
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
            # Limpa TDC - Codigo IMS - Origem e Codigo IMS Destino
            if 'Codigo IMS - Origem' in df.columns:
                df['Codigo IMS - Origem'] = pd.to_numeric(df['Codigo IMS - Origem'], errors='coerce').fillna(0).astype('Int64').astype(str).replace('0', '').replace('<NA>', '')
            
            if 'Codigo IMS Destino' in df.columns:
                df['Codigo IMS Destino'] = pd.to_numeric(df['Codigo IMS Destino'], errors='coerce').fillna(0).astype('Int64').astype(str).replace('0', '').replace('<NA>', '')
            
            # Mantém compatibilidade com formato antigo se existir
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
            "Codigo IMS - Origem","Codigo IMS Destino", "Transportadora", "Pedagio", 
            "Cod. Rota", "Fluxo Viagem", "KM","Veiculo","Trip","CrossDock"
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
    
    def _load_mdr_files(self):
        """Carrega arquivos BD_CADASTRO_MDR com cache Parquet para performance"""
        if not self.db_folder:
            return None
        
        db_path = Path(self.db_folder)
        if not db_path.exists():
            return None
        
        # Define colunas específicas para MDR
        mdr_columns = ["MDR", "FONTE DIMENSÕES", "MDR PESO", "VOLUME"]
        
        dataframes = []
        
        # Procura arquivos BD_CADASTRO_MDR
        for file in db_path.glob("*BD_CADASTRO_MDR*.xlsx"):
            if '~$' in file.name:
                continue
            print(f"Loading MDR file: {file.name}")
            
            parquet_path = file.with_suffix('.parquet')
            
            try:
                if self._needs_parquet_conversion(file, parquet_path):
                    # Converte para parquet (header na linha 0 por padrão)
                    print(f"  Converting to Parquet for faster loading...")
                    df = pd.read_excel(file, usecols=lambda x: x in mdr_columns, engine='openpyxl', header=0)
                    
                    # Limpa dados
                    for col in df.columns:
                        if df[col].dtype == 'object':
                            try:
                                df[col] = df[col].astype(str).replace('nan', '')
                            except:
                                pass
                    
                    df.to_parquet(parquet_path, engine='pyarrow', compression='snappy')
                    print(f"  ✓ Parquet cache created")
                
                if parquet_path.exists():
                    print(f"  Loading from Parquet cache (fast mode)...")
                    df = pd.read_parquet(parquet_path, engine='pyarrow')
                else:
                    print(f"  Loading from Excel (slow mode)...")
                    df = pd.read_excel(file, usecols=lambda x: x in mdr_columns, engine='openpyxl', header=0)
                
                dataframes.append(df)
                
            except Exception as e:
                print(f"Error reading {file.name}: {e}")
        
        # Também procura arquivos .xlsm
        for file in db_path.glob("*BD_CADASTRO_MDR*.xlsm"):
            if '~$' in file.name:
                continue
            print(f"Loading MDR file: {file.name}")
            
            parquet_path = file.with_suffix('.parquet')
            
            try:
                if self._needs_parquet_conversion(file, parquet_path):
                    print(f"  Converting to Parquet for faster loading...")
                    df = pd.read_excel(file, usecols=lambda x: x in mdr_columns, engine='openpyxl', header=0)
                    
                    for col in df.columns:
                        if df[col].dtype == 'object':
                            try:
                                df[col] = df[col].astype(str).replace('nan', '')
                            except:
                                pass
                    
                    df.to_parquet(parquet_path, engine='pyarrow', compression='snappy')
                    print(f"  ✓ Parquet cache created")
                
                if parquet_path.exists():
                    print(f"  Loading from Parquet cache (fast mode)...")
                    df = pd.read_parquet(parquet_path, engine='pyarrow')
                else:
                    print(f"  Loading from Excel (slow mode)...")
                    df = pd.read_excel(file, usecols=lambda x: x in mdr_columns, engine='openpyxl', header=0)
                
                dataframes.append(df)
                
            except Exception as e:
                print(f"Error reading {file.name}: {e}")
        
        if dataframes:
            self.mdr_data = pd.concat(dataframes, ignore_index=True)
            print(f"Loaded {len(self.mdr_data)} rows from MDR files")
            return self.mdr_data
        
        return None
    
    def _load_nprc_files(self):
        """Carrega arquivos NPRC_Geral (sheet NPRC_Monthly) com cache Parquet para performance"""
        if not self.db_folder:
            return None
        
        db_path = Path(self.db_folder)
        if not db_path.exists():
            return None
        
        dataframes = []
        
        # Procura arquivos NPRC_Geral
        for file in db_path.glob("*NPRC_Geral*.xlsx"):
            if '~$' in file.name:
                continue
            print(f"Loading NPRC file: {file.name}")
            
            parquet_path = file.parent / (file.stem + "_NPRC_Monthly.parquet")
            
            try:
                if self._needs_parquet_conversion(file, parquet_path):
                    # Converte para parquet - lê da linha 6 (header=5 para índice 0-based)
                    print(f"  Converting to Parquet for faster loading...")
                    df = pd.read_excel(file, sheet_name='NPRC_Monthly', engine='openpyxl', header=5)
                    
                    # Remove colunas vazias (sem nome ou todas vazias)
                    df = df.dropna(axis=1, how='all')
                    df = df.loc[:, df.columns.notna()]
                    
                    # Limpa coluna PN para remover .0 postfix tratando como string
                    if 'PN' in df.columns:
                        df['PN'] = pd.to_numeric(df['PN'], errors='coerce').fillna(0).astype('Int64').astype(str).replace('0', '').replace('<NA>', '')
                    
                    # Limpa dados
                    for col in df.columns:
                        if df[col].dtype == 'object':
                            try:
                                df[col] = df[col].astype(str).replace('nan', '')
                            except:
                                pass
                    
                    df.to_parquet(parquet_path, engine='pyarrow', compression='snappy')
                    print(f"  ✓ Parquet cache created")
                
                if parquet_path.exists():
                    print(f"  Loading from Parquet cache (fast mode)...")
                    df = pd.read_parquet(parquet_path, engine='pyarrow')
                else:
                    print(f"  Loading from Excel (slow mode)...")
                    df = pd.read_excel(file, sheet_name='NPRC_Monthly', engine='openpyxl', header=5)
                    df = df.dropna(axis=1, how='all')
                    df = df.loc[:, df.columns.notna()]
                
                dataframes.append(df)
                
            except Exception as e:
                print(f"Error reading {file.name}: {e}")
        
        # Também procura arquivos .xlsm
        for file in db_path.glob("*NPRC_Geral*.xlsm"):
            if '~$' in file.name:
                continue
            print(f"Loading NPRC file: {file.name}")
            
            parquet_path = file.parent / (file.stem + "_NPRC_Monthly.parquet")
            
            try:
                if self._needs_parquet_conversion(file, parquet_path):
                    print(f"  Converting to Parquet for faster loading...")
                    df = pd.read_excel(file, sheet_name='NPRC_Monthly', engine='openpyxl', header=5)
                    print(f"  Raw Excel load: {len(df)} rows, {len(df.columns)} columns")
                    
                    df = df.dropna(axis=1, how='all')
                    df = df.loc[:, df.columns.notna()]
                    print(f"  After dropna: {len(df)} rows, {len(df.columns)} columns")
                    
                    # Limpa coluna PN para remover .0 postfix tratando como string
                    if 'PN' in df.columns:
                        df['PN'] = pd.to_numeric(df['PN'], errors='coerce').fillna(0).astype('Int64').astype(str).replace('0', '').replace('<NA>', '')
                        print(f"  PN column cleaned: {df['PN'].head(10).tolist()}")
                    else:
                        print(f"  WARNING: PN column not found! Available columns: {df.columns.tolist()[:10]}")
                    
                    for col in df.columns:
                        if df[col].dtype == 'object':
                            try:
                                df[col] = df[col].astype(str).replace('nan', '')
                            except:
                                pass
                    
                    # Remove rows where PN is empty or '0'
                    if 'PN' in df.columns:
                        df = df[df['PN'].notna() & (df['PN'] != '') & (df['PN'] != '0')]
                        print(f"  After filtering empty PNs: {len(df)} rows")
                    
                    df.to_parquet(parquet_path, engine='pyarrow', compression='snappy')
                    print(f"  ✓ Parquet cache created with {len(df)} rows")
                
                if parquet_path.exists():
                    print(f"  Loading from Parquet cache (fast mode)...")
                    df = pd.read_parquet(parquet_path, engine='pyarrow')
                else:
                    print(f"  Loading from Excel (slow mode)...")
                    df = pd.read_excel(file, sheet_name='NPRC_Monthly', engine='openpyxl', header=5)
                    df = df.dropna(axis=1, how='all')
                    df = df.loc[:, df.columns.notna()]
                
                dataframes.append(df)
                
            except Exception as e:
                print(f"Error reading {file.name}: {e}")
        
        if dataframes:
            self.nprc_data = pd.concat(dataframes, ignore_index=True)
            print(f"Loaded {len(self.nprc_data)} rows from NPRC files")
            return self.nprc_data
        
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
            
            # === CACHE: Verifica se já temos resultado PFEP+NPRC para este SAP code ===
            cache_key = f"{filter_column}_{cod_sap_str}"
            
            # Se já temos dados em cache para este SAP, reutiliza
            if cache_key in self.sap_cache:
                cached = self.sap_cache[cache_key]
                pfep_result = cached['pfep_result']
                cod_ims_for_tdc = cached['cod_ims_for_tdc']
                nprc_result = cached['nprc_result']
                print(f"PFEP lookup for {filter_column}={cod_sap_str}: using cached results ({cached['pfep_count']} matches)")
                print(f"NPRC lookup: using cached results ({cached['nprc_count']} matches)")
            else:
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
                
                # Busca nos dados NPRC usando PNs encontrados no PFEP
                nprc_result = None
                if self.nprc_data is not None and pfep_result:
                    # Obtém todos os PNs relacionados ao SAP/IMS code
                    # Filtra PFEP para obter todos os PNs correspondentes
                    if filter_column == "COD IMS":
                        mask = (self.pfep_data['COD IMS'] == cod_sap_str)
                    else:  # COD SAP
                        mask = (self.pfep_data['COD SAP'] == cod_sap_str)
                    
                    related_pns = self.pfep_data[mask]['Part Number'].astype(str).str.strip().tolist()
                    
                    print(f"NPRC lookup: filtering by {len(related_pns)} PNs from PFEP...")
                    print(f"  Sample PNs from PFEP: {related_pns[:5]}")
                    
                    # Verifica se coluna PN existe no NPRC
                    if 'PN' in self.nprc_data.columns:
                        # Mostra sample de PNs no NPRC para comparação
                        nprc_pns_sample = self.nprc_data['PN'].astype(str).str.strip().unique().tolist()[:10]
                        print(f"  Sample PNs in NPRC database: {nprc_pns_sample}")
                        
                        # Filtra NPRC pelos PNs encontrados
                        nprc_mask = self.nprc_data['PN'].astype(str).str.strip().isin(related_pns)
                        nprc_filtered_df = self.nprc_data[nprc_mask]
                        
                        print(f"NPRC lookup: found {len(nprc_filtered_df)} matches")
                        
                        if nprc_filtered_df.empty and len(related_pns) > 0:
                            # Debug: verifica se há match sem strip
                            nprc_mask_no_strip = self.nprc_data['PN'].astype(str).isin(related_pns)
                            test_match = self.nprc_data[nprc_mask_no_strip]
                            print(f"  DEBUG: Without strip: {len(test_match)} matches")
                            
                            # Debug: mostra tipos de dados
                            print(f"  DEBUG: NPRC PN dtype: {self.nprc_data['PN'].dtype}")
                            print(f"  DEBUG: PFEP Part Number dtype: {self.pfep_data['Part Number'].dtype}")
                        
                        if not nprc_filtered_df.empty:
                            # Retorna o DataFrame filtrado completo (não apenas primeira linha)
                            # Isso permite usar todos os dados nas calculações
                            nprc_result = nprc_filtered_df
                
                # Armazena em cache para reutilizar depois
                self.sap_cache[cache_key] = {
                    'pfep_result': pfep_result,
                    'cod_ims_for_tdc': cod_ims_for_tdc,
                    'nprc_result': nprc_result,
                    'pfep_count': len(pfep_match) if pfep_result else 0,
                    'nprc_count': len(nprc_result) if nprc_result is not None else 0
                }
            
            # Busca nos dados TDC usando COD IMS Origem e Destino
            tdc_result = None
            tdc_needs_destino = False
            crossdock_value = None
            tdc_options = None  # Para armazenar múltiplas opções de TDC
            
            # Limpa e prepara o código IMS Destino
            cod_ims_destino = None
            if cidade_destino and str(cidade_destino).strip():
                # Tenta converter cidade_destino para IMS code (se for numérico)
                destino_str = str(cidade_destino).strip().replace('.0', '')
                if destino_str and destino_str.isdigit():
                    cod_ims_destino = destino_str
            
            if self.tdc_data is not None and cod_ims_for_tdc:
                # Verifica se temos AMBOS origem e destino IMS
                if cod_ims_destino:
                    # Filtra TDC por AMBOS: Codigo IMS - Origem E Codigo IMS Destino
                    mask = (
                        (self.tdc_data['Codigo IMS - Origem'].astype(str).str.strip() == cod_ims_for_tdc) &
                        (self.tdc_data['Codigo IMS Destino'].astype(str).str.strip() == cod_ims_destino)
                    )
                    tdc_match = self.tdc_data[mask]
                    
                    print(f"TDC lookup for Origem={cod_ims_for_tdc} AND Destino={cod_ims_destino}: found {len(tdc_match)} matches")
                    
                    if not tdc_match.empty:
                        # Retorna a primeira linha como resultado padrão
                        tdc_result = tdc_match.iloc[0].to_dict()
                        
                        # Extrai opções únicas para dropdowns
                        tdc_options = {
                            'Transportadora': tdc_match['Transportadora'].dropna().unique().tolist() if 'Transportadora' in tdc_match.columns else [],
                            'Veiculo': tdc_match['Veiculo'].dropna().unique().tolist() if 'Veiculo' in tdc_match.columns else [],
                            'Fluxo': tdc_match['Fluxo Viagem'].dropna().unique().tolist() if 'Fluxo Viagem' in tdc_match.columns else [],
                            'Trip': tdc_match['Trip'].dropna().unique().tolist() if 'Trip' in tdc_match.columns else [],
                            'all_rows': tdc_match.to_dict('records')  # Todas as linhas para referência
                        }
                else:
                    # Destino IMS não foi fornecido - sinaliza que é necessário
                    tdc_needs_destino = True
                    print(f"TDC lookup: Destino IMS required. Only Origem={cod_ims_for_tdc} available.")
                    
                    # Busca CrossDock do TDC apenas com origem (para mostrar ao usuário)
                    mask = (self.tdc_data['Codigo IMS - Origem'].astype(str).str.strip() == cod_ims_for_tdc)
                    tdc_match = self.tdc_data[mask]
                    
                    if not tdc_match.empty and 'CrossDock' in tdc_match.columns:
                        crossdock_value = tdc_match.iloc[0].get('CrossDock', None)

            
            # Combina resultados PFEP e TDC
            if pfep_result or tdc_result:
                combined_data = {}
                if pfep_result:
                    combined_data.update(pfep_result)
                if tdc_result:
                    combined_data.update(tdc_result)
                
                # Armazena o último resultado para reutilizar nos cálculos
                self.last_lookup_result = combined_data
                
                result = {
                    "status": "success",
                    "data": combined_data,
                    "filter_used": filter_column,
                    "tdc_needs_destino": tdc_needs_destino,
                    "cod_ims_origem": cod_ims_for_tdc
                }
                
                # Adiciona opções de TDC se disponíveis
                if tdc_options:
                    result["tdc_options"] = tdc_options
                
                # Adiciona CrossDock se disponível (mesmo sem TDC completo)
                if crossdock_value:
                    result["data"]["CrossDock"] = crossdock_value
                
                # NPRC data is cached and retrieved via get_last_nprc_filtered() when needed
                # Don't add DataFrame to result (not JSON serializable)
                
                return result
            
            # Mesmo sem match completo, se temos PFEP result mas precisa destino
            if pfep_result and tdc_needs_destino:
                combined_data = {}
                combined_data.update(pfep_result)
                
                self.last_lookup_result = combined_data
                
                result = {
                    "status": "success",
                    "data": combined_data,
                    "filter_used": filter_column,
                    "tdc_needs_destino": True,
                    "cod_ims_origem": cod_ims_for_tdc
                }
                
                # Adiciona CrossDock se disponível
                if crossdock_value:
                    result["data"]["CrossDock"] = crossdock_value
                
                # NPRC data is cached and retrieved via get_last_nprc_filtered() when needed
                # Don't add DataFrame to result (not JSON serializable)
                
                return result
            
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
        self.mdr_data = None
        self.nprc_data = None
        
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
        
        if progress_callback:
            progress_callback("Loading MDR files...")
        self._load_mdr_files()
        
        if progress_callback:
            progress_callback("Loading NPRC files...")
        self._load_nprc_files()
        
        print("="*60)
        print("✓ Database ready! You can now perform searches.")
        print("="*60 + "\n")
        
        if progress_callback:
            progress_callback("Database ready!")
    
    def reload_data(self):
        """Recarrega os dados dos arquivos"""
        self.pfep_data = None
        self.tdc_data = None
        self.mdr_data = None
        self.nprc_data = None
        self._load_pfep_files()
        self._load_tdc_files()
        self._load_mdr_files()
        self._load_nprc_files()
    
    def clear_cache(self):
        """Limpa o cache de dados SAP"""
        self.sap_cache.clear()
        self.pfep_data = None
        self.tdc_data = None
        self.mdr_data = None
        self.nprc_data = None
        self.last_lookup_result = None
    
    def get_last_lookup_result(self):
        """Retorna o último resultado de lookup para reutilizar nos cálculos"""
        return self.last_lookup_result
    
    def get_pfep_data(self):
        """Retorna o DataFrame completo de dados PFEP"""
        return self.pfep_data
    
    def get_mdr_data(self):
        """Retorna o DataFrame completo de dados MDR"""
        return self.mdr_data
    
    def get_nprc_data(self):
        """Retorna o DataFrame completo de dados NPRC"""
        return self.nprc_data
    
    def get_cached_nprc_data(self):
        """Retorna o DataFrame filtrado de NPRC do último lookup SAP (se disponível)"""
        # Verifica se há dados em cache
        if not self.sap_cache:
            return None
        
        # Pega o último cache (assumindo que só temos um SAP code ativo por vez)
        for cache_key, cached_data in self.sap_cache.items():
            nprc_result = cached_data.get('nprc_result')
            if nprc_result is not None:
                return nprc_result
        
        return None
