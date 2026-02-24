"""
Módulo para cálculo de QME (AS IS e TO BE)
"""

class QMECalculator:
    def __init__(self):
        self.asis_data = None
        self.last_results = None
    
    def set_asis_data(self, data):
        """Define os dados AS IS/TO BE carregados"""
        self.asis_data = data
    
    def calculate(self, data, pfep_data=None, nprc_data=None, mdr_data=None):
        """
        Calcula QME baseado nos dados TO BE (propose file) e AS IS (PFEP)
        
        Args:
            data: Dicionário com parâmetros de cálculo
            pfep_data: DataFrame com dados PFEP completos - fonte de AS IS data
            nprc_data: DataFrame com dados NPRC filtrados (opcional)
            mdr_data: DataFrame com dados MDR para lookup de volumes
            
        Returns:
            Dicionário com resultados da simulação
        """
        # print("Dados recebidos:", data)
        
        if self.asis_data is None:
            return {
                "status": "error",
                "message": "Carregue o arquivo AS IS/TO BE antes de simular!"
            }
        
        # Log informações sobre os dados recebidos
        print(f"\n{'='*60}")
        print("QME CALCULATION STARTING")
        print(f"{'='*60}")
        print(f"PROPOSE file (TO BE data): {len(self.asis_data)} rows")
        if pfep_data is not None:
            print(f"PFEP data (AS IS source): {len(pfep_data)} total rows")
        else:
            print("PFEP data: Not provided")
        if nprc_data is not None:
            print(f"NPRC data available: {len(nprc_data)} rows")
        else:
            print("NPRC data: Not provided")
        if mdr_data is not None:
            print(f"MDR data (volume lookup): {len(mdr_data)} rows")
        else:
            print("MDR data: Not provided")
        print(f"{'='*60}\n")
        
        # STEP 1: Aggregate NPRC data by PN (sum monthly volumes for duplicate PNs)
        nprc_aggregated = {}
        if nprc_data is not None and 'PN' in nprc_data.columns:
            print(f"Aggregating NPRC data by PN...")
            print(f"  NPRC raw rows: {len(nprc_data)}")
            
            month_cols = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']
            
            for idx, row in nprc_data.iterrows():
                pn = str(row.get('PN', '')).strip()
                
                if pn not in nprc_aggregated:
                    # Initialize entry for this PN
                    nprc_aggregated[pn] = {
                        'PN': pn,
                        'Plant': row.get('Plant', ''),
                        'Model': row.get('Model', ''),
                        'rows_aggregated': 0
                    }
                    # Initialize monthly volumes to 0
                    for col in month_cols:
                        nprc_aggregated[pn][col] = 0
                
                # Aggregate (sum) monthly volumes
                for col in month_cols:
                    if col in row:
                        try:
                            vol = float(row[col]) if row[col] and str(row[col]).lower() != 'nan' else 0
                            nprc_aggregated[pn][col] += vol
                        except:
                            pass
                
                nprc_aggregated[pn]['rows_aggregated'] += 1
            
            # Check how many PNs had duplicates
            duplicates = sum(1 for pn_data in nprc_aggregated.values() if pn_data['rows_aggregated'] > 1)
            print(f"  NPRC unique PNs: {len(nprc_aggregated)}")
            print(f"  PNs with duplicates (aggregated): {duplicates}")
            
            # Show example of aggregated PN
            example_pn = next((pn_data for pn_data in nprc_aggregated.values() if pn_data['rows_aggregated'] > 1), None)
            if example_pn:
                print(f"\n  Example aggregated PN: {example_pn['PN']}")
                print(f"    Rows aggregated: {example_pn['rows_aggregated']}")
                print(f"    Total volume (all months): {sum(example_pn[col] for col in ['1','2','3','4','5','6','7','8','9','10','11','12'])}")
            print()
        
        # STEP 2: Find PNs that exist in BOTH PFEP and NPRC (intersection)
        # This creates the BASE DATASET for calculations
        
        pfep_pn_set = set()
        nprc_pn_set = set()
        
        if pfep_data is not None:
            pfep_pn_set = set(pfep_data['Part Number'].astype(str).str.strip().unique().tolist())
            print(f"PFEP filtered PNs: {len(pfep_pn_set)}")
        
        if nprc_aggregated:
            nprc_pn_set = set(nprc_aggregated.keys())
            print(f"NPRC aggregated PNs: {len(nprc_pn_set)}")
        
        # Find intersection: PNs that exist in BOTH PFEP and NPRC
        matched_pns = pfep_pn_set.intersection(nprc_pn_set)
        
        print(f"\n{'='*60}")
        print(f"PFEP + NPRC INTERSECTION")
        print(f"{'='*60}")
        print(f"PNs in PFEP (filtered by SAP): {len(pfep_pn_set)}")
        print(f"PNs in NPRC (aggregated): {len(nprc_pn_set)}")
        print(f"PNs in BOTH (intersection): {len(matched_pns)}")
        print(f"Sample matched PNs: {list(matched_pns)[:10]}")
        print(f"{'='*60}\n")
        
        # STEP 3: Create propose file lookup for TO BE values
        propose_lookup = {}
        propose_pns_in_dataset = []
        propose_pns_not_in_dataset = []
        
        if self.asis_data is not None:
            for idx, row in self.asis_data.iterrows():
                pn = str(row.get('PN', '')).strip()
                propose_lookup[pn] = {
                    'qme_tobe': row.get('TO_BE_QME', 0),
                    'mdr_tobe': str(row.get('TO_BE_MDR', '')).strip(),
                    'row_index': idx
                }
                
                if pn in matched_pns:
                    propose_pns_in_dataset.append(pn)
                else:
                    propose_pns_not_in_dataset.append(pn)
            
            print(f"Propose file analysis:")
            print(f"  Total PNs in propose file: {len(propose_lookup)}")
            print(f"  PNs found in PFEP+NPRC dataset: {len(propose_pns_in_dataset)} - {propose_pns_in_dataset}")
            print(f"  PNs NOT in dataset: {len(propose_pns_not_in_dataset)} - {propose_pns_not_in_dataset}")
            print()

        # STEP 3: Process ALL matched PNs (PFEP + NPRC intersection)
        results = []
        row_num = 1
        
        for pn in sorted(matched_pns):  # Process all PNs that exist in both PFEP and NPRC
            
            # Check if this PN has TO BE data in propose file
            has_propose_data = pn in propose_lookup
            qme_tobe = 0
            mdr_tobe = ''
            
            if has_propose_data:
                qme_tobe = propose_lookup[pn]['qme_tobe']
                mdr_tobe = propose_lookup[pn]['mdr_tobe']
                
                # Convert to numeric if needed
                try:
                    qme_tobe = int(qme_tobe) if qme_tobe else 0
                except:
                    qme_tobe = 0
            
            # Get PFEP data (AS IS source) - guaranteed to exist since pn is in matched_pns
            pfep_info = {}
            qme_asis = 0
            mdr_asis = ''
            
            if pfep_data is not None:
                pn_match = pfep_data[pfep_data['Part Number'].astype(str).str.strip() == pn]
                if not pn_match.empty:
                    pfep_info = pn_match.iloc[0].to_dict()
                    # AS IS data from PFEP
                    qme_asis = pfep_info.get('QME (Pecas/Embalagem)', 0)
                    mdr_asis = str(pfep_info.get('MDR', '')).strip()
                    
                    # Convert QME AS IS to numeric
                    try:
                        qme_asis = int(qme_asis) if qme_asis else 0
                    except:
                        qme_asis = 0
            
            # Get NPRC data - guaranteed to exist since pn is in matched_pns
            # Use aggregated NPRC data (monthly volumes already summed for duplicate PNs)
            nprc_info = {}
            monthly_volumes = {}
            
            if pn in nprc_aggregated:
                nprc_info = nprc_aggregated[pn]
                
                # Extract monthly volumes from aggregated NPRC
                # Columns: '2' (Feb), '3' (Mar), '4' (Apr), '5' (May), '6' (Jun), '7' (Jul),
                #          '8' (Aug), '9' (Sep), '10' (Oct), '11' (Nov), '12' (Dec), '1' (Jan)
                # NOTE: Next month these will shift - start from '3' (Mar) to next year '3'
                month_mapping = {
                    '1': 'Jan', '2': 'Fev', '3': 'Mar', '4': 'Abr',
                    '5': 'Mai', '6': 'Jun', '7': 'Jul', '8': 'Ago',
                    '9': 'Set', '10': 'Out', '11': 'Nov', '12': 'Dez'
                }
                
                for col, month_name in month_mapping.items():
                    if col in nprc_info:
                        monthly_volumes[month_name] = nprc_info[col]  # Already aggregated
                    else:
                        monthly_volumes[month_name] = 0
            
            # Lookup volumes in MDR using MDR codes (AS IS and TO BE)
            vol_asis_m3 = 0  # Volume in m³ AS IS
            vol_tobe_m3 = 0  # Volume in m³ TO BE
            peso_asis_kg = 0  # Weight in kg AS IS
            peso_tobe_kg = 0  # Weight in kg TO BE
            
            if mdr_data is not None and 'MDR' in mdr_data.columns:
                # Lookup AS IS volume using AS IS MDR
                if mdr_asis:
                    mdr_match_asis = mdr_data[mdr_data['MDR'].astype(str).str.strip() == mdr_asis]
                    if not mdr_match_asis.empty:
                        vol_asis_m3 = mdr_match_asis.iloc[0].get('VOLUME', 0)
                        peso_asis_kg = mdr_match_asis.iloc[0].get('MDR PESO', 0)
                        try:
                            vol_asis_m3 = float(vol_asis_m3) if vol_asis_m3 else 0
                            peso_asis_kg = float(peso_asis_kg) if peso_asis_kg else 0
                        except:
                            vol_asis_m3 = 0
                            peso_asis_kg = 0
                
                # Lookup TO BE volume using TO BE MDR (if propose data exists)
                if mdr_tobe:
                    mdr_match_tobe = mdr_data[mdr_data['MDR'].astype(str).str.strip() == mdr_tobe]
                    if not mdr_match_tobe.empty:
                        vol_tobe_m3 = mdr_match_tobe.iloc[0].get('VOLUME', 0)
                        peso_tobe_kg = mdr_match_tobe.iloc[0].get('MDR PESO', 0)
                        try:
                            vol_tobe_m3 = float(vol_tobe_m3) if vol_tobe_m3 else 0
                            peso_tobe_kg = float(peso_tobe_kg) if peso_tobe_kg else 0
                        except:
                            vol_tobe_m3 = 0
                            peso_tobe_kg = 0
            
            # Calculate savings (to be implemented)
            savings = 0
            
            # Status: All PNs in results are matched (exist in both PFEP and NPRC)
            # Highlight if this PN has propose data (TO BE)
            status = "Matched - In Dataset"
            if has_propose_data:
                if qme_tobe > qme_asis:
                    status = "Matched - TO BE Improvement"
                else:
                    status = "Matched - TO BE No Change"
            
            results.append({
                "row": row_num,
                "pn": pn,
                "qme_asis": qme_asis,
                "mdr_asis": mdr_asis,
                "qme_tobe": qme_tobe,
                "mdr_tobe": mdr_tobe,
                "vol_asis_m3": vol_asis_m3,
                "vol_tobe_m3": vol_tobe_m3,
                "peso_asis_kg": peso_asis_kg,
                "peso_tobe_kg": peso_tobe_kg,
                "vol_asis": vol_asis_m3,  # Backward compat
                "vol_tobe": vol_tobe_m3,  # Backward compat
                "monthly_volumes": monthly_volumes,  # Monthly volumes from NPRC
                "savings": savings,
                "status": status,
                "has_pfep_match": True,  # All PNs in results are matched
                "has_nprc_data": True,   # All PNs in results have NPRC data
                "has_propose_data": has_propose_data,  # Flag if this PN is in propose file
                "pfep_data": pfep_info,
                "nprc_data": nprc_info
            })
            
            row_num += 1
        
        # Aggregate ACTUAL monthly volumes from NPRC (not divide by 12!)
        # Sum monthly volumes across all PNs for each month
        months = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 
                  'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
        
        # Initialize monthly aggregations
        monthly_volumes_total = {month: 0 for month in months}
        
        # Sum volumes from all PNs for each month
        for result in results:
            for month in months:
                monthly_volumes_total[month] += result.get('monthly_volumes', {}).get(month, 0)
        
        # Create monthly dictionaries with actual NPRC volumes
        monthly_qme_asis = {}
        monthly_qme_tobe = {}
        
        for month in months:
            # AS IS: actual NPRC volume for this month
            monthly_qme_asis[month] = monthly_volumes_total[month]
            
            # TO BE: for now same as AS IS (will be different if propose data changes volumes)
            # TODO: Apply TO BE QME adjustments if needed
            monthly_qme_tobe[month] = monthly_volumes_total[month]
        
        # Total annual volumes (sum of all months)
        total_qme_asis = sum(monthly_qme_asis.values())
        total_qme_tobe = sum(monthly_qme_tobe.values())
        
        # Total annual volumes transported (from MDR)
        total_asis_anual = sum(r['vol_asis'] for r in results) * 12
        total_tobe_anual = sum(r['vol_tobe'] for r in results) * 12
        
        print(f"\nMonthly Volume Totals (from NPRC):")
        for month in months:
            print(f"  {month}: {monthly_volumes_total[month]:.0f}")
        print(f"  Total Annual: {total_qme_asis:.0f}\n")
        
        # Count PNs with propose data
        pns_with_propose = sum(1 for r in results if r.get('has_propose_data', False))
        pns_without_propose = len(results) - pns_with_propose
        
        print(f"\n{'='*60}")
        print(f"FINAL DATASET SUMMARY")
        print(f"{'='*60}")
        print(f"Total PNs in dataset (PFEP+NPRC intersection): {len(results)}")
        print(f"PNs with TO BE data (from propose file): {pns_with_propose}")
        print(f"PNs without TO BE data: {pns_without_propose}")
        print(f"{'='*60}\n")
        
        response = {
            "status": "success",
            "message": f"Dataset created with {len(results)} PNs (PFEP+NPRC intersection).",
            "results": results,
            "summary": {
                "total_rows": len(results),  # Total PNs in dataset (PFEP+NPRC intersection)
                "total_savings": sum(r['savings'] for r in results),
                "matched_rows": len(results),  # All rows are matched (PFEP+NPRC)
                "unmatched_rows": 0,  # No unmatched rows in dataset
                "pns_with_propose": pns_with_propose,  # PNs that have TO BE data
                "pns_without_propose": pns_without_propose,  # PNs without TO BE data
                "monthly_qme_asis": monthly_qme_asis,
                "monthly_qme_tobe": monthly_qme_tobe,
                "total_qme_asis": total_qme_asis,
                "total_qme_tobe": total_qme_tobe,
                "total_asis_anual": total_asis_anual,
                "total_tobe_anual": total_tobe_anual,
                "saving_12_meses": total_asis_anual - total_tobe_anual
            },
            "matching": {
                "total_matched_pns": len(results),  # Total PNs in PFEP+NPRC
                "propose_pns_in_dataset": propose_pns_in_dataset,  # Which propose PNs are in dataset
                "propose_pns_not_in_dataset": propose_pns_not_in_dataset  # Which propose PNs are NOT in dataset
            }
        }
        
        self.last_results = response
        return response
    
    
    
    def get_last_results(self):
        """Retorna os últimos resultados calculados"""
        return self.last_results
    
    def has_data(self):
        """Verifica se há dados AS IS carregados"""
        return self.asis_data is not None
