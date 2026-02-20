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
    
    def calculate(self, data, pfep_data=None):
        """
        Calcula QME baseado nos dados AS IS/TO BE e inputs do usuário
        
        Args:
            data: Dicionário com parâmetros de cálculo
            pfep_data: DataFrame com dados PFEP completos (opcional)
            
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
        print(f"AS IS/TO BE data: {len(self.asis_data)} rows")
        if pfep_data is not None:
            print(f"PFEP data available: {len(pfep_data)} total rows")
        else:
            print("PFEP data: Not provided")
        print(f"{'='*60}\n")
        
        # Filtra PFEP data pelos PNs do arquivo Astobe
        pfep_filtered = None
        matched_rows = []
        unmatched_rows = []
        
        if pfep_data is not None:
            # Obtém TODOS os PNs únicos do arquivo Astobe (para filtrar PFEP)
            unique_astobe_pns = self.asis_data['PN'].astype(str).str.strip().unique().tolist()
            
            # Filtra PFEP data para incluir apenas esses PNs
            pfep_filtered = pfep_data[pfep_data['Part Number'].astype(str).str.strip().isin(unique_astobe_pns)]
            
            # Cria um set de PNs que existem na PFEP para busca rápida
            pfep_pn_set = set(pfep_filtered['Part Number'].astype(str).str.strip().unique().tolist())
            
            # Para cada LINHA do Astobe, verifica se o PN tem match na PFEP
            for idx, row in self.asis_data.iterrows():
                pn = str(row.get('PN', '')).strip()
                if pn in pfep_pn_set:
                    matched_rows.append(idx)
                else:
                    unmatched_rows.append(idx)
            
            print(f"\n{'='*60}")
            print(f"PN MATCHING RESULTS")
            print(f"{'='*60}")
            print(f"Total ROWS in Astobe file: {len(self.asis_data)}")
            print(f"Matched ROWS (found in PFEP): {len(matched_rows)}")
            print(f"Unmatched ROWS (not in PFEP): {len(unmatched_rows)}")
            print(f"\n✓ Matched row indices: {matched_rows[:10]}{'...' if len(matched_rows) > 10 else ''}")
            if unmatched_rows:
                print(f"\n✗ Unmatched row indices: {unmatched_rows}")
                print(f"   Unmatched PNs: {[str(self.asis_data.iloc[i]['PN']).strip() for i in unmatched_rows]}")
            print(f"{'='*60}\n")

        # Processa cada linha do arquivo AS IS/TO BE
        results = []
        for idx, row in self.asis_data.iterrows():
            # Extrai os dados de cada linha
            pn = str(row.get('PN', f'PN-{idx+1}'))
            
            # AS IS data
            qme_asis = row.get('AS_IS_QME', 0)
            mdr_asis = row.get('AS_IS_MDR', '')
            
            # TO BE data - usa o valor do arquivo, ou o valor padrão do input se não especificado
            qme_tobe = row.get('TO_BE_QME', data.get('qme_tobe', 0))
            mdr_tobe = row.get('TO_BE_MDR', '')
            
            # Converte para numérico se necessário
            try:
                qme_asis = int(qme_asis) if qme_asis else 0
                qme_tobe = int(qme_tobe) if qme_tobe else 0
            except:
                qme_asis = 0
                qme_tobe = 0
            
            # Busca dados do PFEP filtrado para este PN específico
            pfep_info = {}
            if pfep_filtered is not None:
                pn_match = pfep_filtered[pfep_filtered['Part Number'].astype(str).str.strip() == str(pn).strip()]
                if not pn_match.empty:
                    pfep_info = pn_match.iloc[0].to_dict()
            
            # Cálculos (você pode ajustar conforme a lógica real usando dados PFEP)
            # Aqui estamos fazendo um cálculo simples para demonstração
            vol_asis = 10  # Volume calculado AS IS (você pode adicionar lógica real)
            vol_tobe = 8   # Volume calculado TO BE (você pode adicionar lógica real)
            savings = (vol_asis - vol_tobe) * 100  # Economia fictícia
            
            # Exemplo de uso dos dados PFEP nos cálculos:
            if pfep_info:
                # Pode usar dados como: Metro Cúbico Semanal, Pecas por semana, etc
                metro_cubico = pfep_info.get('Metro Cúbico Semanal', 0)
                pecas_semana = pfep_info.get('Pecas por semana', 0)
                fornecedor = pfep_info.get('Nome Fornecedor', '')
                qme_pfep = pfep_info.get('QME (Pecas/Embalagem)', 0)
                # Ajusta cálculos baseado em dados PFEP...
            
            status = "OK" if qme_tobe > qme_asis else "Sem melhoria"
            
            results.append({
                "row": idx + 1,
                "pn": pn,
                "qme_asis": qme_asis,
                "mdr_asis": mdr_asis,
                "qme_tobe": qme_tobe,
                "mdr_tobe": mdr_tobe,
                "vol_asis": vol_asis,
                "vol_tobe": vol_tobe,
                "savings": savings,
                "status": status
            })
        
        # Calcula agregações mensais de QME
        # Distribui o total de QME por 12 meses (exemplo simplificado - ajustar conforme necessidade)
        total_qme_asis = sum(r['qme_asis'] for r in results)
        total_qme_tobe = sum(r['qme_tobe'] for r in results)
        qme_asis_mensal = total_qme_asis / 12
        qme_tobe_mensal = total_qme_tobe / 12
        
        months = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
                  'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        
        monthly_qme_asis = {month: qme_asis_mensal for month in months}
        monthly_qme_tobe = {month: qme_tobe_mensal for month in months}
        
        # Total anual de volumes transportados
        total_asis_anual = sum(r['vol_asis'] for r in results) * 12
        total_tobe_anual = sum(r['vol_tobe'] for r in results) * 12
        
        response = {
            "status": "success",
            "message": f"Simulação concluída para {len(results)} linhas.",
            "results": results,
            "summary": {
                "total_rows": len(self.asis_data),  # Total de linhas no arquivo Astobe
                "total_savings": sum(r['savings'] for r in results),
                "matched_rows": len(matched_rows),  # Linhas com match na PFEP
                "unmatched_rows": len(unmatched_rows),  # Linhas sem match na PFEP
                "monthly_qme_asis": monthly_qme_asis,
                "monthly_qme_tobe": monthly_qme_tobe,
                "total_qme_asis": total_qme_asis,
                "total_qme_tobe": total_qme_tobe,
                "total_asis_anual": total_asis_anual,
                "total_tobe_anual": total_tobe_anual,
                "saving_12_meses": total_asis_anual - total_tobe_anual
            },
            "matching": {
                "matched_rows": matched_rows,
                "unmatched_rows": unmatched_rows,
                "matched_count": len(matched_rows),
                "unmatched_count": len(unmatched_rows)
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
