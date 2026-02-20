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
        matched_pns = []
        unmatched_pns = []
        if pfep_data is not None:
            # Obtém todos os PNs do arquivo Astobe
            astobe_pns = self.asis_data['PN'].astype(str).str.strip().tolist()
            
            # Filtra PFEP data para incluir apenas esses PNs
            pfep_filtered = pfep_data[pfep_data['Part Number'].astype(str).str.strip().isin(astobe_pns)]
            
            # Lista de PNs que tiveram match
            matched_pns = pfep_filtered['Part Number'].astype(str).str.strip().unique().tolist()
            
            # Lista de PNs que NÃO tiveram match
            unmatched_pns = [pn for pn in astobe_pns if pn not in matched_pns]
            
            print(f"\n{'='*60}")
            print(f"PN MATCHING RESULTS")
            print(f"{'='*60}")
            print(f"Total PNs in Astobe file: {len(astobe_pns)}")
            print(f"Matched PNs in PFEP: {len(matched_pns)}")
            print(f"Unmatched PNs: {len(unmatched_pns)}")
            print(f"\n✓ Matched PNs:")
            for pn in matched_pns:
                print(f"  ✓ {pn}")
            if unmatched_pns:
                print(f"\n✗ Unmatched PNs (not found in PFEP):")
                for pn in unmatched_pns:
                    print(f"  ✗ {pn}")
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
        
        # Calcula agregações mensais (exemplo)
        monthly_asis = {month: 0 for month in ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
                                                  'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']}
        monthly_tobe = monthly_asis.copy()
        
        # Total anual
        total_asis_anual = sum(r['vol_asis'] for r in results) * 12  # Exemplo
        total_tobe_anual = sum(r['vol_tobe'] for r in results) * 12  # Exemplo
        
        response = {
            "status": "success",
            "message": f"Simulação concluída para {len(results)} PNs.",
            "results": results,
            "summary": {
                "total_rows": len(results),
                "total_savings": sum(r['savings'] for r in results),
                "matched_pns": len(matched_pns),
                "unmatched_pns": len(unmatched_pns),
                "monthly_asis": monthly_asis,
                "monthly_tobe": monthly_tobe,
                "total_asis_anual": total_asis_anual,
                "total_tobe_anual": total_tobe_anual,
                "saving_12_meses": total_asis_anual - total_tobe_anual
            },
            "matching": {
                "matched": matched_pns,
                "unmatched": unmatched_pns
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
