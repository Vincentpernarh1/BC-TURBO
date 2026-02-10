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
    
    def calculate(self, data):
        """
        Calcula QME baseado nos dados AS IS/TO BE e inputs do usuário
        
        Args:
            data: Dicionário com parâmetros de cálculo
            
        Returns:
            Dicionário com resultados da simulação
        """
        # print("Dados recebidos:", data)
        
        if self.asis_data is None:
            return {
                "status": "error",
                "message": "Carregue o arquivo AS IS/TO BE antes de simular!"
            }

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
            
            # Cálculos de exemplo (você pode ajustar conforme a lógica real)
            # Aqui estamos fazendo um cálculo simples para demonstração
            vol_asis = 10  # Volume calculado AS IS (você pode adicionar lógica real)
            vol_tobe = 8   # Volume calculado TO BE (você pode adicionar lógica real)
            savings = (vol_asis - vol_tobe) * 100  # Economia fictícia
            
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
        
        response = {
            "status": "success",
            "message": f"Simulação concluída para {len(results)} PNs.",
            "results": results,
            "summary": {
                "total_rows": len(results),
                "total_savings": sum(r['savings'] for r in results)
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
