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
        print("Dados recebidos:", data)
        
        if self.asis_data is None:
            return {
                "status": "error",
                "message": "Carregue o arquivo AS IS/TO BE antes de simular!"
            }

        # AQUI ENTRARÁ O CÁLCULO REAL USANDO O self.asis_data E OS INPUTS
        # Por enquanto, retornamos mock data para teste visual
        
        results = []
        for idx, row in self.asis_data.iterrows():
            # Mock calculation - será substituído pela lógica real
            results.append({
                "row": idx + 1,
                "pn": row.get('PN', f'PN-{idx+1}'),
                "qme_asis": row.get('QME_ASIS', 100),
                "qme_tobe": row.get('QME_TOBE', data.get('qme_tobe', 150)),
                "vol_asis": row.get('VOL_ASIS', 10),
                "vol_tobe": 0,  # Será calculado
                "savings": 0,  # Será calculado
                "status": "OK"
            })
        
        response = {
            "status": "success",
            "message": f"Simulação concluída para {len(results)} linhas.",
            "results": results,
            "summary": {
                "total_rows": len(results),
                "total_savings": sum(r['savings'] for r in results)
            }
        }
        
        self.last_results = results
        return response
    
    def get_last_results(self):
        """Retorna os últimos resultados calculados"""
        return self.last_results
    
    def has_data(self):
        """Verifica se há dados AS IS carregados"""
        return self.asis_data is not None
