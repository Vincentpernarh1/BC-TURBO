import webview
from api import Api

# Inicializa API
api = Api()

# Cria a janela
window = webview.create_window(
    'BC Turbo - System', 
    'assets/index.html', 
    js_api=api,
    width=1200, 
    height=800,
    resizable=True
)

if __name__ == '__main__':
    # debug=True permite clicar com botão direito -> Inspecionar Elemento (útil para dev)
    webview.start(debug=False)
