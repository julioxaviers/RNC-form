import flet as ft
from dotenv import load_dotenv
from form import App

# Carregar variáveis de ambiente do arquivo .env
load_dotenv()

# Iniciando a aplicação Flet
app_instance = App()
ft.app(target=app_instance.main, view=ft.WEB_BROWSER, host="192.168.0.88add", port=8000)
