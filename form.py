import flet as ft
import datetime
from excel_utils import save_to_excel

# Get the current date
today = datetime.date.today()

class App:
    def __init__(self):
        self.NC = None
        self.Area = None
        self.Numpal = None
        self.Fornecedor = None
        self.Transportador = None
        self.Placa = None
        self.Motorista = None
        self.Tipo_de_NC = None
        self.Descricao = None
        self.Emitente = None
        self.Turno = None
        self.NF = None
        self.Qualidade = None
        self.uploaded_files = []

    def main(self, page: ft.Page):
        page.title = "Formulário Completo"
        page.horizontal_alignment = "left"
        page.vertical_alignment = "center"
        
        self.NC = ft.RadioGroup(
            content=ft.Column([
                ft.Radio(value="Transportador", label="Transportador"),
                ft.Radio(value="Fornecedor", label="Fornecedor"),
            ])
        )
        self.Area = ft.RadioGroup(
            content=ft.Column([
                ft.Radio(value="Armazém Interno (Matéria-prima)", label="Armazém Interno (Matéria-prima)"),
                ft.Radio(value="Armazém Interno (Embalagens)", label="Armazém Interno (Embalagens)"),
            ])
        )
        self.Numpal = ft.TextField(label="Numpal")
        self.Fornecedor = ft.TextField(label="Fornecedor")
        self.Transportador = ft.TextField(label="Transportador")
        self.Placa = ft.TextField(label="Placa")
        self.Motorista = ft.TextField(label="Motorista")
        self.Tipo_de_NC = ft.TextField(label="Tipo de NC")
        self.Descricao = ft.TextField(label="Descricao")

        self.Emitente = ft.RadioGroup(
            content=ft.Column([
                ft.Radio(value="Júlio xavier", label="Júlio xavier"),
                ft.Radio(value="Paulo Salvador", label="Paulo Salvador"),
                ft.Radio(value="Bruno Martins", label="Bruno Martins"),
            ])
        )
        self.Turno = ft.RadioGroup(
            content=ft.Column([
                ft.Radio(value="Turno 1", label="Turno 1"),
                ft.Radio(value="Turno 2", label="Turno 2"),
                ft.Radio(value="ADM", label="ADM"),
            ])
        )

        self.NF = ft.TextField(label="Nota Fiscal")
        self.Qualidade = ft.TextField(label="Qualidade acionada")
        self.Descricao = ft.TextField(label="Descrição")
        
        file_picker = ft.FilePicker(on_result=self.on_file_selected)
        page.overlay.append(file_picker)

        def pick_files(e):
            file_picker.pick_files(allow_multiple=True)

        upload_button = ft.ElevatedButton("Imagens", on_click=pick_files)

        def submit(e):
            data = {
                "NC": self.NC.value,
                "Area": self.Area.value,
                "Numpal": self.Numpal.value,
                "Fornecedor": self.Fornecedor.value,
                "Transportador": self.Transportador.value,
                "Placa": self.Placa.value,
                "Motorista": self.Motorista.value,
                "Tipo_de_NC": self.Tipo_de_NC.value,
                "Descricao": self.Descricao.value,
                "Emitente": self.Emitente.value,
                "Turno": self.Turno.value,
                "NF": self.NF.value,
                "Qualidade": self.Qualidade.value,
                "Data": today,
            }
            save_to_excel(data, self.uploaded_files)
            print("Planilha salva com sucesso.")

        submit_button = ft.ElevatedButton(text="Enviar", on_click=submit)

        # Adicionando todos os componentes ao layout da página dentro de um ListView
        page.add(
            ft.ListView(
                controls=[
                    ft.Container(content=ft.Text("NC relacionada a:"), padding=ft.padding.Padding(10, 5, 10, 5)),
                    ft.Container(content=self.NC, padding=ft.padding.Padding(25, 5, 5, 5)),
                    ft.Container(content=ft.Text("Área:"), padding=ft.padding.Padding(10, 5, 10, 5)),
                    ft.Container(content=self.Area, padding=ft.padding.Padding(25, 5, 5, 5)),
                    ft.Container(content=self.Numpal, padding=ft.padding.Padding(10, 10, 10, 10)),
                    ft.Container(content=self.Fornecedor, padding=ft.padding.Padding(10, 10, 10, 10)),
                    ft.Container(content=self.Transportador, padding=ft.padding.Padding(10, 10, 10, 10)),
                    ft.Container(content=self.Placa, padding=ft.padding.Padding(10, 10, 10, 10)),
                    ft.Container(content=self.Motorista, padding=ft.padding.Padding(10, 10, 10, 10)),
                    ft.Container(content=self.Tipo_de_NC, padding=ft.padding.Padding(10, 10, 10, 10)),
                    ft.Container(content=self.Descricao, padding=ft.padding.Padding(10, 10, 10, 10)),
                    ft.Container(content=ft.Text("Emitente:"), padding=ft.padding.Padding(10, 5, 10, 5)),
                    ft.Container(content=self.Emitente, padding=ft.padding.Padding(25, 5, 5, 5)),
                    ft.Container(content=ft.Text("Turno:"), padding=ft.padding.Padding(10, 5, 10, 5)),
                    ft.Container(content=self.Turno, padding=ft.padding.Padding(25, 5, 5, 5)),
                    ft.Container(content=self.NF, padding=ft.padding.Padding(10, 10, 10, 10)),
                    ft.Container(content=self.Qualidade, padding=ft.padding.Padding(10, 10, 10, 10)),
                    upload_button,
                    ft.Container(content=submit_button, padding=ft.padding.Padding(10, 10, 10, 10)),
                ],
                expand=True
            )
        )
    

    def on_file_selected(self, e: ft.FilePickerResultEvent):
        if e.files:
            self.uploaded_files.extend(e.files)
