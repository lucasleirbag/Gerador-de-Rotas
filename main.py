import googlemaps
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLineEdit, QLabel, QProgressBar, QCheckBox, QComboBox, QCompleter, QMessageBox, QHBoxLayout
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5 import QtCore
from PyQt5.QtGui import QIcon
from pandas import read_excel
import xlsxwriter
import subprocess
import sys
import os

chave_api = 'AIzaSyALZGyVtICuk8rlvPcWXH_IBngvZbzLvrc'

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

application_path = os.path.dirname(__file__) if __file__ else os.path.dirname(sys.executable)
base_completa_path = resource_path('base_completa.xlsx')

# Leitura inicial do Excel e armazenamento dos dados
df = read_excel(base_completa_path)
mapa_cidades_ufs = {cidade: uf for cidade, uf in zip(df['MUNICIPIO'], df['UF'])}
cidades = sorted(df['MUNICIPIO'].unique())
ufs = sorted(df['UF'].unique())
cidades_por_estado = {uf: df[df['UF'] == uf]['MUNICIPIO'].unique() for uf in ufs}

class Worker(QThread):
    progressSignal = pyqtSignal(int)
    finishSignal = pyqtSignal(list, list, list)

    def __init__(self, chave_api, origens, destinos, modo):
        super().__init__()
        self.chave_api = chave_api
        self.origens = origens
        self.destinos = destinos
        self.modo = modo

    def run(self):
        gmaps = googlemaps.Client(key=self.chave_api)
        resultado = []
        links_google_maps = []
        total = len(self.destinos)
        for i, destino in enumerate(self.destinos):
            destino_completo = f"{destino[0]}, {destino[1]}, Brazil"
            resultado_distancia = gmaps.distance_matrix(self.origens, [destino_completo], mode=self.modo)
            if resultado_distancia['rows'][0]['elements'][0]['status'] == 'OK':
                distancia = resultado_distancia['rows'][0]['elements'][0]['distance']['text']
                duracao = resultado_distancia['rows'][0]['elements'][0]['duration']['text']
                resultado.append((distancia, duracao))
                links_google_maps.append(f"https://www.google.com/maps/dir/{self.origens[0]}/{destino_completo}")
            else:
                resultado.append(('N/A', 'N/A')) 
                links_google_maps.append('')
            self.progressSignal.emit(int((i+1)*100/total))  # Corrigido para enviar um valor inteiro
        self.finishSignal.emit(resultado, self.destinos, links_google_maps)

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'Calculadora de Distâncias'
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(800, 450, 400, 300)

        # Defina o ícone da janela da aplicação
        icon_path = resource_path('rota.ico')
        self.setWindowIcon(QIcon(icon_path))

        layout = QVBoxLayout()
        self.setLayout(layout)

        self.entrada_cidade = QLineEdit(self)
        self.completer = QCompleter(cidades)  # Usa a lista de cidades carregada previamente
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)  
        self.entrada_cidade.setCompleter(self.completer)  
        self.entrada_cidade.textChanged.connect(self.atualizar_uf)
        layout.addWidget(QLabel("Cidade Origem:"))
        layout.addWidget(self.entrada_cidade)

        # Muda para QComboBox
        self.entrada_uf = QComboBox(self)
        self.entrada_uf.addItems(ufs)  # Preenche com UF's
        layout.addWidget(QLabel("UF Origem:"))
        layout.addWidget(self.entrada_uf)

        self.entrada_modo = QComboBox(self)
        self.entrada_modo.addItems(['driving', 'walking', 'bicycling', 'transit'])
        layout.addWidget(QLabel("Modo de Transporte:"))
        layout.addWidget(self.entrada_modo)

        # dentro da função initUI
        layout_estado_especifico = QHBoxLayout()
        self.estado_especifico = QCheckBox(self)
        layout_estado_especifico.addWidget(self.estado_especifico)
        self.valor_estado_especifico = QComboBox(self)
        self.valor_estado_especifico.addItems(ufs)
        layout_estado_especifico.addWidget(QLabel("Estado Específico:"))
        layout_estado_especifico.addWidget(self.valor_estado_especifico)
        layout.addLayout(layout_estado_especifico)

        self.progresso = QProgressBar(self)
        self.progresso.hide()  # Esconde a barra de progresso inicialmente
        layout.addWidget(self.progresso)

        self.botao = QPushButton("Pesquisar", self)
        self.botao.clicked.connect(self.pesquisar_distancia)
        layout.addWidget(self.botao)

        self.signature_label = QLabel("Desenvolvido por Lucas Gabriel©", self)
        self.signature_label.setStyleSheet("color: gray;")
        self.signature_label.setAlignment(QtCore.Qt.AlignCenter)
        layout.addWidget(self.signature_label)

        self.show()

    # a função de slot para o evento textChanged
    def atualizar_uf(self, text):
        uf = mapa_cidades_ufs.get(text)
        if uf is not None:
            index = self.entrada_uf.findText(uf)
            if index >= 0:
                self.entrada_uf.setCurrentIndex(index)

    def pesquisar_distancia(self):
        self.progresso.show()  # Mostra a barra de progresso antes de iniciar a thread
        cidade_origem = self.entrada_cidade.text()
        uf_origem = self.entrada_uf.currentText()  # Linha corrigida
        origem = [f"{cidade_origem}, {uf_origem}, Brazil"]
        modo = self.entrada_modo.currentText()

        # Verifica se a opção "Estado Específico" está marcada
        if self.estado_especifico.isChecked():
            estado_especifico = self.valor_estado_especifico.currentText()
            destinos = [(cidade, estado_especifico) for cidade in cidades_por_estado[estado_especifico]]
        else:
            destinos = [(cidade, uf) for cidade, uf in zip(cidades, ufs)]

        self.worker = Worker(chave_api, origem, destinos, modo)
        self.worker.progressSignal.connect(self.progresso.setValue)
        self.worker.finishSignal.connect(self.escrever_excel)
        self.worker.start()

    def escrever_excel(self, resultado, destinos, links_google_maps):
        with xlsxwriter.Workbook('resultado.xlsx') as workbook:
            worksheet = workbook.add_worksheet()

            # Defina a largura das colunas
            worksheet.set_column('A:E', 20)

            # Adicione os nomes das colunas
            worksheet.write(0, 0, "Cidade")
            worksheet.write(0, 1, "UF")
            worksheet.write(0, 2, "Distância")
            worksheet.write(0, 3, "Duração")
            worksheet.write(0, 4, "Link")
            
            # Comece a escrever os dados na segunda linha (índice 1)
            for i, ((distancia, duracao), (cidade, uf), link) in enumerate(zip(resultado, destinos, links_google_maps), start=1):
                worksheet.write(i, 0, cidade)
                worksheet.write(i, 1, uf)
                worksheet.write(i, 2, distancia)
                worksheet.write(i, 3, duracao)
                worksheet.write(i, 4, link)

        self.mostrar_mensagem_sucesso()


    def mostrar_mensagem_sucesso(self):
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setWindowTitle("Sucesso!")
        msg_box.setText('O arquivo resultado.xlsx foi criado com sucesso!')
        abrir_button = msg_box.addButton('Abrir Arquivo', QMessageBox.ActionRole)
        msg_box.addButton(QMessageBox.Ok)

        abrir_button.clicked.connect(self.abrir_arquivo)
        msg_box.exec_()

    def abrir_arquivo(self):
        if sys.platform.startswith('darwin'):  # Caso seja MacOS
            subprocess.call(('open', 'resultado.xlsx'))
        elif os.name == 'nt':  # Caso seja Windows
            os.startfile('resultado.xlsx')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
