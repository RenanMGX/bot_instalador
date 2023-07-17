import os
import shutil
import time
import subprocess
import sys
from time import sleep
from PyQt6.QtWidgets import (QApplication, QCheckBox, QDialog, QPushButton, 
QStyleFactory, QVBoxLayout, QMainWindow, QLabel)
import pyautogui
import pygetwindow as gw

#cliando
def clicando(img=None,programa=None, exept=None, cancel=None):
    bt = pyautogui.locateOnScreen(img)
    if bt != None:
        print(bt)
        pyautogui.click(bt)
        return 1
    elif exept != None:
        if pyautogui.locateOnScreen(exept) != None:
            print(exept)
            if cancel != None:
                if pyautogui.locateOnScreen(cancel) != None:
                    print(cancel)
                    pyautogui.click(cancel)
                    return 999
    return 0
#função para trazer um programa para frente de todos os programas ele vai procurar o programa pelo nome dele no gerenciador de tarefas
def trazer_programa_frente(programa):
    try:
        window = gw.getWindowsWithTitle(programa)[0]
        window.activate()
        window.maximize()
        window.restore()
        pyautogui.click(window.left + 10, window.top + 10)
    except IndexError:
        print(f"Janela do {programa} não encontrada!")

# funcoes para instalar os softwares
def install_popen(path, clicar=None, programa=None, exept=None, cancel=None):
    install = subprocess.Popen(path, shell=True)
    if isinstance(clicar, str):
        clicar = [clicar]
    etapa = 0
    contador_finalizar = 0
    verificar_janela = 0
    while (install.poll() is None) or (etapa < len(clicar)):
        if (etapa < len(clicar)) == False:
            break
        etapa += clicando(clicar[etapa],programa, exept, cancel)
        contador_finalizar += 1
        if contador_finalizar >= 25*60:
            break
        verificar_janela +=1
        if verificar_janela >= 60:
            trazer_programa_frente(programa)
        time.sleep(1)
    return True
def intall_run(path):
    command = None
    if type(path) == list:
        command = [path[0], path[1]]
    else:
        command = path
    try:
        subprocess.run(command, shell=True)
       #veri_program(result)
        return True
    except PermissionError:
        print("sem permissão")
        return False
#verificar instalação
def veri_program(result):
    if result.returncode == 0:
        print("Instalação concluída com sucesso!")
        # Caminho para o arquivo de execução do programa
        path_program = r"C:\Programas\program.exe"
        # Executando o programa
        subprocess.run([path_program], capture_output=True)
    else:
        print("A instalação falhou com código de erro:", result.returncode)
# Função para a instalação do winrar
def instalar_programas(parametro):
    # Função para a instalação dos programas padrões Adobe Reader, Winrar, Google Chrome
    if parametro == "padrao":
        path = ["\\\\patrimar089\\e$\\Programas\\Outros\\winrar\\winrar-x64-620br.exe",
                "\\\\patrimar089\\e$\\Programas\\Outros\\Chrome\\ChromeSetup.exe"
        ]
        install_popen(path[1])
        return intall_run([path[0], "/S"])
    #função para instalar o adobe reader
    if parametro == "adobe_reader":
        original_file = "\\\\patrimar089\\e$\\Programas\\Outros\\adobe_reader\\adobe_reader.exe"
        copy_file = "\\\\patrimar089\\e$\\Programas\\Outros\\adobe_reader\\adobe_reader_temp.exe"
        shutil.copy(original_file, copy_file)
        path = "\\\\patrimar089\\e$\\Programas\\Outros\\adobe_reader\\adobe_reader_temp.exe"
        return install_popen(path, clicar=r"onde_clicar\bt_concluir.PNG", programa="Adobe Acrobat Reader DC Instalador")
    # Função para a instalação do sap_770
    if parametro == "sap770":
        path = "\\\\patrimar089\\e$\\\Programas\\\Outros\\\SAP_770\\SetupAll.exe"
        return install_popen(path, clicar=[r"onde_clicar\bt_sap_next.PNG", r"onde_clicar\bt_sap_checkbox.PNG", r"onde_clicar\bt_sap_next.PNG", r"onde_clicar\bt_sap_next.PNG", r"onde_clicar\bt_sap_next.PNG", r"onde_clicar\bt_sap_next.PNG", r"onde_clicar\bt_sap_next.PNG", r"onde_clicar\bt_sap_close.PNG"],programa="SAP Front End Installer", exept=r"onde_clicar\sap_instalado.PNG", cancel=r"onde_clicar\br_sap_cancel.PNG")
    # Função para a instalação do open_vpn
    if parametro == "open_vpn":
        path = "\\\\patrimar089\\e$\\Programas\\Outros\\VPN\\Open VPN\\OpenVPN-2.5.1-I601-amd64.msi"
        retorno = intall_run([path, "/passive"])
        try:
            original_file = "\\\\patrimar089\\e$\\Programas\\Outros\\Open_VPN\\nao_mexer\\Patrimar.ovpn"
            copy_file = "C:\\Users\\Public\\Downloads\\Patrimar.ovpn"
            shutil.copy(original_file, copy_file)
            copy_file2 = "C:\Program Files\OpenVPN\config\Patrimar.ovpn"
            command = f'powershell -Command "Start-Process cmd -ArgumentList \'/C, copy {copy_file} {copy_file2} /y\' -Verb runAs"'
            subprocess.run(command, shell=True)
            shutil.copy(copy_file, copy_file2)
            print("arquivo copiado")
        except Exception as error:
            print(f"open vpn error: {error}")
        return  retorno
    # Função para a instalação do office_2016
    if parametro == "office_2016":
        path = ["\\\\patrimar089\\e$\\Programas\\Outros\\Office\\2016 office\\Office - 2016 - Standard\\setup.exe",
                "\\\\patrimar089\\e$\\Programas\\Outros\\Office\\2016 office\\Atualização outlook 2016\\outlook2016-kb4011123-fullfile-x64-glb.exe"
        ]
        intall_run(path[0])
        return intall_run([path[1], "/quiet"])
    # Função para a instalação do office_365
    if parametro == "office365":
        path = "\\\\patrimar089\\e$\\Programas\\Outros\\Office\\365 on-line\\Office365 64bit Setup.exe"
        return intall_run(path)
    # Função para a instalação do visualizador_sketchup   
    if parametro == "sketchup_viewer_2022":
        path = "\\\\patrimar089\\e$\\Programas\\Outros\\Sketchup 2022\\Sketchup Visualizador\\SketchUpViewer-2022-0-354-126.exe"
        return intall_run([path, "/silent"])
    # Função para a instalação do para os drivers da HP  
    if parametro == "driver_hp":
        path = ["\\\\patrimar089\\e$\\Programas\\Outros\\drive_hp_g8_250\\audio.exe",
                "\\\\patrimar089\\e$\\Programas\\Outros\\drive_hp_g8_250\\video.exe"
        ]
        intall_run(path[0])
        return intall_run(path[0])
    # Função para a instalação do Project
    if parametro == "project":
        path = "\\\\patrimar089\\e$\\Programas\\Outros\\Office\\MS Project\\MSProject_ptbr_64bits.exe"
        return intall_run(path)
        # Função para a instalação do para o adaptador WIFI
    if parametro == "driver_wifi":
        path = "\\\\patrimar089\\e$\\Programas\\Outros\\Drivers\\Adaptador Wifi\\DWA-131_E1_V5.11b03\\Setup.exe"
        #return intall_run(path)
        return install_popen(path, clicar=[r"onde_clicar\bt_setup.PNG", r"onde_clicar\bt_complete.PNG"])
    else:
        return "não encontrado"

# classe do instalador
class Interface(QDialog,QMainWindow):
    def __init__(self, parent=None):
        super(Interface, self).__init__(parent)
        #titulo do Programa
        self.setWindowTitle("Instalado de Programas Automatico")
        #define o layout para Fusion
        self.changeStyle('Fusion')   
        #tamanho minimo da janela
        self.setGeometry(300, 300, 600, 150)
        #lista de programa para instalar salvos em dicionario chave =  parametro para ser instalado, valor =  nome do programa que será exibido para o usuario
        self.programas = {
            "padrao":"Programas Padrão",
            "adobe_reader":"Adobe Reader", 
            "driver_wifi":"Driver para Adaptador WIFI", 
            "sap770": "SAP Versão 770", 
            "open_vpn": "Open VPN", 
            "office_2016": "Office 2016 Escritorio", 
            "office365" : "Office 365", 
            "sketchup_viewer_2022" : "Visualizador do Sketchup 2022", 
            "driver_hp": "Drive Video e Audio do HP G8 250 *Carregamento Demorado*", 
            "project" : "Project", 
        }
        self.programas2 = {
            "totvs": "Instalar TOTVS"
        }
        #cria uma lista para salvar cada programa em uma check box
        self.checkBox = []
        for chave,valor in self.programas.items():
            self.checkBox.append(QCheckBox(valor))
        self.checkBox2 = []
        for chave,valor in self.programas2.items():
            self.checkBox2.append(QCheckBox(valor))

        self.paginas = 1

        #botão instalar com uma chamada para o metodo executar()
        botao = QPushButton('Instalar')
        botao.setDefault(True)
        botao.clicked.connect(self.executar)
        self.texto = QLabel()

        # self.mudar_pagina = QPushButton('TOTVS')
        # self.mudar_pagina.setFlat(True)
        # self.mudar_pagina.clicked.connect(self.pagina)
        # self.mudar_pagina.setStyleSheet("QPushButton { text-align: Right; }")


        #criando o designer definidos nas variaveis e instancias acima
        layout = QVBoxLayout() 
        for x in self.checkBox:
            layout.addWidget(x)
        for x in self.checkBox2:
            layout.addWidget(x)
        # layout.addWidget(self.mudar_pagina)
        layout.addWidget(botao)
        layout.addWidget(self.texto)
        self.setLayout(layout)

        for x in self.checkBox2:
            x.setVisible(False)

    def pagina(self):
        if self.paginas == 1:
            for x in self.checkBox:
                x.setVisible(False)
                x.setChecked(False)
            for x in self.checkBox2:
                x.setVisible(True)
                x.setChecked(False)
            self.mudar_pagina.setText("Programas")
            self.paginas = 2
        elif self.paginas == 2:
            for x in self.checkBox:
                x.setVisible(True)
                x.setChecked(False)
            for x in self.checkBox2:
                x.setVisible(False)
                x.setChecked(False)
            self.mudar_pagina.setText("TOTVS")
            self.paginas = 1

    #metodo para a o estilo Fusion
    def changeStyle(self, styleName):
        QApplication.setStyle(QStyleFactory.create(styleName))
        QApplication.setPalette(QApplication.style().standardPalette())
    # metodo que executara os programas que estiverem com a checkbox marcada
    def executar(self):
        contador = 0
        instalados = []
        errors = []
        procurando = []
        self.programas.update(self.programas2)
        self.checkBox.extend(self.checkBox2)
        #descarregando todos os comandos do dicionario 'programas' e validando qual está com a checkbox marcada e vai listar qual programa teve exito ao instalar e qual deu error
        for chave,programa in self.programas.items():
            procurando.append(self.checkBox[contador].isChecked())
            if self.checkBox[contador].isChecked():
                print(f"instalação do '{programa}' iniciada")
                gallery.showMinimized()
                try:
                    verificar = instalar_programas(chave)
                    print(f"instalação do '{programa}' encerrada")
                    if verificar == True:
                        instalados.append(programa)
                    else:
                        errors.append(programa)
                except PermissionError:
                    print("Sem permissão para executar o arquivo")
                    errors.append(programa)
                except Exception as error:
                    print(error)
                    errors.append(programa)
                gallery.showNormal()
            contador += 1
        #exibe quais programas instalaram e quais deram error
        if True in procurando:
            self.texto.setText(f"os seguintes programas foram instalados com exito! \n {str(instalados)[1:-1]} \n\n os seguintes pragamas não puderam ser instalados: \n{str(errors)[1:-1]}")
        else:
            self.texto.setText("Programa não Encontrado")

#chamando a classe com o __main__
if __name__ == '__main__':
    instalado_semproblema = False
    app = QApplication(sys.argv)
    gallery = Interface()
    gallery.show()
    gallery.showNormal()
    sys.exit(app.exec())
