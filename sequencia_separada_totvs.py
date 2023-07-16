import winreg
import pyautogui
import os
import shutil
from tkinter import filedialog
import tkinter as tk
from tkinter import messagebox
import pyautogui
from time import sleep
import win32gui
import pygetwindow as gw


############# remover depois de montar #############
import subprocess
def intall_run(path):
    command = None
    if type(path) == list:
        command = [path[0], path[1]]
    else:
        command = path
    try:
        #result = subprocess.run(command, shell=True)
        install = subprocess.Popen(path, shell=True)

       #veri_program(result)
        return True
    except PermissionError:
        print("sem permissão")
        return False

parametro = "totvs"
if True:



############# remover depois de montar #############
    if parametro == "totvs":
####### Funcões ####################################################################
        #função para abrir um pop up de alerta
        def popup_completed(mensagem, e=" "):
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo(e, mensagem)

        # função para criar um diretorio se ele não existir
        def criar_diretorio_se_nao_existir(directory):
            if not os.path.exists(directory):
                os.makedirs(directory)
                return directory + "\\"
            return directory + "\\"
        # função cria um arquivo TXT para gravar em qual parte a instalação está para caso der algum problema e for necessario executar o script novamente não começar do 0 algo que já está avançado
        def ponto_de_controle(condi,conteudo = "0"):
            if condi == "ler":
                try:
                    with open(arquivo_controle + "Controle_do_passo_a_passo.txt", "r") as f:
                       return f.read()
                except FileNotFoundError:
                    with open(arquivo_controle + "Controle_do_passo_a_passo.txt", "w") as f:
                        f.write("0")
                        return "0"
            if condi == "escrever":
                with open(arquivo_controle + "Controle_do_passo_a_passo.txt", "w") as f:
                    f.write(conteudo)
                    return conteudo
        # função para verificar se o Controle de usuario "UAC" está ativado
        def is_uac_enabled():
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System')
                value = winreg.QueryValueEx(key, 'EnableLUA')[0]
                return bool(value)
            except:
                return False
        # função para desabilitar o controle de usuario "UAC"
        def disable_uac():
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System', 0, winreg.KEY_WRITE)
                winreg.SetValueEx(key, 'EnableLUA', 0, winreg.REG_DWORD, 0)
                print("UAC Desabilitado - Reinicie a Maquina")
                return True
            except Exception as error:
                print(error)
                return False
       
        #salvar o caminho dos instaladores na lista "setup" na seguinte ordem
        #1- BIBLIOTECARM-12.1.2209.MSI
        #2- BIBLIOTECARM-12.1.2209.EXE
        #3- BibliotecaRM (Patch) 12.1.2209.164.exe
        #4- sqlncli.msi
        #6- odbcad32.exe
        def copiar_arquivos(original, copia):
            shutil.copy(original, copia)
            #print("Copiado o arquivo: " + str(original))
        # função para listar os arquivos e retorna os caminhos como uma "lista"
        def listar_arquivos(lista_bruta,caminho):
            lista_tratada = []
            for lista in lista_bruta:
                if ("1" in lista[0]):
                   lista_tratada.append(caminho +"\\"+ lista)
                if ("2" in lista[0]):
                   lista_tratada.append(caminho +"\\"+ lista)
                if ("3" in lista[0]):
                   lista_tratada.append(caminho +"\\"+ lista)
                if ("4" in lista[0]):
                   lista_tratada.append(caminho +"\\"+ lista)
                if ("5" in lista[0]):
                   lista_tratada.append(caminho +"\\"+ lista)
                if ("6" in lista[0]):
                   lista_tratada.append(caminho +"\\"+ lista)
            return lista_tratada
        #função função para listar os arquivos dentro de um diretorio
        def list_files(directory):
            with os.scandir(directory) as entries:
                lista = []
                for entry in entries:
                    if entry.is_file():
                        lista.append(entry.name)
                return lista
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
        #função um pouco complexa server para clicar nos botoes avancar de um programa e tem algums parametros:
        #"imagem" - será a imagem que o script vai procurar para clicar
        #"nome_programa" - vai acionar uma função para trazer um programa para frente de todos atraves do nome
        #"imagem2" - será uma imagem secundaria caso o arquivo n acha a primeira imagem ele vai procurar pela segunda logo em seguida
        #"exept" - uma exeção geralmente utilizada para procurar o proximo botão exemplo: caso o programa avance mas o script esta procurando a imagem do estagio anterior o execpt vai procurar a imagem para o estagio atual ou seja sempre
        #"atraso" - define um delay para cada verificação
        def clicar_botao(imagem, nome_programa=None, imagem2=None, exept=None, final=False, atraso=1):
            freio_emergencia = 0
            while True:
                sleep(atraso)
                if nome_programa == None:
                    pass                
                else:
                    trazer_programa_frente(nome_programa)
                # Trazer o programa para a frente usando o PyWin32
                general_button_pos = pyautogui.locateOnScreen(imagem)  # substitua pelo arquivo de imagem do botão "Geral"
                if imagem2 != None:
                    general_button_pos2 = pyautogui.locateOnScreen(imagem2)  # substitua pelo arquivo de imagem do botão "Geral"
                else:
                    general_button_pos2 = None
                if exept != None:
                    general_button_pos3 = pyautogui.locateOnScreen(exept)  # substitua pelo arquivo de imagem do botão "Geral"
                else:
                    general_button_pos3 = None
                ################ condicional para achar o botão
                if general_button_pos != None:
                    pyautogui.click(general_button_pos)
                    if final == True:
                        return "finalizou"
                    else:
                        break
                elif general_button_pos2 != None:
                    pyautogui.click(general_button_pos2)
                    break
                elif general_button_pos3 != None:
                    break
                elif freio_emergencia >= 200:
                    popup_completed("Error reinicie o programa de instalação")
                    exit()
                else:
                    print(f"não encontrou {imagem}")
                    print(general_button_pos)
                    sleep(1)
                    freio_emergencia += 1
        #função para escrever em campos que apareceu no programa
        def escrever(texto, imagem, nome_programa=None):
            while True:
                if nome_programa == None:
                    pass                
                else:
                    trazer_programa_frente(nome_programa)
                # Trazer o programa para a frente usando o PyWin32
                sleep(1)
                general_button_pos = pyautogui.locateOnScreen(imagem)  # substitua pelo arquivo de imagem do botão "Geral"
                if general_button_pos != None:
                    pyautogui.click(general_button_pos)
                    sleep(1)
                    pyautogui.typewrite(texto)
                    break
                else:
                    print(f"não encontrou {imagem}")
                    print(general_button_pos)
                    sleep(1)


####### Parametros #################################################################
        servidor = "\\\\patrimar089\\"
        caminho_ondeclicar = servidor + "e$\\Programas\\Scripts do Renan\\onde_clicar_instalador\\"
        arquivo_controle = criar_diretorio_se_nao_existir(r"C:\Controle_TOVS_Instalação")
        caminho_totvs = "\\\\patrimar089\\e$\\Programas\\Outros\\RM - Instaladores\\Instalador RM"
        servidor_totvs = "server017"
        empresa_totvs = "corpore"
####### PARTE 0 ######################################################################################
        if ponto_de_controle("ler") == "0":

            if is_uac_enabled() == True:
                #print("UAC ATIVADO AQUI A FUNÇÂO IRIA DESATIVAR")
########################################################################
                disable_uac()
                popup_completed("o UAC foi desabilitado o computador será reiniciado", e="Alerta")
                os.system("shutdown /r /t 0")
                exit()
            else:
                ponto_de_controle("escrever","1")
########################################################################
####### PARTE 1 ######################################################################################
        
        if ponto_de_controle("ler") == "1":
            setup = []

            local_temp = criar_diretorio_se_nao_existir(r'C:\TOTVS_temp')
            setup_temp = listar_arquivos(list_files(caminho_totvs),caminho_totvs)
            cont = 0
            for programa in setup_temp:
                copiar_arquivos(programa, local_temp)
                setup_temp[cont] = local_temp + setup_temp[cont]
                cont += 1
            #print("copias concluidas")
            setup = listar_arquivos(list_files(local_temp),local_temp)
            
            ponto_de_controle("escrever","2")
####### PARTE 2 ######################################################################################
        if ponto_de_controle("ler") == "2":
            local_temp = criar_diretorio_se_nao_existir(r'C:\TOTVS_temp')
            setup = listar_arquivos(list_files(local_temp),local_temp)
            print(setup[0])
            intall_run(setup[0])
            ##### Clica em "não" na primeira pergunda
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_1_botao_1.PNG",nome_programa="Pergunta", exept=f"{caminho_ondeclicar}totvs_arquivo_1_botao_2.PNG")
            ##### Clica em "não" na segudna pergunta
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_1_botao_1.PNG",nome_programa="Pergunta", exept=f"{caminho_ondeclicar}totvs_arquivo_1_botao_2.PNG")
            ##### Clica em "Avançar"
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_1_botao_2.PNG",nome_programa="BibliotecaRM - 12.1.2209 - InstallShield Wizard", exept=f"{caminho_ondeclicar}totvs_arquivo_1_botao_3.PNG")
            ##### Clica em "Aceito"
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_1_botao_3.PNG",nome_programa="BibliotecaRM - 12.1.2209 - InstallShield Wizard", exept=f"{caminho_ondeclicar}totvs_arquivo_1_botao_4.PNG")
            ##### Clica em "Avançar"
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_1_botao_4.PNG",nome_programa="BibliotecaRM - 12.1.2209 - InstallShield Wizard", exept=f"{caminho_ondeclicar}totvs_arquivo_1_botao_5.PNG")
            ##### Clica em "Avançar" novamente
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_1_botao_4.PNG",nome_programa="BibliotecaRM - 12.1.2209 - InstallShield Wizard", exept=f"{caminho_ondeclicar}totvs_arquivo_1_botao_5.PNG")
            ##### Clica em "Avançar" mais uma vez
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_1_botao_4.PNG",nome_programa="BibliotecaRM - 12.1.2209 - InstallShield Wizard", exept=f"{caminho_ondeclicar}totvs_arquivo_1_botao_5.PNG")
            ##### Clica em "Instalar"
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_1_botao_5.PNG",nome_programa="BibliotecaRM - 12.1.2209 - InstallShield Wizard", exept=f"{caminho_ondeclicar}totvs_arquivo_1_botao_6.PNG")
            ##### Clica em "Avançar" na nova Aba
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_1_botao_6.PNG",nome_programa="Escolha o ambiente que você deseja utilizar nos sistemas RM", exept=f"{caminho_ondeclicar}totvs_arquivo_1_botao_7.PNG")
            ##### Clica em "Avançar" novamente
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_1_botao_6.PNG",nome_programa="Escolha o ambiente que você deseja utilizar nos sistemas RM", exept=f"{caminho_ondeclicar}totvs_arquivo_1_botao_7.PNG")
            ##### Clica em "Avançar" mais uma vez
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_1_botao_6.PNG",nome_programa="Escolha o ambiente que você deseja utilizar nos sistemas RM", exept=f"{caminho_ondeclicar}totvs_arquivo_1_botao_7.PNG")
            ##### Clica em "Salvar"
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_1_botao_7.PNG",nome_programa="Escolha o ambiente que você deseja utilizar nos sistemas RM", imagem2=f"{caminho_ondeclicar}totvs_arquivo_1_botao_6.PNG", exept=f"{caminho_ondeclicar}totvs_arquivo_1_botao_8.PNG")
            ##### Clica em "OK"
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_1_botao_8.PNG",nome_programa="RM", atraso=10, exept=f"{caminho_ondeclicar}totvs_arquivo_1_botao_9.PNG")
            #### Clica em "OK"
            escrever(servidor_totvs,f"{caminho_ondeclicar}totvs_arquivo_1_botao_9.PNG")
            ##### Clica em "OK"
            escrever(empresa_totvs,f"{caminho_ondeclicar}totvs_arquivo_1_botao_10.PNG")
            sleep(1)
            ##### Clica em "OK"
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_1_botao_11.PNG",nome_programa="Conexão com Banco de Dados - Alias", exept=f"{caminho_ondeclicar}totvs_arquivo_1_botao_12.PNG")
            ##### Clica em "OK"
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_1_botao_12.PNG",nome_programa="Conexão com Banco de Dados - Alias", exept=f"{caminho_ondeclicar}totvs_arquivo_1_botao_13.PNG")
            ##### Clica em "OK"
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_1_botao_13.PNG",nome_programa="Conexão com Banco de Dados - Alias", exept=f"{caminho_ondeclicar}totvs_arquivo_1_botao_14.PNG")
            #### Clica em "Concluir"
            if clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_1_botao_14.PNG",nome_programa="BibliotecaRM - 12.1.2209 - InstallShield Wizard",final=True) == "finalizou":
                print("instalação programa 1 finalizado com sucesso!")
                ponto_de_controle("escrever","3")
            else:
                popup_completed("Error na instalação tente novamente")
####### PARTE 3 ######################################################################################
        if ponto_de_controle("ler") == "3":
            local_temp = criar_diretorio_se_nao_existir(r'C:\TOTVS_temp')
            setup = listar_arquivos(list_files(local_temp),local_temp)
            print(setup[1])
            intall_run(setup[1])
            ##### Clica em "OK" 
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_2_botao_1.PNG", nome_programa="Instalação",exept=f"{caminho_ondeclicar}totvs_arquivo_2_botao_2.PNG")
            ##### Clica em "Avançar" 
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_2_botao_2.PNG", nome_programa="Instalação",exept=f"{caminho_ondeclicar}totvs_arquivo_2_botao_3.PNG")
            ##### Clica em "Não" 
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_2_botao_3.PNG", nome_programa="Instalação",exept=f"{caminho_ondeclicar}totvs_arquivo_2_botao_4.PNG")
            ##### Clica em "Não" novamente
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_2_botao_3.PNG", nome_programa="Instalação",exept=f"{caminho_ondeclicar}totvs_arquivo_2_botao_4.PNG")
            ##### Clica em "aceito os termos"
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_2_botao_4.PNG", nome_programa="Instalação")
            ##### Clica em "avançar"
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_2_botao_5.PNG", exept=f"{caminho_ondeclicar}totvs_arquivo_2_botao_6.PNG")
            ##### Clica em "avançar"
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_2_botao_6.PNG", nome_programa="Instalação", exept=f"{caminho_ondeclicar}totvs_arquivo_2_botao_7.PNG")
            ##### Clica em "Não"
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_2_botao_7.PNG", nome_programa="Instalação", exept=f"{caminho_ondeclicar}totvs_arquivo_2_botao_8.PNG")
            ##### Clica em "Instalar"
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_2_botao_8.PNG", nome_programa="Instalação", exept=f"{caminho_ondeclicar}totvs_arquivo_2_botao_9.PNG")
            ##### Clica em "OK"
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_2_botao_9.PNG", nome_programa="Instalação", atraso=10, exept=f"{caminho_ondeclicar}totvs_arquivo_2_botao_10.PNG")
            ##### Clica em "OK"
            clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_2_botao_10.PNG", nome_programa="Instalação", exept=f"{caminho_ondeclicar}totvs_arquivo_2_botao_11.PNG")
            ##### Clica em "OK"
            if clicar_botao(f"{caminho_ondeclicar}totvs_arquivo_2_botao_11.PNG", nome_programa="Instalação", final=True) == "finalizou":
                print("instalação programa 1 finalizado com sucesso!")
                ponto_de_controle("escrever","4")
            else:
                popup_completed("Error na instalação tente novamente")





