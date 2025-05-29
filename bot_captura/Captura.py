import pyautogui
import time
import telebot
import os
import configparser
import threading
from datetime import datetime
from tkinter import ttk, messagebox
import tkinter as tk
import pygetwindow as gw
from PIL import ImageGrab
import mss
import mss.tools

# Constantes fixas
BOT_TOKEN = ""
CHAT_ID = 
NOME_EXECUTAVEL = "chrome"
CAPTURAR_TELA_INTEIRA = True

class CaptureBotApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Capture Bot")

        # Configurações
        self.CONFIG_FILE = "config.ini"
        self.bot_token, self.chat_id, self.caminho_pasta, self.mensagem, self.capturar_tela_inteira, self.nome_executavel = self.ler_arquivo_config(self.CONFIG_FILE)

        self.bot = telebot.TeleBot(self.bot_token)

        # Cria a interface
        self.create_widgets()

        # Variáveis de controle
        self.captura_ativa = False
        self.envios_contador = 0  # Contador de envios

    def create_widgets(self):
        # Configuração da aba principal
        self.frame = ttk.Frame(self.master)
        self.frame.pack(padx=10, pady=10)

        self.start_button = ttk.Button(self.frame, text="Iniciar Captura", command=self.start_capture)
        self.start_button.grid(row=0, column=0, padx=5, pady=5)

        self.stop_button = ttk.Button(self.frame, text="Parar Captura", command=self.stop_capture)
        self.stop_button.grid(row=0, column=1, padx=5, pady=5)

        self.config_button = ttk.Button(self.frame, text="Configurações", command=self.show_config)
        self.config_button.grid(row=0, column=2, padx=5, pady=5)

        self.status_label = ttk.Label(self.frame, text="Status: Aguardando...") 
        self.status_label.grid(row=1, column=0, columnspan=3, pady=10)

        # Caixa de texto para mostrar detalhes do processo
        self.output_text = tk.Text(self.frame, width=50, height=10, state='disabled')
        self.output_text.grid(row=2, column=0, columnspan=3, pady=10)

    def start_capture(self):
        self.captura_ativa = True
        self.status_label.config(text="Status: Capturando...")
        self.output_text.config(state='normal')
        self.output_text.delete(1.0, tk.END)
        threading.Thread(target=self.capture_process).start()

    def stop_capture(self):
        self.captura_ativa = False
        self.status_label.config(text="Status: Aguardando...")
        self.output_text.config(state='normal')
        self.output_text.insert(tk.END, "Captura parada.\n")
        self.output_text.config(state='disabled')

    def capture_process(self):
        with mss.mss() as sct:
            try:
                horario_captura = time.strftime('%Y-%m-%d %H:%M:%S')

                if self.capturar_tela_inteira:
                    # Captura a tela inteira
                    screenshot = sct.grab(sct.monitors[1])
                else:
                    # Captura a área da janela do programa (mesmo não estando ativa)
                    janelas_programa = gw.getWindowsWithTitle(self.nome_executavel)
                    if not janelas_programa:
                        self.exibir_mensagem(f"Janela '{self.nome_executavel}' não encontrada.\n")
                        return

                    janela_programa = janelas_programa[0]
                    bbox = {
                        'left': janela_programa.left,
                        'top': janela_programa.top,
                        'width': janela_programa.width,
                        'height': janela_programa.height
                    }
                    screenshot = sct.grab(bbox)

                # Salva a nova captura como arquivo PNG
                nome_arquivo = f"screenshot_{time.strftime('%Y%m%d_%H%M%S')}.png"
                caminho_completo = os.path.join(self.caminho_pasta, nome_arquivo)
                mss.tools.to_png(screenshot.rgb, screenshot.size, output=caminho_completo)

                # Envia o arquivo via Telegram imediatamente após a captura
                with open(caminho_completo, "rb") as doc:
                    self.bot.send_document(chat_id=self.chat_id, document=doc, caption=f"{self.mensagem.decode('utf-8')}\nHorário da captura: {horario_captura}")

                self.exibir_mensagem(f"Captura realizada e enviada com sucesso! Data e hora: {horario_captura}\n")

                # Incrementa o contador de envios
                self.envios_contador += 1

                # Limpa a pasta a cada 2 envios
                if self.envios_contador >= 2:
                    self.apagar_arquivos_antigos(self.caminho_pasta)
                    self.envios_contador = 0  # Reseta o contador após a limpeza

            except Exception as e:
                self.exibir_mensagem(f"Erro ao capturar e enviar a screenshot: {str(e)}\n")

    def apagar_arquivos_antigos(self, caminho_pasta):
        # Verifica se a pasta existe
        if os.path.exists(caminho_pasta):
            # Lista os arquivos na pasta
            arquivos = os.listdir(caminho_pasta)
            # Filtra os arquivos PNG
            arquivos_png = [arquivo for arquivo in arquivos if arquivo.endswith(".png")]
            # Remove todos os arquivos PNG encontrados
            for arquivo in arquivos_png:
                caminho_arquivo = os.path.join(caminho_pasta, arquivo)
                os.remove(caminho_arquivo)
                print(f"Arquivo {caminho_arquivo} apagado com sucesso.")

    def exibir_mensagem(self, mensagem):
        self.output_text.config(state='normal')
        self.output_text.insert(tk.END, mensagem)
        self.output_text.config(state='disabled')

    def ler_arquivo_config(self, caminho_arquivo):
        config = configparser.ConfigParser()
        if not os.path.exists(caminho_arquivo):
            self.criar_arquivo_config(caminho_arquivo)

        config.read(caminho_arquivo)

        bot_token = config.get("Telegram", "bot_token")
        chat_id = config.get("Telegram", "chat_id")
        caminho_pasta = config.get("Capturas", "caminho_pasta")
        mensagem = config.get("Capturas", "mensagem").encode("utf-8")
        capturar_tela_inteira = config.getboolean("Capturas", "capturar_tela_inteira")
        nome_executavel = config.get("Capturas", "nome_executavel")

        return bot_token, chat_id, caminho_pasta, mensagem, capturar_tela_inteira, nome_executavel

    def criar_arquivo_config(self, caminho_arquivo):
        config = configparser.ConfigParser()
        config["Telegram"] = {
            "bot_token": BOT_TOKEN,
            "chat_id": CHAT_ID
        }
        config["Capturas"] = {
            "caminho_pasta": "C:\\Automacoes\\checkauto\\capture",
            "mensagem": "Captura de tela",
            "capturar_tela_inteira": "True",
            "nome_executavel": NOME_EXECUTAVEL
        }
        with open(caminho_arquivo, "w") as config_file:
            config.write(config_file)

    def show_config(self):
        # Abre uma nova janela para mostrar e editar as configurações
        config_window = tk.Toplevel(self.master)
        config_window.title("Configurações")

        ttk.Label(config_window, text="Mensagem:").grid(row=0, column=0, padx=5, pady=5)
        self.mensagem_entry = ttk.Entry(config_window)
        self.mensagem_entry.insert(0, self.mensagem.decode("utf-8"))  # Decodifica a mensagem para exibir
        self.mensagem_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(config_window, text="Caminho da Pasta:").grid(row=1, column=0, padx=5, pady=5)
        self.caminho_pasta_entry = ttk.Entry(config_window)
        self.caminho_pasta_entry.insert(0, self.caminho_pasta)
        self.caminho_pasta_entry.grid(row=1, column=1, padx=5, pady=5)

        # Botão para salvar configurações
        save_button = ttk.Button(config_window, text="Salvar", command=self.salvar_config)
        save_button.grid(row=3, column=0, columnspan=2, pady=10)

    def salvar_config(self):
        # Atualiza as configurações e salva no arquivo de configuração
        self.mensagem = self.mensagem_entry.get().encode("utf-8")
        self.caminho_pasta = self.caminho_pasta_entry.get()

        config = configparser.ConfigParser()
        config.read(self.CONFIG_FILE)
        config["Capturas"]["mensagem"] = self.mensagem.decode("utf-8")
        config["Capturas"]["caminho_pasta"] = self.caminho_pasta

        with open(self.CONFIG_FILE, "w") as config_file:
            config.write(config_file)

        messagebox.showinfo("Configurações", "Configurações salvas com sucesso!")

# Cria a aplicação e inicia o loop principal do Tkinter
root = tk.Tk()
app = CaptureBotApp(root)
root.mainloop()
