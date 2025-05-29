import pandas as pd
import pywhatkit as kit
import pyautogui
import time
import threading
from tkinter import Tk, Button, Label, Text, END, Toplevel
from tkinter.filedialog import askopenfilename, asksaveasfilename


class WhatsAppAutomation:
    def __init__(self, status_text_widget, root):
        self.planilha = None
        self.df = None
        self.stop_flag = False
        self.paused_index = 0
        self.process_running = False
        self.status_text_widget = status_text_widget
        self.root = root
        self.custom_message = (
            "Bom dia, {nome} tudo bem?\n\n"
            "Sou o David, do Recrutamento, Seleção e Treinamento - PAP da TAHTO! 💙💜🚀\n\n"
            "Recebi seu currículo pelo {honting} e temos uma proposta para você:\n\n"
            "{vaga}  🎊\n\n"
            "REQUISITOS:\n"
            "📒 - Ensino Fundamental Completo, Ensino Médio Cursando ou Completo;\n"
            "📆 Maior de 18 anos;\n"
            "💰 Salário atrativo: Fixo + Remuneração variável (Comissão) + Benefícios (vale alimentação, plano de saúde, vale transporte);\n"
            "📆 Treinamento Remunerado;\n"
            "⏰ Escala: de Segunda a Sábado - 44 horas semanais;\n"
            "📍 Local de trabalho: {praça}\n"
            "🖥️📱 Processo Seletivo online\n\n"
            "Podemos conversar sobre essa oportunidade?"
        )
        self.salvar_retorno_arquivo = None
        self.resultados = []

    def print_status(self, message):
        self.status_text_widget.config(state='normal')
        self.status_text_widget.insert(END, message + "\n")
        self.status_text_widget.see(END)
        self.status_text_widget.config(state='disabled')

    def selecionar_planilha(self):
        self.planilha = askopenfilename(
            title="Selecione a planilha Excel",
            filetypes=[("Arquivo Excel", "*.xlsx")]
        )
        if self.planilha:
            self.print_status(f"Planilha '{self.planilha}' selecionada com sucesso.")
        else:
            self.print_status("Nenhuma planilha foi selecionada.")
        return self.planilha

    def selecionar_arquivo_retorno(self):
        self.salvar_retorno_arquivo = asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Arquivo de Texto", "*.txt")],
            title="Selecione onde salvar o arquivo de retorno"
        )
        if self.salvar_retorno_arquivo:
            self.print_status(f"Arquivo de retorno será salvo em: {self.salvar_retorno_arquivo}")
        else:
            self.print_status("Nenhum arquivo foi selecionado para salvar o retorno.")

    def salvar_retorno(self, numero_formatado, status):
        if self.salvar_retorno_arquivo:
            try:
                with open(self.salvar_retorno_arquivo, "a") as f:
                    f.write(f"Mensagem para {numero_formatado}: {status}\n")
            except Exception as e:
                self.print_status(f"Erro ao salvar o retorno: {e}")

    def salvar_relatorio_csv(self):
        if self.resultados:
            relatorio_df = pd.DataFrame(self.resultados, columns=["Nome", "Número", "Status"])
            arquivo_relatorio = asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("Arquivo CSV", "*.csv")],
                title="Selecione onde salvar o relatório"
            )

            if arquivo_relatorio:
                try:
                    relatorio_df.to_csv(arquivo_relatorio, index=False, sep=';', encoding='utf-8')
                    self.print_status(f"Relatório CSV salvo em: {arquivo_relatorio}")
                except Exception as e:
                    self.print_status(f"Erro ao salvar o relatório: {e}")
            else:
                self.print_status("O relatório não foi salvo.")
        else:
            self.print_status("Nenhum dado para gerar relatório.")

    def enviar_mensagem(self, numero_formatado, mensagem, nome, honting, vaga, praça):
        try:
            kit.sendwhatmsg_instantly(numero_formatado, mensagem)
            self.print_status(f"Mensagem enviada para {numero_formatado}.")
            self.resultados.append([nome, numero_formatado, "Mensagem enviada com sucesso"])
            self.salvar_retorno(numero_formatado, "Mensagem enviada com sucesso")
            time.sleep(2)
            pyautogui.hotkey('ctrl', 'w')
        except Exception as e:
            self.print_status(f"Erro ao enviar a mensagem para {numero_formatado}: {e}")
            self.resultados.append([nome, numero_formatado, f"Erro: {e}"])
            self.salvar_retorno(numero_formatado, f"Erro: {e}")

    def enviar_mensagens(self):
        if self.df is None or self.df.empty:
            self.print_status("Nenhuma planilha carregada ou a planilha está vazia. Por favor, selecione uma planilha válida primeiro.")
            return

        for index in range(self.paused_index, len(self.df)):
            if self.stop_flag:
                self.paused_index = index
                self.process_running = False
                self.print_status("Processo pausado.")
                return

            row = self.df.iloc[index]
            nome = row.iloc[0]
            numero = row.iloc[1]
            honting = row.iloc[2]
            vaga = row.iloc[3]
            praça = row.iloc[4]

            numero_formatado = f'+{numero}'
            mensagem = self.custom_message.format(nome=nome, honting=honting, vaga=vaga, praça=praça)

            self.print_status(f"Enviando mensagem para {nome} no número {numero_formatado}...")
            self.enviar_mensagem(numero_formatado, mensagem, nome, honting, vaga, praça)

            time.sleep(5)

        self.process_running = False
        self.print_status("Todas as mensagens foram enviadas.")
        self.salvar_relatorio_csv()

    def iniciar_processo(self):
        if not self.process_running:
            self.stop_flag = False

            if not self.planilha:
                self.print_status("Nenhuma planilha foi selecionada. O processo foi encerrado.")
                return

            try:
                self.df = pd.read_excel(self.planilha)
                total_contatos = len(self.df)
                self.print_status(f"Total de contatos encontrados: {total_contatos}")
                self.process_running = True
                threading.Thread(target=self.enviar_mensagens).start()
            except Exception as e:
                self.print_status(f"Ocorreu um erro ao processar a planilha: {e}")

    def stop_process(self):
        self.stop_flag = True

    def abrir_janela_mensagem_personalizada(self):
        def salvar_mensagem():
            self.custom_message = mensagem_entry.get("1.0", END).strip()
            self.print_status("Mensagem personalizada salva.")
            janela_msg.destroy()

        janela_msg = Toplevel()
        janela_msg.title("Editar Mensagem")
        x = self.root.winfo_x()
        y = self.root.winfo_y()
        janela_msg.geometry(f"400x300+{x}+{y + 50}")

        mensagem_label = Label(janela_msg, text="Edite a mensagem:")
        mensagem_label.pack(pady=10)

        mensagem_entry = Text(janela_msg, height=10, width=40)
        mensagem_entry.insert(END, self.custom_message)
        mensagem_entry.pack(pady=10)

        salvar_btn = Button(janela_msg, text="Salvar", command=salvar_mensagem)
        salvar_btn.pack(pady=10)

    def adicionar_botoes(self):
        selecionar_btn = Button(self.root, text="Selecionar Planilha", command=self.selecionar_planilha)
        selecionar_btn.pack(pady=10)

        iniciar_btn = Button(self.root, text="Iniciar Envio", command=self.iniciar_processo)
        iniciar_btn.pack(pady=10)

        parar_btn = Button(self.root, text="Parar", command=self.stop_process)
        parar_btn.pack(pady=10)

        editar_mensagem_btn = Button(self.root, text="Editar Mensagem", command=self.abrir_janela_mensagem_personalizada)
        editar_mensagem_btn.pack(pady=10)

        salvar_relatorio_btn = Button(self.root, text="Salvar Relatório CSV", command=self.salvar_relatorio_csv)
        salvar_relatorio_btn.pack(pady=10)


# Configuração da interface gráfica
root = Tk()
root.title("Automação WhatsApp")
root.geometry("500x500")

status_text = Text(root, height=15, width=50)
status_text.pack(pady=20)

automacao = WhatsAppAutomation(status_text, root)
automacao.adicionar_botoes()

root.mainloop()
