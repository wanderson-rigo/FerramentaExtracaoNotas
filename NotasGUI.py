import tkinter as tk
from tkinter import filedialog, messagebox
import sys
import json
import ExtrairNotasSIGAA  # Importe o módulo ExtrairNotasSIGAA.py
import DigitarDrive  # Importe o módulo DigitarDrive.py

class InterfaceGrafica:
    def __init__(self, root):
        self.root = root
        self.root.title("Extrair Notas do SIGAA e Exportar para Google Sheets")
        self.root.geometry("800x600")  # Largura x Altura

        self.log_text = tk.Text(self.root, wrap="word")
        self.log_text.pack(expand=True, fill="both")

        # Redireciona sys.stdout para o widget de texto
        self.log_redirector = LogRedirector(self.log_text)
        sys.stdout = self.log_redirector
        sys.stderr = self.log_redirector

        self.menu = tk.Menu(root)
        self.root.config(menu=self.menu)

        # Menu de ações
        self.file_menu = tk.Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="Ações", menu=self.file_menu)
        self.file_menu.add_command(label="1 - Carregar Configurações", command=self.carregar_configuracoes)
        self.file_menu.add_command(label="2 - Extrair Notas do SIGAA para CSV", command=self.extrair_notas)
        self.file_menu.add_command(label="3 - Exportar Dados CSV para Google Sheets", command=self.exportar_google_sheets)
        self.file_menu.add_command(label="4 - Resumir Trimestres no Google Sheets", command=self.resumir_trimestres)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Limpar Log", command=self.limpar_log)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Sair", command=root.quit)

        # Menu de opções com checkbox
        self.extrair_optativas = tk.BooleanVar(value=False)  # Estado inicial do checkbox (não marcado por padrão)
        self.options_menu = tk.Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="Disciplinas Optativas", menu=self.options_menu)
        self.options_menu.add_checkbutton(
            label="Extrair disciplinas optativas",
            variable=self.extrair_optativas,
            onvalue=True,
            offvalue=False
        )

        # Menu de ajuda
        self.help_menu = tk.Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="Ajuda", menu=self.help_menu)
        self.help_menu.add_command(label="Instruções", command=self.mostrar_ajuda)

        # Menu sobre
        self.about_menu = tk.Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="Sobre", menu=self.about_menu)
        self.about_menu.add_command(label="Sobre o Programa", command=self.mostrar_sobre)

        self.configuracoes = None

    def carregar_configuracoes(self):
        try:
            with open("config.json", "r") as config_file:
                self.configuracoes = json.load(config_file)
            print("Configurações carregadas com sucesso!")
            # imprimir as configurações carregadas
            print("----------------------------------------------------------")
            print("Configurações atuais:")
            print(json.dumps(self.configuracoes, indent=4))
            print("----------------------------------------------------------")             
        except Exception as e:
            print("Erro ao carregar as configurações:", e)

    def extrair_notas(self):
        print("Abrindo Firefox...aguarde...")
        if self.configuracoes:
            try:
                incluir_optativas = self.extrair_optativas.get()
                print(f"Incluir disciplinas optativas: {incluir_optativas}")
                ExtrairNotasSIGAA.extrair_notas_sigaa(self.configuracoes, incluir_optativas)
                print("----------------------------------------------------------")
                print("Notas extraídas com sucesso do SIGAA!")
            except Exception as e:
                print("Ocorreu um erro ao extrair as notas do SIGAA:", e)
        else:
            print("As configurações não foram carregadas. Carregue as configurações antes de extrair as notas.")

    def exportar_google_sheets(self):
        if self.configuracoes:
            try:
                DigitarDrive.atualizar_drive(self.configuracoes)
                print("----------------------------------------------------------")
                print("Notas dos alunos atualizadas com sucesso no Google Sheets!")
            except Exception as e:
                print("Ocorreu um erro ao exportar os dados para o Google Sheets:", e)
        else:
            print("As configurações não foram carregadas. Carregue as configurações antes de exportar para o Google Sheets.")

    def resumir_trimestres(self):
        if self.configuracoes:
            try:
                DigitarDrive.resumir_trimestres(self.configuracoes)
                print("----------------------------------------------------------")
                print("Trimestres resumidos com sucesso no Google Sheets!")
            except Exception as e:
                print("Ocorreu um erro ao resumir trimestres no Google Sheets:", e)
        else:
            print("As configurações não foram carregadas. Carregue as configurações antes de exportar para o Google Sheets.")

    def limpar_log(self):
        self.log_text.delete(1.0, tk.END)

    def mostrar_ajuda(self):
        help_text = (
            "Instruções de uso para cada Turma:\n\n"
            "1. Editar o arquivo \"alunos.txt\" com os nomes dos alunos da turma.\n"
            "2. Editar o arquivo \"config.json\" definindo o ano da turma e qual o link do Google Sheets.\n"
            "3. Carregar Configurações.\n"
            "4. Extrair Notas do SIGAA para CSV.\n"
            "5. Exportar Dados CSV para Google Sheets.\n"
            "6. Resumir Trimestres no Google Sheets.\n"
            "7. Use o menu 'Opções' para incluir/excluir disciplinas optativas da extração.\n"
        )
        messagebox.showinfo("Ajuda - Instruções de Uso", help_text)

    def mostrar_sobre(self):
        about_text = (
            "Sobre o Programa:\n\n"
            "Este é um software educacional experimental desenvolvido por Wanderson Rigo, professor do IFC - Campus Videira.\n\n"
            "O objetivo deste software é extrair as notas dos alunos do SIGAA e exportar para o Google Sheets, poupando trabalho manual e repetitivo.\n\n"
            "Versão: 1.0\n"
            "Desenvolvido com entusiasmo em: 2024\n"
            "Contato: wanderson.rigo@ifc.edu.br\n\n"
            "Este software é gratuito e de código aberto. Você pode modificar e distribuir este software, desde que mantenha a referência ao autor original.\n\n"
            "Use por sua conta e risco. Este software é fornecido sem garantia de qualquer tipo. O autor não se responsabiliza por qualquer dano causado pelo uso deste software."
        )
        messagebox.showinfo("Sobre o Programa", about_text)


class LogRedirector:
    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        self.widget.insert("end", text)
        self.widget.see("end")


def main():
    root = tk.Tk()
    app = InterfaceGrafica(root)
    root.mainloop()


if __name__ == "__main__":
    main()
