import tkinter as tk
from tkinter import filedialog, messagebox
import sys
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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
        
        self.file_menu = tk.Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="Ações", menu=self.file_menu)
        self.file_menu.add_command(label="1 - Carregar Configurações", command=self.carregar_configuracoes)
        self.file_menu.add_command(label="2 - Extrair Notas do SIGAA para CSV", command=self.extrair_notas)
        self.file_menu.add_command(label="3 - Exportar Dados CSV para Google Sheets", command=self.exportar_google_sheets)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Limpar Log", command=self.limpar_log)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Sair", command=root.quit)

        # Adiciona o menu de Ajuda
        self.help_menu = tk.Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="Ajuda", menu=self.help_menu)
        self.help_menu.add_command(label="Instruções", command=self.mostrar_ajuda)
        
        # Adiciona o menu de Sobre
        self.about_menu = tk.Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="Sobre", menu=self.about_menu)
        self.about_menu.add_command(label="Sobre o Programa", command=self.mostrar_sobre)

        self.configuracoes = None
    
    def carregar_configuracoes(self):
        try:
            with open("config.json", "r") as config_file:
                self.configuracoes = json.load(config_file)
            print("Configurações carregadas com sucesso!")
        except Exception as e:
            print("Erro ao carregar as configurações:", e)
    
    def extrair_notas(self):
        print("Abrindo Firefox...aguarde...")
        if self.configuracoes:
            try:
                ExtrairNotasSIGAA.extrair_notas_sigaa(self.configuracoes)
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

    def limpar_log(self):
        self.log_text.delete(1.0, tk.END)

    def mostrar_ajuda(self):
        help_text = (
            "Instruções de uso para cada Turma:\n\n"
            "1. Editar o arquivo \"alunos.txt\" com os nomes dos alunos da turma:\n\n"
            "2. Editar o arquivo \"config.json\" definindo o ano da turma e qual o link do Google Sheets.\n\n"
            "3. Carregar Configurações:\n"
            "   Clique em 'Ações' -> '1 - Carregar Configurações' para carregar as configurações do arquivo config.json.\n\n"
            "4. Extrair Notas do SIGAA para CSV:\n"
            "   Após carregar as configurações, clique em 'Ações' -> '2 - Extrair Notas do SIGAA para CSV' para iniciar a extração das notas.\n\n"
            "5. Exportar Dados CSV para Google Sheets:\n"
            "   Após extrair as notas, clique em 'Ações' -> '3 - Exportar Dados CSV para Google Sheets' para exportar os dados para o Google Sheets.\n\n"
            "6. Limpar Log:\n"
            "   Clique em 'Ações' -> 'Limpar Log' para limpar a área de log.\n\n"
            "7. Sair:\n"
            "   Clique em 'Ações' -> 'Sair' para fechar o aplicativo.\n\n\n"
            "As configurações no arquivo \"config.json\": são\n\n"
            "- \"URL\": \"https://sig.ifc.edu.br/sigaa/verTelaLogin.do\", URL do SIGAA.\n"
            "- \"USERNAME\": \"fulano.sobrenome\", nome de usuário do SIGAA.\n"
            "- \"PASSWORD\": \"senha\", a senha do SIGAA. Se não preencher aqui, uma caixa de diálogo vai pedir a senha.\n"
            "- \"STUDANTS_NAMES\": \"alunos.txt\", nome do arquivo que contém os nomes dos estudantes.\n"
            "- \"STUDANTS_CLASS\": \"PRIMEIRO_ANO\", o ano dos estudantes. Pode ser: PRIMEIRO_ANO, SEGUNDO_ANO, TERCEIRO_ANO.\n"
            "- \"STUDANTS_NOTES\": \"notas_alunos.csv\", o arquivo CSV onde as notas dos alunos serão salvas.\n"
            "- \"SHEET_URL\": \"https://docs.google.com/spreadsheets/etc.\", o link para a planilha de notas no Google Drive."
        )
        # Exibir uma caixa de mensagem com as instruções
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
        # Exibir uma caixa de mensagem com as informações sobre o programa
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
