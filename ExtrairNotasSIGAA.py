import csv
import json
from tkinter import simpledialog
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

def pedir_senha():
        senha = simpledialog.askstring("Senha", "Digite sua senha do SIGAA:", initialvalue="", show='*')
        return senha

def extrair_notas_sigaa(config, incluir_optativas, status_lista):
    URL = config.get("URL")
    USERNAME = config.get("USERNAME")
    PASSWORD = config.get("PASSWORD")
    STUDANTS_CLASS = config.get("STUDANTS_CLASS")

    #se não tem password, deve pedir
    if not PASSWORD:
        PASSWORD = pedir_senha()

    STUDANTS_NAMES = config.get("STUDANTS_NAMES")
    STUDANTS_NOTES = config.get("STUDANTS_NOTES")
    REAV_NOTES = config.get("REAV_NOTES")


    # Defina as variáveis
    '''
    URL = "https://sig.ifc.edu.br/sigaa/verTelaLogin.do"
    USERNAME = "wanderson.rigo"
    PASSWORD = ""
    STUDANTS_NAMES = "alunos3.txt"
    STUDANTS_NOTES = "notas_alunos.csv"
    '''

    ALUNOS = []
    # Carregar lista de alunos do diretório nomes/alunos.txt
    with open(STUDANTS_NAMES, "r", encoding="utf-8") as arquivo:
        for linha in arquivo:
            ALUNOS.append(linha.strip())

    # Inicialize o driver do navegador
    browser = webdriver.Firefox()
    browser.maximize_window()  # Maximize a janela do navegador
    browser.get(URL)

    # Faça login
    username_field = browser.find_element(By.NAME, "user.login")
    username_field.send_keys(USERNAME)
    password_field = browser.find_element(By.NAME, "user.senha")
    password_field.send_keys(PASSWORD)
    password_field.send_keys(Keys.RETURN)

    # Aguarde até que a página da tabela seja carregada
    WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.XPATH, "//table[@class='listagem table tabela-selecao-vinculo']")))

    # Navegue até a página desejada
    browser.find_element(By.XPATH, "//a[@class='withoutFormat' and contains(text(),'Secretário')]").click()
    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "//li[@class='tecnicoIntegrado on']")))
    browser.find_element(By.XPATH, "//a[@href='/sigaa/verMenuTecnicoIntegrado.do']").click()

    # Ciente
    browser.find_element(By.CSS_SELECTOR, "button.btn.btn-primary").click()

    try:   
        # problema no perfil da Rosicler: Secretario
        browser.find_element(By.XPATH, "//span[text()='Aluno']").click()
        #browser.find_element(By.ID, "menuTecnicoForm:emitirBoletimTecnicoDisc").click()
        browser.find_element(By.XPATH,"//a[normalize-space(text())='Emitir Boletim']").click()
    except NoSuchElementException:
        # meu perfil de secretário
        #browser.find_element(By.ID, "menuTecnicoForm:emitirBoletimTecnicoDiscMenuPedagogico").click()
        print("Erro: não conseguiu clicar em Aluno/Emitir Boletim!")


    def pegar_notas_aluno(aluno):

        # está tando erro arqui! Tem um voltar para o segundo ano!
        
        #Se não for do primeiro ano
        '''
        if browser.find_element(By.ID, "form:voltar"):
            browser.find_element(By.ID, "form:voltar").click()
        '''

        browser.find_element(By.ID, "formulario:nomeDiscente").clear()
        browser.find_element(By.ID, "formulario:nomeDiscente").send_keys(aluno)
        browser.find_element(By.ID, "formulario:buscar").click()

        # Selecionar todas as linhas da tabela com as classes 'linhaPar' ou 'linhaImpar',
        # excluindo 'curso' e linhas que contêm um <td> com a classe 'detalhesDiscente'
        linhas = browser.find_elements(By.XPATH, 
        "//table[@class='listagem']/tbody/tr[not(contains(@class, 'curso')) and " \
        "(contains(@class, 'linhaPar') or contains(@class, 'linhaImpar')) and " \
        "not(td[contains(@class, 'detalhesDiscente')])]")

        # Iterar sobre cada linha da tabela
        for linha in linhas:
            # Obter todas as células (td) da linha
            celulas = linha.find_elements(By.XPATH, ".//td")

            #se o tamanho for maior que 1, tem dados
            if len(celulas) > 1:
                nome = celulas[3].text # nome do aluno
                print(nome)

                status = celulas[5].text # situação do aluno
                print(status)

                ação = celulas[6] # ação de clicar no aluno

                #se o nome do aluno for igual ao nome do aluno que está sendo procurado, e ele for ATIVO clicar nele
                if (nome == aluno) and (status in  status_lista):
                    botoes_selecionar = ação.find_element(By.XPATH, ".//input[@title='Selecionar Discente']")
                    botoes_selecionar.click()
                    break
                else:
                    print("Aluno não encontrado ou não está ativo!")

        '''

        ativos = browser.find_elements(By.XPATH, "//td[contains(text(), 'ATIVO')]")
        tamanho = len(ativos)

        #se tamanho maior que 1, tem mais de um aluno com o mesmo nome ou bem parecido
        if tamanho > 1:
            # Iterar sobre cada elemento "ATIVO"
            for ativo in ativos:
                # Encontrar todos os irmãos anteriores do elemento "ATIVO"
                td_irmaos_anteriores = ativo.find_elements(By.XPATH, "preceding-sibling::*")

                nome = td_irmaos_anteriores[3].text # nome do aluno

                if nome == aluno:
                    print("Clicando no botão de ação")
                    acao = ativo.find_elements(By.XPATH, "//following-sibling::td/input")
                    acao.click()
                    break
        else:
            browser.find_element(By.XPATH, "//td[contains(text(), 'ATIVO')]/following-sibling::td/input").click()

        '''

        print(f"Extraindo notas do aluno: {aluno}")

        '''
        if STUDANTS_CLASS   == "PRIMEIRO_ANO":
            print("1º ano!")

           
            1º ano só tem um voltar no SIGAA via

            <td class="voltar" align="left"><a href="javascript:history.back();"> Voltar </a></td>

            ---------------------------------------------

            2º ano NORMAL tem 2 voltar no SIGAA via

            <td class="voltar" align="left"><a href="javascript:history.back();"> Voltar </a></td>

            <input id="form:voltar" type="submit" name="form:voltar" value="<< Voltar">

            ---------------------------------------------

            2º ano VINDA DE FORA tem 1 voltar no SIGAA via

            <td class="voltar" align="left"><a href="javascript:history.back();"> Voltar </a></td>

            -------------------------------------------------

            3º ano NORMAL tem 2 voltar no SIGAA via

            <td class="voltar" align="left"><a href="javascript:history.back();"> Voltar </a></td>


            <input id="form:voltar" type="submit" name="form:voltar" value="<< Voltar">


            
        elif STUDANTS_CLASS == "SEGUNDO_ANO":
            print("2º ano!")
        elif STUDANTS_CLASS == "TERCEIRO_ANO":
            print("3º ano!")
        '''

        try:
            # se veio de outro IF, não tem MATRICULADO, só ATIVO, então já vai para o boletim
            relatorio = browser.find_element(By.ID, "relatorio")                    
        except NoSuchElementException:
            # se é REPROVADO
            elemento_matriculado = browser.find_element(By.XPATH, "//td[contains(text(), 'MATRICULADO')]/following-sibling::td/a")
            elemento_matriculado.click()
        else:
            print("Veio de outro IF")

        '''

        try:
            #se for do primeiro ano abre o boletim direto, senão precisa selecionar o ano / MATRICULADO 
            if browser.find_element(By.CLASS_NAME, "infoAltRem"):
                print("não é do primeiro ano! Precisa selecionar o ano!")
                # Localize o elemento com base no texto "MATRICULADO" dentro da tag <td>
                elemento_matriculado = browser.find_element(By.XPATH, "//td[contains(text(), 'MATRICULADO')]/following-sibling::td/a")
                elemento_matriculado.click()
            else:
                print("é do primeiro ano!")
        except Exception as e:
            print("Erro ao verificar se é do primeiro ano:", e)

        '''

        # Aguarde até que a página da tabela seja carregada
        WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.listagem[border='1']")))

        # Obtenha as linhas da tabela
        tabelasDeNotas = browser.find_elements(By.CSS_SELECTOR, "table.listagem[border='1']")

        notas_aluno = {'Nome': aluno}

        num_tabelas = len(tabelasDeNotas)

        if num_tabelas == 3:
            tabelaObrigatorias = tabelasDeNotas[0]
            tabelaOptativas = tabelasDeNotas[1]
            tabelaOutras = tabelasDeNotas[2]
        elif num_tabelas == 2:
            tabelaObrigatorias = tabelasDeNotas[0]
            tabelaOptativas = tabelasDeNotas[1]
            tabelaOutras = None
        else:  
            tabelaObrigatorias = tabelasDeNotas[0]
            tabelaOptativas = None
            tabelaOutras = None

        linhasTabelaObrigatorias = tabelaObrigatorias.find_elements(By.TAG_NAME, "tr")
        linhasTabelaObrigatorias = linhasTabelaObrigatorias[1:-2]  # Ignorar cabeçalho e linhas de faltas

        notasObrigatorias = {}
        for linha in linhasTabelaObrigatorias:
            tds = linha.find_elements(By.TAG_NAME, "td")
            disciplina = tds[0].text

            # se há aprovação anterior, deve ignorar as notas
            # aqui testa se o indice tds[3] existe, que seria a nota do 1º tri
            try:
                tds[3] #se não existir, igorna as notas disciplina
            except IndexError:
                continue
            
            # se no boletim tem reav a cada avaliação, como em Fraiburgo
            if REAV_NOTES == "durante":
                print("Reav a cada avaliação, como em Fraiburgo")
                # se não há reprovação
                n1tri = tds[7].text
                n2tri = tds[14].text
                n3tri = tds[21].text

            # se tem reav só no final, como Videira
            if REAV_NOTES == "final":
                print("Reav ao final do trimestre, como em Videira")
                # se não há reprovação
                n1tri = tds[3].text
                n2tri = tds[6].text
                n3tri = tds[9].text

                '''' colunas no boletim de Videira

                    0 - COMPONENTE CURRICULAR
                    1 - Trimestre 1
                    2 - Trimestre 1 - Reavaliação
                    3 - Trimestre 1 - Média Parcial
                    4 - Trimestre 2 	
                    5 - Trimestre 2 - Reavaliação
                    6 - Trimestre 2 - Média Parcial
                    7 - Trimestre 3 	
                    8 - Trimestre 3 - Reavaliação
                    9 - Trimestre 3 - Média Parcial
                    10 - Média Final 	
                    11 - Faltas
                    12 - Situação

                '''

            notasObrigatorias[disciplina] = {'Nota 1º trimestre': n1tri, 'Nota 2º trimestre': n2tri, 'Nota 3º trimestre': n3tri}

        
        notasOptativas = {}

        '''  inicio optativas '''

        # se tem optativas
        if tabelaOptativas != None:
            linhasTabelaOptativas = tabelaOptativas.find_elements(By.TAG_NAME, "tr")
            linhasTabelaOptativas = linhasTabelaOptativas[1:]  # Ignorar cabeçalho
        
            for linha in linhasTabelaOptativas:
                tds = linha.find_elements(By.TAG_NAME, "td")
                disciplina = tds[0].text
                ths = linha.find_elements(By.TAG_NAME, "th")

                # se o nome da disciplina for INB0702-1 - INGLÊS - BÁSICO I if disciplina.startswith("INB0702-1 - "):
                if len(tds) == 3:
                    n1tri = ths[9].text
                    notasOptativas[disciplina] = {'Nota 1º trimestre': n1tri}
                else:
                    n1tri = tds[1].text
                    n2tri = tds[4].text
                    n3tri = tds[7].text
                    notasOptativas[disciplina] = {'Nota 1º trimestre': n1tri, 'Nota 2º trimestre': n2tri, 'Nota 3º trimestre': n3tri}
        
       
        ''' fim optativas '''

        notas_aluno['Optativas'] = notasOptativas
        notas_aluno['Obrigatorias'] = notasObrigatorias

        return notas_aluno

    # Pegar notas de todos os alunos
    notas_alunos = []
    for aluno in ALUNOS:
        notas_aluno = pegar_notas_aluno(aluno)
        #pegou as notas do 
        
        print(f'Notas do aluno {aluno} capturadas com sucesso!')

        # imprimindo as optativas do aluno
        #imprimirOptativas(notas_aluno)


        notas_alunos.append(notas_aluno)

        if STUDANTS_CLASS == "PRIMEIRO_ANO":
            #print("voltando das notas do 1º ano!")
            browser.find_element(By.XPATH, "//td[@class='voltar']").click()

            try:
                #se reprovou tem mais um voltar
                voltar = browser.find_element(By.ID, "form:voltar")                    
            except NoSuchElementException:
                pass
            else:
                voltar.click()

        elif STUDANTS_CLASS == "SEGUNDO_ANO" or STUDANTS_CLASS == "TERCEIRO_ANO":
            #print("voltando 2º ou 3º ano!")

            #se veio de outro IF, só tem um voltar
            browser.find_element(By.XPATH, "//td[@class='voltar']").click()

            try:
                #mas se não veio de outro IF, é só de videira, tem mais um voltar
                voltar = browser.find_element(By.ID, "form:voltar")                    
            except NoSuchElementException:
                pass
            else:
                #print("não veio de outro IF, é do IFC videira")
                voltar.click()

        '''
        try:
            #se for 2º ou 3º ano volta 2 vezes até página de seleção de aluno
            if browser.find_element(By.XPATH, "//td[@class='voltar']"):
                browser.find_element(By.XPATH, "//td[@class='voltar']").click()
                browser.find_element(By.ID, "form:voltar").click()
            else:
                #se for do 1º ano volta 1 vez até página de seleção de aluno
                browser.find_element(By.ID, "form:voltar").click()
        except Exception as e:
            print("Erro ao verificar se é do primeiro ano:", e)

        '''
        '''
        #Se for do primeiro ano
        browser.find_element(By.XPATH, "//td[@class='voltar']").click()

        #Se não for do primeiro ano
        browser.find_element(By.ID, "form:voltar").click()
        '''

    # Salvar as notas dos alunos em um arquivo CSV
    with open(STUDANTS_NOTES, 'w', newline='') as arquivo_csv:
        escritor_csv = csv.DictWriter(arquivo_csv, fieldnames=['Nome', 'Disciplina', 'Nota 1º trimestre', 'Nota 2º trimestre', 'Nota 3º trimestre'])
        escritor_csv.writeheader()
        
        for aluno in notas_alunos:

            for disciplina, notas in aluno['Obrigatorias'].items():
                escritor_csv.writerow({'Nome': aluno['Nome'], 'Disciplina': disciplina, **notas})# desempacotamento de dicionário

            if incluir_optativas:
                for disciplina, notas in aluno['Optativas'].items():
                    # Marcar as optativas um '#' no início
                    nome_da_disciplina = ' # ' + disciplina
                    escritor_csv.writerow({'Nome': aluno['Nome'], 'Disciplina': nome_da_disciplina, **notas})# desempacotamento de dicionário

    print("Notas salvas com sucesso no arquivo 'notas_alunos.csv'!")
    browser.quit()

def imprimirOptativas(notas_aluno):
    #nome do aluno
    print(f'Nome do aluno: {notas_aluno["Nome"]}')
    for disciplina, notas in notas_aluno['Optativas'].items():
        print(f'{disciplina}: {notas}')

# Verifica se o arquivo está sendo executado diretamente
if __name__ == "__main__":
    try:
        with open("config.json", "r") as config_file:
            config = json.load(config_file)
            print("Configurações carregadas com sucesso!")
            extrair_notas_sigaa(config)
        print("Notas extraídas do SIGAA com sucesso!")
    except Exception as e:
        print("Erro ao carregar as configurações:", e)