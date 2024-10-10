import gspread
from oauth2client.service_account import ServiceAccountCredentials
import csv
import json

def atualizar_drive(config):
    # Escopo da API e credenciais
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)

    SHEET_URL = config.get("SHEET_URL")
    STUDANTS_NOTES = config.get("STUDANTS_NOTES")

    print(f'URL da planilha: {SHEET_URL}')
    print(f'Notas no CSV: {STUDANTS_NOTES}')

    '''
    URL_DA_PLANILHA_1 = "https://docs.google.com/spreadsheets/d/1gdAZspevEDd-Nd58um7eEeCd16d9z4P390en9hqC31w"
    URL_DA_PLANILHA_2 = "https://docs.google.com/spreadsheets/d/1X8NEizq6U2OlOPGTH0oyKHWG68qTi3k6UAPqoWyraqE"
    URL_DA_PLANILHA_3 = "https://docs.google.com/spreadsheets/d/1h9hdgnKjfX4YC-8GRa9nomsZbayPBQnRYH3A7v2PQ38"
    '''

    # Autenticar e acessar a planilha
    gc = gspread.authorize(credentials)
    planilha = gc.open_by_url(SHEET_URL)

    #carregar lista de alunos do arquivo alunos.txt
    ALUNOS = {}
    # Abrir o arquivo CSV em modo de leitura
    with open(STUDANTS_NOTES, 'r') as arquivo_csv:
        leitor_csv = csv.DictReader(arquivo_csv)

        # Iterar sobre cada linha do arquivo CSV
        for linha in leitor_csv:
            nome_aluno = linha['Nome']
            disciplina = linha['Disciplina']
            nota1 = linha['Nota 1º trimestre']
            nota2 = linha['Nota 2º trimestre']
            nota3 = linha['Nota 3º trimestre']

            #troca . por , nas notas
            nota1 = nota1.replace('.', ',')
            nota2 = nota2.replace('.', ',')
            nota3 = nota3.replace('.', ',')
                    
            # Verificar se o aluno já está no dicionário
            if nome_aluno in ALUNOS:
                # Se o aluno já existe, adicione a disciplina e as notas à lista de disciplinas desse aluno
                ALUNOS[nome_aluno].append({'Disciplina': disciplina, 'Nota 1º trimestre': nota1, 'Nota 2º trimestre': nota2, 'Nota 3º trimestre': nota3})
            else:
                # Se o aluno não existe, crie uma nova entrada no dicionário com a disciplina e notas como a primeira entrada da lista
                ALUNOS[nome_aluno] = [{'Disciplina': disciplina, 'Nota 1º trimestre': nota1, 'Nota 2º trimestre': nota2, 'Nota 3º trimestre': nota3}]
        
        # Fechar o arquivo CSV
        arquivo_csv.close()

    #print(f'Lista de alunos carregada com sucesso! Total de {len(ALUNOS)} alunos.')
    #print(ALUNOS)


    #imprimir notas de todos os alunos
    def imprimir_notas_alunos():
        for aluno in ALUNOS:
            nome = aluno
            print(f'Aluno: {aluno}')
            print(f'Notas:')
            notas = ALUNOS[aluno]

            for nota in notas:
                print(f'Disciplina: {nota["Disciplina"]}')
                print(f'Nota 1º trimestre: {nota["Nota 1º trimestre"]}')
                print(f'Nota 2º trimestre: {nota["Nota 2º trimestre"]}')
                print(f'Nota 3º trimestre: {nota["Nota 3º trimestre"]}')
                print('---')

            print('---')

    def atualizarDrive():
        #para cada aluno em ALUNOS, recuperar as notas
        for aluno in ALUNOS:
            nome = aluno
            notas = ALUNOS[aluno]
            inserirNotas = []

            for nota in notas:
                disciplina = nota["Disciplina"]
                nota1 = nota["Nota 1º trimestre"]
                nota2 = nota["Nota 2º trimestre"]
                nota3 = nota["Nota 3º trimestre"]
                inserirNotas.append([nota1, nota2, nota3])

            #atualizar a planilha
            aba = planilha.worksheet(nome)
            intervalo = f'D6:F{5 + len(inserirNotas)}'
            aba.update(range_name=intervalo, values=inserirNotas, value_input_option='USER_ENTERED')
            print(f'Notas do(a) aluno(a) {nome} atualizadas com sucesso!')

    atualizarDrive()
    #print("Notas dos alunos atualizadas com sucesso no Google Sheets!")
    #imprimir_notas_alunos()

# Verifica se o arquivo está sendo executado diretamente
if __name__ == "__main__":
    try:
        with open("config.json", "r") as config_file:
            config = json.load(config_file)
            print("Configurações carregadas com sucesso!")
            atualizar_drive(config)
            print("----------------------------------------------------------")
            print("Notas dos alunos atualizadas com sucesso no Google Sheets!")
    except Exception as e:
        print("Erro:", e)

# email de compartilhamento da planilha com o serviço de credenciais é notas-24@planar-sunrise-422311-h9.iam.gserviceaccount.com