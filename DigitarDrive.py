import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import *
import csv
import json

ALUNOS = {}
planilha = None

def resumir_trimestres(config):

    #criar uma aba com o resumo das notas para cada aluno. 
    # Notas de cada aluno na mesma linha. 
    # Disciplins em colunas
    def elaborarResumo(aba, trimestre):
        # Extrair os nomes das disciplinas para montar o cabeçalho
        disciplinas = []
        abreviaturaDisciplinas = []
        
        for info_disciplinas in ALUNOS.values():
            for info in info_disciplinas:
                if info['Disciplina'] not in disciplinas:
                    #considera só o texto depois - na disciplina
                    abreviatura = info['Disciplina'].split(' - ')[1].strip()
                    abreviaturaDisciplinas.append(abreviatura)
                    disciplinas.append(info['Disciplina'])

        # Cabeçalhos com 'Aluno' seguido das disciplinas
        header = ['Aluno'] + abreviaturaDisciplinas

        # Lista para armazenar todas as linhas que serão enviadas para a planilha
        dados_para_inserir = [header]  # Adiciona o cabeçalho como a primeira linha

        # Iterando sobre os alunos e suas disciplinas para preencher as notas
        for aluno, info_disciplinas in ALUNOS.items():
            linha = [aluno]  # Inicia a linha com o nome do aluno
            notas_aluno = {disc['Disciplina']: disc[f'Nota {trimestre}º trimestre'] for disc in info_disciplinas}

            # Adiciona as notas para as disciplinas na ordem do cabeçalho
            for disciplina in disciplinas:
                linha.append(notas_aluno.get(disciplina, ''))  # Se não houver nota, coloca em branco

            dados_para_inserir.append(linha)


        aba.update(values=dados_para_inserir, range_name='A1', value_input_option='USER_ENTERED')


    def criar_grafico_pizza_notas(aba):
        # Obtém os dados na faixa 'B1:P36'
        notas = aba.get('B2:P36')
    
        # Contagem dos alunos em cada faixa de nota
        acima_5_9 = 0
        entre_4_e_5_9 = 0
        abaixo_4 = 0
    
        for linha in notas:
            for nota in linha:
                try:
                    nota = float(nota.replace(',', '.'))
                    if nota > 5.9:
                        acima_5_9 += 1
                    elif 5 <= nota <= 5.9:
                        entre_4_e_5_9 += 1
                    elif nota < 5:
                        abaixo_4 += 1
                except ValueError:
                    continue  # Ignora células vazias ou valores não numéricos
    
        # Prepara os dados para o gráfico
        categorias = ["Acima de 5.9", "Entre 5 e 5.9", "Abaixo de 5"]
        valores = [acima_5_9, entre_4_e_5_9, abaixo_4]
    
        # Insere os dados na planilha para o gráfico
        aba.update(values=[["Categoria", "Quantidade"]] + list(zip(categorias, valores)), range_name='R1')
        
        # Configura o gráfico de pizza
        chart = {
            'spec': {
                'title': 'Distribuição de Notas dos Alunos',
                'pieChart': {
                    'legendPosition': 'RIGHT_LEGEND',
                    'domain': {
                        'sourceRange': {
                            'sources': [{'sheetId': aba.id, 'startRowIndex': 1, 'endRowIndex': 4, 'startColumnIndex': 17, 'endColumnIndex': 18}]
                        }
                    },
                    'series': {
                        'sourceRange': {
                            'sources': [{'sheetId': aba.id, 'startRowIndex': 1, 'endRowIndex': 4, 'startColumnIndex': 18, 'endColumnIndex': 19}]
                        }
                    }
                }
            },
            'position': {
                'overlayPosition': {
                    'anchorCell': {'sheetId': aba.id, 'rowIndex': 5, 'columnIndex': 19}
                }
            }
        }
        
        # Adiciona o gráfico ao sheet
        aba.spreadsheet.batch_update({
            'requests': [{
                'addChart': {'chart': chart}
            }]
        })

    def definir_formatacao(aba):
        # Define o intervalo em que deseja aplicar a formatação condicional (exemplo: notas da célula B2 até F10)
        range_notas = 'B1:P36'
        # Criando as regras de formatação condicional
        # 1. Formatação para células com valor > 7 (Verde)
        rule_mais_6 = ConditionalFormatRule(
            ranges=[GridRange.from_a1_range(range_notas, aba)],
            booleanRule=BooleanRule(
                condition=BooleanCondition('NUMBER_GREATER', ['5,9']),# Cor de fundo verde                
                format=CellFormat(backgroundColor=Color(0, 1, 0))
        ))

        # 2. Formatação para células com valor < 5 (Vermelho)
        rule_menos_5 = ConditionalFormatRule(
            ranges=[GridRange.from_a1_range(range_notas, aba)],
            booleanRule=BooleanRule(
                condition=BooleanCondition('NUMBER_LESS', ['5']), # Cor de fundo vermelho forte                
                format=CellFormat(backgroundColor=Color(1, 0, 0))
            )
        )

        # 3. Formatação para células com valor entre 5 e 5.9 (Amarelo)
        rule_entre_5_59 = ConditionalFormatRule(
            ranges=[GridRange.from_a1_range(range_notas, aba)],
            booleanRule=BooleanRule(
                condition=BooleanCondition('NUMBER_BETWEEN', ['5', '5,90']),  # Note que o segundo valor é '5.90'
                format=CellFormat(backgroundColor=Color(1, 1, 0))  # Cor de fundo amarelo claro
            )
        )

        # Define o formato de alinhamento para centralizar
        center_format = CellFormat(
            horizontalAlignment='CENTER'  # Centraliza o conteúdo horizontalmente
        )

        # Aplica o formato de centralização ao intervalo especificado
        format_cell_range(aba, range_notas, center_format)

        # Adicionando as regras à planilha
        rules = get_conditional_format_rules(aba)
        rules.append(rule_mais_6)
        rules.append(rule_entre_5_59)
        rules.append(rule_menos_5)
        rules.save()

        # Obtém o ID da planilha
        sheet_id = aba.id
        
        # Configurações para orientação vertical
        request = {
            "requests": [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 0,  # Linha do cabeçalho
                            "endRowIndex": 1,    # Apenas a primeira linha
                            "startColumnIndex": 1,  # Coluna B
                            "endColumnIndex": 16    # Coluna P
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "textRotation": {
                                    "angle": 90  # Orienta o texto verticalmente
                                },
                                "horizontalAlignment": "CENTER",  # Centraliza o texto horizontalmente
                                "verticalAlignment": "BOTTOM",  # Alinha o texto na parte inferior da célula
                                "textFormat": {
                                    "bold": True  # Define o texto em negrito
                                }
                            }
                        },
                        "fields": "userEnteredFormat(textRotation,horizontalAlignment,verticalAlignment)"
                    }
                }
            ]
        }
        
        # Aplica a formatação
        aba.spreadsheet.batch_update(request)

    # Função para garantir a existência de uma aba
    def obter_ou_criar_aba(nome_aba):
        try:
            aba = planilha.worksheet(nome_aba)
            #remover a aba para criar uma nova
            planilha.del_worksheet(aba)
            aba = planilha.add_worksheet(title=nome_aba, rows="50", cols="30")
        except gspread.exceptions.WorksheetNotFound:
            aba = planilha.add_worksheet(title=nome_aba, rows="50", cols="30")
        return aba

    def resumir_trimestres():
        # repita para as 3 abas resumo1, resumo2 e resumo3
        for i in range(1, 4):
            aba = obter_ou_criar_aba(f"resumo{i}")
            elaborarResumo(aba,i)
            definir_formatacao(aba)
            criar_grafico_pizza_notas(aba)
            aba.update_index(i-1)
            print(f"Resumo {i} atualizado com sucesso!")

    atualizar_drive(config)
    resumir_trimestres()

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
    global planilha
    planilha = gc.open_by_url(SHEET_URL)

    #carregar lista de alunos do arquivo alunos.txt
    global ALUNOS
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
    

# Verifica se o arquivo está sendo executado diretamente
if __name__ == "__main__":
    try:
        with open("config.json", "r") as config_file:
            config = json.load(config_file)
            print("Configurações carregadas com sucesso!")
            #atualizar_drive(config)
            #resumir_trimestres(config)
            print("----------------------------------------------------------")
            print("Notas dos alunos atualizadas com sucesso no Google Sheets!")
    except Exception as e:
        print("Erro:", e)

# email de compartilhamento da planilha com o serviço de credenciais é notas-24@planar-sunrise-422311-h9.iam.gserviceaccount.com