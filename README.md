# Notas

Este sistema foi desenvolvido pelo prof. Wanderson Rigo do IFC-Videira para capturar as notas dos alunos no SIGAA, gerar um arquivo CSV e então repassar tais notas para planilhas do Google Drive, poupando trabalho repetivivo de acesso e digitação manual de cada nota.

O arquivo CSV gerado pode ser manipulado de forma independente para geração de estatíticas, por exemplo.

Para usar o sistema é preciso ter acesso de *Secretário* no SIGAA e também se definir algumas configurações.

O arquivo credentials.json foi gerado pelo Google Console.

Para a extração das notas de cada turma, não esqueça de atualizar os nomes dos alunos, e o link para a planilha que será usada no conselho de classe.

## Detalhes técnicos

Programado em Python valendo-se principalmente das APIs selenium e gspread.

Criei em https://console.cloud.google.com/apis/credentials?authuser=1 uma conta de serviço e credenciais do cliente OAuth 2.0  para acessar as APIs do Google. 

O e-mail da conta de serviço criado é notas-24@planar-sunrise-422311-h9.iam.gserviceaccount.com Ele deve ser incluído com acesso de editor na planilha do Google Drive para que as notas possam ser inseridas.

Para maiores informações de como criar o arquivo *credentials.json* ver [aqui](https://github.com/wanderson-rigo/FerramentaExtracaoNotas/blob/main/credentials.md)

As configurações a serem definidas no arquivo *config.json* são:

- "URL": "https://sig.ifc.edu.br/sigaa/verTelaLogin.do", do SIGAA
- "USERNAME": "fulano.sobrenome", nome de usuário do SIGAA
- "PASSWORD": "senha", a senha do SIGAA. Se não preencher aqui, uma caixa de diálogo vai pedir a senha.
- "STUDANTS_NAMES": "alunos.txt", nome do arquivo que contém os nomes dos estudantes.
- "STUDANTS_CLASS": "PRIMEIRO_ANO", o ano dos estudantes. Pode ser: PRIMEIRO_ANO, SEGUNDO_ANO, TERCEIRO_ANO.
- "STUDANTS_NOTES": "notas_alunos.csv", 
- "REAV_NOTES": "durante", define quando ocorre a Reavaliação. Pode ser: "final" ou "durante". Ex: em Videira é no final e em Fraiburgo é durante o trimestre. Esta configuração orienta o programa a buscar as notas nas colunas corretas no boletim, acessando, se "final", a 3ª, 6ª e 9ª colunas ou se "durante", as colunas 7ª, 14ª e 21ª do boletim. A coluna 0 (zero) é do COMPONENTE CURRICULAR.
- "SHEET_URL": "https://docs.google.com/spreadsheets/d/1gdAZspevEDd-Nd58um7eEeCd16d9z4P390en9hqC31w", o link para a planilha de notas no Google Drive

As cópias das planilhas estão em https://drive.google.com/drive/u/1/folders/1wUuauEh7SkCT5xZWckkaOf3wm87MgNXJ