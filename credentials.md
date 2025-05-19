Para alterar dados em planilhas do Google via um programa, o programa precisa de autorização para agir como um usuário com poderes de edição. Para isso, um e-mail de conta de serviço que representa o programa deve ser incluído com permissão de editor na planilha do Google Drive. Esse e-mail é gerado quando criamos uma conta de serviço no Google Cloud.
Passos:

    1) Acesse o Google Cloud Console em  https://console.cloud.google.com/apis/credentials?authuser=1

    2) Crie uma conta de serviço:

        Acessar na lateral esquerda o menu "Credenciais"

        Depois clicar em "+ CRIAR CREDENCIAIS"

        E ainda clica em "Contas de Serviço"

        Daí preencha os campos necessários.

    2.5) Ative a Google Sheets API para o projeto:

        Acesse o link: https://console.developers.google.com/apis/api/sheets.googleapis.com/overview?project=[ID do projeto]

        Substitua [ID do projeto] pelo ID do seu projeto no Google Cloud (geralmente aparece no topo da página do Console).

        Clique no botão “Ativar” para ativar a API.

        Este passo é essencial para que o programa consiga acessar planilhas usando a API.

    3) Gere uma chave de credenciais no formato JSON:

        Depois de criar a conta de serviço, clique nela.

        Vai aparecer algumas abas, clique em "CHAVES"

        Vá na caixa de combinação escrita "ADICIONAR CHAVE". Clique nela e selecione "Criar nova chave"

        O tipo de chave é JSON

        Aí vai gerar um arquivo com um nome enorme e terminado em ".json"

        Este arquivo deve ser renomeado para "credentials.json" e colocando junto do arquivo ".exe" do programa

        Ele será usada pelo programa para se autenticar com as APIs do Google.

    4) Use o e-mail associado à conta de serviço (visível na página de Contas de Serviço) e adicione-o como editor na sua planilha do Google Drive:
        Abra a planilha no Google Drive.

        Clique no botão Compartilhar no canto superior direito.

        Adicione o e-mail da conta de serviço com permissão de Editor.

Agora o programa estará autorizado a acessar e modificar a planilha.

Se seguir os passos corretamente, o programa poderá usar as APIs do Google (como Google Sheets API) para inserir, editar ou ler dados da planilha.