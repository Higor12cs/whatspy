## whatspy
Usando Python para automatizar o envio de mensagens no WhatsApp de modo oculto.

## Objetivo
Basicamente meu objetivo era automatizar o envio de imagens armazenadas no Excel,  tudo isso de forma oculta, para que não atrapalhe o uso do computador no tempo de execução do script.

## Configurações Iniciais
Nesse programa eu automatizei o envio de imagens no WhatsApp usando o Excel para obter os dados. Dentro do arquivo .xls eu tenho uma lista de contatos, e ao lado de cada contato eu tenho o nome da figura que será enviada, tudo na mesma planilha. 

Dito isso, o programa precisa dos seguintes dados:

## Caminho do GeckoDriver

Isso já está pré-definido dentro do código como: `driver_path = str(Path(__file__).parent.absolute()) + r"\geckodriver.exe"`

Mas caso precise alterar o esse caminho por efeitos de atualizações ou outras necessidades, essa é a variável a ser alterada.

**Observação:**

Deixei uma cópia do GeckoDriver dentro da pasta código, até a data dessa publicação, essa versão estava funcionado perfeitamente, tentarei manter esse repositório sempre atualizado para demais situações.

## Caminho do perfil no Firefox
Aqui é onde fixamos um perfil do navegador que o código vai usar sempre. Isso dribla a necessidade de inserir o QR code do WhatsAppWeb toda vez que o programa rodar. Automatizar esse processo foi um pouco difícil, tentei várias vezes com o Chrome, mas a minha ultima solução foi o Firefox. 

Fazer essa configuração é simples, primeiro você instala o Firefox normalmente. Feito isso, pressione as teclas `Windows + R` e digite `firefox.exe -p` Isso ira abrir o gerenciador de perfis do Firefox. Crie um novo perfil com o nome de sua preferencia e inicie o navegador, deixe a opção `Usar o perfil selecionado sem perguntar ao iniciar` marcada. Após abrir o navegador, faça login no WhatsAppWeb normalmente.

Agora, pressione novamente `Windows + R` e digite `%appdata%`. Dentro da pasta `\Mozilla\Firefox\Profiles` você ira encontrar uma pasta com o nome de perfil que voce acabou de criar. Entre nela e copie o caminho e substitua a variável `profile_path` no código. Ela ficara mais ou menos assim: `profile_path = r"C:\Users\<seu_user>\AppData\Roaming\Mozilla\Firefox\Profiles\<seu_profile>"`

## Caminho do arquivo Excel
Dentro do código nós temos a variável: `planilha_contatos = str(Path(__file__).parent.absolute()) + r"\<sua_planilha>.xlsx"` Essa variável aponta pra um arquivo .xls dentro da mesma pasta do código, sinta-se a vontade para alterar o caminho até o seu arquivo, mas para manter isso simples, é só colocar a sua planilha no mesmo diretório e trocar o `<sua_planilha>.xlsx` pelo nome do seu arquivo Excel.

Agora precisamos indicar em qual "aba" o programa ira procurar pelos dados. Então seguindo a mesma metodologia de antes, substitua o atributo da variável `aba_planilha` pelo nome da sua aba dentro da sua planilha. 

Essa variável deve ficar parecida com isso: `aba_planilha = "<aba>"`

## Guiando o programa dentro do arquivo do Excel
Aqui iremos indicar para o programa onde começar a ler o nome das figuras e dos contatos que serão usados para enviar as imagens. Essa parte funciona mais ou menos assim: 
|-|*ColunaA*|*ColunaB*|
|--|--|--|
|***Linha1*** |**InicioContatos**|**InicioFiguras**|
|***Linha2***| Contato1 | Figura1 |
|***Linha3***| Contato2 | Figura2 |
|***Linha4***| Contato3 | Figura3 |
|***Linha5***| | |

Pegando a planilha acima de exemplo, nossa lista de contatos começa na coluna A, na linha 2. Com isso, informamos essas coordenadas para o código e ele seguira lendo essa coluna até achar um espaço vazio, esse espaço vazio definirá o fim da lista. 

O mesmo acontece com as figuras, que por sua vez, começam na coluna B e na linha 2.

O nosso programa entende o número de linhas perfeitamente, mas não entende as colunas como letras, e sim, como números também. Então nesse caso, substituímos a coluna A por 1, a coluna B por 2 e assim por diante.

E então, dentro do código, devemos atribuir as variáveis `coluna_contatos`, `linha_contatos`, `coluna_figuras`, `linha_figuras` seus respectivos valores.

Exemplo:

    coluna_contatos = 1  #coluna A
    linha_contatos = 2  #linha 2
    coluna_figuras = 2  #coluna B
    linha_figuras = 2  #linha 2

Observação:

O número de figuras deve ser correspondente ao número de contatos, e visse versa. Caso não seja o caso, encontraremos erros no código.

## Rodando o código
Para rodar o código é simples, basta executar o arquivo e aguardar sua execução terminar. O código também gera um arquivo `historico.log` que guarda todo o log do processo. 

## É isso!
Sei que meu código está longe de ser bom, mas ainda estou aperfeiçoando, e também é bom lembrar que sou apenas um iniciante em Python. 
Qualquer dúvida ou sugestão é só entrar em contato. Obrigado!
