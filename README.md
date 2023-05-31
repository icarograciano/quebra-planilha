Para acessar o proejeto no ar acesse: https://quebra-planilha.netlify.app/

#Introdução

Este é o README do projeto "Quebra Planilha", que contém o código-fonte para um aplicativo web que permite quebrar uma planilha do Excel em várias planilhas menores com base em um número específico de linhas. Neste documento, vamos explorar os recursos utilizados, explicar o que o código faz e como ele funciona.

#Métodos

O código utiliza os seguintes métodos e recursos:

HTML e CSS: O código utiliza a linguagem de marcação HTML para estruturar o conteúdo da página e o CSS para aplicar estilos e layout. A estrutura do aplicativo é organizada em uma div principal com classes e IDs específicos para os elementos interativos.

JavaScript: O código incorpora JavaScript para adicionar funcionalidades ao aplicativo. Ele utiliza a biblioteca XLSX.js para manipular arquivos do Excel e quebrar a planilha em planilhas menores.

Bootstrap: O código faz uso do framework CSS Bootstrap para fornecer um estilo visual moderno e responsivo ao aplicativo. Os links para os arquivos CSS e JavaScript do Bootstrap são incluídos para carregar os estilos e funcionalidades correspondentes.

#Desenvolvimento

O aplicativo permite que o usuário selecione um arquivo da planilha do Excel (.xlsx) por meio de um campo de entrada de arquivo. Em seguida, o usuário informa a quantidade desejada de linhas para quebra da planilha e o nome da aba específica da planilha. Ao clicar no botão "Quebrar Planilha", o JavaScript é acionado.

No JavaScript, o código verifica se um arquivo foi selecionado. Se sim, ele utiliza a biblioteca FileReader para ler o arquivo como um array de bytes. Em seguida, o código utiliza a biblioteca XLSX.js para analisar o arquivo e extrair a planilha especificada pela aba informada.

Após obter a planilha, o código calcula o número de planilhas necessárias com base na quantidade de linhas informadas pelo usuário e na quantidade total de linhas na planilha. Em seguida, ele itera por cada planilha necessária, copia a quantidade de linhas especificada para uma nova planilha e cria um novo arquivo do Excel (.xlsx) com o conteúdo da planilha menor.

Cada novo arquivo do Excel é disponibilizado para download por meio de um link simulado. O arquivo é gerado dinamicamente como um arquivo binário (Blob) e o link de download é criado com o URL correspondente. O usuário pode clicar no link para baixar cada arquivo gerado.

#Conclusão

O aplicativo "Quebra Planilha" é uma ferramenta simples, mas útil, para dividir uma planilha do Excel em planilhas menores. Ele facilita o processo de trabalhar com grandes volumes de dados, permitindo que o usuário especifique a quantidade desejada de linhas para cada planilha gerada.

#Autor

Este aplicativo foi desenvolvido por Icaro Graciano, um entusiasta de programação apaixonado por desenvolvimento web. Para entrar em contato, envie um e-mail para icaro.graciano@gmail.com.
