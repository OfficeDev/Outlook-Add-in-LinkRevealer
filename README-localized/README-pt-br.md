---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 8/27/2015 1:08:49 PM
---
# Suplemento do Outlook: Suplemento de e-mail para um cenário de leitura que localiza e analisa todos os links no corpo de um e-mail. 

**Sumário**

* [Resumo](#summary)
* [Pré-requisitos](#prerequisites)
* [Componentes principais do exemplo](#components)
* [Descrição do código](#codedescription)
* [Criar e depurar](#build)
* [Solução de problemas](#troubleshooting)
* [Perguntas e comentários](#questions)
* [Colaboração](#contribute)
* [Recursos adicionais](#additional-resources)

<a name="summary"></a>
##Resumo

Neste exemplo mostraremos como usar a [API JavaScript para Office](https://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298(v=office.15)) para criar um suplemento do Outlook que analisa o corpo de um e-mail procurando hiperlinks. Veja uma imagem do cenário em questão (no Outlook Web App).

 ![](/readme-images/screen2.PNG)
 
Esse suplemento está configurado para usar os [comandos suplemento](https://msdn.microsoft.com/EN-US/library/office/mt267547.aspx), portanto, quando você estiver lendo seus e-mails no cliente da área de trabalho, inicie o suplemento escolhendo este botão de comando na faixa de opções:

![](/readme-images/commandbutton.png)

 Já aconteceu com todos nós durante nossas vidas úteis de e-mail - recebemos algo que parece um e-mail regular do que parece ser uma fonte confiável que contém hiperlinks. Clicamos em um desses links sem pensar e corremos o risco de ter o nosso computador, nossos sistemas ou negócios comprometidos. Esse é um cenário de phishing clássico em que os hiperlinks em um e-mail não são o que parecem ser. Este exemplo mostra uma maneira alternativa de verificar hiperlinks. Em vez de passar o mouse sobre um link para ver qual é a URL de destino real por trás do texto do link e talvez arriscar um clique acidental no link em questão, esse suplemento localiza todos os links em um e-mail e os exibe em um formato decomposto de texto de link e URL de link. Dessa maneira, o usuário pode ver claramente o endereço que está por trás do texto do link. O exemplo vai um pouco mais além. Se um link tem uma URL como o texto do link e essa URL não corresponde ao href subjacente do link, o link é sinalizado em vermelho no suplemento para garantir que o usuário veja esse link potencialmente de phishing. 

<a name="prerequisites"></a>
##Prerequisites
Este exemplo exige o seguinte:  

  - Visual Studio 2013 com Atualização 5 ou Visual Studio 2015.  
  - Um computador executando o Exchange 2013 com pelo menos uma conta de e-mail ou uma conta do Office 365. [Participar do Programa de Desenvolvedores do Office 365 e obter uma assinatura gratuita de 1 ano do Office 365](https://aka.ms/devprogramsignup).
  - Internet Explorer 9 ou posterior, que deve estar instalado, mas não precisa ser o navegador padrão. Para oferecer suporte aos Suplementos do Office, o cliente do Office que atua como host usa os componentes do navegador que fazem parte do Internet Explorer 9 ou posterior.
  - Um dos seguintes como o navegador padrão: Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13 ou uma versão mais recente de um desses navegadores.
  - Familiaridade com programação em JavaScript e serviços Web.

<a name="components"></a>
Componentes ##Key

Essa solução foi criada no [Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS). Consiste de dois projetos – LinkRevealer e LinkRevealerWeb. Veja uma lista dos principais arquivos dentro desses projetos. 
#### Projeto LinkRevealer

* [```LinkRevealer.xml```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealer/LinkRevealerManifest/LinkRevealer.xml) O [arquivo manifesto](https://msdn.microsoft.com/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp) do suplemento do Word.

#### Projeto LinkRevealerWeb

* [```Home.html```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.html) interface do usuário HTML para o suplemento do Word.
* [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.js) o código JavaScript usado por Home.html para interagir com o Word usando o a API JavaScript para Office. 


<a name="codedescription"></a>
##Descrição do código

A lógica principal deste exemplo está no arquivo [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.js) no projeto LinkRevealerWeb. Uma vez que o suplemente é inicializado, o método [```getAsync()```](https://msdn.microsoft.com/library/office/mt269089.aspx)do objeto Corpo é então usado para recuperar o corpo do e-mail no formato Texto. Quando essa operação assíncrona for concluída, a função de retorno de chamada em linha é ativada. Esta função carrega primeiro o conteúdo recuperado do corpo em um DomParser. Essa árvore de objeto é então analisada usando o método getElementsByTagName ("a") para localizar todos os hiperlinks. Por fim, cada hiperlink é exibido na interface do usuário e analisado para verificar se há links de phishing. 

Usar o body.getAsync() para recuperar o corpo de um e-mail tem inúmeras vantagens em relação a soluções anteriores. Nas versões anteriores do Office.js, a única maneira de obter o corpo de um e-mail em um cenário de leitura era chamar [```makeEWSRequest```](https://msdn.microsoft.com/library/office/fp161019.aspx) no objeto da caixa de correio. Nem só a construção desta solicitação SOAP foi mais envolvida, mas também exigia que um suplemento tivesse permissões do ReadWriteMailbox. A solução getAsync() exige que o suplemento tenha permissões ReadItem.  

<a name="build"></a>
##Criar e depurar
1. Abra o arquivo [```LinkRevealer.sln```](LinkRevealer.sln) no Visual Studio.
2. Pressione F5 para compilar e implantar o suplemento de exemplo
3. Quando o Outlook iniciar, escolha um email de sua caixa de entrada
4. Inicie o suplemento selecionando-o na barra de aplicativo do suplemento

![](readme-images/screen1.PNG)


5. Quando o suplemento for iniciado, ele verificará o corpo da mensagem de e-mail selecionada buscando por hiperlinks. Todos os links encontrados serão exibidos em uma tabela no painel principal do suplemento. Se o suplemento achar que um link é suspeito, ele marcará essa linha na tabela em vermelho. Um link suspeito é definido como um que tenha uma URL no texto do link que não corresponde à URL no href real do link. 


<a name="troubleshooting"></a>
## Solução de problemas

- Se o suplemento não for exibido no painel de tarefas, escolha **Inserir > Meus Suplementos > Revelador de Link **.

<a name="questions"></a>
## Perguntas e comentários

- Se você tiver problemas para executar este exemplo, [relate um problema](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/issues).
- Perguntas sobre o desenvolvimento de Suplementos do Office em geral devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Não deixe de marcar as perguntas ou comentários com [office-addins].

<a name="contribute"></a>
## Colaboração ##
Recomendamos que você contribua para nossos exemplos. Para obter diretrizes sobre como proceder, confira nosso [guia de contribuição](./Contributing.md)

Este projeto adotou o [Código de Conduta de Código Aberto da Microsoft](https://opensource.microsoft.com/codeofconduct/).  Para saber mais, confira as [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou entre em contato pelo [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.


<a name="additional-resources"></a>
## Recursos adicionais ##

- [Mais exemplos de Suplementos](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Suplementos do Office](http://msdn.microsoft.com/library/office/jj220060.aspx)
- [Anatomia de um Suplemento](https://msdn.microsoft.com/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp)
- [Criando um suplemento do Office com o Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS)


## Direitos autorais
Copyright © 2015 Microsoft. Todos os direitos reservados.

