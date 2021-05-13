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
# Надстройка Outlook: почтовая надстройка для сценария чтения, которая находит и анализирует все ссылки в тексте сообщения электронной почты. 

**Содержание**

* [Сводка](#summary)
* [Предварительные требования](#prerequisites)
* [Ключевые компоненты примера](#components)
* [Описание кода](#codedescription)
* [Сборка и отладка](#build)
* [Устранение неполадок](#troubleshooting)
* [Вопросы и комментарии](#questions)
* [Участие](#contribute)
* [Дополнительные ресурсы](#additional-resources)

<a name="summary"></a>
##Сводка

В этом примере показано, как с помощью [API JavaScript для Office](https://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298(v=office.15)) создать надстройку Outlook, анализирующую текст сообщения на предмет ссылок. На следующем изображении показан рассматриваемый сценарий (в Outlook Web App).

 ![](/readme-images/screen2.PNG)
 
Эта надстройка настроена на использование [команд надстроек](https://msdn.microsoft.com/EN-US/library/office/mt267547.aspx), поэтому при чтении электронной почты в классическом клиенте для запуска надстройки используется следующая кнопка на ленте:

![](/readme-images/commandbutton.png)

 Такое случается со всеми пользователями электронной почты: нам приходит, казалось бы, обычное письмо из, казалось бы, надежного источника, но с гиперссылками. Мы не думая щелкаем какую-нибудь из этих ссылок и подвергаем свой компьютер, системы и бизнес риску взлома. Это классический сценарий фишинга, когда гиперссылки в сообщении электронной почты — совсем не то, чем кажутся. В данном примере показан альтернативный способ проверки гиперссылок. Вам не нужно наводить указатель мыши на ссылку, чтобы увидеть настоящий URL-адрес, на который указывает ссылка, стараясь случайно не щелкнуть ее — эта надстройка найдет все ссылки в сообщении и отобразит их в виде текста ссылки и ее URL-адреса. Таким образом пользователь ясно увидит, какой адрес скрывается за текстом ссылки. Однако наш пример этим не ограничивается. Если текст ссылки представляет собой URL-адрес, который не совпадает с фактическим URL-адресом в атрибуте href ссылки, такая ссылка помечается красным цветом как потенциально обманная, чтобы пользователь обязательно заметил ее. 

<a name="prerequisites"></a>
##Предварительные
требования Для этого примера требуется следующее:  

  - Visual Studio 2013 с обновлением 5 или Visual Studio 2015.  
  - Компьютер с Exchange 2013 и по крайней мере одной учетной записью электронной почты или учетной записью Office 365. [Примите участие в программе Office 365 для разработчиков и получите бесплатную подписку на Office 365 сроком на 1 год](https://aka.ms/devprogramsignup).
  - Браузер Internet Explorer 9 или более поздней версии, который должен быть установлен, но может не использоваться по умолчанию. Для поддержки надстроек Office клиент Office, выступающий в роли ведущего приложения, использует компоненты браузера, которые входят в состав Internet Explorer 9 или более поздней версии.
  - По умолчанию используется один из следующих браузеров: Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13 или более поздние версии этих браузеров.
  - Опыт программирования на JavaScript и работы с веб-службами.

<a name="components"></a>
##Ключевые компоненты

Это решение было создано в [Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS). Оно состоит из двух проектов — LinkRevealer и LinkRevealerWeb. Ниже представлен список основных файлов в этих проектах. 
#### Проект LinkRevealer

* [```LinkRevealer.xml```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealer/LinkRevealerManifest/LinkRevealer.xml) [Файл манифеста](https://msdn.microsoft.com/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp) для надстройки Word.

#### Проект LinkRevealerWeb

* [```Home.html```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.html) Пользовательский интерфейс на HTML для надстройки Word.
* [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.js) Код JavaScript, используемый на странице Home.html для взаимодействия с Word при помощи API JavaScript для Office. 


<a name="codedescription"></a>
##Описание кода

Базовая логика этого примера находится в файле [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.js) проекта LinkRevealerWeb. После инициализации надстройки метод [```getAsync()```](https://msdn.microsoft.com/library/office/mt269089.aspx) объекта Body обеспечивает получение текста сообщения в формате HTML. По окончании этой асинхронной операции вызывается функция обратного вызова processHtmlBody. Эта функция сначала загружает полученное содержимое сообщения в DomParser. Затем дерево объектов анализируется методом getElementsByTagName("a"), который находит все гиперссылки. Наконец, каждая гиперссылка отображается в пользовательском интерфейсе и проверяется на отсутствие фишинга. 

Использование метода body.getAsync() для получения текста сообщения дает многочисленные преимущества перед другими решениями. В прежних версиях Office.js получить текст сообщения в сценарии чтения можно было только путем вызова метода [```makeEWSRequest```](https://msdn.microsoft.com/library/office/fp161019.aspx) в объекте почтового ящика. Это не только усложняло построение запроса SOAP, но и требовало наличия разрешений ReadWriteMailbox для надстройки. Решение getAsync() требует лишь, чтобы у надстройки были разрешения ReadItem.  

<a name="build"></a>
##Сборка и отладка
1. Откройте файл [```LinkRevealer.sln```](LinkRevealer.sln) в Visual Studio.
2. Нажмите клавишу F5, чтобы собрать и развернуть пример надстройки
3. После запуска Outlook выберите сообщение в папке "Входящие"
4. Запустите надстройку, выбрав ее на панели приложения.

![](readme-images/screen1.PNG)


5. После запуска надстройки она проверит текст выбранного сообщения электронной почты на наличие гиперссылок. Все найденные гиперссылки будут отображены в таблице основной области надстройки. Если ссылка окажется подозрительной, она будет помечена в таблице красным цветом. Подозрительной считается ссылка, URL-адрес в тексте которой не совпадает с URL-адресом в атрибуте href. 


<a name="troubleshooting"></a>
## Устранение неполадок

- Если надстройка не отображается в области задач, выберите **Вставка > Мои надстройки > Link Revealer**.

<a name="questions"></a>
## Вопросы и комментарии

- Если у вас возникли проблемы с запуском этого примера, [сообщите о неполадке](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/issues).
- Общие вопросы о разработке надстроек Office следует задавать на сайте [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Обязательно помечайте свои вопросы и комментарии тегом [office-addins].

<a name="contribute"></a>
## Участие ##
Мы приветствуем ваше участие в создании примеров. Сведения о дальнейших действиях см. в [руководстве по участию](./Contributing.md).

Этот проект соответствует [Правилам поведения разработчиков открытого кода Майкрософт](https://opensource.microsoft.com/codeofconduct/). Дополнительные сведения см. в разделе [часто задаваемых вопросов о правилах поведения](https://opensource.microsoft.com/codeofconduct/faq/). Если у вас возникли вопросы или замечания, напишите нам по адресу [opencode@microsoft.com](mailto:opencode@microsoft.com).


<a name="additional-resources"></a>
## Дополнительные ресурсы ##

- [Дополнительные примеры надстроек](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Надстройки Office](http://msdn.microsoft.com/library/office/jj220060.aspx)
- [Структура надстройки](https://msdn.microsoft.com/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp)
- [Создание надстройки Office в Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS)


## Авторские права
(c) Корпорация Майкрософт (Microsoft Corporation), 2015. Все права защищены.

