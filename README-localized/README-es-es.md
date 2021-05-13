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
# Complemento de Outlook: Complemento de correo para un escenario de lectura que encuentra y analiza todos los vínculos en el cuerpo de un correo electrónico. 

**Tabla de contenido**

* [Resumen](#summary)
* [Requisitos previos](#prerequisites)
* [Componentes clave del ejemplo](#components)
* [Descripción del código](#codedescription)
* [Compilar y depurar](#build)
* [Solución de problemas](#troubleshooting)
* [Preguntas y comentarios](#questions)
* [Colaboradores](#contribute)
* [Recursos adicionales](#additional-resources)

<a name="summary"></a>
##Resumen

En este ejemplo se muestra cómo usar la [API de JavaScript para Office](https://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298(v=office.15)) para crear un complemento de Outlook que analice el cuerpo de un correo electrónico para encontrar hipervínculos. Esta es una imagen del escenario en cuestión (en Outlook Web App).

 ![](/readme-images/screen2.PNG)
 
Este complemento se configura para usar [comandos de complemento](https://msdn.microsoft.com/EN-US/library/office/mt267547.aspx), por lo que, cuando lee el correo electrónico en el cliente de escritorio, se inicia el complemento al elegir este botón de comando en la cinta de opciones:

![](/readme-images/commandbutton.png)

 Nos ha ocurrido a todos durante la vida del correo electrónico, recibimos un correo electrónico aparentemente normal de lo que parece un origen de confianza y que contiene hipervínculos. Hacemos clic en uno de los vínculos sin pensar y corremos el riesgo de que nuestro equipo, nuestros sistemas o empresas se vean comprometidos. Este es un escenario clásico de phising, en el que los hipervínculos de un correo electrónico no son lo que parecen. Este ejemplo muestra una forma alternativa de comprobar los hipervínculos. En lugar de pasar el ratón por encima de un vínculo para ver cuál es la URL de destino real que subyace al texto del vínculo, y suponer un riesgo de hacer clic por accidente en un vínculo, este complemento encuentra todos los vínculos en un correo electrónico y los muestra en un formato descompuesto de texto del vínculo y la dirección URL del vínculo. De esta forma, el usuario puede ver claramente las direcciones que aparecen detrás del texto del vínculo. El ejemplo va un poco más allá. Si un vínculo tiene una dirección URL como el texto del vínculo y esa dirección URL no coincide con el href subyacente del vínculo, el vínculo se marca en rojo en el complemento para asegurarse de que el usuario vea este vínculo que potencialmente es un caso de phishing. 

<a name="prerequisites"></a>
##Requisitos previos
Este ejemplo necesita lo siguiente:  

  - Visual Studio 2013 con Update 5 o Visual Studio 2015.  
  - Un equipo que ejecute Exchange 2013 y, como mínimo, una cuenta de correo electrónico o una cuenta de Office 365. [Únase al programa Office 365 Developer y consiga una suscripción gratuita de 1 año a Office 365.](https://aka.ms/devprogramsignup)
  - Internet Explorer 9 o posterior, que debe estar instalado, pero no es necesario que sea el navegador predeterminado. Para admitir los complementos de Office, el cliente de Office que actúa como host utiliza componentes del explorador que forman parte de Internet Explorer 9 o posterior.
  - Uno de los siguientes como explorador predeterminado: Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13 o una versión posterior de estos exploradores.
  - Familiaridad con los servicios web y la programación de JavaScript.

<a name="components"></a>
##Componentes clave

Esta solución se creó en [Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS). Consta de dos proyectos: LinkRevealer y LinkRevealerWeb. Esta es una lista de los archivos de claves en esos proyectos. 
#### Proyecto LinkRevealer

* [```LinkRevealer.XML```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealer/LinkRevealerManifest/LinkRevealer.xml) el [archivo de manifiesto](https://msdn.microsoft.com/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp) para el complemento de Word.

#### Proyecto LinkRevealerWeb

* [```Home.html```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.html) La interfaz de usuario HTML para el complemento de Word.
* [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.js) El código de JavaScript que ha usado Home.html para interactuar con Word mediante la API de JavaScript para Office. 


<a name="codedescription"></a>
##Descripción del código

La lógica básica de este ejemplo está en el archivo [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.js) en el proyecto LinkRevealerWeb. Una vez que se inicialice el complemento, se usará el método [```getAsync()```](https://msdn.microsoft.com/library/office/mt269089.aspx) del objeto de cuerpo para recuperar el cuerpo del mensaje de correo electrónico en formato HTML. Cuando se completa esta operación asincrónica, se invoca la función de devolución de llamada, processHtmlBod. Esta función carga en primer lugar el contenido de cuerpo recuperado en un DomParser. Este árbol de objetos se analiza con el método getElementsByTagName("a") para buscar todos los hipervínculos. Por último, se muestran todos los hipervínculos en la interfaz de usuario y se analiza para ver si algún vínculo es phishing. 

Usar Body.getAsync() para recuperar el cuerpo de un correo electrónico tiene numerosas ventajas respecto a las soluciones anteriores. En versiones anteriores de Office.js, la única manera de recibir el cuerpo de un mensaje de correo electrónico en un escenario de lectura era llamar a [```makeEWSRequest```](https://msdn.microsoft.com/library/office/fp161019.aspx) en el objeto buzón. No solo era la construcción de esta solicitud SOAP más complicada, sino que también necesitaba que un complemento tuviera permisos ReadWriteMailbox. La solución getAsync() solo requiere que el complemento tenga permisos ReadItem.  

<a name="build"></a>
##Compilar y depurar
1. Abra el archivo [```LinkRevealer.sln```](LinkRevealer.sln) en Visual Studio.
2. Pulse F5 para crear e implementar el complemento de ejemplo
3. Cuando se inicie Outlook, seleccione un correo electrónico de la Bandeja de entrada
4. Inicie el complemento seleccionándolo en la barra de la aplicación del complemento

![](readme-images/screen1.PNG)


5. Cuando se inicia el complemento, explora el cuerpo del mensaje de correo electrónico seleccionado para comprobar si hay hipervínculos. Se mostrarán los vínculos que encuentre en una tabla en el panel principal del complemento. Si el complemento cree que un vínculo es sospechoso, marcará dicha fila en la tabla en rojo. Un vínculo sospechoso se define como uno que tiene una dirección URL en el texto del vínculo que no coincide con la dirección URL de la href real del vínculo. 


<a name="troubleshooting"></a>
## Solución de problemas

- Si el complemento no se muestra en el panel de tareas, elija **Insertar > Mis complementos > Link Revealer**.

<a name="questions"></a>
## Preguntas y comentarios

- Si tiene algún problema para ejecutar este ejemplo, [registre un problema](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/issues).
- Las preguntas sobre el desarrollo de complementos para Office en general deben enviarse a [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Asegúrese de que sus preguntas o comentarios se etiquetan con [office-addins].

<a name="contribute"></a>
## Contribuciones ##
Le animamos a contribuir a nuestros ejemplos. Para obtener instrucciones sobre cómo continuar, consulte nuestra [guía de contribución](./Contributing.md)

Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, vea las [Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.


<a name="additional-resources"></a>
## Recursos adicionales ##

- [Más complementos de ejemplo](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Complementos de Office](http://msdn.microsoft.com/library/office/jj220060.aspx)
- [Anatomía de un complemento](https://msdn.microsoft.com/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp)
- [Crear un complemento de Office con Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS)


## Copyright
Copyright (c) 2015 Microsoft. Todos los derechos reservados.

