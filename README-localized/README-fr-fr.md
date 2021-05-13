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
# Complément Outlook : Complément courrier pour un scénario de lecture permettant de rechercher et d’analyser tous les liens dans le corps d’un courrier électronique. 

**Table des matières**

* [Résumé](#summary)
* [Conditions préalables](#prerequisites)
* [Composants clés de l’exemple](#components)
* [Description du code](#codedescription)
* [Création et débogage](#build)
* [Résolution des problèmes](#troubleshooting)
* [Questions et commentaires](#questions)
* [Contribution](#contribute)
* [Ressources supplémentaires](#additional-resources)

<a name="summary"></a>
##Résumé

Cet exemple vous présente comment utiliser l’[interface API JavaScript pour Office](https://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298(v=office.15)) afin de créer un complément Outlook qui analyse le corps d’un message électronique pour rechercher des liens hypertextes. Voici une image du scénario en question (dans Outlook Web App).

 ![](/readme-images/screen2.PNG)
 
Ce complément est configuré pour utiliser des [commandes de complément](https://msdn.microsoft.com/EN-US/library/office/mt267547.aspx). Ainsi, lorsque vous lisez votre courrier électronique dans le client de bureau, vous lancez le complément en cliquant sur ce bouton de commande dans le ruban :

![](/readme-images/commandbutton.png)

 Cela nous est arrivé à tous : nous recevons un courrier électronique standard d’une source approuvée contenant des liens hypertexte. Nous cliquons sur l’un de ces liens, sans réfléchir, et risquons de compromettre notre ordinateur, nos systèmes ou notre entreprise. Il s’agit d’un scénario classique de hameçonnage dans lequel les liens hypertexte dans un courrier électronique ne sont pas ce qu’ils paraissent. Cet exemple illustre une autre façon de vérifier les liens hypertexte. Au lieu de pointer au-dessus d’un lien pour voir l’URL cible réelle qui se trouve derrière le lien, et risquer d’avoir un clic accidentel sur ce lien, ce complément recherche tous les liens dans un message électronique et les affiche dans un format décomposé de texte de lien et d’URL de lien. Ainsi, l’utilisateur peut voir clairement l’adresse qui se trouve derrière le texte du lien. L’exemple va un peu plus loin. Si un lien possède une URL comme texte du lien et que celle-ci ne correspond pas à la href de la liaison sous-jacente, le lien est marqué en rouge dans le complément pour s’assurer que l’utilisateur voit ce lien potentiellement hameçon. 

<a name="prerequisites"></a>
##Conditions préalables :
cet exemple nécessite les éléments suivants :  

  - Visual Studio 2013 avec la mise à jour 5 ou Visual Studio 2015.  
  - Un ordinateur exécutant Exchange 2013 avec au moins un compte de messagerie ou un compte Office 365. [Participer au programme pour les développeurs Office 365 et obtenir un abonnement gratuit d’un an à Office 365](https://aka.ms/devprogramsignup).
  - Internet Explorer 9 ou version ultérieure, qui doit être installé, mais ne doit pas être le navigateur par défaut. Pour prendre en charge les compléments Office, le client Office qui s’exécute en tant qu’hôte utilise des composants de navigateur qui font partie d’Internet Explorer 9 ou version ultérieure.
  - L’un des éléments suivants en tant que navigateur par défaut : Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13 ou une version ultérieure de l’un de ces navigateurs.
  - Être familiarisé avec les services web et de programmation JavaScript.

<a name="components"></a>
##Composants clés

Cette solution a été créée dans [Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS). Il se compose de deux projets : LinkRevealer et LinkRevealerWeb. Voici la liste des fichiers clés compris dans ces projets. 
#### Projet LinkRevealer

* [```LinkRevealer.xml```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealer/LinkRevealerManifest/LinkRevealer.xml) le [fichier manifeste](https://msdn.microsoft.com/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp) pour le complément Word.

#### Projet LinkRevealerWeb

* [```Home.html```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.html) l’interface utilisateur HTML pour le complément Word.
* [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.js) code JavaScript utilisé par Home.html pour interagir avec Word à l’aide de l’API JavaScript pour Office. 


<a name="codedescription"></a>
##Description du code

La logique de base de cet exemple se trouve dans le fichier [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.js) dans le projet LinkRevealerWeb. Une fois le complément initialisé, la méthode [```getAsync ()```](https://msdn.microsoft.com/library/office/mt269089.aspx) de l’objet Corps est utilisée pour récupérer le corps du message au format HTML. Lorsque cette opération asynchrone est terminée, la fonction de rappel en ligne processHtmlBody est invoquée. Cette fonction charge tout d’abord le contenu de corps récupéré dans une DomParser. Cet arbre d’objets est ensuite analysé à l’aide de la méthode getElementsByTagName (« a ») pour rechercher tous les liens hypertexte. Enfin, chaque lien hypertexte s’affiche sur l’interface utilisateur et est analysé pour voir s’il existe des liens hameçons. 

L’utilisation de la fonction body.getAsync () pour récupérer le corps d’un courrier électronique présente de nombreux avantages par rapport aux solutions précédentes. Dans les versions précédentes d’Office.js, la seule façon d’obtenir le corps d’un courrier électronique dans un scénario de lecture était d’appeler [```makeEWSRequest```](https://msdn.microsoft.com/library/office/fp161019.aspx) sur l’objet boîte aux lettres. Non seulement la construction de cette demande SOAP était plus impliquée, mais il était également nécessaire qu’un complément dispose des autorisations ReadWriteMailbox. La solution getAsync () requiert uniquement que le complément dispose des autorisations ReadItem.  

<a name="build"></a>
##Création et débogage
1. Ouvrez le fichier [```LinkRevealer.sln```](LinkRevealer.sln) dans Visual Studio.
2. Appuyez sur F5 pour créer et déployer l’exemple de complément 
3. Au démarrage d'Outlook, sélectionnez un e-mail de votre boîte de réception
4. Lancez le complément en le sélectionnant dans la barre d’application du complément.

![](readme-images/screen1.PNG)


5. Lorsque le complément démarre, il analyse le corps du message électronique sélectionné pour les liens hypertexte. Les liens trouvés sont affichés dans un tableau dans le volet principal du complément. Si le complément pense qu’un lien est suspect, il marque cette ligne en rouge dans le tableau. Un lien suspect est reconnu lorsqu’une URL est présente dans le texte du lien qui ne correspond pas à l’URL dans la href du lien. 


<a name="troubleshooting"></a>
## Résolution des problèmes

- Si le complément n’apparaît pas dans le volet des tâches, **Insertion > Mes compléments > Link Revealer**.

<a name="questions"></a>
## Questions et commentaires

- Si vous rencontrez des difficultés pour exécuter cet exemple, veuillez [consigner un problème](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/issues).
- Si vous avez des questions sur les compléments d’Office, envoyez-les sur [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Posez vos questions ou envoyez vos commentaires en incluant la balise [office-addins].

<a name="contribute"></a>
## Contribution ##
Nous vous invitons à contribuer à nos exemples. Pour obtenir des instructions sur la façon de procéder, consultez notre [guide de contribution](./Contributing.md).

Ce projet a adopté le [Code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/). Pour en savoir plus, reportez-vous à la [FAQ relative au Code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.


<a name="additional-resources"></a>
## Ressources supplémentaires ##

- [Autres exemples de compléments](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Compléments Office](http://msdn.microsoft.com/library/office/jj220060.aspx)
- [Structure d’un complément](https://msdn.microsoft.com/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp)
- [Création d’un complément Office avec Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS)


## Copyright
Copyright (c) 2015 Microsoft. Tous droits réservés.

