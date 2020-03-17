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
# Outlook 加载项：阅读场景“邮件”加载项，可查找和分析电子邮件正文中的所有链接。 

**目录**

* [摘要](#summary)
* [先决条件](#prerequisites)
* [示例主要组件](#components)
* [代码说明](#codedescription)
* [构建和调试](#build)
* [疑难解答](#troubleshooting)
* [问题和意见](#questions)
* [参与](#contribute)
* [其他资源](#additional-resources)

<a name="summary"></a>
##摘要

本示例介绍如何使用[适用于 Office 的 JavaScript API](https://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298(v=office.15))来创建可分析查找超链接的电子邮件正文的 Outlook 加载项。下面是有问题的方案图像（Outlook Web 应用中）。

 ![](/readme-images/screen2.PNG)
 
此加载项的配置用于使用“[加载项命令](https://msdn.microsoft.com/EN-US/library/office/mt267547.aspx)”，因此使用桌面版客户端阅读电子邮件时，可通过选择功能区中的此命令按钮：

![](/readme-images/commandbutton.png)

 电子邮件生命周期期间，会发生在我们所有人身上 - 我们从包含超链接的看似可信来源那里收到看起来像普通电子邮件的邮件。如果我们未加思索地点击这些链接中的一个，随后会产生影响计算机、系统或业务的风险。这是一个经典的钓鱼情景，其中电子邮件中的超链接与看到的不一样。此示例显示验证超链接的替代方式。无需将鼠标悬停在链接来查看链接文本后面的实际目标 URL，并且意外点击链接会导致危险发生，因此加载项查找电子邮件中的所有链接，并采用分解的格式显示链接文本和链接 URL。采用这种方式，用户清楚看到链接文本后面的地址。此示例将更进一步。如果链接的 URL 为链接文本，且 URL 与 链接的超文本应用不一致，此链接在加载项中采用红色标记，以确保用户能够查看此潜在可疑的链接。 

<a name="prerequisites"></a>
##先决条件
此示例需要下列内容：  

  - Visual Studio 2013 Update 5 或 Visual Studio 2015。  
  - 运行至少具有一个电子邮件帐户或 Office 365 帐户的 Exchange 2013 的计算机。[参加 Office 365 开发人员计划并获取为期 1 年的免费 Office 365 订阅](https://aka.ms/devprogramsignup)。
  - 必须安装的 Internet Explorer 9 或更高版本无需是默认浏览器。为了支持 Office 加载项，充当主机的 Office 客户端所使用的浏览器组件是 Internet Explorer 9 或更高版本的一部分。
  - 使用以下任一浏览器作为默认浏览器：Internet Explorer 9、Safari 5.0.6、Firefox 5、Chrome 13 或其中一个浏览器的更高版本。
  - 熟悉 JavaScript 编程和 Web 服务。

<a name="components"></a>
##主要组件

此解决方案在 [Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS) 中创建。它包含 LinkRevealer 和 LinkRevealerWeb 两个项目组成。以下是这些项目中关键文件的列表。 
#### LinkRevealer 项目

* [```LinkRevealer.xml```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealer/LinkRevealerManifest/LinkRevealer.xml) Word 加载项的[清单文件](https://msdn.microsoft.com/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp)。

#### LinkRevealerWeb 项目

* [```Home.html```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.html) Word 外接程序的 HTML 用户界面。
* [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.js) 由 Home.html 使用的与使用适用于 Office 的 JavaScript API 的 Word 进行交互的 JavaScript 代码。 


<a name="codedescription"></a>
##代码说明

此示例的核心逻辑位于 ScanForMeWeb 项目中的 [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.js) 文件。加载项启动后，正文对象的 [```getAsync()```](https://msdn.microsoft.com/library/office/mt269089.aspx) 方法用于采用 HTML 格式获取电子邮件的正文。在此异步操作完成时，我们的内联回叫函数 processHtmlBody会得到调用。此函数首先载入获取的正文内容至 DomParser 中。此对象树随后使用 getElementsByTagName("a") 方法进行解析，以查找所有超链接。最后各超链接在用户界面上显示并进行分析，查看任何链接是否可疑。 

相比之前的解决方案，使用 body.getAsync() 获取电子邮件正文有诸多优势。在早期版本的 Office.js 中，获取阅读场景中电子邮件正文的唯一方式是调用邮箱对象上的 [```makeEWSRequest```](https://msdn.microsoft.com/library/office/fp161019.aspx)。不仅更多涉及创建此 SOAP 请求，而且还需要加载项拥有 ReadWriteMailbox 权限。getAsync() 解决方案仅要求加载项拥有 ReadItem 权限。  

<a name="build"></a>
##构建和调试
1.在 Visual Studio 中打开 [```LinkRevealer.sln```](LinkRevealer.sln) 文件。
2.按 F5 生成并部署示例加载项
3.当 Outlook 启动时，从收件箱中选择一封电子邮件
4.通过从加载项应用栏选择加载项，启动加载项

![](readme-images/screen1.PNG)


5. 启动加载项时，扫描选定电子邮件正文中的超链接。找到的任何链接将在加载项主窗格的表格中显示。如果加载项认为链接可疑，则将表中的该行标记为红色。可疑链接是链接文本中的 URL 与链接真实超文本引用 URL 不匹配的链接。 


<a name="troubleshooting"></a>
## 疑难解答

- 如果任务窗格中未显示加载项，请选择“**插入 > 我的外接程序 > 链接显示器**”。

<a name="questions"></a>
## 问题和意见

- 如果你在运行此示例时遇到任何问题，请[记录问题](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/issues)。
- 与 Office 加载项开发相关的问题一般应发布到 [堆栈溢出](http://stackoverflow.com/questions/tagged/office-addins)。确保你的问题或意见使用 [Office 加载项] 进行了标记。

<a name="contribute"></a>
## 参与 ##
我们鼓励你参与我们的示例。有关如何继续的指南，请参阅我们的[参与指南](./Contributing.md)

此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。


<a name="additional-resources"></a>
## 其他资源 ##

- [更多加载项示例](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Office 加载项](http://msdn.microsoft.com/library/office/jj220060.aspx)
- [加载项解析](https://msdn.microsoft.com/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp)
- [使用 Visual Studio 创建 Office 加载项](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS)


## 版权信息
版权所有 (c) 2015 Microsoft。保留所有权利。

