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
# Outlook アドイン:メールの本文内のすべてのリンクを検出して解析する読み取りシナリオのメール アドイン。 

**目次**

* [概要](#summary)
* [前提条件](#prerequisites)
* [サンプルの主要なコンポーネント](#components)
* [コードの説明](#codedescription)
* [ビルドとデバッグ](#build)
* [トラブルシューティング](#troubleshooting)
* [質問とコメント](#questions)
* [投稿](#contribute)
* [その他の技術情報](#additional-resources)

<a name="summary"></a>
##要約

このサンプルでは、[JavaScript API for Office](https://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298(v=office.15)) を使用して、ハイパーリンクを参照する電子メールの本文を解析する Outlook アドインの作成方法を示します。次に、質問のシナリオの画像を示します (Outlook Web App で)。

 ![](/readme-images/screen2.PNG)
 
このアドインは[アドイン コマンド](https://msdn.microsoft.com/EN-US/library/office/mt267547.aspx)を使用するように構成されているので、デスクトップ クライアントでメールを閲覧するときに、リボンのこのコマンド ボタンを選択してアドインを起動します。

![](/readme-images/commandbutton.png)

 これは、メールの有効期間中に全員に起こりました。ハイパーリンクを含む信頼できるソースのように見えるものから、通常のメールのように見えるものを受け取ります。何も疑わずこれらのリンクのいずれかをクリックすると、マシン、システム、またはビジネスが危険にさらされます。これは典型的なフィッシングのシナリオで、メール内のハイパーリンクは見た目とは異なるものです。このサンプルでは、ハイパーリンクを確認する別の方法を示します。このアドインは、リンク テキストの背後にある実際のターゲット URL を確認するためにリンク上にマウスを移動したり、そのリンクを誤ってクリックしたりする代わりに、メール内のすべてのリンクを見つけてリンク テキストとリンク URL の分解形式で表示します。これで、リンク テキストの背後にあるアドレスがはっきりわかるようになります。サンプルはもう少し進んだ機能をお見せします。リンクにリンク テキストとして URL があり、その URL がリンクの基礎となる href と一致しない場合、そのリンクには、ユーザーがこのフィッシングが疑われるリンクを確認できるように、アドインで赤色のフラグが付けられます。 

<a name="prerequisites"></a>
##前提条件
このサンプルを実行するには次のものが必要です。  

  - Visual Studio 2013 更新プログラム 5 または Visual Studio 2015。  
  - 少なくとも 1 つのメール アカウントまたは Office 365 アカウントがある Exchange 2013 を実行するコンピューター。[Office 365 Developer プログラムに参加して、Office 365 の 1 年間無料のサブスクリプションを取得します](https://aka.ms/devprogramsignup)。
  - Internet Explorer 9 以降をインストールする必要がありますが、必ずしも既定のブラウザーにする必要はありません。Office アドインをサポートするために、ホストとして動作する Office のクライアントは、Internet Explorer 9 以降に組み込まれているブラウザー コンポーネントを使用します。
  - 既定のブラウザーとして次のいずれか:Internet Explorer 9、Safari 5.0.6、Firefox 5、Chrome 13、これらのブラウザーのいずれかの最新バージョン。
  - JavaScript プログラミングと Web サービスに関する知識。

<a name="components"></a>
##主要なコンポーネント

このソリューションは、[Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS) で作成されました。これは、LinkRevealer と LinkRevealerWeb の 2 つのプロジェクトで構成されています。以下に、これらのプロジェクト内のキー ファイルの一覧を示します。 
#### LinkRevealer プロジェクト

* [```LinkRevealer.xml```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealer/LinkRevealerManifest/LinkRevealer.xml) [Word アドインの](https://msdn.microsoft.com/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp)マニフェスト ファイル。

#### LinkRevealerWeb プロジェクト

* [```Home.html```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.html) Word アドインの HTML ユーザー インターフェイス。
* [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.js) Office API の JavaScript を使用して Word と対話するために Home.html によって使用される JavaScript コード。 


<a name="codedescription"></a>
##コードの説明

このサンプルのコア ロジックは、LinkRevealerWeb プロジェクトの [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/blob/master/LinkRevealerWeb/AppRead/Home/Home.js) ファイルです。アドインが初期化されると、Body オブジェクトの [```getAsync()```](https://msdn.microsoft.com/library/office/mt269089.aspx) メソッドは、HTML 形式のメールの本文を取得するために使用されます。この非同期操作が完了すると、コールバック関数 processHtmlBody が呼び出されます。この関数は、最初に取得した本文のコンテンツを DomParser に読み込みます。このオブジェクト ツリーは、getElementsByTagName("a") メソッドを使用して解析され、すべてのハイパーリンクを検索します。最後に、各ハイパーリンクが UI に表示され、リンクがフィッシングかどうかが分析されます。 

body.getAsync() を使用してメールの本文を取得すると、以前のソリューションよりも多くの利点が得られます。以前のバージョンの Office .js では、読み取りのシナリオでメールの本文を取得する唯一の方法は、メールボックス オブジェクトの [```makeEWSRequest```](https://msdn.microsoft.com/library/office/fp161019.aspx) を呼び出すことでした。この SOAP 要求の構造がより複雑になっただけでなく、アドインに ReadWriteMailbox 権限が必要です。getAsync() ソリューションでは、アドインに ReadItem 権限があればいいです。  

<a name="build"></a>
##ビルドとデバッグ
1.Visual Studio で [```LinkRevealer.sln```](LinkRevealer.sln) ファイルを開きます。
2.F5 キーを押して、サンプル アドインをビルドおよび展開します。
3.Outlook が起動したら、受信トレイから電子メールを選択します。
4.アドイン アプリ バーからアドインを選択して、起動します。

![](readme-images/screen1.PNG)


5. アドインが起動すると、選択したメール メッセージの本文でハイパーリンクがスキャンされます。見つかったリンクは、アドインのメイン ウィンドウのテーブルに表示されます。アドインは、リンクが疑わしいと判断すると、テーブル内のその行に赤のマークを付けます。不審なリンクは、リンクの実際の href のURLと一致しない URL をリンク テキストに含むものとして定義されます。 


<a name="troubleshooting"></a>
## トラブルシューティング

- アドインが作業ウィンドウに表示されない場合、**[挿入]、[個人用アドイン]、[Link Revealer]** の順に選択します。

<a name="questions"></a>
## 質問とコメント

- このサンプルの実行について問題がある場合は、[問題をログに記録](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer/issues)してください。
- Office アドイン開発全般の質問については、「[Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins)」に投稿してください。質問やコメントには、必ず "office-addins" のタグを付けてください。

<a name="contribute"></a>
## 投稿 ##
当社のサンプルに是非貢献してください。投稿方法のガイドラインについては、[投稿ガイド](./Contributing.md)を参照してください。

このプロジェクトでは、[Microsoft オープン ソース倫理規定](https://opensource.microsoft.com/codeofconduct/) が採用されています。詳細については、「[Code of Conduct の FAQ (倫理規定の FAQ)](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。


<a name="additional-resources"></a>
## その他の技術情報 ##

- [その他のアドイン サンプル](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Office アドイン](http://msdn.microsoft.com/library/office/jj220060.aspx)
- [アドインの構造](https://msdn.microsoft.com/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp)
- [Visual Studio で Office アドインを作成する](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS)


## 著作権
Copyright (c) 2015 Microsoft.All rights reserved.

