---
title: Visual Studio での Office アドインの作成とデバッグ
description: ''
ms.date: 03/14/2018
ms.openlocfilehash: 2e5c08a72ec97e26000d6ea7e53dd1d0f2c9e6dc
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945356"
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a>Visual Studio での Office アドインの作成とデバッグ

この記事では、Visual Studio を使用して、最初の Office アドインを作成する方法について説明します。ここに示す手順は Visual Studio 2015 に基づいたものです。別のバージョンの Visual Studio を使用している場合は、わずかに手順が異なることがあります。

> [!NOTE]
> OneNote 用のアドインを使い始めるには、「[最初の OneNote アドインをビルドする](../onenote/onenote-add-ins-getting-started.md)」を参照してください。

## <a name="create-an-office-add-in-project-in-visual-studio"></a>Visual StudioでOfficeアドインプロジェクトを作成する


作業を開始するために、[Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) がインストールされていることと、Microsoft Office のバージョンを確認します。[Office 365 Developer プログラム](https://developer.microsoft.com/office/dev-program)に参加するか、以下の手順を実行して[最新バージョン](../develop/install-latest-office-version.md)を取得できます。


1. [Visual Studio] メニュー バーで、**[ファイル]** > **[新規作成]** > **[プロジェクト]** の順に選択します。
    
2. プロジェクトの種類の一覧で、**[Visual C#]** または **[Visual Basic]** の下にある **[Office/SharePoint]** を展開し、**[Web アドイン]** を選択してからアドイン プロジェクトのいずれかを選択します。  
    
3. プロジェクトに名前を付けて、プロジェクトを作成するために **[OK]** を選択します。
    
4. Visual Studio によってソリューションとその 2 つのプロジェクトが作成され、**ソリューション エクスプローラー**に表示されます。既定の Home.html ページが Visual Studio に開かれます。
    
Visual Studio 2015 で、追加機能を反映するために、次に示す一部のアドイン プロジェクト テンプレートが更新されました。


- コンテンツのアドインは、Excel スプレッドシートに加えて、Access ドキュメントと PowerPoint ドキュメントの本文に表示されます。[Basic Project] オプションを選択すると、最小限のスタート コードで基本のコンテンツ アドイン プロジェクトを作成できます。または、[Document Visualization Project] オプション (Access および Excel のみ) を選択すると、データの視覚化とバインドを行うためのスタート コードが組み込まれているフル機能のコンテンツ アドインを作成できます。
    
- Outlook アドインには、電子メール メッセージや予定内にアドインを組み込むオプションだけでなく、電子メール メッセージや予定の閲覧時や新規作成時にアドインが使用可能かどうか指定するオプションも含まれています。
    

> [!NOTE]
> Visual Studio では、ほとんどのオプションは、説明を読んで理解できますが、**[電子メール メッセージ]** チェック ボックスは例外です。このチェック ボックスは、メール アイテムだけでなく、会議出席依頼、返信、キャンセルでも表示される Outlook アドインを作成する場合に使用します。

ウィザードの完了後、Visual Studio によって 2 つのプロジェクトを含むソリューションが作成されます。



|**プロジェクト**|**説明**|
|:-----|:-----|
|アドイン プロジェクト|アドインを記述するすべての設定を含む XML マニフェスト ファイルのみが含まれます。これらの設定は、Office ホストがアドインをアクティブ化するタイミングと、アドインの表示場所を決定するのに役立ちます。すぐにプロジェクトを実行し、アドインを使用できるように、Visual Studio によってこのファイルのコンテンツが生成されます。これらの設定は、マニフェスト エディターを使用していつでも変更できます。|
|Web アプリケーション プロジェクト|Office 対応の HTML および JavaScript ページを開発するために必要なすべてのファイルとファイル参照を含むアドインのコンテンツ ページが含まれます。アドインを開発している間、Visual Studio は Web アプリケーションをローカル IIS サーバー上でホストします。発行する準備が整ったら、このプロジェクトをホストするサーバーを見つける必要があります。ASP.NET Web アプリケーション プロジェクトの詳細については、「 [ASP.NET Web プロジェクト](http://msdn.microsoft.com/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx)」を参照してください。|

## <a name="modify-your-add-in-settings"></a>アドイン設定の変更


アドインの設定を変更するには、プロジェクトの XML マニフェスト ファイルを編集します。[**ソリューション エクスプローラー**] で、アドイン プロジェクト ノードを展開し、XML マニフェストを格納するフォルダーを展開して、XML マニフェストを選択します。ファイル内の任意の要素をポイントして、要素の目的を説明するヒントを表示できます。マニフェスト ファイルの詳細については、「[Office アドイン XML マニフェスト](../develop/add-in-manifests.md)」をご覧ください。


## <a name="develop-the-contents-of-your-add-in"></a>アドインのコンテンツの開発


アドイン プロジェクトはアドインを説明する設定を変更でき、Web アプリケーションはアドインに表示されるコンテンツを提供します。 

Web アプリケーション プロジェクトには、作業の開始時に使用できる既定の HTML ページと JavaScript ファイルが含まれています。また、プロジェクトに追加するすべてのページに共通の JavaScript ファイルも含まれています。JavaScript API for Office などの他の JavaScript ライブラリへの参照が含まれるので、これらのファイルは便利です。 

アドインが高度になるにつれて、追加する HTML ファイルや JavaScript ファイルの数が多くなります。アドインと連動させるためにプロジェクト内の他のページに追加できる参照の種類の例として、既定の HTML ファイルと JavaScript ファイルのコンテンツがあります。次の表は、既定の HTML ファイルと JavaScript ファイルを示しています。



|**ファイル**|**説明**|
|:-----|:-----|
|**Home.html**|アドインの既定の HTML ページで、プロジェクトの  **[Home]** フォルダーに存在します。アドインがドキュメント、電子メール メッセージ、または予定のアイテムでアクティブ化されると、このページがアドイン内の最初のページとして表示されます。このファイルには作業を開始する際に必要となるファイル参照がすべて含まれていて便利です。最初のアドインを作成する準備ができたら、このファイルに HTML コードを追加するだけで済みます。|
|**Home.js**|Home.js ページに関連付けられた JavaScript ファイルで、プロジェクトの  **[Home]** フォルダーに存在します。Home.js ファイルには Home.html ページの動作にとって固有のコードを組み込むことができます。Home.js ファイルには、作業の開始に利用できるサンプル コードが含まれています。|
|**App.js**|アドイン全体の既定の JavaScript ファイルで、プロジェクトの  **[Add-in]** フォルダーに存在します。App.js ファイルにはアドインの複数ページの動作にとって共通のコードを組み込むことができます。App.js ファイルには作業を開始する際に使用できるサンプル コードがいくつか含まれています。|

> [!NOTE]
> これらのファイルを必ず使用する必要はありません。他のファイルをプロジェクトに追加して代わりに使用してもかまいません。別の HTML ファイルをアドインの初期ページとして表示する場合は、マニフェスト エディターを開き、**SourceLocation** プロパティにそのファイルの名前を設定します。


## <a name="debug-your-add-in"></a>アドインのデバッグ


アドインを起動する準備ができたら、ビルドとデバッグに関連するプロパティを確認してください。確認が終了したら、ソリューションを起動します。


### <a name="review-the-build-and-debug-properties"></a>ビルドおよびデバッグ プロパティの確認

ソリューションを起動する前に、Visual Studio で目的のホスト アプリケーションが開けることを確認します。この情報は、アドインのビルドとデバッグに関連する他のプロパティと共に、プロジェクトのプロパティ ページに表示されます。


### <a name="to-open-the-property-pages-of-a-project"></a>プロジェクトのプロパティ ページを開くには


1. **ソリューション エクスプ ローラー**では、Web プロジェクトではなく、基本的なアドイン プロジェクトを選択します。
    
2. メニュー バーで、[ **表示**]、[ **プロパティ ウィンドウ**] の順に選択します。
    
次の表に、プロジェクトのプロパティを示します。



|**プロパティ**|**説明**|
|:-----|:-----|
|**開始動作**|Office デスクトップ クライアントまたは指定のブラウザー内の Office Online クライアントのどちらでアドインをデバッグするか指定します。|
|**開始ドキュメント** (コンテンツ アドインと作業ウィンドウ アドインのみ)|プロジェクトの開始時に開くドキュメントを指定します。|
|**Web プロジェクト**|アドインに関連付けられている Web プロジェクトの名前を指定します。|
|**電子メール アドレス** (Outlook アドインのみ)|Outlook アドインのテストに使用する Exchange Server か Exchange Online のユーザー アカウントの電子メール アドレスを指定します。|
|**EWS の URL** (Outlook アドインのみ)|Exchange Web サービスの URL (例: https://www.contoso.com/ews/exchange.aspx)。 |
|**OWA の URL** (Outlook アドインのみ)|Outlook Web App の URL (例: https://www.contoso.com/owa)。|
|**ユーザー名** (Outlook アドインのみ)|Exchange Server または Exchange Online のユーザー アカウントの名前を指定します。|
|**プロジェクト ファイル**|ビルド、構成、およびその他のプロジェクト情報が含まれているファイルの名前を指定します。|
|**プロジェクト フォルダー**|プロジェクト ファイルの場所です。|

### <a name="use-an-existing-document-to-debug-the-add-in-content-and-task-pane-add-ins-only"></a>既存のドキュメントを使用してアドインをデバッグする (コンテンツ アドインと作業ウィンドウ アドインのみ)


アドイン プロジェクトにドキュメントを追加できます。アドインで使用するテスト データを含むドキュメントがある場合、プロジェクトの開始時に Visual Studio によってそのドキュメントが開かれます。


### <a name="to-use-an-existing-document-to-debug-the-add-in"></a>既存のドキュメントを使用してアドインをデバッグするには


1. **ソリューション エクスプローラー**で、アドイン プロジェクト フォルダーを選択します。
    
    > [!NOTE]
    > Web アプリケーション プロジェクトではなく、アドイン プロジェクトを選択します。

2. **[プロジェクト]** メニューで、**[既存の項目の追加]** を選択します。
    
3. [ **既存の項目の追加**] ダイアログ ボックスで、追加するドキュメントを探して選択します。
    
4. [ **追加**] を選択して、ドキュメントをプロジェクトに追加します。
    
5. **ソリューション エクスプローラー**で、プロジェクトのショートカット メニューを開き、[ **プロパティ**] を選択します。
    
    プロジェクトのプロパティ ページが表示されます。
    
6. [ **開始ドキュメント**] 一覧で、プロジェクトに追加したドキュメントを選択し、[ **OK**] を選択してプロパティ ページを閉じます。
    

### <a name="start-the-solution"></a>ソリューションの起動


Visual Studio は、起動すると、自動的にソリューションをビルドします。ソリューションを起動するには、 **メニュー** バーから **[デバッグ]**、 **[開始]** の順に選択します。 


> [!NOTE]
> Internet Explorer でスクリプトのデバッグが有効になっていない場合は、Visual Studio でデバッガーを起動することはできません。スクリプトのデバッグを有効にするには、**[インターネット オプション]** ダイアログ ボックスを開いて、**[詳細設定]** タブをクリックし、**[スクリプトのデバッグを使用しない (Internet Explorer)]** チェック ボックスおよび **[スクリプトのデバッグを使用しない (その他)]** チェック ボックスをオフにします。

Visual Studio はプロジェクトをビルドし、次の操作を実行します。


1. XML マニフェスト ファイルのコピーを作成し、それを  _プロジェクト名_\Output ディレクトリに追加します。このコピーは、Visual Studio を起動してアドインをデバッグするときにホスト アプリケーションで使用されます。
    
2. アドインをホストアプリケーションに表示するための一連のレジストリエントリをコンピューターに作成します。
    
3. Web アプリケーション プロジェクトをビルドし、ローカルの IIS Web サーバー (http://localhost)) に展開します。 
    
次に、Visual Studio は次の操作を実行します。


1. ~remoteAppUrlトークンを開始ページの完全修飾アドレス (例: http://localhost/MyAgave.html)) で置き換えることによって、XML マニフェストファイルの  [SourceLocation](https://docs.microsoft.com/javascript/office/manifest/sourcelocation?view=office-js)  要素を変更します。
    
2. IIS Express で Web アプリケーション プロジェクトを起動します。
    
3. ホスト アプリケーションを開きます。 
    
プロジェクトをビルドする際、Visual Studio は **出力**ウィンドウに検証エラーを表示しません。Visual Studio は、エラーと警告を、発生時に  **ERRORLIST** ウィンドウ内で報告します。Visual Studio は、コードおよびテキスト エディター内で検証エラーを別の色の波形の下線 (波線と呼びます) で示します。このようなマークにより、Visual Studio がコード内で検出した問題が通知されます。詳細については、「 [コードおよびテキスト エディター](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)」を参照してください。検証を有効化または無効化する方法の詳細については、次のトピックを参照してください。 

- [[オプション]、[テキスト エディター]、[JavaScript]、[IntelliSense]](https://docs.microsoft.com/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2015)
    
- [方法:Visual Web Developer で HTML 編集用の検証オプションを設定する](https://msdn.microsoft.com/library/0byxkfet(v=vs.100).aspx)
    
- [[検証] ([オプション] ダイアログ ボックス - [テキスト エディター] - [CSS])](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)
    
プロジェクト内の XML マニフェスト ファイルの検証ルールを確認するには、「[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。


### <a name="show-an-add-in-in-excel-word-or-project-and-step-through-your-code"></a>Excel、Word、または Project にアドインを表示し、コードをステップ実行する


アドイン プロジェクトの **開始ドキュメント** プロパティを Excel または Word に設定した場合、Visual Studio はドキュメントを新規作成し、アドインが表示されます。アドイン プロジェクトの **開始ドキュメント** プロパティを既存のドキュメントを使用するように設定した場合、Visual Studio はドキュメントを開きますが、アドインは手動で挿入する必要があります。 **開始ドキュメント** を **Microsoft Project** に設定した場合も、アドインを手動で挿入する必要があります。


### <a name="to-show-an-office-add-in-in-excel-or-word"></a>Office アドイン を Excel または Word で表示するには


1. Excel または Word の  **[挿入]** タブで、 **[Office アドイン]** を選択します。
    
2. 表示される一覧で、アドインを選択します。
    

### <a name="to-show-an-office-add-in-in-project"></a>Project で Office アドインを表示するには


1. Project の  **[プロジェクト]** タブで、 **[Office アドイン]** を選択します。
    
2. 表示される一覧で、アドインを選択します。
    
Visual Studio でブレークポイントを設定できます。ブレークポイントを設定したら、アドインを操作し、HTML、JavaScript、および C# か VB のコード ファイル内のコードをステップ実行します。


### <a name="show-the-outlook-add-in-in-outlook-and-step-through-your-code"></a>Outlook で Outlook のアドインを表示し、コードをステップ実行する


Outlook でアドインを表示するには、電子メール メッセージまたは予定アイテムを開きます。

Outlook は、アクティブ化の基準を満たしていれば、アイテムの アドイン をアクティブ化します。アドイン バーが [インスペクタ] ウィンドウまたは閲覧ウィンドウの上部に表示され、Outlook アドインがアドイン バーにボタンとして表示されます。アドインにアドイン コマンドがある場合は、リボンの既定のタブまたは指定されたカスタム タブのいずれかにボタンが表示され、アドイン バーにはアドインは表示されません。

Outlook アドインを表示するには、Outlook アドインのボタンを選択します。

Visual Studio では、ブレークポイントを設定できます。ブレークポイントを設定した後、Outlook アドインを操作して、HTML、JavaScript、および C# または VB のコード ファイルのコードをステップ実行します。 

また、コードを変更してから、Office アドイン を終了してプロジェクトを再度起動しなくても、Outlook アドインへの影響を確認することができます。Outlook で Outlook アドインのショートカット メニューを開き、 **[再読み込み]** を選択します。


### <a name="modify-code-and-continue-to-debug-the-add-in-without-having-to-start-the-project-again"></a>コードを変更した後、プロジェクトを再び開始することなくアドインのデバッグを続行する


コードを変更したら、ホスト アプリケーションを閉じてプロジェクトを再起動しなくても、アドインに対する変更の影響を確認することができます。コードを変更した後、アドインのショートカット メニューを開き、 **[再読み込み]** を選択します。アドインを再読み込みすると、アドインは Visual Studio デバッガーと切断された状態になります。そのため、変更の影響を確認することはできても、利用できるすべての Iexplore.exe プロセスに Visual Studio デバッガーをアタッチするまでは、コードをステップ実行していくことはできません。


### <a name="to-attach-the-visual-studio-debugger-to-all-of-the-available-iexploreexe-processes"></a>使用可能な Iexplore.exe プロセスのすべてに Visual Studio デバッガーをアタッチするには


1. Visual Studio で、 **[デバッグ]**、 **[プロセスにアタッチ]** の順に選択します。
    
2. [ **プロセスにアタッチ**] ダイアログ ボックスで、利用可能なすべての  **Iexplore.exe** プロセスを選択して、 [ **アタッチ**] を選択します。
    

## <a name="next-steps"></a>次の手順

- [Office アドインを展開し、発行する](../publish/publish.md)
    
