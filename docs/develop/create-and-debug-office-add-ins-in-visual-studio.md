---
title: Visual Studio での Office アドインの作成とデバッグ
description: ''
ms.date: 10/01/2018
ms.openlocfilehash: 0bbc1b167924ce4b7a1310f04b41751173312cd6
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506127"
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a>Visual Studio での Office アドインの作成とデバッグ

この記事では、Visual Studio を使用して、最初の Office アドインを作成する方法について説明します。ここに示す手順は Visual Studio 2015 に基づいたものです。別のバージョンの Visual Studio を使用している場合は、わずかに手順が異なることがあります。

> [!NOTE]
> OneNote 用のアドインを使い始めるには、「[最初の OneNote アドインをビルドする](../onenote/onenote-add-ins-getting-started.md)」を参照してください。

## <a name="create-an-office-add-in-project-in-visual-studio"></a>Visual Studio での Office アドイン プロジェクトの作成


作業を開始するために、[Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) がインストールされていることと、Microsoft Office のバージョンを確認します。[Office 365 Developer プログラム](https://developer.microsoft.com/office/dev-program)に参加するか、以下の手順を実行して[最新バージョン](../develop/install-latest-office-version.md)を取得できます。

1. [Visual Studio] メニュー バーで、**[ファイル]**  >  **[新規作成]**  >  **[プロジェクト]** の順に選択します。
2. プロジェクトの種類の一覧で、**[Visual C#]** または **[Visual Basic]** の下にある **[Office/SharePoint]** を展開し、**[Web アドイン]** を選択してからアドイン プロジェクトのいずれかを選択します。
3. プロジェクトに名前を付けて、プロジェクトを作成するために **[OK]** を選択します。

Visual Studio 2017 で **[OK]** を選択した後、次のアドイン プロジェクト テンプレートに追加の選択肢があります 。

**PowerPoint**
- 作業ウィンドウ アドインを作成する **PowerPoint の新しい機能を追加** することができます。
- または **PowerPoint スライドにコンテンツを挿入する** コンテンツを追加で作成することもできます。

**Excel** 
- 作業ウィンドウ アドインを作成する **Excel の新しい機能を追加** することができます。
- または、 **Excel のスプレッドシートにコンテンツを挿入する** コンテンツを追加で作成することもできます。
    - コンテンツを追加で作成する場合の **基本的なアドインを** 最低限のスタート コードとコンテンツの追加のプロジェクトを作成する追加の選択肢があります。
    - または **アドインがドキュメントのビジュアル化** を視覚化し、データにバインドする初期のコードを含むことができます。

ウィザードを完了した後 Visual Studio の 2 つのプロジェクトを含むソリューションを作成します。 Home.html の既定のページが開くことがわかります。

|**プロジェクト**|**説明**|
|:-----|:-----|
|アドイン プロジェクト|アドインを記述するすべての設定を含む XML マニフェスト ファイルのみが含まれます。これらの設定は、Office ホストがアドインをアクティブ化するタイミングと、アドインの表示場所を決定するのに役立ちます。すぐにプロジェクトを実行し、アドインを使用できるように、Visual Studio によってこのファイルのコンテンツが生成されます。これらの設定は、マニフェスト エディターを使用していつでも変更できます。|
|Web アプリケーション プロジェクト|Office 対応 HTML および JavaScript ページを開発するために必要なすべてのファイルおよびファイル参照を含む、アドインのコンテンツ ページが含まれています。 アドインを開発している間、Visual Studio はローカル IIS サーバー上の Web アプリケーションをホストします。 発行する準備ができたら、このプロジェクトをホストするサーバーを検索する必要があります。  ASP.NET Web アプリケーション プロジェクト の詳細については、「[ ASP.NET Web プロジェクト](http://msdn.microsoft.com/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx) 」を参照してください。|

## <a name="modify-your-add-in-settings"></a>アドイン設定の変更


アドインの設定を変更するには、プロジェクトの XML マニフェスト ファイルを編集します。 [**ソリューション エクスプローラー**] で、アドイン プロジェクト ノードを展開し、XML マニフェストを格納するフォルダーを展開して、XML マニフェストを選択します。 ファイル内の任意の要素をポイントして、要素の目的を説明するヒントを表示できます。 マニフェスト ファイルの詳細については、「[Office アドイン XML マニフェスト](../develop/add-in-manifests.md)」をご覧ください。


## <a name="develop-the-contents-of-your-add-in"></a>アドインのコンテンツの開発

アドイン プロジェクトはアドインを説明する設定を変更でき、Web アプリケーションはアドインに表示されるコンテンツを提供します。 

Web アプリケーション プロジェクトには、既定の HTML ページとを開始するのに使用できる JavaScript ファイルが含まれています。 これらのファイルには、Office 用の JavaScript API を含む他の JavaScript ライブラリへの参照が含まれています。 これらのファイルを更新し、さらに HTML と JavaScript ファイルを追加することによって、アドインを開発できます。 次の表は、既定の HTML や JavaScript ファイルについて説明します。

> [!NOTE]
> Web プロジェクトのルート フォルダーで **ホーム** フォルダーを使用してプロジェクト テンプレートの種類に応じて次の表のファイルがあります。

|**ファイル**|**説明**|
|:-----|:-----|
|**Home.html**|アドインの既定の HTML ページです。 アクティブ化すると、ドキュメント、電子メール メッセージ、または予定アイテムでは、アドイン内の最初のページとしてこのページが表示されます。 このファイルには、すべてのファイル参照を開始する必要があるが含まれています。 このファイルに HTML コードを追加することによって、アドインの開発を開始できます。|
|**Home.js**|Home.html ページに関連付けられた JavaScript ファイルです。 Home.js ファイルで Home.html ページの動作に固有のコードを配置することができます。 Home.js ファイルには、開始するためのいくつかのコード例が含まれています。|
|**home.css**|アドインに適用する既定のスタイルを定義します。 デザインとスタイルの Office UI のファブリックを使用することをお勧めします。 詳細については、 [Office アドインの Office UI Fabric](../design/office-ui-fabric.md)を参照してください。|

> [!NOTE]
> これらのファイルを使用する必要はありません。 他のファイルをプロジェクトに自由に追加し、代わりに使用することができます。 別の HTML ファイルをアドインの最初のページとして表示する場合は、マニフェスト エディターを開き、およびファイルの名前に **SourceLocation** プロパティを設定します。

## <a name="debug-your-add-in"></a>アドインのデバッグ

Visual Studio のビルドを提供して、アドインのデバッグを支援するためのプロパティをデバッグします。

### <a name="review-the-build-and-debug-properties"></a>ビルドおよびデバッグ プロパティの確認

ソリューションを起動する前に、Visual Studio で目的のホスト アプリケーションが開けることを確認します。この情報は、アドインのビルドとデバッグに関連する他のプロパティと共に、プロジェクトのプロパティ ページに表示されます。

### <a name="to-open-the-property-pages-of-a-project"></a>プロジェクトのプロパティ ページを開くには

1. **ソリューション エクスプ ローラー**では、Web プロジェクトではなく、基本的なアドイン プロジェクトを選択します。    
2. メニュー バーで、[ **表示**] >   [ **プロパティ ウィンドウ**] の順に選択します。
    
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

1. **ソリューション エクスプローラ**で、アドイン プロジェクト フォルダーを選択します。
    
    > [!NOTE]
    > Web アプリケーション プロジェクトではなく、アドイン プロジェクトを選択します。

2. **[プロジェクト]** メニューで、**[既存の項目の追加]** を選択します。
    
3. [ **既存の項目の追加**] ダイアログ ボックスで、追加するドキュメントを探して選択します。
    
4. [ **追加**] を選択して、ドキュメントをプロジェクトに追加します。
    
5. **ソリューション エクスプローラ**で、アドイン プロジェクト フォルダーを選択します。
6. メニュー バーで、[ **表示**] >  [ **プロパティ ウィンドウ**] の順に選択します。
7. [プロパティ] ウィンドウでは、 **ドキュメントの開始** ] ボックスの一覧を選択し、プロジェクトに追加したドキュメントを選択します。 今すぐプロジェクトを構成して、既存の文書でアドインを起動します。

### <a name="start-the-solution"></a>ソリューションの起動

 **デバッグ**を選択して、メニュー バーからソリューションを開始 > **デバッグを開始**します。 Visual Studio は自動的にソリューションをビルドし、アドインをホストするための Office を起動します。

Visual Studio プロジェクトをビルドするときは、次のタスクを実行します。

1. XML マニフェスト ファイルのコピーを作成し、それを  _プロジェクト名_\Output ディレクトリに追加します。このコピーは、Visual Studio を起動してアドインをデバッグするときにホスト アプリケーションで使用されます。
    
2. アドインをホスト アプリケーションに表示するための一連のレジストリ エントリをコンピューターに作成します。
    
3. Web アプリケーション プロジェクトをビルドし、ローカルの IIS Web サーバー (http://localhost)) に展開します。 
    
次に、Visual Studio は次の操作を実行します。

1. ~remoteAppUrlトークンを開始ページの完全修飾アドレス (例: http://localhost/MyAgave.html)) で置き換えることによって、XML マニフェストファイルの  [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation?view=office-js)  要素を変更します。
    
2. IIS Express で Web アプリケーション プロジェクトを起動します。
    
3. ホスト アプリケーションを開きます。 
    
プロジェクトをビルドする際、Visual Studio は **出力**ウィンドウに検証エラーを表示しません。Visual Studio は、エラーと警告を、発生時に  **ERRORLIST** ウィンドウ内で報告します。Visual Studio は、コードおよびテキスト エディター内で検証エラーを別の色の波形の下線 (波線と呼びます) で示します。このようなマークにより、Visual Studio がコード内で検出した問題が通知されます。詳細については、「 [コードおよびテキスト エディター](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)」を参照してください。検証を有効化または無効化する方法の詳細については、次のトピックを参照してください。 

- [[オプション]、[テキスト エディター]、[JavaScript]、[IntelliSense]](https://docs.microsoft.com/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2015)
    
- [方法: Visual Web Developer で HTML 編集用の検証オプションを設定する](https://msdn.microsoft.com/library/0byxkfet(v=vs.100).aspx)
    
- [[検証] ([オプション] ダイアログ ボックス - [テキスト エディター] - [CSS])](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)
    
プロジェクト内の XML マニフェスト ファイルの検証ルールを確認するには、「[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。

### <a name="show-an-add-in-in-excel-or-word-and-step-through-your-code"></a>Excel または Word でアドインを表示して、コードをステップ実行

アドイン プロジェクトの **開始ドキュメント** プロパティを Excel または Word に設定した場合、Visual Studio はドキュメントを新規作成し、アドインが表示されます。 アドイン プロジェクトの**  開始ドキュメント** プロパティを既存のドキュメントを使用するように設定した場合、Visual Studio はドキュメントを開きますが、アドインは手動で挿入する必要があります。

1. Excel または Word の [ **挿入** ] タブで、ドロップダウン リストの **[アドイン]** を選択します。 ボタン自体ではなくドロップダウン リストから[**Office の アドイン** ] ダイアログを開きます。
2.  **アドインの開発者**下の、アドインを選択します。

Visual Studio は、ブレーク ポイントを設定し、アドインと対話し、HTML や JavaScript ファイルにコードをステップ実行します。

### <a name="show-the-outlook-add-in-in-outlook-and-step-through-your-code"></a>Outlook で Outlook のアドインを表示し、コードをステップ実行する

Outlook でアドインを表示するには、電子メール メッセージまたは予定アイテムを開きます。

Outlook は、アクティブ化の基準を満たしていれば、アイテムの アドイン をアクティブ化します。アドイン バーが [インスペクタ] ウィンドウまたは閲覧ウィンドウの上部に表示され、Outlook アドインがアドイン バーにボタンとして表示されます。アドインにアドイン コマンドがある場合は、リボンの既定のタブまたは指定されたカスタム タブのいずれかにボタンが表示され、アドイン バーにはアドインは表示されません。

Outlook アドインを表示するには、Outlook アドインのボタンを選択します。

Visual Studio は、ブレーク ポイントを設定し、アドインと対話し、HTML や JavaScript ファイルにコードをステップ実行します。

また、コードを変更してから、Office アドイン を終了してプロジェクトを再度起動しなくても、Outlook アドインへの影響を確認することができます。Outlook で Outlook アドインのショートカット メニューを開き、 **[再読み込み]** を選択します。


### <a name="modify-code-and-continue-to-debug-the-add-in-without-having-to-start-the-project-again"></a>コードを変更した後、プロジェクトを再び開始することなくアドインのデバッグを続行する

ホスト アプリケーションを終了してもう一度プロジェクトを開始することなく、コードを変更してアドインのこれらの変更の効果を確認できます。 コードを変更して保存した後、アドインのショートカット メニューを開いて **再読み込み**を選択します。
    

## <a name="next-steps"></a>次の手順

- [Office アドインを展開し、発行する](../publish/publish.md)
    
