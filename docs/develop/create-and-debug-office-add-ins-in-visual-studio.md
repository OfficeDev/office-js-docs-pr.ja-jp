---
title: Visual Studio での Office アドインの作成とデバッグ
description: Visual Studio を使用して、Windows 用の Office デスクトップ クライアントで Office アドインを作成し、デバッグします
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: f9a52719ed7990063ed3f2dbb7d6bd5866e73760
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870717"
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a>Visual Studio での Office アドインの作成とデバッグ

この記事では、Visual Studio 2017 を使用して、Excel、Word、PowerPoint、または Outlook の Office アドインを作成し、Windows 上の Office デスクトップ クライアントでそのアドインをデバッグする方法について説明します。 別のバージョンの Visual Studio を使用している場合は、わずかに手順が異なることがあります。

> [!NOTE]
> Visual Studio では、OneNote または Project 用の Office アドインの作成はサポートされていませんが、[Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)を使用してこれらの種類のアドインを作成できます。
> - OneNote 用のアドインを使い始めるには、「[最初の OneNote アドインをビルドする](../quickstarts/onenote-quickstart.md)」を参照してください。
>
> - Project 用のアドインを使い始めるには、「[最初の Project アドインをビルドする](../quickstarts/project-quickstart.md)」を参照してください。

## <a name="prerequisites"></a>前提条件

- **Office/SharePoint 開発**ワークロードがインストールされている [Visual Studio 2017](https://www.visualstudio.com/vs/)

    > [!TIP]
    > 既に Visual Studio 2017 がインストールされている場合は、[Visual Studio インストーラー](/visualstudio/install/modify-visual-studio)を使用して、**Office/SharePoint 開発**ワークロードがインストールされていることを確認してください。 このワークロードがまだインストールされていない場合は、Visual Studio インストーラーを使用して[インストール](/visualstudio/install/modify-visual-studio?view=vs-2017#modify-workloads)してください。

- Office 2013 以降

    > [!TIP]
    > まだ Office をお持ちでない場合は、[Office 365 開発者プログラム](https://developer.microsoft.com/office/dev-program)に参加して Office 365 サブスクリプションを取得するか、[1 か月間無料試用に登録](https://products.office.com/en-US/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735)することができます。

## <a name="create-the-add-in-project-in-visual-studio"></a>Visual Studio でアドイン プロジェクトを作成する

最初に次の 3 つの手順を完了して、以下の作成しているアドインの種類に対応するセクションの手順を完了します。 

1. Visual Studio を開いて、Visual Studio のメニュー バーから、[**ファイル**]、[**新規作成**]、[**プロジェクト**] の順に選択します。

2. [**Visual C#**] または [**Visual Basic**] の下にあるプロジェクトの種類のリストで、[**Office/SharePoint**] を展開し、[**アドイン**] を選択して、作成するアドイン プロジェクトの種類を選択します。 

3. プロジェクトに名前を付けて、[**OK**] を選択します。

### <a name="word-web-add-in-or-outlook-web-add-in"></a>Word Web アドインまたは Outlook Web アドイン

**Word Web アドイン**または **Outlook Web アドイン**を作成することを選択した場合、Visual Studio でソリューションが作成され、その 2 つのプロジェクトが**ソリューション エクスプローラー**に表示されます。 次に、[Visual Studio ソリューションを調べる](#explore-the-visual-studio-solution)ことができます。 

### <a name="powerpoint-web-add-in"></a>PowerPoint Web アドイン

**PowerPoint Web アドイン**を作成することを選択した場合、[**Office アドインの作成**] ダイアログが表示されます。 

- 作業ウィンドウ アドインを作成するには、[**新機能を PowerPoint に追加する**] を選択して [**完了**] ボタンを選び、Visual Studio ソリューションを作成します。

- コンテンツ アドインを作成するには、[**コンテンツを PowerPoint スライドに挿入する**] を選択して [**完了**] ボタンを選び、Visual Studio ソリューションを作成します。

次に、[Visual Studio ソリューションを調べる](#explore-the-visual-studio-solution)ことができます。

### <a name="excel-web-add-in"></a>Excel Web アドイン

**Excel Web アドイン**を作成することを選択した場合、[**Office アドインの作成**] ダイアログが表示されます。 

- 作業ウィンドウ アドインを作成するには、[**新機能を Excel に追加する**] を選択して [**完了**] ボタンを選び、Visual Studio ソリューションを作成します。

- コンテンツ アドインを作成するには、[**コンテンツを Excel スプレッドシートに挿入する**] を選択して [**次へ**] ボタンを選び、次のいずれかのオプションを選択して [**完了**] ボタンを選び、次の Visual Studio ソリューションを作成します。

    - **基本的なアドイン** - 最小のスタート コードでコンテンツ アドイン プロジェクトを作成します

    - **ドキュメント視覚化アドイン** - スタート コードでコンテンツ アドイン プロジェクトを作成し、視覚化してデータにバインドします  

### <a name="explore-the-visual-studio-solution"></a>Visual Studio ソリューションを調べる

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

## <a name="modify-your-add-in-settings"></a>アドイン設定の変更

アドインの設定を変更するには、アドイン プロジェクトの XML マニフェスト ファイルを編集します。 **ソリューション エクスプローラー**で、アドイン プロジェクト ノードを展開し、XML マニフェストを含めるフォルダーを展開して、XML マニフェストを選択します。 ファイル内の任意の要素をポイントして、要素の目的を説明するヒントを表示できます。 マニフェスト ファイルの詳細については、「[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。

## <a name="develop-the-contents-of-your-add-in"></a>アドインのコンテンツの開発

アドイン プロジェクトはアドインを説明する設定を変更でき、Web アプリケーションはアドインに表示されるコンテンツを提供します。 

Web アプリケーション プロジェクトには、開始するために使用できる既定の HTML ファイル、JavaScript ファイル、および CSS ファイルが含まれています。 これらのファイルの一部には、JavaScript API for Office など、他の JavaScript ライブラリへの参照が含まれています。 これらのファイルを更新したり、さらに HTML ファイルと JavaScript ファイルを追加したりすることによって、アドインを開発できます。 次の表では、Visual Studio ソリューションが作成されるときに Web アプリケーション プロジェクトに含まれる既定のファイルについて説明します。

|**ファイル名**|**説明**|
|:-----|:-----|
|**Home.html**<br/>(Excel、PowerPoint、Word)<br/><br/>**MessageRead.html**<br/>(Outlook)|アドインの既定の HTML ページです。 ドキュメント、電子メール メッセージ、または予定アイテム内でアクティブ化されると、このページはアドイン内の最初のページとして表示されます。 このファイルには、開始するのに必要なファイル参照がすべて含まれています。 このファイルに HTML コードを追加することによって、ご自身のアドインの開発を開始できます。|
|**Home.js**<br/>(Excel、PowerPoint、Word)<br/><br/>**MessageRead.js**<br/>(Outlook)|**Home.html** ページ (Excel、PowerPoint、Word) または **MessageRead.html** ページ (Outlook) に関連付けられた JavaScript ファイルです。 このファイルには、**Home.html** ページ (Excel、PowerPoint、Word) または **MessageRead.html** ページ (Outlook) の動作に固有のコードを含める必要があります。 このファイルには、開始するためのいくつかのコード例が含まれています。|
|**Home.css**<br/>(Excel、PowerPoint、Word)<br/><br/>**MessageRead.css**<br/>(Outlook)|ご自身のアドインに適用する既定のスタイルを定義します。 デザインとスタイルには Office UI Fabric を使用することをお勧めします。 詳細については、「[Office アドインでの Office UI Fabric](../design/office-ui-fabric.md)」を参照してください。|

> [!NOTE]
> これらのファイルを使用する必要はありません。 自由に他のファイルをプロジェクトに追加し、代わりに使用してください。 別の HTML ファイルをアドインの最初のページとして表示する必要がある場合は、マニフェスト エディターを開き、ファイルの名前に **SourceLocation** プロパティを設定します。

## <a name="debug-your-add-in"></a>アドインのデバッグ

次のセクションで説明されているように、Visual Studio を使用して、Windows 上の Office デスクトップ クライアントでご自身のアドインをデバッグできます。

- [ビルドとデバッグのプロパティの確認](#review-the-build-and-debug-properties)
- [既存のドキュメントを使用してアドインをデバッグする](#use-an-existing-document-to-debug-the-add-in)
- [プロジェクトの開始](#start-the-project)
- [Excel、PowerPoint、または Word アドイン用のコードのデバッグ](#debug-the-code-for-an-excel-powerpoint-or-word-add-in)
- [Outlook アドイン用のコードのデバッグ](#debug-the-code-for-an-outlook-add-in)

> [!NOTE]
> Visual Studio を使用して、Office Online または Office for Mac で Office アドインをデバッグすることはできません。 これらのプラットフォームのデバッグについては、[Office Online での Office アドインのデバッグ](../testing/debug-add-ins-in-office-online.md)に関するページ、または「[iPad と Mac で Office アドインをデバッグする](../testing/debug-office-add-ins-on-ipad-and-mac.md)」を参照してください。

### <a name="review-the-build-and-debug-properties"></a>ビルドとデバッグのプロパティの確認

デバッグを開始する前に、各プロジェクトのプロパティを確認し、Visual Studio で目的のホスト アプリケーションが開くことと、他のビルドとデバッグのプロパティが適切に設定されていることを確認します。

#### <a name="add-in-project-properties"></a>アドイン プロジェクトのプロパティ

アドイン プロジェクトの [**プロパティ**] ウィンドウを開き、プロジェクト プロパティを確認します。

1. **ソリューション エクスプローラー**で、(Web アプリケーション プロジェクトでは*なく*) アドイン プロジェクトを選択します。

2. メニュー バーで、[**表示**]、[**プロパティ ウィンドウ**] の順に選択します。

次の表では、アドイン プロジェクトのプロパティについて説明します。

|**プロパティ**|**説明**|
|:-----|:-----|
|**開始動作**|ご自身のアドインに対してデバッグ モードを指定します。 現在、Office アドイン プロジェクトでは、**Office デスクトップ クライアント** モードのみがサポートされています。|
|**開始ドキュメント**<br/>(Excel、PowerPoint、Word アドインのみ)|プロジェクトの開始時に開くドキュメントを指定します。|
|**Web プロジェクト**|アドインに関連付けられている Web プロジェクトの名前を指定します。|
|**メール アドレス**<br/>(Outlook アドインのみ)|Outlook アドインのテストに使用する Exchange Server または Exchange Online のユーザー アカウントのメール アドレスを指定します。|
|**EWS の URL**<br/>(Outlook アドインのみ)|Exchange Web サービスの URL (例: `https://www.contoso.com/ews/exchange.aspx`)。 |
|**OWA の URL**<br/>(Outlook アドインのみ)|Outlook Web App の URL (例: `https://www.contoso.com/owa`)。|
|**多要素認証を使用する**<br/>(Outlook アドインのみ)|多要素認証を使用する必要があるかどうかを示すブール値。|
|**ユーザー名**<br/>(Outlook アドインのみ)|Outlook アドインのテストに使用する Exchange Server または Exchange Online のユーザー アカウントの名前を指定します。|
|**プロジェクト ファイル**|ビルド、構成、およびその他のプロジェクト情報が含まれているファイルの名前を指定します。|
|**プロジェクト フォルダー**|プロジェクト ファイルの場所です。|

> [!NOTE]
> Outlook アドインの場合、[**プロパティ**] ウィンドウで 1 つまたは複数の *Outlook アドインのみ*のプロパティを指定できますが、指定する必要はありません。

#### <a name="web-application-project-properties"></a>Web アプリケーション プロジェクトのプロパティ

Web アプリケーション プロジェクトの [**プロパティ**] ウィンドウを開き、次のプロジェクト プロパティを確認します。

1. **ソリューション エクスプローラー**で、アプリケーション プロジェクトを選択します。

2. メニュー バーで、[**表示**]、[**プロパティ ウィンドウ**] の順に選択します。

次の表では、Office アドイン プロジェクトに最も関連する Web アプリケーション プロジェクトのプロパティについて説明します。

|**プロパティ**|**説明**|
|:-----|:-----|
|**SSL 有効**|サイトで SSL を有効にするかどうかを指定します。 Office アドイン プロジェクトの場合、このプロパティを **True** に設定する必要があります。|
|**SSL URL**|サイトにセキュリティで保護された HTTPS URL を指定します。 読み取り専用です。|
|**URL**|サイトに HTTP URL を指定します。 読み取り専用です。|
|**プロジェクト ファイル**|ビルド、構成、およびその他のプロジェクト情報が含まれているファイルの名前を指定します。|
|**プロジェクト フォルダー**|プロジェクト ファイルの場所を指定します。 読み取り専用です。 Visual Studio で実行時に生成されるマニフェスト ファイルは、この場所の `bin\Debug\OfficeAppManifests` フォルダーに書き込まれます。|

### <a name="use-an-existing-document-to-debug-the-add-in"></a>既存のドキュメントを使用してアドインをデバッグする

Excel、PowerPoint、または Word アドインのデバッグ時に使用するテスト データを含むドキュメントがある場合、プロジェクトの開始時にドキュメントが開くように、Visual Studio を構成できます。 アドインのデバッグ時に使用する既存のドキュメントを指定するには、次の手順を完了します。

1. **ソリューション エクスプローラー**で、(Web アプリケーション プロジェクトでは*なく*) アドイン プロジェクトを選択します。

2. メニュー バーから [**プロジェクト**]、[**既存のアイテムを追加**] の順に選択します。

3. [**既存のアイテムを追加**] ダイアログ ボックスで、追加するドキュメントを探して選択します。

4. [**追加**] を選択して、ドキュメントをプロジェクトに追加します。

5. **ソリューション エクスプローラー**で、(Web アプリケーション プロジェクトでは*なく*) アドイン プロジェクトを選択します。

6. メニュー バーから [**表示**]、[**プロパティ ウィンドウ**] の順に選択します。

7. [**プロパティ**] ウィンドウで、[**ドキュメントの開始**] リストを選択して、プロジェクトに追加したドキュメントを選択します。 現在、このプロジェクトは、そのドキュメントでアドインを起動するように構成されています。

### <a name="start-the-project"></a>プロジェクトの開始

メニュー バーから [**デバッグ**]、[**デバッグの開始 **] の順に選択し、プロジェクトを開始します。 Visual Studio では、自動的にソリューションがビルドされ、Office が起動されてアドインがホストされます。

> [!NOTE]
> Outlook アドイン プロジェクトを開始すると、ログインの資格情報を求めるメッセージが表示されます。 繰り返しログインを求められた場合、ご利用の Office 365 テナントのアカウントで基本認証が無効になる場合があります。 この場合、代わりに Microsoft アカウントを使用してみます。

Visual Studio によってプロジェクトがビルドされると、次のタスクが実行されます。

1. XML マニフェスト ファイルのコピーを作成し、`_ProjectName_\bin\Debug\OfficeAppManifests` ディレクトリに追加します。 Visual Studio を起動してアドインをデバッグするときに、ホスト アプリケーションでこのコピーが使用されます。

2. アドインをホスト アプリケーションに表示するための一連のレジストリ エントリをお使いのコンピューターに作成します。

3. Web アプリケーション プロジェクトをビルドし、ローカルの IIS Web サーバー (https://localhost)) に展開します。

次に、Visual Studio で次の操作が行われます。

1. `~remoteAppUrl` トークンを開始ページの完全修飾アドレス (例: `https://localhost:44302/Home.html`) で置き換えることによって、XML マニフェスト ファイルの [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) 要素を変更します。

2. IIS Express で Web アプリケーション プロジェクトを起動します。

3. ホスト アプリケーションを開きます。

プロジェクトをビルドするときに、Visual Studio では**出力**ウィンドウに検証エラーは表示されません。 Visual Studio では、エラーと警告が発生すると **ERRORLIST** ウィンドウ内で報告されます。 また、Visual Studio では、検証エラーは、コードおよびテキスト エディター内で別の色の波形の下線 (波線と呼ばれる) で報告されます。 このようなマークにより、Visual Studio によってご自身のコード内で検出された問題が通知されます。 詳細については、[コードとテキスト エディター](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)に関するページを参照してください。 検証を有効または無効にする方法の詳細については、「[[オプション]、[テキスト エディター]、[JavaScript]、[IntelliSense]](/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2017)」を参照してください。

プロジェクト内の XML マニフェスト ファイルの検証ルールを確認するには、「[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。

### <a name="debug-the-code-for-an-excel-powerpoint-or-word-add-in"></a>Excel、PowerPoint、または Word アドイン用のコードのデバッグ

[プロジェクトを開始](#start-the-project)した後に、ホスト アプリケーション (Excel、PowerPoint、または Word) に表示されたドキュメント内にご利用のアドインが表示されない場合、ホスト アプリケーションでアドインを手動で起動します。 たとえば、[**ホーム**] タブのリボンで [**作業ウィンドウの表示**] ボタンを選択して作業ウィンドウを起動します。ご利用のアドインが Excel、PowerPoint、または Word 内に表示されたら、次の操作を行うことでご自身のコードをデバッグできます。

1. Excel、PowerPoint、または Word で、[**挿入**] タブを選択し、[**個人用アドイン**] の右側に配置された下向き矢印を選択します。

    ![[個人用アドイン] の矢印が強調表示された Windows 用の Excel の [挿入] リボン](../images/excel-cf-register-add-in-1b.png)

2. 使用可能なアドインのリストから **[開発者向けアドイン]** セクションを見つけ、ご利用のアドインを選択して登録します。

3. Visual Studio でコードにブレークポイントを設定します。

4. Excel、PowerPoint、または Word でご利用のアドインを操作します。

5. Visual Studio でブレークポイントに達したときは、必要に応じて、コードのステップ実行を行います。

コードを変更し、ご利用のアドインでこれらの変更の影響を確認できます。ホスト アプリケーションを閉じて、プロジェクトを再起動する必要はありません。 コードに対する変更を保存した後に、ホスト アプリケーションでアドインを再読み込みするだけです。 たとえば、[パーソナリティ メニュー](../design/task-pane-add-ins.md#personality-menu)をアクティブにして、[**再読み込み**] を選択するには、作業ウィンドウの右上隅を選択して、作業ウィンドウ アドインを再読み込みします。

### <a name="debug-the-code-for-an-outlook-add-in"></a>Outlook アドイン用のコードのデバッグ

[プロジェクトを開始](#start-the-project)して、Visual Studio で Outlook を起動してご利用のアドインをホストした後、メール メッセージまたは予定アイテムを開きます。 

Outlook は、アクティブ化の基準を満たしていれば、アイテムの アドイン をアクティブ化します。アドイン バーが [インスペクタ] ウィンドウまたは閲覧ウィンドウの上部に表示され、Outlook アドインがアドイン バーにボタンとして表示されます。アドインにアドイン コマンドがある場合は、リボンの既定のタブまたは指定されたカスタム タブのいずれかにボタンが表示され、アドイン バーにはアドインは表示されません。

Outlook アドインを表示するには、Outlook アドインのボタンを選択します。 ご利用のアドインが Outlook に表示された後、以下の操作を行うことでコードをデバッグできます。

1. Visual Studio でコードにブレークポイントを設定します。

2. Outlook で、ご利用のアドインを操作します。

3. Visual Studio でブレークポイントに達したときは、必要に応じて、コードのステップ実行を行います。

コードを変更し、ご利用のアドインでこれらの変更の影響を確認できます。Outlook を閉じて、プロジェクトを再起動する必要はありません。 コードへの変更を保存した後、(Outlook で) アドインのショートカット メニューを開いて、[**再読み込み**] を選択するだけです。

## <a name="next-steps"></a>次のステップ

アドインが意図したとおりに動作した後、アドインをユーザーに配布する方法については、「[Office アドインを展開し、発行する](../publish/publish.md)」を参照してください。
