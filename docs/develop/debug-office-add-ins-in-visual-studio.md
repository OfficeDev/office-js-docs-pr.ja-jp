---
title: Visual Studio で Office アドインをデバッグする
description: Visual Studio を使用して、Windows 上の Office デスクトップ クライアントで Office アドインをデバッグする
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 08f8b48666955db413e3bdaa6c329326f80bdb07
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340261"
---
# <a name="debug-office-add-ins-in-visual-studio"></a>Visual Studio で Office アドインをデバッグする

この記事では、Office 2022 の Office Visual Studio アドイン プロジェクト テンプレートの 1 つで作成された Office アドインでクライアント側コードをデバッグする方法について説明します。  Office アドインでのサーバー側コードのデバッグの詳細については、「Office アドインのデバッグの概要 [- サーバー](../testing/debug-add-ins-overview.md#server-side-or-client-side)側またはクライアント側?」を参照してください。

> [!NOTE]
> Mac でアドインVisual Studioをデバッグする場合は、Office使用することはできません。 Mac でのデバッグの詳細については、「Debug [Officeアドイン」を参照してください](../testing/debug-office-add-ins-on-ipad-and-mac.md)。

## <a name="review-the-build-and-debug-properties"></a>ビルドとデバッグのプロパティの確認

デバッグを開始する前に、各プロジェクトのプロパティを確認して、Visual Studio が目的の Office アプリケーションを開き、他のビルドプロパティとデバッグ プロパティが適切に設定されているのを確認します。

### <a name="add-in-project-properties"></a>アドイン プロジェクトのプロパティ

アドイン プロジェクトの **[プロパティ** ] ウィンドウを開き、プロジェクトのプロパティを確認します。

1. **ソリューション エクスプローラー** で、(Web アプリケーション プロジェクトでは *なく*) アドイン プロジェクトを選択します。

2. メニュー バーから [**表示**]、[**プロパティ ウィンドウ**] の順に選択します。

次の表では、アドイン プロジェクトのプロパティについて説明します。

|プロパティ|説明|
|:-----|:-----|
|**開始動作**|ご自身のアドインに対してデバッグ モードを指定します。 これは、**アドインのMicrosoft Edge** にOutlookする必要があります。 他のすべてのアプリケーションOfficeデスクトップ クライアントに設定するOffice **必要があります**。|
|**開始ドキュメント**<br/>(Excel、PowerPoint、Word アドインのみ)|プロジェクトの開始時に開くドキュメントを指定します。 新しいプロジェクトでは、[ブックの新しいExcel **]**、[新しい Word ドキュメント]、または [新しい文書のプレゼンテーション **]** **にPowerPointされます**。 特定のドキュメントを指定するには、「既存のドキュメントを使用してアドインをデバッグする」の手順 [に従います](#use-an-existing-document-to-debug-the-add-in)。|
|**Web プロジェクト**|アドインに関連付けられている Web プロジェクトの名前を指定します。|
|**メール アドレス**<br/>(Outlook アドインのみ)|Outlook アドインのテストに使用する Exchange Server または Exchange Online のユーザー アカウントのメール アドレスを指定します。 空白の場合は、デバッグを開始するときに電子メール アドレスの入力を求めるメッセージが表示されます。|
|**EWS の URL**<br/>(Outlook アドインのみ)|Web サービス URL Exchange指定します (例: `https://www.contoso.com/ews/exchange.aspx`)。 このプロパティは空白のままにできます。|
|**OWA の URL**<br/>(Outlook アドインのみ)|URL のOutlook on the web指定します (次に例を示します`https://www.contoso.com/owa`)。 このプロパティは空白のままにできます。|
|**多要素認証を使用する**<br/>(Outlook アドインのみ)|多要素認証を使用するかどうかを示すブール値を指定します。 既定値は **false ですが**、プロパティは実用的な効果はありません。 通常、電子メール アカウントにログインする第 2 の要素を指定する必要がある場合は、デバッグを開始するときにメッセージが表示されます。 |
|**ユーザー名**<br/>(Outlook アドインのみ)|Outlook アドインのテストに使用する Exchange Server または Exchange Online のユーザー アカウントの名前を指定します。 このプロパティは空白のままにできます。|
|**プロジェクト ファイル**|ビルド、構成、およびその他のプロジェクト情報が含まれているファイルの名前を指定します。|
|**プロジェクト フォルダー**|プロジェクト ファイルの場所を指定します。|

> [!NOTE]
> Outlook アドインの場合、[**プロパティ**] ウィンドウで 1 つまたは複数の *Outlook アドインのみ* のプロパティを指定できますが、指定する必要はありません。

### <a name="web-application-project-properties"></a>Web アプリケーション プロジェクトのプロパティ

Web アプリケーション **プロジェクトの [** プロパティ] ウィンドウを開き、プロジェクトのプロパティを確認します。

1. ソリューション **エクスプローラーで、** Web アプリケーション プロジェクトを選択します。

2. メニュー バーから [**表示**]、[**プロパティ ウィンドウ**] の順に選択します。

次の表では、Office アドイン プロジェクトに最も関連する Web アプリケーション プロジェクトのプロパティについて説明します。

|プロパティ|説明|
|:-----|:-----|
|**SSL 有効**|サイトで SSL を有効にするかどうかを指定します。 Office アドイン プロジェクトの場合、このプロパティを **True** に設定する必要があります。|
|**SSL URL**|サイトにセキュリティで保護された HTTPS URL を指定します。 読み取り専用です。|
|**URL**|サイトに HTTP URL を指定します。 読み取り専用です。|
|**プロジェクト ファイル**|ビルド、構成、およびその他のプロジェクト情報が含まれているファイルの名前を指定します。|
|**プロジェクト フォルダー**|プロジェクト ファイルの場所を指定します。 読み取り専用です。 Visual Studio で実行時に生成されるマニフェスト ファイルは、この場所の `bin\Debug\OfficeAppManifests` フォルダーに書き込まれます。|

## <a name="debug-an-excel-powerpoint-or-word-add-in-project"></a>アドイン プロジェクトExcel、PowerPoint、または Word アドイン プロジェクトをデバッグする

このセクションでは、Word アドイン、Excel、PowerPointを開始およびデバッグする方法について説明します。

### <a name="start-the-excel-powerpoint-or-word-add-in-project"></a>Word アドイン Excel、PowerPoint、または Word アドイン プロジェクトを開始する

メニュー バーから [**DebugStart** >  **Debuging**] を選択するか、F5 ボタンを押してプロジェクトを開始します。 Visual Studioソリューションが自動的にビルドされ、ホスト アプリケーションOffice開始されます。

プロジェクトVisual Studio、次のタスクを実行します。

1. XML マニフェスト ファイルのコピーを作成し、ディレクトリに追加  `_ProjectName_\bin\Debug\OfficeAppManifests` します。 アドインOfficeホストするアプリケーションは、アドインのインストールとデバッグを開始Visual Studioこのコピーを使用します。

2. アドインをアプリケーションに表示できるWindowsコンピューターにレジストリ エントリのセットをOfficeします。

3. Web アプリケーション プロジェクトをビルドし、ローカル IIS Web サーバー () に展開します`https://localhost`。

4. これがローカル IIS Web サーバーに展開した最初のアドイン プロジェクトである場合は、Self-Signed 証明書を現在のユーザーの信頼されたルート証明書ストアにインストールするように求めるメッセージが表示される場合があります。 これは、IIS Express がアドインの内容を正しく表示するために必要です。

> [!NOTE]
> エッジ レガOffice Web ビュー コントロール (EdgeHTML) を使用して Windows コンピューターでアドインを実行する場合、Visual Studio はローカル ネットワーク ループバックの除外を追加するように求めるメッセージを表示する場合があります。 これは、Webview コントロールがローカル IIS Web サーバーに展開されている Web サイトにアクセスするために必要です。 この設定は、Visual Studio の **[ツール]** > **[オプション]** > **[Office ツール (Web)]** > **[Web アドインのデバッグ]** の順に選択して変更することもできます。 お使いのコンピューターで使用されているブラウザー コントロールWindows、アドインで使用されるブラウザー [Office参照してください](../concepts/browsers-used-by-office-web-add-ins.md)。

次に、Visual Studio で次の操作が行われます。

1. トークンをスタート ページの完全修飾アドレス (たとえば) に置き換え、XML マニフェスト ファイル (`_ProjectName_\bin\Debug\OfficeAppManifests`ディレクトリにコピーされた) `~remoteAppUrl` の [SourceLocation](../reference/manifest/sourcelocation.md) 要素を変更します`https://localhost:44302/Home.html`。

2. IIS Express で Web アプリケーション プロジェクトを起動します。

3. マニフェストを検証します。 プロジェクト内の XML マニフェスト ファイルの検証ルールを確認するには、「[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。 

   > [!IMPORTANT]
   > インストールOfficeマニフェスト XSD ファイルVisual Studioは古いものです。 マニフェストの検証エラーが発生した場合は、最初のトラブルシューティング手順として、これらのファイルの 1 つ以上を最新バージョンに置き換える必要があります。 詳細な手順については、「マニフェスト スキーマ検証エラー[」を参照Visual Studioしてください](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects)。

4. アプリケーションをOfficeし、アドインをサイドロードします。

### <a name="debug-the-excel-powerpoint-or-word-add-in"></a>Word アドインExcel、PowerPoint、または Word アドインをデバッグする

1. アドインをアプリケーションで起動Officeします。 たとえば、作業ウィンドウ アドインの場合、ホーム リボンにボタンが追加されます (たとえば、[タスクウィンドウの表示] **ボタンなど)。** リボンのボタンを選択します。 

   > [!NOTE]
   > アドインがユーザーによってサイドロードされていないVisual Studio手動でサイドロードできます。 [Excel、PowerPoint Word で、[挿入] タブを選択し、[マイ アドイン] の右側にある下矢印 **を選択します**。
   >
   > ![[マイ アドイン] 矢印Excel強調表示Windows上にリボンを挿入するを示すスクリーンショット。](../images/excel-cf-register-add-in-1b.png)
   >
   > 使用可能なアドインのリストから **[開発者向けアドイン]** セクションを見つけ、ご利用のアドインを選択して登録します。

   > [!TIP]
   > 作業ウィンドウが最初に開くと、作業ウィンドウが空白に表示される場合があります。 その場合は、後の手順でデバッグ ツールを起動すると、正しく表示されます。

3. [パーソ [ナリティ] メニューを開き、[](../design/task-pane-add-ins.md#personality-menu) デバッガーの接続 **] を選択します**。 これにより、webview コントロールのデバッグ ツールが開Officeコンピューターでアドインを実行するために使用Windowsされます。 次のいずれかの記事で説明するように、ブレークポイントを設定し、コードをステップ実行できます。

    - [Internet Explorer の開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-f12-tools-ie.md)
    - [Edge レガシー用の開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-devtools-edge-legacy.md)
    - [Microsoft Edge (Chromium ベース)で開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-devtools-edge-chromium.md)

4. コードを変更するには、最初にデバッグ セッションを停止し、Visual StudioアプリケーションをOfficeします。 変更を加え、新しいデバッグ セッションを開始します。

## <a name="debug-an-outlook-add-in-project"></a>アドイン プロジェクトOutlookデバッグする

このセクションでは、アドインを起動してデバッグするOutlook説明します。

### <a name="start-the-outlook-add-in-project"></a>アドイン プロジェクトOutlookを開始する

メニュー バーから [**DebugStart** >  **Debuging**] を選択するか、F5 ボタンを押してプロジェクトを開始します。 Visual Studioソリューションが自動的にビルドされ、テナントの [Outlook] ページがMicrosoft 365されます。

プロジェクトVisual Studio、次のタスクを実行します。

1. ログイン資格情報の入力を求めるメッセージが表示されます。 繰り返しサインインする必要がある場合や、承認されていないというエラーが表示された場合は、Microsoft 365 テナントのアカウントに対して Basic Auth が無効になる可能性があります。 この場合、代わりに Microsoft アカウントを使用してみます。 また、[Web アドイン プロジェクトのプロパティ **] ウィンドウの** [複数要素認証を使用する] プロパティを **True** Outlook設定することもできます。 「 [アドイン プロジェクトのプロパティ」を参照してください](#add-in-project-properties)。

1. XML マニフェスト ファイルのコピーを作成し、ディレクトリに追加 `_ProjectName_\bin\Debug\OfficeAppManifests` します。 Outlookを開始し、アドインをデバッグVisual Studioこのコピーを使用します。

2. Web アプリケーション プロジェクトをビルドし、ローカル IIS Web サーバー () に展開します`https://localhost`。

3. これがローカル IIS Web サーバーに展開した最初のアドイン プロジェクトである場合は、Self-Signed 証明書を現在のユーザーの信頼されたルート証明書ストアにインストールするように求めるメッセージが表示される場合があります。 これは、IIS Express がアドインの内容を正しく表示するために必要です。

> [!NOTE]
> エッジ レガOffice Web ビュー コントロール (EdgeHTML) を使用して Windows コンピューターでアドインを実行する場合、Visual Studio はローカル ネットワーク ループバックの除外を追加するように求めるメッセージを表示する場合があります。 これは、Webview コントロールがローカル IIS Web サーバーに展開されている Web サイトにアクセスするために必要です。 この設定は、Visual Studio の **[ツール]** > **[オプション]** > **[Office ツール (Web)]** > **[Web アドインのデバッグ]** の順に選択して変更することもできます。 お使いのコンピューターで使用されているブラウザー コントロールWindows、アドインで使用されるブラウザー [Office参照してください](../concepts/browsers-used-by-office-web-add-ins.md)。

次に、Visual Studio で次の操作が行われます。

1. トークンをスタート ページの完全修飾アドレス (たとえば) に置き換え、XML マニフェスト ファイル (`_ProjectName_\bin\Debug\OfficeAppManifests`ディレクトリにコピーされた) `~remoteAppUrl` の [SourceLocation](../reference/manifest/sourcelocation.md) 要素を変更します`https://localhost:44302/Home.html`。

2. IIS Express で Web アプリケーション プロジェクトを起動します。

3. マニフェストを検証します。 プロジェクト内の XML マニフェスト ファイルの検証ルールを確認するには、「[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。 

   > [!IMPORTANT]
   > インストールOfficeマニフェスト XSD ファイルVisual Studioは古いものです。 マニフェストの検証エラーが発生した場合は、最初のトラブルシューティング手順として、これらのファイルの 1 つ以上を最新バージョンに置き換える必要があります。 詳細な手順については、「マニフェスト スキーマ検証エラー[」を参照Visual Studioしてください](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects)。

4. テナントのOutlookページをMicrosoft 365開Microsoft Edge。

### <a name="debug-the-outlook-add-in"></a>アドインOutlookデバッグする

1. [メッセージOutlook] ページで、電子メール メッセージまたは予定アイテムを選択して、独自のウィンドウで開きます。 

2. F12 キーを押してエッジ デバッグ ツールを開きます。

3. ツールを開いた後、アドインを起動します。 たとえば、メッセージの上部にあるツール バーで、[その他のアプリ]  ボタンを選択し、開く吹き出しからアドインを選択します。

   ![[その他のアプリ] ボタンと、アドインの名前とアイコンが他のアプリ アイコンと共に表示された吹き出しを示すスクリーンショット。](../images/outlook-more-apps-button.png)

4. 次のいずれかの記事の手順を使用して、ブレークポイントを設定し、コードをステップ実行します。 それぞれに詳細なガイダンスへのリンクがあります。

   - [Edge レガシー用の開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-devtools-edge-legacy.md)
   - [Microsoft Edge (Chromium ベース)で開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-devtools-edge-chromium.md)

   > [!TIP]
   > メソッドまたはアドインが開くと`Office.initialize``Office.onReady`実行されるメソッドで実行されるコードをデバッグするには、ブレークポイントを設定し、アドインを閉じて再度開きます。 これらのメソッドの詳細については、「Initialize [your your Officeアドイン」を参照してください](../develop/initialize-add-in.md)。

5. コードを変更するには、最初にデバッグ セッションを停止し、Visual Studioページを閉Outlookします。 変更を加え、新しいデバッグ セッションを開始します。

## <a name="use-an-existing-document-to-debug-the-add-in"></a>既存のドキュメントを使用してアドインをデバッグする

Excel、PowerPoint、または Word アドインのデバッグ時に使用するテスト データを含むドキュメントがある場合、プロジェクトの開始時にドキュメントが開くように、Visual Studio を構成できます。 アドインのデバッグ時に使用する既存のドキュメントを指定するには、次の手順を完了します。

1. **ソリューション エクスプローラー** で、(Web アプリケーション プロジェクトでは *なく*) アドイン プロジェクトを選択します。

2. メニュー バーから [**プロジェクト**]、[**既存のアイテムを追加**] の順に選択します。

3. [**既存のアイテムを追加**] ダイアログ ボックスで、追加するドキュメントを探して選択します。

4. [**追加**] を選択して、ドキュメントをプロジェクトに追加します。

5. **ソリューション エクスプローラー** で、(Web アプリケーション プロジェクトでは *なく*) アドイン プロジェクトを選択します。

6. メニュー バーから [**表示**]、[**プロパティ ウィンドウ**] の順に選択します。

7. [**プロパティ**] ウィンドウで、[**ドキュメントの開始**] リストを選択して、プロジェクトに追加したドキュメントを選択します。 現在、このプロジェクトは、そのドキュメントでアドインを起動するように構成されています。

## <a name="next-steps"></a>次のステップ

アドインが意図したとおりに動作した後、アドインをユーザーに配布する方法については、「[Office アドインを展開し、発行する](../publish/publish.md)」を参照してください。
