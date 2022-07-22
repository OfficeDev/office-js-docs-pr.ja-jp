---
title: Visual Studio で Office アドインをデバッグする
description: Visual Studio を使用して、Windows 上の Office デスクトップ クライアントで Office アドインをデバッグします。
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 09693f81c069aba97740265fa88bf117a937c742
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958714"
---
# <a name="debug-office-add-ins-in-visual-studio"></a>Visual Studio で Office アドインをデバッグする

この記事では、Visual Studio 2022 の Office アドイン プロジェクト テンプレートのいずれかで作成された Office アドインでクライアント側コードをデバッグする方法について説明します。  Office アドインでのサーバー側コードのデバッグの詳細については、「 [Office アドインのデバッグの概要 - サーバー側またはクライアント側?](../testing/debug-add-ins-overview.md#server-side-or-client-side)」を参照してください。

> [!NOTE]
> Visual Studio を使用して Office on Mac でアドインをデバッグすることはできません。 Mac でのデバッグの詳細については、「Mac で [Office アドインをデバッグする](../testing/debug-office-add-ins-on-ipad-and-mac.md)」を参照してください。

## <a name="review-the-build-and-debug-properties"></a>ビルドとデバッグのプロパティの確認

デバッグを開始する前に、各プロジェクトのプロパティを確認して、Visual Studio が目的の Office アプリケーションを開き、その他のビルドプロパティとデバッグ プロパティが適切に設定されていることを確認します。

### <a name="add-in-project-properties"></a>アドイン プロジェクトのプロパティ

アドイン プロジェクトの **[プロパティ** ] ウィンドウを開き、プロジェクトのプロパティを確認します。

1. **ソリューション エクスプローラー** で、(Web アプリケーション プロジェクトでは *なく*) アドイン プロジェクトを選択します。

2. メニュー バーから [**表示**]、[**プロパティ ウィンドウ**] の順に選択します。

次の表では、アドイン プロジェクトのプロパティについて説明します。

|プロパティ|説明|
|:-----|:-----|
|**開始動作**|ご自身のアドインに対してデバッグ モードを指定します。 これは、Outlook アドインの **Microsoft Edge** に設定する必要があります。 その他のすべての Office アプリケーションの場合は、 **Office デスクトップ クライアント** に設定する必要があります。|
|**開始ドキュメント**<br/>(Excel、PowerPoint、Word アドインのみ)|プロジェクトの開始時に開くドキュメントを指定します。 新しいプロジェクトでは、[ **新しい Excel ブック]**、 **[新しい Word ドキュメント]**、または **[新しい PowerPoint プレゼンテーション] に** 設定されます。 特定のドキュメントを指定するには、「 [既存のドキュメントを使用してアドインをデバッグする](#use-an-existing-document-to-debug-the-add-in)」の手順に従います。|
|**Web プロジェクト**|アドインに関連付けられている Web プロジェクトの名前を指定します。|
|**メール アドレス**<br/>(Outlook アドインのみ)|Outlook アドインのテストに使用する Exchange Server または Exchange Online のユーザー アカウントのメール アドレスを指定します。 空白のままにすると、デバッグを開始するときに電子メール アドレスの入力を求めるメッセージが表示されます。|
|**EWS の URL**<br/>(Outlook アドインのみ)|Exchange Web Services URL を指定します (例: `https://www.contoso.com/ews/exchange.aspx`)。 このプロパティは空白のままにできます。|
|**OWA の URL**<br/>(Outlook アドインのみ)|Outlook on the web URL を指定します (例: `https://www.contoso.com/owa`)。 このプロパティは空白のままにできます。|
|**多要素認証を使用する**<br/>(Outlook アドインのみ)|多要素認証を使用するかどうかを示すブール値を指定します。 既定値は **false** ですが、プロパティには実用的な効果はありません。 通常、電子メール アカウントにログインするための 2 番目の要素を指定する必要がある場合は、デバッグを開始するときにメッセージが表示されます。 |
|**ユーザー名**<br/>(Outlook アドインのみ)|Outlook アドインのテストに使用する Exchange Server または Exchange Online のユーザー アカウントの名前を指定します。 このプロパティは空白のままにできます。|
|**プロジェクト ファイル**|ビルド、構成、およびその他のプロジェクト情報が含まれているファイルの名前を指定します。|
|**プロジェクト フォルダー**|プロジェクト ファイルの場所を指定します。|

> [!NOTE]
> Outlook アドインの場合、[**プロパティ**] ウィンドウで 1 つまたは複数の *Outlook アドインのみ* のプロパティを指定できますが、指定する必要はありません。

### <a name="web-application-project-properties"></a>Web アプリケーション プロジェクトのプロパティ

Web アプリケーション プロジェクトの **[プロパティ** ] ウィンドウを開き、プロジェクトのプロパティを確認します。

1. **ソリューション エクスプローラー** で、Web アプリケーション プロジェクトを選択します。

2. メニュー バーから [**表示**]、[**プロパティ ウィンドウ**] の順に選択します。

次の表では、Office アドイン プロジェクトに最も関連する Web アプリケーション プロジェクトのプロパティについて説明します。

|プロパティ|説明|
|:-----|:-----|
|**SSL 有効**|サイトで SSL を有効にするかどうかを指定します。 Office アドイン プロジェクトの場合、このプロパティを **True** に設定する必要があります。|
|**SSL URL**|サイトにセキュリティで保護された HTTPS URL を指定します。 読み取り専用です。|
|**URL**|サイトに HTTP URL を指定します。 読み取り専用です。|
|**プロジェクト ファイル**|ビルド、構成、およびその他のプロジェクト情報が含まれているファイルの名前を指定します。|
|**プロジェクト フォルダー**|プロジェクト ファイルの場所を指定します。 読み取り専用です。 Visual Studio で実行時に生成されるマニフェスト ファイルは、この場所の `bin\Debug\OfficeAppManifests` フォルダーに書き込まれます。|

## <a name="debug-an-excel-powerpoint-or-word-add-in-project"></a>Excel、PowerPoint、または Word アドイン プロジェクトをデバッグする

このセクションでは、Excel、PowerPoint、または Word アドインを開始してデバッグする方法について説明します。

### <a name="start-the-excel-powerpoint-or-word-add-in-project"></a>Excel、PowerPoint、または Word アドイン プロジェクトを開始する

プロジェクトを開始するには、メニュー バーから **[デバッグ** > **の開始** ] を選択するか、F5 ボタンを押します。 Visual Studio によってソリューションが自動的にビルドされ、Office ホスト アプリケーションが起動されます。

Visual Studio によってプロジェクトがビルドされると、次のタスクが実行されます。

1. XML マニフェスト ファイルのコピーを作成し、ディレクトリに  `_ProjectName_\bin\Debug\OfficeAppManifests` 追加します。 アドインをホストする Office アプリケーションは、Visual Studio を起動してアドインをデバッグするときにこのコピーを使用します。

2. Office アプリケーションにアドインを表示できるようにする一連のレジストリ エントリを Windows コンピューターに作成します。

3. Web アプリケーション プロジェクトをビルドし、ローカル IIS Web サーバー (`https://localhost`) にデプロイします。

4. これがローカル IIS Web サーバーにデプロイした最初のアドイン プロジェクトである場合は、現在のユーザーの信頼されたルート証明書ストアにSelf-Signed証明書をインストールするように求められる場合があります。 これは、IIS Express がアドインの内容を正しく表示するために必要です。

> [!NOTE]
> Office が Edge Legacy Webview コントロール (EdgeHTML) を使用して Windows コンピューターでアドインを実行する場合、Visual Studio はローカル ネットワーク ループバックの除外を追加するよう求めるメッセージを表示する場合があります。 これは、Webview コントロールがローカル IIS Web サーバーにデプロイされた Web サイトにアクセスできるようにするために必要です。 この設定は、Visual Studio の **[ツール]** > **[オプション]** > **[Office ツール (Web)]** > **[Web アドインのデバッグ]** の順に選択して変更することもできます。 Windows コンピューターで使用されているブラウザー コントロールを確認するには、「 [Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。

次に、Visual Studio で次の操作が行われます。

1. トークンを開始ページの完全修飾アドレス (たとえば`https://localhost:44302/Home.html`、) に置き換えることで`~remoteAppUrl`、XML マニフェスト ファイルの [SourceLocation](/javascript/api/manifest/sourcelocation) 要素 (ディレクトリに`_ProjectName_\bin\Debug\OfficeAppManifests`コピーされました) を変更します。

2. IIS Express で Web アプリケーション プロジェクトを起動します。

3. マニフェストを検証します。 プロジェクト内の XML マニフェスト ファイルの検証ルールを確認するには、「[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。 

   > [!IMPORTANT]
   > Visual Studio がインストールする Office マニフェスト XSD ファイルは最新ではありません。 マニフェストの検証エラーが発生した場合は、最初のトラブルシューティング手順として、これらのファイルの 1 つ以上を最新バージョンに置き換える必要があります。 詳細な手順については、「 [Visual Studio プロジェクトでのマニフェスト スキーマ検証エラー](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects)」を参照してください。

4. Office アプリケーションを開き、アドインをサイドロードします。

### <a name="debug-the-excel-powerpoint-or-word-add-in"></a>Excel、PowerPoint、または Word アドインをデバッグする

1. Office アプリケーションでアドインを起動します。 たとえば、作業ウィンドウ アドインの場合、 **ホーム** リボンにボタンが追加されます (たとえば、[ **タスクウィンドウの表示** ] ボタン)。 リボンのボタンを選択します。 

   > [!NOTE]
   > アドインが Visual Studio によってサイドロードされていない場合は、手動でサイドロードできます。 Excel、PowerPoint、または Word で、[ **挿入** ] タブを選択し、 **マイ アドイン** の右側にある下矢印を選択します。
   >
   > ![[マイ アドイン] 矢印が強調表示された Windows 上の Excel でリボンを挿入するを示すスクリーンショット。](../images/excel-cf-register-add-in-1b.png)
   >
   > 使用可能なアドインのリストから **[開発者向けアドイン]** セクションを見つけ、ご利用のアドインを選択して登録します。

   > [!TIP]
   > 作業ウィンドウが最初に開いたときに空白で表示されることがあります。 その場合は、後の手順でデバッグ ツールを起動するときに正しくレンダリングされます。

3. [[パーソナリティ] メニュー](../design/task-pane-add-ins.md#personality-menu)を開き、[**デバッガーのアタッチ**] を選択します。 これにより、Office が Windows コンピューターでアドインを実行するために使用している Webview コントロールのデバッグ ツールが開きます。 次のいずれかの記事で説明されているように、ブレークポイントを設定し、コードをステップ実行できます。

    - [Internet Explorer の開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-f12-tools-ie.md)
    - [Edge レガシー用の開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-devtools-edge-legacy.md)
    - [Microsoft Edge (Chromium ベース)で開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-devtools-edge-chromium.md)

4. コードを変更するには、まず Visual Studio でデバッグ セッションを停止し、Office アプリケーションを閉じます。 変更を加え、新しいデバッグ セッションを開始します。

## <a name="debug-an-outlook-add-in-project"></a>Outlook アドイン プロジェクトをデバッグする

このセクションでは、Outlook アドインを開始してデバッグする方法について説明します。

### <a name="start-the-outlook-add-in-project"></a>Outlook アドイン プロジェクトを開始する

プロジェクトを開始するには、メニュー バーから **[デバッグ** > **の開始** ] を選択するか、F5 ボタンを押します。 Visual Studio によってソリューションが自動的にビルドされ、Microsoft 365 テナントの Outlook ページが起動します。

Visual Studio によってプロジェクトがビルドされると、次のタスクが実行されます。

1. ログイン資格情報の入力を求めるメッセージが表示されます。 繰り返しサインインするように求められた場合、または承認されていないというエラーが表示された場合は、Microsoft 365 テナントのアカウントに対して Basic Auth が無効になる可能性があります。 この場合、代わりに Microsoft アカウントを使用してみます。 Outlook Web アドイン プロジェクトのプロパティ ウィンドウで、 **多要素認証を使用** するプロパティを **True** に設定することもできます。 「 [アドイン プロジェクトのプロパティ](#add-in-project-properties)」を参照してください。

1. XML マニフェスト ファイルのコピーを作成し、ディレクトリに `_ProjectName_\bin\Debug\OfficeAppManifests` 追加します。 Visual Studio を起動し、アドインをデバッグすると、Outlook はこのコピーを使用します。

2. Web アプリケーション プロジェクトをビルドし、ローカル IIS Web サーバー (`https://localhost`) にデプロイします。

3. これがローカル IIS Web サーバーにデプロイした最初のアドイン プロジェクトである場合は、現在のユーザーの信頼されたルート証明書ストアにSelf-Signed証明書をインストールするように求められる場合があります。 これは、IIS Express がアドインの内容を正しく表示するために必要です。

> [!NOTE]
> Office が Edge Legacy Webview コントロール (EdgeHTML) を使用して Windows コンピューターでアドインを実行する場合、Visual Studio はローカル ネットワーク ループバックの除外を追加するよう求めるメッセージを表示する場合があります。 これは、Webview コントロールがローカル IIS Web サーバーにデプロイされた Web サイトにアクセスできるようにするために必要です。 この設定は、Visual Studio の **[ツール]** > **[オプション]** > **[Office ツール (Web)]** > **[Web アドインのデバッグ]** の順に選択して変更することもできます。 Windows コンピューターで使用されているブラウザー コントロールを確認するには、「 [Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。

次に、Visual Studio で次の操作が行われます。

1. トークンを開始ページの完全修飾アドレス (たとえば`https://localhost:44302/Home.html`、) に置き換えることで`~remoteAppUrl`、XML マニフェスト ファイルの [SourceLocation](/javascript/api/manifest/sourcelocation) 要素 (ディレクトリに`_ProjectName_\bin\Debug\OfficeAppManifests`コピーされました) を変更します。

2. IIS Express で Web アプリケーション プロジェクトを起動します。

3. マニフェストを検証します。 プロジェクト内の XML マニフェスト ファイルの検証ルールを確認するには、「[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。 

   > [!IMPORTANT]
   > Visual Studio がインストールする Office マニフェスト XSD ファイルは最新ではありません。 マニフェストの検証エラーが発生した場合は、最初のトラブルシューティング手順として、これらのファイルの 1 つ以上を最新バージョンに置き換える必要があります。 詳細な手順については、「 [Visual Studio プロジェクトでのマニフェスト スキーマ検証エラー](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects)」を参照してください。

4. Microsoft Edge で Microsoft 365 テナントの Outlook ページを開きます。

### <a name="debug-the-outlook-add-in"></a>Outlook アドインをデバッグする

1. Outlook ページで、電子メール メッセージまたは予定アイテムを選択して、独自のウィンドウで開きます。 

2. F12 キーを押して、Edge デバッグ ツールを開きます。

3. ツールを開いた後、アドインを起動します。 たとえば、メッセージの上部にあるツール バーで、[ **その他のアプリ** ] ボタンを選択し、開いた吹き出しからアドインを選択します。

   ![[その他のアプリ] ボタンと、アドインの名前とアイコンが他のアプリ アイコンと共に表示された吹き出しを示すスクリーンショット。](../images/outlook-more-apps-button.png)

4. ブレークポイントを設定し、コードをステップ実行するには、次のいずれかの記事の手順を使用します。 それぞれ、より詳細なガイダンスへのリンクがあります。

   - [Edge レガシー用の開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-devtools-edge-legacy.md)
   - [Microsoft Edge (Chromium ベース)で開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-devtools-edge-chromium.md)

   > [!TIP]
   > 関数または`Office.onReady`アドインが開いたときに実行される関数で`Office.initialize`実行されるコードをデバッグするには、ブレークポイントを設定し、アドインを閉じてから再度開きます。 これらの関数の詳細については、「 [Office アドインを初期化する](../develop/initialize-add-in.md)」を参照してください。

5. コードを変更するには、まず Visual Studio でデバッグ セッションを停止し、Outlook ページを閉じます。 変更を加え、新しいデバッグ セッションを開始します。

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
