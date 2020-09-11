---
title: Visual Studio で Office アドインをデバッグする
description: Visual Studio を使用して、Windows 上の Office デスクトップ クライアントで Office アドインをデバッグする
ms.date: 12/31/2019
localization_priority: Normal
ms.openlocfilehash: 7c49e3019c22af0b5d44a382b33187e5d2de4ceb
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430479"
---
# <a name="debug-office-add-ins-in-visual-studio"></a>Visual Studio で Office アドインをデバッグする

この記事では、Visual Studio 2019 を使用して、Windows 上の Office デスクトップ クライアントで Office アドインをデバッグする方法について説明します。 別のバージョンの Visual Studio を使用している場合は、手順が少し異なる可能性があります。 

> [!NOTE]
> Office on the web または Office on Mac では、Visual Studio を使用してアドインをデバッグすることはできません。 これらのプラットフォームでのデバッグについては、「[Office on the web での Office アドインのデバッグ](../testing/debug-add-ins-in-office-online.md)」または「[Mac での Office アドインのデバッグ](../testing/debug-office-add-ins-on-ipad-and-mac.md)」を参照してください。

## <a name="enable-debugging-for-add-in-commands-and-ui-less-code"></a>アドイン コマンドと UI のないコードのデバッグを有効にする

Visual Studio が Windows 上の Office をデバッグする場合、アドインは、Microsoft Internet Explorer または Microsoft Edge ブラウザー インスタンスのいずれかにホストされています。 開発用コンピューターで使用されているブラウザーを確認するには、「[Office アドインによって使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。
> [!NOTE]
> 次の手順では、JS_Debug 環境変数が不要になりました。 詳細については、Microsoft Developer Community サポート フォーラムの「[Office Web アドインでのデバッグ動作](https://developercommunity.visualstudio.com/content/problem/740413/office-development-inconsistent-script-debugging-b.html)」を参照してください。

[!include[Enable debugging on Microsoft Edge DevTools](../includes/enable-debugging-on-edge-devtools.md)]

## <a name="review-the-build-and-debug-properties"></a>ビルドとデバッグのプロパティの確認

デバッグを開始する前に、各プロジェクトのプロパティを確認して、Visual Studio が必要な Office アプリケーションを開き、その他のビルドおよびデバッグのプロパティが適切に設定されていることを確認してください。

### <a name="add-in-project-properties"></a>アドイン プロジェクトのプロパティ

アドイン プロジェクトの [**プロパティ**] ウィンドウを開き、プロジェクト プロパティを確認します。

1. **ソリューション エクスプローラー**で、(Web アプリケーション プロジェクトでは*なく*) アドイン プロジェクトを選択します。

2. メニュー バーから [**表示**]、[**プロパティ ウィンドウ**] の順に選択します。

次の表では、アドイン プロジェクトのプロパティについて説明します。

|**プロパティ**|**説明**|
|:-----|:-----|
|**開始動作**|ご自身のアドインに対してデバッグ モードを指定します。 現在、Office アドイン プロジェクトでは、**Office デスクトップ クライアント** モードのみがサポートされています。|
|**開始ドキュメント**<br/>(Excel、PowerPoint、Word アドインのみ)|プロジェクトの開始時に開くドキュメントを指定します。|
|**Web プロジェクト**|アドインに関連付けられている Web プロジェクトの名前を指定します。|
|**メール アドレス**<br/>(Outlook アドインのみ)|Outlook アドインのテストに使用する Exchange Server または Exchange Online のユーザー アカウントのメール アドレスを指定します。|
|**EWS の URL**<br/>(Outlook アドインのみ)|Exchange Web サービスの URL (例: `https://www.contoso.com/ews/exchange.aspx`)。 |
|**OWA の URL**<br/>(Outlook アドインのみ)|Outlook on the web の URL (例: `https://www.contoso.com/owa`)。|
|**多要素認証を使用する**<br/>(Outlook アドインのみ)|多要素認証を使用する必要があるかどうかを示すブール値。|
|**ユーザー名**<br/>(Outlook アドインのみ)|Outlook アドインのテストに使用する Exchange Server または Exchange Online のユーザー アカウントの名前を指定します。|
|**プロジェクト ファイル**|ビルド、構成、およびその他のプロジェクト情報が含まれているファイルの名前を指定します。|
|**プロジェクト フォルダー**|プロジェクト ファイルの場所です。|

> [!NOTE]
> Outlook アドインの場合、[**プロパティ**] ウィンドウで 1 つまたは複数の *Outlook アドインのみ*のプロパティを指定できますが、指定する必要はありません。

### <a name="web-application-project-properties"></a>Web アプリケーション プロジェクトのプロパティ

Web アプリケーション プロジェクトの [**プロパティ**] ウィンドウを開き、次のプロジェクト プロパティを確認します。

1. **ソリューションエクスプローラー**で、web アプリケーションプロジェクトを選択します。

2. メニュー バーから [**表示**]、[**プロパティ ウィンドウ**] の順に選択します。

次の表では、Office アドイン プロジェクトに最も関連する Web アプリケーション プロジェクトのプロパティについて説明します。

|**プロパティ**|**説明**|
|:-----|:-----|
|**SSL 有効**|サイトで SSL を有効にするかどうかを指定します。 Office アドイン プロジェクトの場合、このプロパティを **True** に設定する必要があります。|
|**SSL URL**|サイトにセキュリティで保護された HTTPS URL を指定します。 読み取り専用です。|
|**URL**|サイトに HTTP URL を指定します。 読み取り専用です。|
|**プロジェクト ファイル**|ビルド、構成、およびその他のプロジェクト情報が含まれているファイルの名前を指定します。|
|**プロジェクト フォルダー**|プロジェクト ファイルの場所を指定します。 読み取り専用です。 Visual Studio で実行時に生成されるマニフェスト ファイルは、この場所の `bin\Debug\OfficeAppManifests` フォルダーに書き込まれます。|

## <a name="use-an-existing-document-to-debug-the-add-in"></a>既存のドキュメントを使用してアドインをデバッグする

Excel、PowerPoint、または Word アドインのデバッグ時に使用するテスト データを含むドキュメントがある場合、プロジェクトの開始時にドキュメントが開くように、Visual Studio を構成できます。 アドインのデバッグ時に使用する既存のドキュメントを指定するには、次の手順を完了します。

1. **ソリューション エクスプローラー**で、(Web アプリケーション プロジェクトでは*なく*) アドイン プロジェクトを選択します。

2. メニュー バーから [**プロジェクト**]、[**既存のアイテムを追加**] の順に選択します。

3. [**既存のアイテムを追加**] ダイアログ ボックスで、追加するドキュメントを探して選択します。

4. [**追加**] を選択して、ドキュメントをプロジェクトに追加します。

5. **ソリューション エクスプローラー**で、(Web アプリケーション プロジェクトでは*なく*) アドイン プロジェクトを選択します。

6. メニュー バーから [**表示**]、[**プロパティ ウィンドウ**] の順に選択します。

7. [**プロパティ**] ウィンドウで、[**ドキュメントの開始**] リストを選択して、プロジェクトに追加したドキュメントを選択します。 現在、このプロジェクトは、そのドキュメントでアドインを起動するように構成されています。

## <a name="start-the-project"></a>プロジェクトの開始

メニュー バーから [**デバッグ**]、[**デバッグの開始 **] の順に選択し、プロジェクトを開始します。 Visual Studio では、自動的にソリューションがビルドされ、Office が起動されてアドインがホストされます。

> [!NOTE]
> Outlook アドイン プロジェクトを開始すると、ログインの資格情報を求めるメッセージが表示されます。 繰り返しログインするように求められた場合、または承認されていないというエラーが表示された場合は、Microsoft 365 テナントのアカウントに対して基本認証を無効にすることができます。 この場合、代わりに Microsoft アカウントを使用してみます。 Outlook Web アドイン プロジェクトのプロパティ ダイアログで、[多要素認証を使用する] プロパティを True に設定する必要がある場合もあります。

Visual Studio によってプロジェクトがビルドされると、次のタスクが実行されます。

1. XML マニフェスト ファイルのコピーを作成し、`_ProjectName_\bin\Debug\OfficeAppManifests` ディレクトリに追加します。 アドインをホストする Office アプリケーションは、Visual Studio を起動してアドインをデバッグするときに、このコピーを使用します。

2. Office アプリケーションにアドインが表示されるようにするためのレジストリエントリのセットをコンピューター上に作成します。

3. Web アプリケーション プロジェクトをビルドし、ローカルの IIS Web サーバー (https://localhost)) に展開します。

4. これがローカル IIS Web サーバーに最初に展開したアドイン プロジェクトである場合は、現在のユーザーの信頼されたルート証明書ストアに自己署名証明書をインストールするように求められることがあります。 これは、IIS Express がアドインの内容を正しく表示するために必要です。

> [!NOTE]
> Windows 10 上で実行している場合、最新バージョンの Office では、新しい Web コントロールを使用してアドインの内容を表示することがあります。 この場合、Visual Studio はローカル ネットワークのループバック除外を追加するように促します。 これは、Office クライアントアプリケーションの web コントロールで、ローカルの IIS web サーバーに展開された web サイトにアクセスできるようにするために必要です。 この設定は、Visual Studio の **[ツール]** > **[オプション]** > **[Office ツール (Web)]** > **[Web アドインのデバッグ]** の順に選択して変更することもできます。

次に、Visual Studio で次の操作が行われます。

1. `~remoteAppUrl` トークンを開始ページの完全修飾アドレス (例: `https://localhost:44302/Home.html`) で置き換えることによって、XML マニフェスト ファイルの [SourceLocation](../reference/manifest/sourcelocation.md) 要素を変更します。

2. IIS Express で Web アプリケーション プロジェクトを起動します。

3. Office アプリケーションを開きます。

Visual Studio では、プロジェクトのビルド時の検証エラーは [**出力**] ウィンドウには表示されません。 Visual Studio では、エラーと警告が発生すると **ERRORLIST** ウィンドウ内で報告されます。 また、Visual Studio では、検証エラーは、コードおよびテキスト エディター内で別の色の波形の下線 (波線と呼ばれる) で報告されます。 このようなマークにより、Visual Studio によってご自身のコード内で検出された問題が通知されます。 検証を有効または無効にする方法の詳細については、「[[オプション]、[テキスト エディター]、[JavaScript]、[IntelliSense]](/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2019&preserve-view=true)」を参照してください。

プロジェクト内の XML マニフェスト ファイルの検証ルールを確認するには、「[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。

## <a name="debug-the-code-for-an-excel-powerpoint-or-word-add-in"></a>Excel、PowerPoint、または Word アドイン用のコードのデバッグ

[プロジェクトを開始](#start-the-project)した後に、office アプリケーション (Excel、PowerPoint、または Word) に表示されているドキュメント内にアドインが表示されていない場合は、手動で office アプリケーションでアドインを起動します。 たとえば、[**ホーム**] タブのリボンで [**作業ウィンドウの表示**] ボタンを選択して作業ウィンドウを起動します。ご利用のアドインが Excel、PowerPoint、または Word 内に表示されたら、次の操作を行うことでご自身のコードをデバッグできます。

1. Excel、PowerPoint、または Word で、[**挿入**] タブを選択し、[**個人用アドイン**] の右側に配置された下向き矢印を選択します。

    ![[個人用アドイン] の矢印が強調表示された Windows での Excel の [挿入] リボン](../images/excel-cf-register-add-in-1b.png)

2. 使用可能なアドインのリストから **[開発者向けアドイン]** セクションを見つけ、ご利用のアドインを選択して登録します。

3. Visual Studio でコードにブレークポイントを設定します。

4. Excel、PowerPoint、または Word でご利用のアドインを操作します。

5. Visual Studio でブレークポイントに達したときは、必要に応じて、コードのステップ実行を行います。

Office アプリケーションを閉じてプロジェクトを再起動しなくても、コードを変更し、その変更によるアドインへの影響を確認できます。 コードに加えた変更を保存した後、Office アプリケーションにアドインを再読み込みするだけです。 たとえば、[パーソナリティ メニュー](../design/task-pane-add-ins.md#personality-menu)をアクティブにして、[**再読み込み**] を選択するには、作業ウィンドウの右上隅を選択して、作業ウィンドウ アドインを再読み込みします。

## <a name="debug-the-code-for-an-outlook-add-in"></a>Outlook アドイン用のコードのデバッグ

[プロジェクトを開始](#start-the-project)して、Visual Studio で Outlook を起動してご利用のアドインをホストした後、メール メッセージまたは予定アイテムを開きます。

Outlook は、アクティブ化の基準を満たしていれば、アイテムの アドイン をアクティブ化します。アドイン バーが [インスペクタ] ウィンドウまたは閲覧ウィンドウの上部に表示され、Outlook アドインがアドイン バーにボタンとして表示されます。アドインにアドイン コマンドがある場合は、リボンの既定のタブまたは指定されたカスタム タブのいずれかにボタンが表示され、アドイン バーにはアドインは表示されません。

Outlook アドインを表示するには、Outlook アドインのボタンを選択します。 ご利用のアドインが Outlook に表示された後、以下の操作を行うことでコードをデバッグできます。

1. Visual Studio でコードにブレークポイントを設定します。

2. Outlook で、ご利用のアドインを操作します。

3. Visual Studio でブレークポイントに達したときは、必要に応じて、コードのステップ実行を行います。

コードを変更し、ご利用のアドインでこれらの変更の影響を確認できます。Outlook を閉じて、プロジェクトを再起動する必要はありません。 コードへの変更を保存した後、(Outlook で) アドインのショートカット メニューを開いて、[**再読み込み**] を選択するだけです。

## <a name="next-steps"></a>次のステップ

アドインが意図したとおりに動作した後、アドインをユーザーに配布する方法については、「[Office アドインを展開し、発行する](../publish/publish.md)」を参照してください。
