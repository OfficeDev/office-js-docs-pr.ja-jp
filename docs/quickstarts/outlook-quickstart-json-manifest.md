---
title: Teams マニフェスト (プレビュー) を使用して、Outlook アドインをビルドする
description: JSON マニフェストを使用して単純な Outlook 作業ウィンドウ アドインを構築する方法について説明します。
ms.date: 06/06/2022
ms.prod: outlook
ms.localizationpriority: high
ms.openlocfilehash: 407c4ccd4249008c203c760a01d8579989a12e4c
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467224"
---
# <a name="build-an-outlook-add-in-with-a-teams-manifest-preview"></a>Teams マニフェスト (プレビュー) を使用して、Outlook アドインをビルドする

この記事では、選択したメッセージのプロパティを表示し、閲覧ウィンドウで通知をトリガーし、作成ウィンドウのメッセージにテキストを挿入する Outlook 作業ウィンドウ アドインを作成するプロセスについて説明します。 このアドインは、カスタム タブやメッセージング拡張機能などの Teams 拡張機能が使用する JSON 形式のマニフェストのプレビュー バージョンを使用します。 このマニフェストの詳細については、「[Office アドインの Teams マニフェスト (プレビュー)](../develop/json-manifest-overview.md)」を参照してください。

> [!NOTE]
> 新しいマニフェストはプレビューに使用でき、フィードバックに基づいて変更される可能性があります。 経験豊富なアドイン開発者には、それを試してみることをお勧めします。 プレビュー マニフェストは、運用環境のアドインでは使用しないでください。

プレビューは、Windows 上の Microsoft 365 サブスクリプション Office でのみサポートされます。

> [!TIP]
> XML マニフェストを使用して Outlook アドインを構築する場合は、「[最初の Outlook アドインの構築](outlook-quickstart.md)」を参照してください。

## <a name="create-the-add-in"></a>アドインを作成する

[Office アドイン用の Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)を使用して、JSON マニフェストで Office アドインを作成できます。Yeoman ジェネレーターは、Visual Studio Code またはその他のエディターで管理できる Node.js プロジェクトを作成します。

### <a name="prerequisites"></a>前提条件

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]

- Windows 用の [.NET ランタイム](https://dotnet.microsoft.com/download/dotnet/6.0/runtime)。 プレビューで使用されるツールの 1 つは、.NET で実行されます。

[!INCLUDE [Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- [Visual Studio Code (VS Code)](https://code.visualstudio.com/) またはお好みのコード エディター

- Outlook on Windows (Microsoft 365 アカウントに接続)

### <a name="create-the-add-in-project"></a>アドイン プロジェクトの作成

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Choose a project type: (プロジェクトの種類を選択)** - `Outlook Add-in with Teams Manifest (Developer preview)`

    - **What would you want to name your add-in?: (アドインの名前を何にしますか)** - `Add-in with Teams Manifest`

     ![JSON マニフェスト オプションが選択されたコマンド ライン インターフェイスでの Yeoman ジェネレーターのプロンプトと回答を示すスクリーンショット。](../images/yo-office-outlook-json-manifest.png)

    > [!NOTE]
    > このプレビューでは、アドイン名は 30 文字を超えることはできません。 
    
    ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. Web アプリケーション プロジェクトのルート フォルダーに移動します。

    ```command&nbsp;line
    cd "Add-in with Teams Manifest"
    ```

### <a name="explore-the-project"></a>プロジェクトを探究する

Yeomanジェネレーターで作成したアドインプロジェクトには、原型となる作業ペインアドインのサンプルコードが含まれています。

- プロジェクトのルート ディレクトリにある **./manifest/manifest.json** ファイルで、アドインの機能と設定を定義します。
- **./src/taskpane/taskpane.html** ファイルには、作業ペイン用のHTMLマークアップが含まれています。
- **./src/taskpane/taskpane.css** ファイルには、作業ペインのコンテンツに適用されるCSSが含まれています。
- **./src/taskpane/taskpane.ts** ファイルには、Office JavaScript ライブラリを呼び出して、作業ウィンドウと Outlook の間の対話を容易にするコードが含まれています。
- **./src/command/command.html** ファイルは、ビルド時に WebPack によって編集され、command.ts ファイルからトランスパイルされた JavaScript ファイルを読み込む HTML `<script>` タグが挿入されます。
- **./src/command/command.ts** ファイルには、最初はほとんどコードが含まれていません。 この記事の後半で、Office JavaScript ライブラリを呼び出し、カスタム リボン ボタンが選択されたときに実行されるコードを追加します。

### <a name="update-the-code"></a>コードを更新する

1. VS Codeまたは任意のコード エディターでプロジェクトを開きます。

    > [!TIP]
    > Windows では、コマンド ラインからプロジェクトのルート ディレクトリに移動し、`code .` を入力して VS Code でそのフォルダーを開くことができます。 

1. コードエディタで、**./src/taskpane/taskpane.html** ファイルを開き、全体の **\<main\>** 要素（一部の **\<body\>** 要素）を次のマークアップに置き換えます。 この新しいマークアップは、**./src/taskpane/taskpane.ts** のスクリプトがデータを書き込む場所にラベルを追加します。

    ```html
    <main id="app-body" class="ms-welcome__main" style="display: none;">
        <h2 class="ms-font-xl"> Discover what Office Add-ins can do for you today! </h2>
        <p><label id="item-subject"></label></p>
        <div role="button" id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">Run</span>
        </div>
    </main>
    ```

1. コードエディターで、ファイル **./src/taskpane/taskpane.ts** を開き、**実行** 関数内に次のコードを追加してください。 このコードは、Office JavaScript API を使用して、現在のメッセージへの参照を取得し、その **subject** プロパティの値をタスクペインに書き込むものです。

    ```typescript
    // Get a reference to the current message.
    let item = Office.context.mailbox.item;

    // Write a message property value to the task pane.
    document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
    ```

### <a name="try-it-out"></a>試してみる

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

1. プロジェクトのルート ディレクトリから次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動し、アドインが [サイドロード](../outlook/sideload-outlook-add-ins-for-testing.md) されます。 

    ```command&nbsp;line
    npm start
    ```

1. Outlook でクラシック リボンを使用します。 これらの手順の残りの部分は、これを前提としています。  

1. [閲覧ウィンドウ](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0)でメッセージを表示するか、独自のウィンドウでメッセージを開きます。 **Contoso アドイン** という名前の新しいコントロール グループが Outlook の **[ホーム]** タブ (または、新しいウィンドウでメッセージを開いた場合は **[メッセージ]** タブ) に表示されます。 グループには、**[作業ウィンドウの表示]** という名前のボタンと **[アクションの実行]** という名前のボタンがあります。

    > [!NOTE]
    > 新しいグループが存在しない場合、アドインは自動的にサイドローディングされませんでした。 「[手動でサイドロードする - Windows または Mac の Outlook 2016 以降](../outlook/sideload-outlook-add-ins-for-testing.md#outlook-2016-or-later-on-windows-or-mac)」の指示に従い、Outlook のアドインを手動でサイドロードします。 マニフェスト ファイルをアップロードするように求められたら、`C:\Users\{your_user_name}\AppData\Local\Temp\manifest.xml` ファイルを使用します。 プレビュー期間中に JSON 形式のマニフェストが XML マニフェストに変換され、サイドロードされるため、ファイルには `.xml` 拡張子が付いています。

1. **[アクションの実行]** ボタンを選択します。 [コマンドを実行して](../develop/create-addin-commands.md?branch=outlook-json-manifest#step-5-add-the-functionfile-element)、メッセージ ヘッダーの下部、メッセージ本文のすぐ上に小さな情報通知を生成します。

1. **WebView Stop On Load** ダイアログ ボックスでプロンプトが表示されたら、**[OK]** を選択します。

    [!INCLUDE [Cancelling the WebView Stop On Load dialog box](../includes/webview-stop-on-load-cancel-dialog.md)]

1. アドインの作業ウィンドウを開くには、**[作業ウィンドウの表示]** を選択します。

    > [!NOTE]
    > 作業ウィンドウで、「このアドインを localhost から開くことはできません」 というエラーが表示される場合は、[「トラブルシューティングの記事」](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost) に記載されている手順に従ってください。

1. **WebView Stop On Load** ダイアログ ボックスでプロンプトが表示されたら、**OK** を選択します。

    [!INCLUDE [Cancelling the WebView Stop On Load dialog box](../includes/webview-stop-on-load-cancel-dialog.md)]

1. 作業ウィンドウの一番下までスクロールし、**[実行]** リンクを選択して、メッセージの件名を作業ウィンドウにコピーします。

1. 次のコマンドでデバッグ セッションを終了します。

    ```command&nbsp;line
    npm stop
    ```

    > [!IMPORTANT]
    > Web サーバー ウィンドウを閉じても、Web サーバーが確実にシャットダウンされるわけではありません。 適切にシャットダウンされていない場合、プロジェクトを変更して再実行するときに問題が発生します。

1. Outlook のすべてのインスタンスを閉じます。

## <a name="add-a-custom-button-to-the-ribbon"></a>リボンにカスタム ボタンを追加する

メッセージ本文にテキストを挿入するカスタム ボタンをリボンに追加します。

1. VS Code またはお好みのコード エディターでプロジェクトを開きます。

    > [!TIP]
    > Windows では、コマンド ラインからプロジェクトのルート ディレクトリに移動し、`code .` を入力して VS Code でそのフォルダーを開くことができます。 

1. コード エディターで、ファイル **./src/command/command.ts** を開き、ファイルの最後に次のコードを追加します。 この関数は、メッセージ本文のカーソル ポイントに `Hello World` を挿入します。

    ```typescript
    function insertHelloWorld(event: Office.AddinCommands.Event) {
        Office.context.mailbox.item.body.setSelectedDataAsync("Hello World", {coercionType: Office.CoercionType.Text});

        // Be sure to indicate when the add-in command function is complete
        event.completed();
    }

    // Register the function with Office
    Office.actions.associate("insertHelloWorld", insertHelloWorld);
    ```

1. ファイル **./manifest/manifest.json** を開きます。

    > [!NOTE]
    > ネストされた JSON プロパティを参照する場合、この記事ではドット表記を使用します。 配列内のアイテムが参照される場合、アイテムの括弧で囲まれたゼロベースの番号が使用されます。 

1. メッセージに書き込むには、アドインのアクセス許可を上げる必要があります。 プロパティ `authorization.permissions.resourceSpecific[0].name` までスクロールし、値を `MailboxItem.ReadWrite.User` に変更します。

1. アドイン コマンドが作業ウィンドウを開く代わりにコードを実行する場合は、作業ウィンドウ コードが実行される埋め込み Web ビューとは別のランタイムでコードを実行する必要があります。 したがって、マニフェストは追加のランタイムを指定する必要があります。 プロパティ `extension.runtimes` までスクロールし、次のオブジェクトを `runtimes` 配列に追加します。 既に配列にあるオブジェクトの後には、必ずコンマを入れてください。 このマークアップについて、次の情報にご注意ください。

    - `actions[0].id` プロパティの値は、**commands.ts** ファイルに追加した関数の名前 (この場合は `insertHelloWorld`) と完全に同じである必要があります。 後の手順で、この ID でアイテムを参照します。

    ```json
    {
        "id": "ComposeCommandsRuntime",
        "type": "general",
        "code": {
            "page": "https://localhost:3000/commands.html",
            "script": "https://localhost:3000/commands.js"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "insertHelloWorld",
                "type": "executeFunction",
                "displayName": "insertHelloWorld"
            }
        ]
    }
    ```

1. **[作業ウィンドウの表示]** ボタンは、ユーザーがメールを読んでいるときに表示されますが、テキストを追加するためのボタンは、ユーザーが新しいメールを作成しているとき (または返信しているとき) にのみ表示されます。 したがって、マニフェストは新しいリボン オブジェクトを指定する必要があります。 プロパティ `extension.ribbons` までスクロールし、次のオブジェクトを `ribbons` 配列に追加します。 既に配列にあるオブジェクトの後には、必ずコンマを入れてください。 このマークアップについて、次の点に注意してください。

    - `contexts` 配列の唯一の値は "mailCompose" であるため、ボタンは作成 (または返信) ウィンドウに表示されますが、**[作業ウィンドウの表示]** および **[アクションの実行]** ボタンが表示されるメッセージ読み取りウィンドウには表示されません。 この値を、値が `["mailRead"]` である既存のリボン オブジェクトの `contexts` 配列と比較します。
    - `tabs[0].groups[0].controls[0].actionId` の値は、前の手順で作成したランタイム オブジェクトの `actions[0].id` プロパティの値と完全に同じである必要があります。

    ```json
    {
        "contexts": ["mailCompose"],
        "tabs": [
            {
                "builtInTabId": "TabDefault",
                "groups": [
                    {
                        "id": "msgWriteGroup",
                        "label": "Contoso Add-in",
                        "icons": [
                            { "size": 16, "file": "https://localhost:3000/assets/icon-16.png" },
                            { "size": 32, "file": "https://localhost:3000/assets/icon-32.png" },
                            { "size": 80, "file": "https://localhost:3000/assets/icon-80.png" }
                        ],
                        "controls": [
                            {
                                "id": "HelloWorldButton",
                                "type": "button",
                                "label": "Insert text",
                                "icons": [
                                    { "size": 16, "file": "https://localhost:3000/assets/icon-16.png" },
                                    { "size": 32, "file": "https://localhost:3000/assets/icon-32.png" },
                                    { "size": 80, "file": "https://localhost:3000/assets/icon-80.png" }
                                ],
                                "supertip": {
                                    "title": "Insert text",
                                    "description": "Inserts some text."
                                },
                                "actionId": "insertHelloWorld"
                            }                  
                        ]
                    }
                ]
            }
        ]
    }
    ```

### <a name="try-out-the-updated-add-in"></a>更新されたアドインを試してください

1. プロジェクトのルート ディレクトリから次のコマンドを実行します。

    ```command&nbsp;line
    npm start
    ```

1. Outlook で、新しいメッセージ ウィンドウを開きます (または既存のメッセージに返信します)。 **Contoso アドイン** という名前の新しいコントロール グループが Outlook の **[メッセージ]** タブに表示されます。グループには、**[テキストの挿入]** という名前のボタンがあります。

    > [!NOTE]
    > 新しいグループが存在しない場合、アドインは自動的にサイドローディングされませんでした。 「[手動でサイドロードする - Windows または Mac の Outlook 2016 以降](../outlook/sideload-outlook-add-ins-for-testing.md#outlook-2016-or-later-on-windows-or-mac)」の指示に従い、Outlook のアドインを手動でサイドロードします。 マニフェスト ファイルをアップロードするように求められたら、`C:\Users\{your_user_name}\AppData\Local\Temp\manifest.xml` ファイルを使用します。 プレビュー期間中に JSON 形式のマニフェストが XML マニフェストに変換され、サイドロードされるため、ファイルには `.xml` 拡張子が付いています。

1. メッセージ本文の任意の場所にカーソルを置き、**[テキストの挿入]** ボタンを選択します。

1. **WebView Stop On Load** ダイアログ ボックスでプロンプトが表示されたら、**[OK]** を選択します。

    [!INCLUDE [Cancelling the WebView Stop On Load dialog box](../includes/webview-stop-on-load-cancel-dialog.md)]

    "Hello World" という語句がカーソルに挿入されます。

1. 次のコマンドでデバッグ セッションを終了します。

    ```command&nbsp;line
    npm stop
    ```

## <a name="see-also"></a>関連項目

- [Office アドインの Teams マニフェスト (プレビュー)](../develop/json-manifest-overview.md)
- [Visual Studio コードを使用して発行する](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)