# <a name="build-your-first-word-add-in"></a>最初の Word アドインをビルドする

_適用対象:Word 2016、Word for iPad、Word for Mac_

Word アドインは、Word の内部で動作し、Word JavaScript API を使用して文書のコンテンツを操作できます。この API は、Office アプリケーションを拡張するための Office アドイン プログラミング モデルに組み込まれています。 このアドイン プログラミング モデルでは、任意のプラットフォームと言語を使用して、Word に対する拡張機能をホストする Web アプリケーションを作成できます。その設定と機能は、アドインの[マニフェスト](../../docs/overview/add-in-manifests.md)使用して定義できます。

この記事では、jQuery と Word JavaScript API を使用して Word アドインを構築する手順について説明します。 

> **注**: Word 2013 のアドインを開発するには、共有の [Office Javascript API]( https://dev.office.com/docs/add-ins/word/word-add-ins-programming-overview#javascript-apis-for-word) を使用する必要があります。 使用可能なプラットフォームと各種 API の詳細については、「[Office アドインを使用できるホストおよびプラットフォーム](https://dev.office.com/add-in-availability)」を参照してください。 

## <a name="create-the-web-app"></a>Web アプリを作成する 

1. ローカル ドライブにフォルダーを作成して、**BoilerplateAddin** という名前を付けます。 アプリのファイルは、ここに作成します。

2. アプリ フォルダーで、**home.html** という名前のファイルを作成して、アドインの作業ウィンドウにレンダリングされる HTML を指定します。 このアドインには、3 つのボタンが表示されます。いずれかのボタンを選択すると、文書に定型句が追加されます。 次のコードを **home.html** に追加して、ファイルを保存します。

    ```html
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Boilerplate text app</title>
        <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
        <script src="home.js" type="text/javascript"></script>
        </head>
        <body>
            <div>
                <h1>Welcome</h1>
            </div>
            <div>
                <p>This sample shows how to add boilerplate text to a document by using the Word JavaScript API.</p>
                <br />
                <h3>Try it out</h3>
                <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                <button id="checkhov">Add quote from Anton Chekhov</button>
                <button id="proverb">Add Chinese proverb</button>
            </div>
            <h3><div id="supportedVersion"/></h3>
        </body>
    </html>
    ```

3. アプリ フォルダーで、**home.js** という名前のファイルを作成して、アドインの jQuery スクリプトを指定します。 このスクリプトには、初期化のコードと、Word 文書に変更を加える (ボタンが選択されたときに、ドキュメントにテキストを挿入する) コードが含まれています。 次のコードを **home.js** に追加して、ファイルを保存します。

    ```javascript
    (function () {
        "use strict";

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or greater.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
                }
            });
        };

        function insertEmersonQuoteAtSelection() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();

                // Queue a command to replace the selected text.
                range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Ralph Waldo Emerson.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChekhovQuoteAtTheBeginning() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the start of the document body.
                body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Anton Chekhov.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChineseProverbAtTheEnd() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the end of the document body.
                body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from a Chinese proverb.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

## <a name="create-the-manifest-file"></a>マニフェスト ファイルを作成する

1. アプリ フォルダーで、**BoilerplateManifest.xml** という名前のファイルを作成して、アドインの設定と機能を定義します。 このファイルに次のコードを追加します。 

    ```xml
    <?xml version="1.0" encoding="UTF-8"?>
        <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xsi:type="TaskPaneApp">
            <Id>2b88100c-656e-4bab-9f1e-f6731d86e464</Id>
            <Version>1.0.0.0</Version>
            <ProviderName>Microsoft</ProviderName>
            <DefaultLocale>en-US</DefaultLocale>
            <DisplayName DefaultValue="Boilerplate content" />
            <Description DefaultValue="Insert boilerplate content into a Word document." />
            <Hosts>
                <Host Name="Document"/>
            </Hosts>
            <DefaultSettings>
                <SourceLocation DefaultValue="\\MyShare\boilerplate\home.html" />
            </DefaultSettings>
            <Permissions>ReadWriteDocument</Permissions>
        </OfficeApp>
    ```

2. 任意のオンライン ジェネレーターを使用して、GUID を生成します。 次に、前述の手順で示した **Id** 要素の値をその GUID に置き換えます。

3. マニフェスト ファイルを保存します。

## <a name="deploy-the-web-app-and-update-the-manifest"></a>Web アプリを展開してマニフェストを更新する

1. 任意の Web サーバーに Web アプリ (アプリ フォルダーのコンテンツ) を展開します。

2. ローカルのアプリ フォルダーで、マニフェスト ファイル (**BoilerplateManifest.xml**) を開きます。 **SourceLocation** 要素内の属性値を編集し、Web サーバー上の **home.html** ファイルの場所を指定して、ファイルを保存します。

## <a name="try-it-out"></a>お試しください

1. アドインの実行に使用するプラットフォームの指示に従って、Word 内でアドインをサイドロードします。

    - Windows: [Windows でのテスト用に Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Word Online: [Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad および Mac:[iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

2. 右側の作業ウィンドウで、いずれかのボタンを選択して文書に定型句を追加します。

![定型句アドインが読み込まれている Word アプリケーションの画像。](../../images/boilerplateAddin.png)

## <a name="next-steps"></a>次のステップ

これで完了です。jQuery を使用して Word アドインが正常に作成されました。 次に、Word アドイン構築の[中心概念](word-add-ins-programming-overview.md)の詳細について説明します。

## <a name="additional-resources"></a>その他の技術資料

* [Word アドインの概要](word-add-ins-programming-overview.md)
* [Script Lab でスニペットを探す](https://store.office.com/en-001/app.aspx?assetid=WA104380862&ui=en-US&rs=en-001&ad=US&appredirect=false)
* [Word アドインのコード サンプル](http://dev.office.com/code-samples#?filters=word,office%20add-ins)
* [Word JavaScript API リファレンス](../../reference/word/word-add-ins-reference-overview.md)