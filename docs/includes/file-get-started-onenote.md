# <a name="build-your-first-onenote-add-in"></a>最初の OneNote アドインをビルドする

この記事では、jQuery と Office JavaScript API を使用して OneNote アドインを作成する手順について説明します。

## <a name="prerequisites"></a>前提条件

- [Node.js](https://nodejs.org)

- [Yeoman](https://github.com/yeoman/yo) の最新バージョンと [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)をグローバルにインストールします。

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a>アドイン プロジェクトの作成

1. ローカル ドライブにフォルダーを作成し、`my-onenote-addin`という名前を付けます。ここにアドインのファイルを作成します。

    ```bash
    mkdir my-onenote-addin
    ```

2. 新しいフォルダーに移動します。

    ```bash
    cd my-onenote-addin
    ```

3. Yeoman を使用して OneNote アドイン プロジェクトを作成するジェネレーターです。次のコマンドを実行し、プロンプトに次のように応答し、します。

    ```bash
    yo office
    ```

    - **Choose a project type:​ (プロジェクト タイプを選択してください)** `Office Add-in project using Jquery framework`
    - **Choose a script type: (スクリプト タイプを選択してください)** `Javascript`
    - **What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`
    - **Which Office client application would you like to support? (サポートする Office クライアント アプリケーションを選んでください):** `Onenote`

    ![Yeoman ジェネレーターのプロンプトと応答のスクリーンショット](../images/yo-office-onenote-jquery.png)
    
    ウィザードが完了すると、ジェネレーターはプロジェクトを作成し、サポートする Node コンポーネントをインストールします。
    
4. Web アプリケーション プロジェクトのルート フォルダーに移動します。

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a>コードを更新する

1. コード エディターで、プロジェクトのルートに**index.html** を開きます。このファイルは、アドインの作業ウィンドウでレンダリングされる HTML を指定します。

2. 要素内の `<body>` 要素を次のマークアップに置き換えて、ファイルを保存します。 

    ```html
    <body class="ms-font-m ms-welcome">
        <header class="ms-welcome__header ms-bgColor-themeDark ms-u-fadeIn500">
            <h2 class="ms-fontSize-xxl ms-fontWeight-regular ms-fontColor-white">OneNote Add-in</h1>
        </header>
        <main id="app-body" class="ms-welcome__main">
            <br />
            <p class="ms-font-m">Enter HTML content here:</p>
            <div class="ms-TextField ms-TextField--placeholder">
                <textarea id="textBox" rows="8" cols="30"></textarea>
            </div>
            <button id="addOutline" class="ms-Button ms-Button--primary">
                <span class="ms-Button-label">Add outline</span>
            </button>
        </main>
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>
    ```

3. ファイル **src\index.js** を開いてアドインのスクリプトを指定します。内容全体を以下のコードで置き換え、ファイルを保存します。

    ```js
    import * as OfficeHelpers from "@microsoft/office-js-helpers";

    Office.initialize = (reason) => {
        $(document).ready(() => {
            $('#addOutline').click(addOutlineToPage);
        });
    };
    
    async function addOutlineToPage() {
        try {
            await OneNote.run(async context => {
                var html = "<p>" + $("#textBox").val() + "</p>";

                // Get the current page.
                var page = context.application.getActivePage();

                // Queue a command to load the page with the title property.
                page.load("title");

                // Add text to the page by using the specified HTML.
                var outline = page.addOutline(40, 90, html);

                // Run the queued commands, and return a promise to indicate task completion.
                return context.sync()
                    .then(function() {
                        console.log("Added outline to page " + page.title);
                    })
                    .catch(function(error) {
                        app.showNotification("Error: " + error);
                        console.log("Error: " + error);
                        if (error instanceof OfficeExtension.Error) {
                            console.log("Debug info: " + JSON.stringify(error.debugInfo));
                        }
                    });
                });
        } catch (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        }
    }
    ```

4. **app.css** ファイルを開いて、アドインのカスタム スタイルを指定します。 すべての内容を次の内容に置き換えて、ファイルを保存します。

    ```css
    html, body {
        width: 100%;
        height: 100%;
        margin: 0;
        padding: 0;
    }

    ul, p, h1, h2, h3, h4, h5, h6 {
        margin: 0;
        padding: 0;
    }

    .ms-welcome {
        position: relative;
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        min-height: 500px;
        min-width: 320px;
        overflow: auto;
        overflow-x: hidden;
    }

    .ms-welcome__header {
        min-height: 30px;
        padding: 0px;
        padding-bottom: 5px;
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        -webkit-align-items: center;
        align-items: center;
        -webkit-justify-content: flex-end;
        justify-content: flex-end;
    }

    .ms-welcome__header > h1 {
        margin-top: 5px;
        text-align: center;
    }

    .ms-welcome__main {
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        -webkit-align-items: center;
        align-items: left;
        -webkit-flex: 1 0 0;
        flex: 1 0 0;
        padding: 30px 20px;
    }

    .ms-welcome__main > h2 {
        width: 100%;
        text-align: left;
    }

    @media (min-width: 0) and (max-width: 350px) {
        .ms-welcome__features {
            width: 100%;
        }
    }
    ```

## <a name="update-the-manifest"></a>マニフェストを更新する

1. [ **one-note-add-in-manifest.xml** ]ファイルを開いて、アドインの設定と機能を定義します。

2.  `ProviderName` 要素にはプレースホルダーの値があります。これを自分の名前で置き換えます。

3. `Description` 要素の `DefaultValue` 属性にはプレースホルダーがあります。これを **Excel の作業ウィンドウ アドイン** で置き換えます。

4. ファイルを保存します。

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for OneNote"/>
    ...
    ```

## <a name="start-the-dev-server"></a>開発用サーバーを起動する

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a>試してみる

1. [OneNote Online](https://www.onenote.com/notebooks) でノートブックを開きます。

2. **[挿入] > [Office アドイン]** の順に選択し、[Office アドイン] ダイアログを開きます。

    - コンシューマー アカウントでサインインしている場合は、**[マイ アドイン]** タブを選択し、**[マイ アドインのアップロード]** を選択します。

    - 職場または学校アカウントでサインインしている場合は、**[自分の所属組織]** タブを選択し、**[マイ アドインのアップロード]** を選択します。 

    次の図は、コンシューマー ノートブックの [ **マイ アドイン** ]タブを示しています。

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. アップロード アドインダイアログで、プロジェクト フォルダー内の [ **manifest.xml** ] を参照し、[ **アップロード** ]を選択します。 

4. [ **ホーム** ] タブから、リボンの [ **作業ウィンドウの表示** ] ボタンを選択します。OneNote のページの横にある iFrame にアドインの作業ウィンドウが開きます。

5. テキスト領域で、次の HTML コンテンツを入力し、 **アウトラインを追加**します。  

    ```html
    <ol>
    <li>Item #1</li>
    <li>Item #2</li>
    <li>Item #3</li>
    <li>Item #4</li>
    </ol>
    ```

    指定したアウトラインは、ページに追加されます。

    ![このチュートリアルでビルドした OneNote アドイン](../images/onenote-first-add-in-3.png)

## <a name="troubleshooting-and-tips"></a>トラブルシューティングとヒント

- ブラウザーの開発者ツールを使ってアドインをデバッグできます。Gulp Web サーバーを使っており、Internet Explorer や Chrome でデバッグしている場合は、ローカルで変更を保存して、アドインの iFrame を更新するだけです。

- OneNote オブジェクトを調べる場合、現在使用可能なプロパティに実際の値が表示されます。読み込む必要のあるプロパティには、*undefined* と表示されます。`_proto_` ノードを展開し、オブジェクトで定義されているものの、まだ読み込まれていないプロパティを確認します。

   ![デバッガーでアンロードされた OneNote オブジェクト](../images/onenote-debug.png)

- アドインで任意の HTTP リソースを使っている場合は、ブラウザーで混在したコンテンツを有効にする必要があります。運用アドインでは、セキュリティで保護された HTTPS リソースのみを使う必要があります。

- 作業ウィンドウ アドインは、任意の場所から開くことができますが、コンテンツアドインは、通常のページ コンテンツ (タイトル、イメージ、iframe などは含まない) の内部にのみ挿入できます。 

## <a name="next-steps"></a>次の手順

これで完了です。OneNote アドインが正常に作成されました。次に、OneNote アドイン構築の中心概念の詳細について説明します。

> [!div class="nextstepaction"]
> [OneNote JavaScript API のプログラミングの概要](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a>関連項目

- [OneNote JavaScript API のプログラミングの概要](../onenote/onenote-add-ins-programming-overview.md)
- [OneNote JavaScript API リファレンス](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference?view=office-js)
- [Rubric Grader のサンプル](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office アドイン プラットフォームの概要](../overview/office-add-ins.md)
