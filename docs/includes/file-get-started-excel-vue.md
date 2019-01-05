# <a name="build-an-excel-add-in-using-vue"></a>Vue を使用して Excel アドインをビルドする

この記事では、Vue と Excel の JavaScript API を使用して Excel アドインを構築する手順について説明します。

## <a name="prerequisites"></a>前提条件

- [Node.js](https://nodejs.org)

- [Vue CLI](https://github.com/vuejs/vue-cli) をグローバルにインストールします。

    ```bash
    npm install -g vue-cli
    ```

- [Yeoman](https://github.com/yeoman/yo) の最新バージョンと [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)をグローバルにインストールします。

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-vue-app"></a>新しい Vue アプリの生成

Vue CLI を使用して新しい Vue アプリを生成します。 端末から次のコマンドを実行し、次に示すとおり、プロンプトに応答します。

```bash
vue init webpack my-add-in
```

前のコマンドで生成されたプロンプトに応答すると、次の 3 つのプロンプトの既定の応答を上書きします。 その他のプロンプトについては、すべて既定の応答を受け入れることができます。

- **Install vue-router? (Vue ルーターをインストールしますか)** `No`
- **Set up unit tests: (単体テストのセットアップ:)** `No`
- **Setup e2e tests with Nightwatch? (Nightwatch とともに e2e テストをセットアップしますか)** `No`

![Vue CLI のプロンプト](../images/vue-cli-prompts.png)

## <a name="generate-the-manifest-file"></a>マニフェスト ファイルを生成する

各アドインには、設定と機能を定義するマニフェスト ファイルが必要です。

1. アプリ フォルダーに移動します。

    ```bash
    cd my-add-in
    ```

2. Yeoman ジェネレーター使用して、アドインのマニフェスト ファイルを生成します。 次のコマンドを実行し、以下に示すプロンプトに応答します。

    ```bash
    yo office 
    ```

    - **Choose a project type: (プロジェクトの種類を選択)** `Office Add-in containing the manifest only`
    - **What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`
    - **Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Excel`

    ![Yeoman ジェネレーター](../images/yo-office.png)
    
    ウィザードを完了すると、ジェネレーターによってマニフェスト ファイルが作成されます。

## <a name="secure-the-app"></a>アプリをセキュリティ保護する

[!include[HTTPS guidance](../includes/https-guidance.md)]

アプリに対して HTTPS を有効にするには、Vue プロジェクトのルート フォルダーにあるファイル **package.json** を開き、`--https` フラグを追加するよう `dev` スクリプトを変更して、ファイルを保存します。

```json
"dev": "webpack-dev-server --https --inline --progress --config build/webpack.dev.conf.js"
```

## <a name="update-the-app"></a>アプリを更新する

1. Yue Office によって Vue プロジェクトのルートに作成されたフォルダー **My Office Add-in** をコード エディターで開きます。 このフォルダーには、アドインの設定が定義されたマニフェスト ファイル **manifest.xml** が表示されます。

2. マニフェスト ファイルを開き、`https://localhost:3000` のすべてを `https://localhost:8080` に置き換えて、ファイルを保存します。

3. (Vue プロジェクトのルートにある) ファイル **index.html** を開き、`</head>` タグの直前に次の `<script>` タグを追加してファイルを保存します。

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

3. **src/main.js** を開き、次のコード ブロックを*削除*します。

    ```js
    new Vue({
        el: '#app',
        components: {App},
        template: '<App/>'
    })
    ```
    
    その後、同じ場所に次のコードを追加し、ファイルを保存します。 
                                                         
    ```js
    const Office = window.Office
    Office.initialize = () => {
      new Vue({
        el: '#app',
        components: {App},
        template: '<App/>'
      })
    }
    ```

4. **src/App.vue** を開き、ファイルの内容を次のコードに置き換え、ファイルの末尾 (つまり、`</style>` タグの直後) に改行を追加して、ファイルを保存します。 

    ```html
    <template>
    <div id="app">
        <div id="content">
        <div id="content-header">
            <div class="padding">
            <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br/>
            <h3>Try it out</h3>
            <button @click="onSetColor">Set color</button>
            </div>
        </div>
        </div>
    </div>
    </template>

    <script>
    export default {
      name: 'App',
      methods: {
        onSetColor () {
          window.Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange()
            range.format.fill.color = 'green'
            await context.sync()
          })
        }
      }
    }
    </script>

    <style>
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px;
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto;
    }

    .padding {
        padding: 15px;
    }
    </style>
    ```

## <a name="start-the-dev-server"></a>開発用サーバーの起動

1. ターミナルから、次のコマンドを実行してデベロッパー サーバーを起動します。

    ```bash
    npm start
    ```

2. Web ブラウザーで `https://localhost:8080` に移動します。 ブラウザーにサイトの証明書が信頼されていないことが示された場合は、その証明書を信頼するようコンピューターを構成する必要があります。 

3. ブラウザーに証明書エラーなしでアドイン ページが読み込まれたら、アドインをテストする準備ができています。 

## <a name="try-it-out"></a>お試しください

1. アドインを実行して、Excel 内のアドインをサイドロードするのに使用するプラットフォームの手順に従います。

    - Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online:[Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)
    - iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

2. Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。

    ![Excel アドイン ボタン](../images/excel-quickstart-addin-2a.png)

3. ワークシート内で任意のセルの範囲を選択します。

4. 作業ウィンドウで、**[色の設定]** ボタンをクリックして、選択範囲の色を緑に設定します。

    ![Excel アドイン](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>次の手順

これで完了です。Vue を使用して Excel アドインが正常に作成されました。次に、Excel アドインの機能の詳細について説明します。Excel アドインのチュートリアルに従って、より複雑なアドインをビルドします。

> [!div class="nextstepaction"]
> [Excel アドインのチュートリアル](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>関連項目

* [Excel アドインのチュートリアル](../tutorials/excel-tutorial-create-table.md)
* [Excel JavaScript API を使用した基本的なプログラミングの概念](../excel/excel-add-ins-core-concepts.md)
* [Excel アドインのコード サンプル](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Excel JavaScript API リファレンス](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
