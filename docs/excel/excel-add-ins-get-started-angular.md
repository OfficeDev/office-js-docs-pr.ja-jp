# <a name="build-an-excel-add-in-using-angular"></a>Angular を使用して Excel のアドインを作成する

この記事では、Angular と Excel の JavaScript API を使用して Excel アドインを構築する手順について説明します。

## <a name="prerequisites"></a>必須コンポーネント

次のツールをインストールします (まだインストールしていない場合)。

1. 既に [Angular CLI の必須コンポーネント](https://github.com/angular/angular-cli#prerequisites)がインストールされているかを確認し、足りない場合は必須コンポーネントをインストールします。

2. [Angular CLI](https://github.com/angular/angular-cli) をグローバルにインストールします。 

    ```bash
    npm install -g @angular/cli
    ```

3. [Yeoman](https://github.com/yeoman/yo) と [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)をグローバルにインストールします。

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-angular-app"></a>新しい Angular アプリを作成する

Angular CLI を使用して Angular アプリを生成します。 ターミナルから、次のコマンドを実行します。

```bash
ng new my-addin
```

## <a name="generate-the-manifest-file-and-sideload-the-add-in"></a>マニフェスト ファイルを生成し、アドインをサイドロードする

アドインのマニフェスト ファイルは、その設定と機能を定義します。

1. アプリ フォルダーに移動します。

    ```bash
    cd my-addin
    ```

2. Yeoman ジェネレーター使用して、アドインのマニフェスト ファイルを生成します。 次のコマンドを実行し、以下のスクリーンショットに示すとおり、プロンプトに応答します。

    ```bash
    yo office
    ```
    ![Yeoman ジェネレーター](../images/yo-office.png)
    > **注**:**package.json** を上書きするメッセージが表示された場合は、**No** (上書きしない) と答えます。

3. マニフェスト ファイル (つまり、アプリのルート ディレクトリ内にある、名前が "manifest.xml" で終わるファイル) を開きます。 `https://localhost:3000` の出現箇所すべてを `http://localhost:4200` で置き換えて、ファイルを保存します。

    > **注**:ポート番号を **4200** に変更し、プロトコルを **http** に変更してください。

4. アドインを実行して、Excel 内のアドインをサイドロードするのに使用するプラットフォームの手順に従います。

    - Windows:[Windows でのテスト用に Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online:[Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad および Mac:[iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

## <a name="update-the-app"></a>アプリを更新する

1. **src/index.html** を開き、`</head>` タグの直前に次の `<script>` タグを追加して、ファイルを保存します。

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

2. **src/main.ts** を開き、`platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` を次のコードで置き換えて、ファイルを保存します。 

    ```typescript 
    declare const Office: any;

    Office.initialize = () => {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    };
    ```

3. **src/polyfills.ts** を開き、存在する他のすべての `import` ステートメントの上に次のコード行を追加して、ファイルを保存します。

    ```typescript
    import 'core-js/client/shim';
    ```

4. **src/polyfills.ts** で、次の行のコメントを解除してファイルを保存します。

    ```typescript
    import 'core-js/es6/symbol';
    import 'core-js/es6/object';
    import 'core-js/es6/function';
    import 'core-js/es6/parse-int';
    import 'core-js/es6/parse-float';
    import 'core-js/es6/number';
    import 'core-js/es6/math';
    import 'core-js/es6/string';
    import 'core-js/es6/date';
    import 'core-js/es6/array';
    import 'core-js/es6/regexp';
    import 'core-js/es6/map';
    import 'core-js/es6/weak-map';
    import 'core-js/es6/set';
    ```

5. **src/app/app.component.html** を開き、ファイルのコンテンツを次の HTML で置き換えて、ファイルを保存します。 

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <button (click)="onColorMe()">Color Me</button>
        </div>
    </div>
    ```

6. **src/app/app.component.css** を開き、ファイルのコンテンツを次の CSS コードで置き換えて、ファイルを保存します。

    ```css
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
    ```

7. **src/app/app.component.ts** を開き、ファイルのコンテンツを次のコードで置き換えて、ファイルを保存します。 

    ```typescript
    import { Component } from '@angular/core';

    declare const Excel: any;

    @Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.css']
    })
    export class AppComponent {
    onColorMe() {
        Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.color = 'green';
        await context.sync();
        });
    }
    }
    ```

## <a name="try-it-out"></a>お試しください

1. ターミナルから、次のコマンドを実行してデベロッパー サーバーを起動します。

    ```bash
    npm start
    ```

2. Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。

    ![Excel アドイン ボタン](../images/excel_quickstart_addin_2a.png)

3. 作業ウィンドウで **[色の選択]** ボタンをクリックして、選択範囲の色を緑に設定します。

    ![Excel アドイン](../images/excel_quickstart_addin_2b.png)

## <a name="next-steps"></a>次の手順

これで完了です。Angular を使用して Excel アドインが正常に作成されました。 次に、Excel アドイン構築の[中心概念](excel-add-ins-core-concepts.md)の詳細について説明します。

## <a name="additional-resources"></a>その他のリソース

* [Excel JavaScript API の中心概念](excel-add-ins-core-concepts.md)
* [Excel アドインのコード サンプル](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Excel JavaScript API リファレンス](../../reference/excel/excel-add-ins-reference-overview.md)
