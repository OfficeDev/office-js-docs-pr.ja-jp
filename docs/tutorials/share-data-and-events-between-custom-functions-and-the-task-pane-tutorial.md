---
title: 'チュートリアル: Excel カスタム関数と作業ウィンドウの間でデータとイベントを共有する'
description: Excel でカスタム関数と作業ウィンドウの間でデータとイベントを共有する方法について説明します。
ms.date: 11/29/2021
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 8bc2ea45588c7e10cd4fbd2fc32ff88a6c3233a2
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746473"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane"></a>チュートリアル: Excel カスタム関数と作業ウィンドウの間でデータとイベントを共有する

グローバル データを共有し、共有ランタイムを使用して、Excel アドインの作業ウィンドウとカスタム関数の間でイベントを送信します。 ほとんどのカスタム関数のシナリオでは、作業ウィンドウのない（非表示の）カスタム関数を使用する特別な理由がない限り、共有ランタイムの使用をお勧めします。 このチュートリアルでは、[Office アドイン用の Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)を使用してアドイン プロジェクトを作成する方法に慣れていることを前提としています。 まだ使い慣れていない場合は、[Excel カスタム関数のチュートリアル](excel-tutorial-create-custom-functions.md)を完了することを検討してください。

## <a name="create-the-add-in-project"></a>アドイン プロジェクトの作成

[Office アドイン用の Yeoman ジェネレーター](../develop/yeoman-generator-overview.md) を使用し、Excel アドイン プロジェクトを作成します。

- カスタム関数を使用して Excel アドインを生成するには、次のコマンドを実行します。
    
    ```command&nbsp;line
    yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js true
    ```

ジェネレーターによってプロジェクトが作成され、サポートしているノード コンポーネントがインストールされます。

## <a name="configure-the-manifest"></a>マニフェストを構成する

次の手順に従ってアドイン プロジェクトを構成し、共有ランタイムを使用します。

1. Visual Studio Code を起動し、生成したアドイン プロジェクトを開きます。
1. **manifest.xml** ファイルを開きます。
1. 次の `<Requirements>` セクション XML を置き換えて (または追加して)、[共有ランタイム要件セット](../reference/requirement-sets/shared-runtime-requirement-sets.md) を要求します。

    ```xml
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    ```

    更新後、マニフェスト XML は次の順序で表示されます。

    ```xml
    <Hosts>
      <Host Name="..."/>
    </Hosts>
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    <DefaultSettings>
    ```

1. `<VersionOverrides>` セクションを検索して次の `<Runtimes>` セクションを追加します。 作業ウィンドウを閉じてもアドイン コードを実行できるように、有効期間は **長く** する必要があります。 `resid` 値は **Taskpane.Url** で、**manifest.xml** ファイルの下部付近の `<bt:Urls>` セクションで指定された **taskpane.html** ファイルの場所を参照します。
    
    ```xml
    <Runtimes>
      <Runtime resid="Taskpane.Url" lifetime="long" />
    </Runtimes>
    ```
    
    > [!IMPORTANT]
    > `<Runtimes>` セクションは、次の XML で示される正確な順序で `<Host xsi:type="...">` 要素の後に入力する必要があります。

    ```xml
    <VersionOverrides ...>
      <Hosts>
        <Host xsi:type="...">
          <Runtimes>
            <Runtime resid="Taskpane.Url" lifetime="long" />
          </Runtimes>
        ...
        </Host>
    ```
    
    > [!NOTE]
    > アドインにマニフェストの `Runtimes` 要素 (共有ランタイムに必要) が含まれており、WebView2 (Chromium ベース) で Microsoft Edge の使用条件が満たされている場合、その WebView2 コントロールが使用されます。 使用条件が満たされていない場合は、Windows または Microsoft 365 のバージョンに関係なく、Internet Explorer 11 が使用されます。 詳細については、「[ランタイム](../reference/manifest/runtimes.md)」および「[Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。

1. `<Page>` 要素を検索します。次に、ソースの場所を **Functions.Page.Url** から **Taskpane.Url** に変更します。

   ```xml
   <AllFormFactors>
   ...
   <Page>
     <SourceLocation resid="Taskpane.Url"/>
   </Page>
   ...
   ```

1. `<FunctionFile ...>` タグを見つけて、`resid` を **Commands.Url** から **Taskpane.Url** に変更します。

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. **manifest.xml** ファイルを保存します。

## <a name="configure-the-webpackconfigjs-file"></a>webpack.config.js ファイルを構成する

**webpack.config.js** は、複数のランタイム ローダーをビルドします。 **taskpane.html** ファイルを介して共有 JavaScript ランタイムのみを読み込むように変更する必要があります。

1. **webpack.config.js** ファイルを開きます。
1. `plugins:` セクションに移動します。
1. セクションが存在する場合は、次の `functions.html` プラグインを削除します。
    
    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

1. セクションが存在する場合は、次の `commands.html` プラグインを削除します。

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

1. `functions` または `commands` プラグインを削除した場合は、`chunks` として追加します。 `functions` または `commands` プラグインを削除した場合は、次の JavaScript に更新されたエントリーが表示されます。
    
    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```
    
1. 変更を保存してプロジェクトを再ビルドします。

   ```command&nbsp;line
   npm run build
   ```
    
    > [!NOTE]
    > **functions.html** ファイルと **commands.html** ファイルを削除することもできます。 **taskpane.html** は、先ほど行った webpack の更新を介して、**functions.js** および **commands.js** コードを共有 JavaScript ランタイムに読み込みます。
    
1. 変更を保存してプロジェクトを実行します。 エラー無しで読み込み、実行が行われるようにします。
    
   ```command&nbsp;line
   npm run start
   ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a>カスタム関数と作業ウィンドウのコードの間で状態を共有する

カスタム関数が作業ウィンドウのコードと同じコンテキストで実行されるようになったため、**ストレージ** オブジェクトを使用せずに状態を直接共有できます。 次の手順は、カスタム関数と作業ウィンドウのコードの間でグローバル変数を共有する方法を示します。

### <a name="create-custom-functions-to-get-or-store-shared-state"></a>共有状態を取得または保存するカスタム関数を作成する

1. Visual Studio Code でファイル **src/functions/functions.js** を開きます。
2. 1 行目で、次のコードを一番上に挿入します。 これにより、**sharedState** という名前のグローバル変数が初期化されます。

   ```js
   window.sharedState = "empty";
   ```

3. 次のコードを追加して、値を **sharedState** 変数に保存するカスタム関数を作成します。

   ```js
   /**
    * Saves a string value to shared state with the task pane
    * @customfunction STOREVALUE
    * @param {string} value String to write to shared state with task pane.
    * @return {string} A success value
    */
   function storeValue(sharedValue) {
     window.sharedState = sharedValue;
     return "value stored";
   }
   ```

4. 次のコードを追加して、**sharedState** 変数の現在の値を取得するカスタム関数を作成します。

   ```js
   /**
    * Gets a string value from shared state with the task pane
    * @customfunction GETVALUE
    * @returns {string} String value of the shared state with task pane.
    */
   function getValue() {
     return window.sharedState;
   }
   ```

5. ファイルを保存します。

### <a name="create-task-pane-controls-to-work-with-global-data"></a>グローバル データを操作する作業ウィンドウのコントロールを作成する

1. ファイル **src/taskpane/taskpane.html** を開きます。
2. 次のスクリプト要素を追加してから、`</head>` 要素を閉じます。

   ```html
   <script src="../functions/functions.js"></script>
   ```

3. 終了 `</main>` 要素の後に、次の HTML を追加します。 HTML は、グローバル データの取得または保存に使用される 2 つのテキスト ボックスとボタンを作成します。

   ```html
   <ol>
     <li>
       Enter a value to send to the custom function and select
       <strong>Store</strong>.
     </li>
     <li>
       Enter <strong>=CONTOSO.GETVALUE()</strong> into a cell to retrieve it.
     </li>
     <li>
       To send data to the task pane, in a cell, enter
       <strong>=CONTOSO.STOREVALUE("new value")</strong>
     </li>
     <li>Select <strong>Get</strong> to display the value in the task pane.</li>
   </ol>

   <p>Store new value to shared state</p>
   <div>
     <input type="text" id="storeBox" />
     <button onclick="storeSharedValue()">Store</button>
   </div>

   <p>Get shared state value</p>
   <div>
     <input type="text" id="getBox" />
     <button onclick="getSharedValue()">Get</button>
   </div>
   ```

4. `</body>` 要素を閉じる前に、次のスクリプトを追加します。このコードは、ユーザーがグローバル データを保存または取得するときにボタンのクリック イベントを処理します。

   ```js
   <script>
   function storeSharedValue() {
     let sharedValue = document.getElementById('storeBox').value;
     window.sharedState = sharedValue;
   }

   function getSharedValue() {
     document.getElementById('getBox').value = window.sharedState;
   }
   </script>
   ```

5. ファイルを保存します。
6. プロジェクトをビルドする

   ```command line
   npm run build
   ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a>カスタム関数と作業ウィンドウの間でデータの共有を試す

- 次のコマンドを使用してプロジェクトを開始します。

  ```command line
  npm run start
  ```

Excel が起動したら、作業ウィンドウのボタンを使用して共有データを保存または取得できます。 カスタム関数のセルに `=CONTOSO.GETVALUE()` を入力して、同じ共有データを取得します。 または `=CONTOSO.STOREVALUE("new value")` を使用して、共有データを新しい値に変更します。

> [!NOTE]
> この記事で示すように、プロジェクトを構成すると、カスタム機能と作業ウィンドウのコンテキストが共有されます。 カスタム関数から一部の Office API を呼び出すことができます。 詳細については、「[カスタム関数からの Microsoft Excel API の呼び出しについて](../excel/call-excel-apis-from-custom-function.md)」を参照してください。

## <a name="see-also"></a>関連項目

- [Office アドインを構成して共有 JavaScript ランタイムを使用する](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
