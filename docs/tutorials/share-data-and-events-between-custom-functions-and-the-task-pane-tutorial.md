---
title: 'チュートリアル: Excel カスタム関数と作業ウィンドウの間でデータとイベントを共有する'
description: Excel でカスタム関数と作業ウィンドウの間でデータとイベントを共有する方法について説明します。
ms.date: 08/13/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 0def8178a06231a866bbb87573f936314ac064f1
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2020
ms.locfileid: "49131781"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane"></a>チュートリアル: Excel カスタム関数と作業ウィンドウの間でデータとイベントを共有する

共有ランタイムを使用するように Excel アドインを構成できます。 これにより、グローバル データを共有したり、作業ウィンドウとユーザー設定の関数の間でイベントを送信したりできます。

ほとんどのカスタム関数のシナリオでは、作業ウィンドウのない（非表示の）カスタム関数を使用する特別な理由がない限り、共有ランタイムの使用をお勧めします。

このチュートリアルでは、Yo Office ジェネレーターを使用してアドイン プロジェクトを作成する方法に慣れていることを前提としています。 まだ使い慣れていない場合は、[Excel カスタム関数のチュートリアル](excel-tutorial-create-custom-functions.md)を完了することを検討してください。

## <a name="create-the-add-in-project"></a>アドイン プロジェクトの作成

Yeoman ジェネレーターを使用して、Excel アドイン プロジェクトを作成します。 次のコマンドを実行し、プロンプトに次の回答を入力します。

```command line
yo office
```

- プロジェクトの種類を選択する:  **Excel カスタム関数アドイン プロジェクト**
- スクリプトの種類を選択する:  **JavaScript**
- アドインの名前を何にしますか?  **個人用 Office アドイン**

![コマンドライン インターフェイスでの Yeoman ジェネレーターのプロンプトと回答を示すスクリーンショット](../images/yo-office-excel-project.png)

ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。

## <a name="configure-the-manifest"></a>マニフェストを構成する

1. Visual Studio Code を開始して [**個人用 Office アドイン**] プロジェクトを開きます。
2. 
            **manifest.xml** ファイルを開きます。
3. `<VersionOverrides>` セクションを探し、次の `<Runtimes>` セクションを追加します。 作業ウィンドウを閉じてもカスタム関数が引き続き機能するように、有効期間は **長く** する必要があります。

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
         <Runtimes>
           <Runtime resid="ContosoAddin.Url" lifetime="long" />
         </Runtimes>
       <AllFormFactors>
   ```

> [!NOTE]
> アドインにマニフェストの `Runtimes` 要素が含まれている場合、Windows または Microsoft 365 のバージョンに関係なく、Internet Explorer 11 が使用されます。 詳細については、「[ランタイム](../reference/manifest/runtimes.md)」を参照してください。

4. `<Page>` 要素で、ソースの場所を **Functions.Page.Url** から **ContosoAddin.Url** に変更します。

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. `<DesktopFormFactor>` セクションで、**ContosoAddin.Url** を使用するように、**Command.Url** から **FunctionFile** を変更します。

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. `<Action>` セクションで、ソースの場所を **Taskpane.Url** から **ContosoAddin.Url** に変更します。

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. **taskpane.html** を指す **ContosoAddin.Url** の新しい **Url ID** を追加します。

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. 変更を保存してプロジェクトを再ビルドします。

   ```command line
   npm run build
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
2. `</head>` 要素の前に、次のスクリプト要素を追加します。

   ```html
   <script src="functions.js"></script>
   ```

3. 終了 `</main>` 要素の後に、次の HTML を追加します。 HTML は、グローバル データの取得または保存に使用される 2 つのテキスト ボックスとボタンを作成します。

   ```html
   <ol>
     <li>
       Enter a value to send to the custom function and select
       <strong>Store</strong>.
     </li>
     <li>
       Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve
       it.
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

4. `<body>` 要素の前に、次のスクリプトを追加します。 このコードは、ユーザーがグローバル データを保存または取得するときにボタンのクリック イベントを処理します。

   ```js
   <script>
   function storeSharedValue() {
   let sharedValue = document.getElementById('storeBox').value;
   window.sharedState = sharedValue;
   }

   function getSharedValue() {
   document.getElementById('getBox').value = window.sharedState;
   }</script>
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
