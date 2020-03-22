---
title: Excel アドインのチュートリアル
description: このチュートリアルでは、Excel アドインを構築します。このアドインでは、テーブルの作成、表示、フィルター処理、並べ替えを行うことができ、グラフの作成、テーブルのヘッダーの固定、ワークシートの保護も可能となります。また、ダイアログを開くこともできます。
ms.date: 01/16/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 8d16369908f5885e20c145581c021ba69855e968
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719792"
---
# <a name="tutorial-create-an-excel-task-pane-add-in"></a>チュートリアル: Excel 作業ウィンドウ アドインを作成する

このチュートリアルでは、以下を実行する Excel 作業ウィンドウ アドインを作成します。

> [!div class="checklist"]
> * テーブルの作成
> * テーブルのフィルター処理と並べ替え
> * グラフの作成
> * テーブルのヘッダーの固定
> * ワークシートの保護
> * ダイアログを開く

> [!TIP]
> 既に [[Excel タスク ウィンドウ アドインのビルド](../quickstarts/excel-quickstart-jquery.md)] の クイックスタートを完​​了しており、このチュートリアルの出発点としてそのプロジェクトを使用する場合は、[[テーブルの作成](#create-a-table)] セクションに直接移動します。

## <a name="prerequisites"></a>前提条件

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-your-add-in-project"></a>アドイン プロジェクトの作成

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project`
- **Choose a script type: (スクリプトの種類を選択)** `Javascript`
- **What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`
- **Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Excel`

![Yeoman ジェネレーター](../images/yo-office-excel.png)

ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="create-a-table"></a>表を作成する

チュートリアルのこの手順では、プログラムによってアドインがユーザーの Excel の現在のバージョンをサポートしているかどうかをテストし、ワークシートにテーブルを追加して、そのテーブルのデータ設定と書式設定を実行します。

### <a name="code-the-add-in"></a>アドインのコードを作成する

1. コード エディターでプロジェクトを開きます。

2. ファイル **./src/taskpane/taskpane.html** を開きます。  このファイルには、作業ウィンドウの HTML マークアップが含まれます。

3. `<main>` 要素を見つけて、開始 `<main>` タグの後、終了 `</main>` タグの前に表示されるすべての行を削除します。

4. 開始 `<main>` タグの後に次のマークアップを追加します。

    ```html
    <button class="ms-Button" id="create-table">Create Table</button><br/><br/>
    ```

5. ファイル **./src/taskpane/taskpane.js** を開きます。 このファイルには、作業ウィンドウと Office のホスト アプリケーションの間のやり取りを容易にする Office JavaScript API コードが含まれています。

6. 次の操作を行って、[`run`] ボタンと [`run()`] 関数へのすべての参照を削除します。

    - `document.getElementById("run").onclick = run;` 行を見つけて削除します。

    - `run()` 関数全体を見つけて削除します。

7. `Office.onReady` メソッドの呼び出しで、`if (info.host === Office.HostType.Excel) {` 行を見つけ、その行の直後に次のコードを追加します。 注:

    - このコードの最初の部分では、ユーザーの Excel のバージョンが、このチュートリアルのシリーズで使用する API をすべて含んでいるバージョンの Excel.js をサポートしているかどうかを調べます。 運用アドインでは、未サポートの API を呼び出す UI を非表示または無効化する条件ブロックの本体を使用してください。 これにより、ユーザーは、そのユーザーの Excel のバージョンでサポートされているアドインの部分を使用できるようになります。

    - 2 番目の部分では、[`create-table`] ボタンのイベント ハンドラーを追加します。

    ```js
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = createTable;
    ```

8. 次の関数をファイルの最後に追加します。 注:

    - Excel .js のビジネスロジックが `Excel.run`に渡される関数に追加されます。このロジックは直ちには実行されません。代わりに、保留中のコマンドのキューに追加されます。

    - `context.sync` メソッドは、キューに登録されたすべてのコマンドを実行するために Excel に送信します。

    - これは、どのような場合にも当てはまるベスト プラクティスです。 

    ```js
    function createTable() {
        Excel.run(function (context) {

            // TODO1: Queue table creation logic here.

            // TODO2: Queue commands to populate the table with data.

            // TODO3: Queue commands to format the table.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

9. `createTable()` 関数で、`TODO1` を次のコードに置き換えます。 注:

    - コードは、ワークシートのテーブルコレクションの `add` メソッドを使用してテーブルを作成します。これは空の場合でも常に存在します。これは、Excel .js オブジェクトが作成される標準的な方法です。クラスコンストラクター Api はありません。 `new` 演算子を使用して Excel オブジェクトを作成することはできません。代わりに、親コレクションオブジェクトに追加します。

    - `add` の最初のパラメーターは、テーブルの一番上の行のみの範囲で、テーブルが最終的に使用する範囲全体ではありません。これは、アドインがデータ行にデータを入力するときに、既存の行のセルに値を入力する代わりに、テーブルに新しい行を追加するためです。テーブルを作成するときに、テーブルに含まれる行の数がわからないことが多いため、これはより一般的なパターンです。

    - テーブルの名前は、ワークシート内だけでなくブック全体で一意にする必要があります。

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

10. `createTable()` 関数で、`TODO2` を次のコードに置き換えます。 注:

    - 範囲に含まれるセルの値は、配列の配列で設定します。

    - テーブルに新しい行が作成されるのは、テーブルの行コレクションの `add` メソッドを呼び出すことです。2番目のパラメーターとして渡される親の配列に複数のセル値の配列を含めることにより、1つの `add` 呼び出しに複数の行を追加できます。

    ```js
    expensesTable.getHeaderRowRange().values =
        [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add at the end*/, [
        ["1/1/2017", "The Phone Company", "Communications", "120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
        ["1/11/2017", "Bellows College", "Education", "350.1"],
        ["1/15/2017", "Trey Research", "Other", "135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
    ]);
    ```

11. `createTable()` 関数で、`TODO3` を次のコードに置き換えます。 注:

    - このコードでは、ゼロから始まるインデックスをテーブルの列コレクションの `getItemAt` メソッドに渡すことで、**Amount** 列への参照を取得します。

        > [!NOTE]
        > Excel.js のコレクション オブジェクト (`TableCollection`、`WorksheetCollection`、`TableColumnCollection` など) には、`items` プロパティがあります。このプロパティは、子オブジェクト タイプ (`Table`、`Worksheet`、`TableColumn` など) の配列ですが、`*Collection` オブジェクト自体は配列ではありません。

    - その次に、コードでは、**Amount** 列の範囲を小数点以下 2 桁までのユーロとして書式設定します。 

    - 最後に、列の幅と行の高さが、最も長い (または最も高い) データ項目の幅になるようにします。コードを書式設定するには `Range` オブジェクトを取得する必要があります。`TableColumn` と `TableRow` オブジェクトには、書式プロパティがありません。

    ```js
    expensesTable.columns.getItemAt(3).getRange().numberFormat = [['&euro;#,##0.00']];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();
    ```

12. プロジェクトに行ったすべての変更が保存されていることを確認します。

### <a name="test-the-add-in"></a>アドインをテストする

1. 以下の手順を実行し、ローカル Web サーバーを起動してアドインのサイドロードを行います。

    > [!NOTE]
    > 開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。 次のいずれかのコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。

    > [!TIP]
    > Mac でアドインをテストする場合は、先に進む前にプロジェクトのルート ディレクトリで次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します。
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - Excel でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。 ローカル Web サーバーが (まだ実行されていない場合) 起動し、アドインが読み込まれた Excel が開きます。

        ```command&nbsp;line
        npm start
        ```

    - Excel on the web でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。

        ```command&nbsp;line
        npm run start:web
        ```

        アドインを使用するには、Excel on the web で新しいドキュメントを開き、「[Office on the web で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)」の手順に従ってアドインをサイドロードします。

2. Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。

    ![Excel アドイン ボタン](../images/excel-quickstart-addin-3b.png)

3. 作業ウィンドウで、[**テーブルの作成**] ボタンを選択します。

    ![Excel チュートリアル - テーブルの作成](../images/excel-tutorial-create-table-2.png)

## <a name="filter-and-sort-a-table"></a>テーブルのフィルター処理と並べ替え

チュートリアルのこの手順では、以前に作成したテーブルをフィルター処理したり並べ替えたりします。

### <a name="filter-the-table"></a>表のフィルター処理

1. ファイル **./src/taskpane/taskpane.html** を開きます。

2. `create-table` ボタンの `<button>` 要素を見つけ、その行の後に次のマークアップを追加します。

    ```html
    <button class="ms-Button" id="filter-table">Filter Table</button><br/><br/>
    ```

3. ファイル **./src/taskpane/taskpane.js** を開きます。

4. `Office.onReady` メソッドの呼び出し内で、クリック ハンドラーを `create-table` ボタンに割り当てる行を見つけ、その行の後に次のコードを追加します。

    ```js
    document.getElementById("filter-table").onclick = filterTable;
    ```

5. 次の関数をファイルの最後に追加します。

    ```js
    function filterTable() {
        Excel.run(function (context) {

            // TODO1: Queue commands to filter out all expense categories except
            //        Groceries and Education.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. `filterTable()` 関数で、`TODO1` を次のコードに置き換えます。 注:

   - このコードは、まず、`createTable` メソッドとして `getItemAt` メソッドにインデックスを渡す代わりに、列名をフィルター処理する必要がある列への参照を取得します。そのためには、最初に列名を `getItem` メソッドに渡す必要があります。ユーザーはテーブル列を移動できるので、テーブルの作成後に、指定されたインデックスの列は変更される場合があります。そのため、列名を使用する方が列を参照できます。このチュートリアルでは、テーブルを作成するのと同じ方法で使用するので、ユーザーが列を移動している可能性はありませんので、このチュートリアルでは、`getItemAt` 安全に使用しました。

   - `applyValuesFilter` メソッドは、`Filter` オブジェクトのフィルター処理方法の 1 つです。

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(['Education', 'Groceries']);
    ``` 

### <a name="sort-the-table"></a>表の並べ替え

1. ファイル **./src/taskpane/taskpane.html** を開きます。

2. `filter-table` ボタンの `<button>` 要素を見つけ、その行の後に次のマークアップを追加します。 

    ```html
    <button class="ms-Button" id="sort-table">Sort Table</button><br/><br/>
    ```

3. ファイル **./src/taskpane/taskpane.js** を開きます。

4. `Office.onReady` メソッドの呼び出し内で、クリック ハンドラーを `filter-table` ボタンに割り当てる行を見つけ、その行の後に次のコードを追加します。

    ```js
    document.getElementById("sort-table").onclick = sortTable;
    ```

5. 次の関数をファイルの最後に追加します。

    ```js
    function sortTable() {
        Excel.run(function (context) {

            // TODO1: Queue commands to sort the table by Merchant name.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. `sortTable()` 関数で、`TODO1` を次のコードに置き換えます。 注:

   - アドインで並べ替えるのは Merchant 列のみであるため、このコードでは、1 つのメンバーだけを含む `SortField` オブジェクトの配列を作成します。

   - `SortField` オブジェクトの `key` プロパティは、並べ替える対象列の 0 から始まるインデックスです。

   - `Table` の `sort` メンバーは、`TableSort` オブジェクトであり、メソッドではありません。 `TableSort` オブジェクトの `apply` メソッドには、`SortField` が渡されます。

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var sortFields = [
        {
            key: 1,            // Merchant column
            ascending: false,
        }
    ];

    expensesTable.sort.apply(sortFields);
    ```

7. プロジェクトに行ったすべての変更が保存されていることを確認します。

### <a name="test-the-add-in"></a>アドインをテストする

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. アドイン タスク ウィンドウが Excel でまだ開いていない場合は、[**ホーム**] タブに移動し、リボンの [**作業ウィンドウを表示**] ボタンを選択して開きます。

3. このチュートリアルで以前に追加したテーブルが、開いているワークシートにない場合は、タスク ウィンドウの [**テーブルの作成**] ボタンを選択します。

4. [**テーブルのフィルター**] ボタンと [**テーブルの並べ替え**] ボタンを任意の順序で選択します。

    ![Excel のチュートリアル - テーブルのフィルター処理と並べ替え](../images/excel-tutorial-filter-and-sort-table-2.png)

## <a name="create-a-chart"></a>グラフの作成

チュートリアルのこの手順では、前の手順で作成したテーブルのデータを使用してグラフを作成して、そのグラフの書式を設定します。

### <a name="chart-a-chart-using-table-data"></a>テーブルのデータを使用してグラフを作成する

1. ファイル **./src/taskpane/taskpane.html** を開きます。

2. `sort-table` ボタンの `<button>` 要素を見つけ、その行の後に次のマークアップを追加します。 

    ```html
    <button class="ms-Button" id="create-chart">Create Chart</button><br/><br/>
    ```

3. ファイル **./src/taskpane/taskpane.js** を開きます。

4. `Office.onReady` メソッドの呼び出し内で、クリック ハンドラーを `sort-table` ボタンに割り当てる行を見つけ、その行の後に次のコードを追加します。

    ```js
    document.getElementById("create-chart").onclick = createChart;
    ```

5. 次の関数をファイルの最後に追加します。

    ```js
    function createChart() {
        Excel.run(function (context) {

            // TODO1: Queue commands to get the range of data to be charted.

            // TODO2: Queue command to create the chart and define its type.

            // TODO3: Queue commands to position and format the chart.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. `createChart()` 関数で、`TODO1` を次のコードに置き換えます。 ヘッダー行を除外するために、このコードでは、`getRange` メソッドではなく `Table.getDataBodyRange` メソッドを使用してグラフを作成するデータの範囲を取得しています。

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var dataRange = expensesTable.getDataBodyRange();
    ```

7. `createChart()` 関数で、`TODO2` を次のコードに置き換えます。 次のパラメーターに注意してください。

   - `add` への最初のパラメーターでは、グラフの種類を指定します。数十種類あります。

   - 2 番目のパラメーターでは、グラフに含めるデータの範囲を指定します。

   - 3 番目のパラメーターでは、テーブルからの一連のデータ ポイントを行方向と列方向のどちらでグラフ化する必要があるかを決定します。 オプション `auto` は、最適な方法を判断するように Excel に指示します。

    ```js
    var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ```

8. `createChart()` 関数で、`TODO3` を次のコードに置き換えます。 このコードのほとんどの部分は、わかりやすく説明不要なものです。 注:
   
   - `setPosition` 方法のパラメーターでは、グラフを挿入するワークシート領域の左上のセルを指定します。Excel では、指定した空間でグラフを見栄えよくするために、線の太さなどの調整ができます。
   
   - "系列" とは、テーブルの 1 つの列にある一連のデータ ポイントのことです。このテーブルには文字列以外の列は 1 列しか含まれていないため、Excel は、グラフ化するデータ ポイントの列は、この列のみであると推測します。その他の列はグラフのラベルであると解釈されます。従って、グラフに含まれる系列は 1 つのみとなり、この系列のインデックスは 0 となります。を含みます。"&euro; での値" というラベルを付ける系列は、 この系列です。

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in &euro;';
    ```

9. プロジェクトに行ったすべての変更が保存されていることを確認します。

### <a name="test-the-add-in"></a>アドインをテストする

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. アドイン タスク ウィンドウが Excel でまだ開いていない場合は、[**ホーム**] タブに移動し、リボンの [**作業ウィンドウを表示**] ボタンを選択して開きます。

3. このチュートリアルで以前に追加したテーブルが、開いているワークシートにない場合は、タスク ウィンドウの [**テーブルの作成**] ボタンを選択します。次に、[**テーブルのフィルター処理**] ボタン、および [**テーブルの並べ替え**] ボタンのいずれかを選択します。

4. [グラフの作成 ****] ボタンを選択します。グラフが作成され、フィルター処理された行のデータのみが含まれます。一番下にあるデータポイントのラベルは、グラフの並べ替え順序になります。つまり、名前の逆アルファベット順での商社の名前です。

    ![Excel チュートリアル - グラフの作成](../images/excel-tutorial-create-chart-2.png)

## <a name="freeze-a-table-header"></a>テーブルのヘッダーの固定

ユーザーがスクロールして一部の行を見る必要があるほど長いテーブルがあると、見出し行がスクロールして見えなくなります。チュートリアルのこの手順では、ユーザーがワークシートを下にスクロールしても表示されるように、前に作成したテーブルの見出し行を固定します。

### <a name="freeze-the-tables-header-row"></a>表のヘッダー行を固定する

1. ファイル **./src/taskpane/taskpane.html** を開きます。

2. `create-chart` ボタンの `<button>` 要素を見つけ、その行の後に次のマークアップを追加します。 

    ```html
    <button class="ms-Button" id="freeze-header">Freeze Header</button><br/><br/>
    ```

3. ファイル **./src/taskpane/taskpane.js** を開きます。

4. `Office.onReady` メソッドの呼び出し内で、クリック ハンドラーを `create-chart` ボタンに割り当てる行を見つけ、その行の後に次のコードを追加します。

    ```js
    document.getElementById("freeze-header").onclick = freezeHeader;
    ```

5. 次の関数をファイルの最後に追加します。

    ```js
    function freezeHeader() {
        Excel.run(function (context) {

            // TODO1: Queue commands to keep the header visible when the user scrolls.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. `freezeHeader()` 関数で、`TODO1` を次のコードに置き換えます。 注:

   - `Worksheet.freezePanes` コレクションは、ワークシートのスクロール操作時に、ワークシート上でピン留めつまり固定される一式のペインのことです。

   - `freezeRows` メソッドでは、上から数えた行数を、ピン留めする位置のパラメーターとして使用します。`1` を渡して最初の行を適所にピン留めします。

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

7. プロジェクトに行ったすべての変更が保存されていることを確認します。

### <a name="test-the-add-in"></a>アドインをテストする

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. アドイン タスク ウィンドウが Excel でまだ開いていない場合は、[**ホーム**] タブに移動し、リボンの [**作業ウィンドウを表示**] ボタンを選択して開きます。

3. このチュートリアルに以前追加したテーブルがワークシートに存在する場合は、それを削除します。

4. 作業ウィンドウで、[**テーブルの作成**] ボタンを選択します。

5. 作業ウィンドウで、[**ヘッダーを固定**] ボタンを選択します。

6. ヘッダー以降の行が画面の外に出て見えなくなるまでワークシートを十分下にスクロールしても、表のヘッダーが最上部に表示されていることを確認します。

    ![Excel のチュートリアル - ヘッダーの固定](../images/excel-tutorial-freeze-header-2.png)

## <a name="protect-a-worksheet"></a>ワークシートの保護

チュートリアルのこの手順では、リボンに別のボタンを追加します。このボタンをクリックすると、ワークシートの保護のオン/オフが切り替わるように定義した関数が実行されるようにします。

### <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a>2 つ目のリボン ボタンを追加するようにマニフェストを構成する

1. マニフェスト ファイル **./manifest.xml** を開きます。

2. `<Control>` 要素を見つけます。 この要素では、アドインの起動に使用している **[ホーム]** リボンの **[作業ウィンドウの表示]** ボタンを定義しています。 ここでは、**[ホーム]** リボンの同じグループに 2 つ目のボタンを追加します。 Control 終了タグ (`</Control>`) と Group 終了タグ (`</Group>`) の間に、次のマークアップを追加します。

    ```xml
    <Control xsi:type="Button" id="<!--TODO1: Unique (in manifest) name for button -->">
        <Label resid="<!--TODO2: Button label -->" />
        <Supertip>            
            <Title resid="<!-- TODO3: Button tool tip title -->" />
            <Description resid="<!-- TODO4: Button tool tip description -->" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Icon.16x16"/>
            <bt:Image size="32" resid="Icon.32x32"/>
            <bt:Image size="80" resid="Icon.80x80"/>
        </Icon>
        <Action xsi:type="<!-- TODO5: Specify the type of action-->">
            <!-- TODO6: Identify the function.-->
        </Action>
    </Control>
    ```

3. マニフェスト ファイルに追加した XML 内で `TODO1` を文字列に置き換えて、このマニフェスト ファイル内で一意の ID をボタンに割り当てます。 このボタンでは、ワークシートの保護のオン/オフを切り替える予定なので、「ToggleProtection」を使用することにします。 完了すると、`Control` 要素の開始タグは次のようになります。

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. その次の 3 つの `TODO` では、"resid" を設定します ("resid" はリソース ID の略号です)。 リソースは文字列です。これら 3 つの文字列は、この後の手順で作成します。 ここでは、そのリソースに ID を割り当てる必要があります。 ボタンのラベルは「Toggle Protection」と表示されるようにしますが、この文字列の *ID* は「ProtectionButtonLabel」にします。そのため、`Label` 要素は次のようになります。

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. `SuperTip` 要素では、このボタンのツール ヒントを定義します。 ツール ヒントのタイトルはボタンのラベルと同じにする必要があるため、リソース ID にはまったく同じ "ProtectionButtonLabel" を使用することにします。 ツール ヒントの説明は、"Click to turn protection of the worksheet on and off" にする予定です。 ただし、`ID` は "ProtectionButtonToolTip" にします。 したがって、完了すると、`SuperTip` 要素は次のようになります。 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > 運用アドインでは、2つの異なるボタンに同じアイコンを使用することはできません。ただし、このチュートリアルを簡素化するには、このチュートリアルを行います。新しい `Control` の `Icon` マークアップは、既存の `Control`の `Icon` 要素のコピーにすぎません。 

6. 既にマニフェストに存在している元の `Control` 要素の内側にある `Action` 要素では、その要素のタイプが `ShowTaskpane` に設定されていますが、新しいボタンで作業ウィンドウを開く予定はありません。このボタンでは、この後の手順で作成するカスタム関数を実行する予定です。 そのため、`TODO5` は、カスタム関数をトリガーするボタンのアクション タイプである `ExecuteFunction` に置き換えます。 `Action` 要素の開始タグは次のようになります。
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. 元の `Action` 要素には、作業ウィンドウの ID と、作業ウィンドウで開くページの URL を指定する子要素があります。ただし、`ExecuteFunction` の種類の `Action` の要素には、そのコントロールが実行している関数に名前を付けた単一の子要素があります。この関数は、後の手順で作成し、`toggleProtection`と呼ばれます。`TODO6` は次のマークアップに置き換えます。
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    `Control` マークアップの全体は、次のようになりました。

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
        <Label resid="ProtectionButtonLabel" />
        <Supertip>            
            <Title resid="ProtectionButtonLabel" />
            <Description resid="ProtectionButtonToolTip" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Icon.16x16"/>
            <bt:Image size="32" resid="Icon.32x32"/>
            <bt:Image size="80" resid="Icon.80x80"/>
        </Icon>
        <Action xsi:type="ExecuteFunction">
           <FunctionName>toggleProtection</FunctionName>
        </Action>
    </Control>
    ```

8. マニフェストの `Resources` セクションまで下にスクロールします。

9. `bt:ShortStrings` 要素の子として、次のマークアップを追加します。

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. `bt:LongStrings` 要素の子として、次のマークアップを追加します。

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. ファイルを保存します。

### <a name="create-the-function-that-protects-the-sheet"></a>シートを保護する関数を作成する

1. **.\commands\commands.js** ファイルを開きます。

2. `action` 関数の直後に次の関数を追加します。 関数に `args` パラメーターを指定していることと、関数の最後のほうの行で `args.completed` を呼び出していることに注目してください。 **ExecuteFunction** タイプのすべてのアドイン コマンドでは、これが要件になります。 これにより、関数が終了したことと、UI が再度応答可能になることを Office ホスト アプリケーションに通知します。

    ```js
    function toggleProtection(args) {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to reverse the protection status of the current worksheet.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```

3. 次の行を、ファイルの最後に追加します。

    ```js
    g.toggleProtection = toggleProtection;
    ```

4. `toggleProtection` 関数で、`TODO1` を次のコードに置き換えます。 このコードでは、標準の切り替えパターンで、ワークシート オブジェクトの protection プロパティを使用します。 `TODO2` については、次のセクションで説明します。

    ```js
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // TODO2: Queue command to load the sheet's "protection.protected" property from
    //        the document and re-synchronize the document and task pane.

    if (sheet.protection.protected) {
        sheet.protection.unprotect();
    } else {
        sheet.protection.protect();
    }
    ``` 

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a>ドキュメントのプロパティを作業ウィンドウのスクリプト オブジェクトにフェッチするコードを追加する

これまでこのチュートリアルで作成した各関数で、Office ドキュメントに*書き込む*コマンドをキューに入れました。 各関数は、キューに登録されたコマンドを実行対象のドキュメントに送信する `context.sync()` メソッドを呼び出すことで終了しています。 ただし、最後の手順で追加したコードでは、`sheet.protection.protected` プロパティを呼び出しています。このことが、これまでに作成した関数とは大きく異なります。`sheet` オブジェクトは、この作業ウィンドウのスクリプトに存在する単なるプロキシ オブジェクトなので、 ドキュメントの実際の保護の状態を認識できません。そのため、その `protection.protected` プロパティでは実際の値が保持できません。 まず、ドキュメントから保護の状態をフェッチする必要があり、その状態を使用して `sheet.protection.protected` の値を設定します。 そのようにした場合にのみ、例外がスローされることなく `sheet.protection.protected` を呼び出せるようになります。 このフェッチ処理には、3 つの手順があります。

   1. コードで読み取る必要があるプロパティをロードする (つまりフェッチする) コマンドをキューに登録します。

   2. コンテキスト オブジェクトの `sync` メソッドを呼び出します。このメソッドは、キューに登録されたコマンドを実行対象のドキュメントに送信して、要求された情報を返します。

   3. `sync` メソッドは非同期であるため、フェッチされたプロパティをコードで呼び出す前に、そのメソッドが完了していることを確認します。

こうした手順は、コードで Office ドキュメントから情報を*読み取る*必要がある場合には必ず完了する必要があります。

1. `toggleProtection` 関数で、`TODO2` を次のコードに置き換えます。 注:
   
   - すべての Excel オブジェクトに `load` の方法があります。パラメーターで読み取るオブジェクトのプロパティをコンマ区切り名前の文字列として指定します。この場合、必要なプロパティは、`protection` プロパティのサブプロパティです。サブプロパティは、コード内の他の場所とほぼ同じ方法で参照します。ただし、"." 文字の代わりにスラッシュ ('/') を使用します。

   - `sync` が完了してドキュメントからフェッチされた適切な値が `sheet.protection.protected` に割り当てられるまで、`sheet.protection.protected` を読み取る切り替えロジックが実行されないようにするために、そのロジックを `sync` が完了するまで実行されない `then` 関数に (この後の手順で) 移動します。 

    ```js
    sheet.load('protection/protected');
    return context.sync()
        .then(
            function() {
                // TODO3: Move the queued toggle logic here.
            }
        )
        // TODO4: Move the final call of `context.sync` here and ensure that it
        //        does not run until the toggle logic has been queued.
    ``` 

2. 分岐していない同一のコード パスに 2 つの `return` ステートメントを含めることはできないため、`Excel.run` の最後にある最終行の `return context.sync();` を削除します。新しい最後の `context.sync` は、このチュートリアルの後の方で追加します。

3. `toggleProtection` 関数内の `if ... else` 構造を切り取って、`TODO3` の代わりに貼り付けます。

4. `TODO4` を次のコードに置き換えます。次の点に注意してください。

   - `sync` メソッドを `then` 関数に渡すことで、`sheet.protection.unprotect()` または `sheet.protection.protect()` のどちらかがキューに登録されるまで、そのメソッドが実行されないようにします。

   - `then` メソッドは渡された関数を呼び出します。`sync` が 2 回呼び出されないように、`context.sync` の末尾の "()" は省略します。

    ```js
    .then(context.sync);
    ```

   作業が完了すると、関数の全体は次のようになります。

    ```js
    function toggleProtection(args) {
        Excel.run(function (context) {            
          var sheet = context.workbook.worksheets.getActiveWorksheet();          
          sheet.load('protection/protected');

          return context.sync()
              .then(
                  function() {
                    if (sheet.protection.protected) {
                        sheet.protection.unprotect();
                    } else {
                        sheet.protection.protect();
                    }
                  }
              )
              .then(context.sync);
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```

5. プロジェクトに行ったすべての変更が保存されていることを確認します。

### <a name="test-the-add-in"></a>アドインをテストする

1. Excel も含めて、すべての Office アプリケーションを閉じます。 

2. [キャッシュ] フォルダーの内容を削除して、Office キャッシュを削除します。これは、アドインの以前のバージョンをホストから完全に消去するために必要です。 

    - Windows の場合: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`。

    - Mac の場合: `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`。 
    
        > [!NOTE]
        > そのフォルダーが存在しない場合は、次のフォルダーを確認し、見つかった場合はフォルダーの内容を削除します。
        >    - `{host}` が Office ホスト (例: `Excel`) の `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/`
        >    - `{host}` が Office ホスト (例: `Excel`) の `~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`
        >    - `com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`

3. ローカル Web サーバーが既に実行中の場合は、ノード コマンド ウィンドウを閉じて終了します。

4. マニフェスト ファイルが更新されているため、更新されたマニフェスト ファイルを使用してアドインを再度サイドロードする必要があります。 ローカル Web サーバーを起動し、アドインのサイドロードを行います。 

    - Excel でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。 ローカル Web サーバーが (まだ実行されていない場合) 起動し、アドインが読み込まれた Excel が開きます。

        ```command&nbsp;line
        npm start
        ```

    - Excel on the web でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。

        ```command&nbsp;line
        npm run start:web
        ```

        アドインを使用するには、Excel on the web で新しいドキュメントを開き、「[Office on the web で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)」の手順に従ってアドインをサイドロードします。

5. Excel の [**ホーム**] タブで、[**ワークシート保護を切り換える**] ボタンを選択します。 次のスクリーンショットに示すように、リボンのほとんどのコントロールは、無効化 (淡色表示) されます。 

    ![Excel チュートリアル - 保護がオンになっているリボン](../images/excel-tutorial-ribbon-with-protection-on-2.png)

6. セルの内容を変更する場合は、そのセルを選択します。 Excel にワークシートが保護されていることを示すエラー メッセージが表示されます。

7. もう一度 [**ワークシート保護を切り換える**] ボタンを選択すると、コントロールが再有効化され、再びセルの値を変更できるようになります。

## <a name="open-a-dialog"></a>ダイアログを開く

チュートリアルの最後の手順では、アドインのダイアログを開いて、ダイアログプロセスのメッセージを作業ウィンドウのプロセスに渡し、ダイアログボックスを閉じます。Office アドインのダイアログには *非モーダル*があります。ユーザーは、ホスト Office アプリケーションでも、作業ウィンドウのホストページでも、ドキュメントの操作を続行できます。

### <a name="create-the-dialog-page"></a>ダイアログ ページを作成する

1. プロジェクトのルートにある **./src** フォルダーで、**ダイアログ**という名前の新しいフォルダーを作成します。

2. **./src/dialogs** フォルダー に **popup.html** という名前の新しいファイルを作成します。

3. **popup.html** に、次のコードを追加します。 注:

   - このページには、ユーザーが自分の名前を入力する `<input>` と、その名前を作業ウィンドウ内のページ (入力した名前が表示されるページ) に送信するボタンが含まれています。

   - このマークアップでは、**popup.js** という名前のスクリプトを読み込みます。このスクリプトは、この後の手順で作成します。

   - また、**popup.js** で使用することになる Office.JS ライブラリも読み込みます。

    ```html
    <!DOCTYPE html>
    <html>
        <head lang="en">
            <title>Dialog for My Office Add-in</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">

            <!-- For more information on Office UI Fabric, visit https://developer.microsoft.com/fabric. -->
            <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css"/>

            <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
            <script type="text/javascript" src="popup.js"></script>

        </head>
        <body style="display:flex;flex-direction:column;align-items:center;justify-content:center">
            <p class="ms-font-xl">ENTER YOUR NAME</p>
            <input id="name-box" type="text"/><br/><br/>
            <button id="ok-button" class="ms-Button">OK</button>
        </body>
    </html>
    ```

4. **./src/dialogs** フォルダーで、**popup.js** という名前の新しいファイルを作成します。

5. **popup.js** に、次のコードを追加します。 このコードについては、次の点に注意してください。

   - *Office.js ライブラリ内の API を呼び出すすべてのページでは、まずライブラリが完全に初期化されていることを確認する必要があります。* これを行う最善の方法は `Office.onReady()` メソッドを呼び出すことです。 アドインに独自の初期化タスクがある場合、コードを `Office.onReady()` の呼び出しにチェーンされている `then()` メソッドに含める必要があります。 `Office.onReady()` の呼び出しは、Office.js を呼び出す前に実行する必要があります。そのため、この例で示すように、割り当てはページによって読み込まれるスクリプト ファイル内に入れてあります。

    ```js
    (function () {
    "use strict";

        Office.onReady()
            .then(function() {

                // TODO1: Assign handler to the OK button.

            });

        // TODO2: Create the OK button handler

    }());
    ```

6. `TODO1` は次のコードで置き換えます。`sendStringToParentPage` 関数は、次の手順で作成します。

    ```js
    document.getElementById("ok-button").onclick = sendStringToParentPage;
    ```

7. `TODO2` は次のコードで置き換えます。`messageParent` メソッドは、パラメーターを親ページ (この場合は作業ウィンドウのページ) に渡します。パラメーターには、ブール値または文字列を指定できます。これには、XML や JSON など、文字列としてシリアル化できるものが含まれます。

    ```js
    function sendStringToParentPage() {
        var userName = document.getElementById("name-box").value;
        Office.context.ui.messageParent(userName);
    }
    ```

> [!NOTE]
> **popup.html** ファイルと、そのファイルで読み込む **popup.js** ファイルは、アドインの作業ウィンドウとは完全に別な Microsoft Edge または Internet Explorer プロセスで実行されます。 **popup.js** が **app.js** ファイルと同じ **bundle.js** ファイルからトランスパイルされていた場合、アドインでは **bundle.js** の 2 つのコピーを読み込むことが必要になり、バンドル化の意味がなくなります。 そのため、このアドインでは **popup.js** を一切トランスパイルしていません。

### <a name="update-webpack-config-settings"></a>Webpackの機能設定を更新する

プロジェクトのルートディレクトリにあるファイル **webpack.config.js** を開き、以下の手順を実行します。

1. `config`オブジェクト内で`entry`オブジェクトを探し、`popup`の新しいエントリーを追加します。

    ```js
    popup: "./src/dialogs/popup.js"
    ```

    これを実行すると、新しい`entry`オブジェクトは次のようになります。

    ```js
    entry: {
      polyfill: "@babel/polyfill",
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
      popup: "./src/dialogs/popup.js"
    },
    ```
  
2. `config`オブジェクト内で`plugins`配列を探し、次の新しいオブジェクトをその配列の末尾に追加します。

    ```js
    new HtmlWebpackPlugin({
      filename: "popup.html",
      template: "./src/dialogs/popup.html",
      chunks: ["polyfill", "popup"]
    })
    ```

    これを実行すると、新しい`plugins`配列は次のようになります。

    ```js
    plugins: [
      new CleanWebpackPlugin(),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ['polyfill', 'taskpane']
      }),
      new CopyWebpackPlugin([
      {
        to: "taskpane.css",
        from: "./src/taskpane/taskpane.css"
      }
      ]),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      }),
      new HtmlWebpackPlugin({
        filename: "popup.html",
        template: "./src/dialogs/popup.html",
        chunks: ["polyfill", "popup"]
      })
    ],
    ```

3. ローカル Web サーバーが実行中の場合は、ノード コマンド ウィンドウを閉じて終了します。

4. 次のコマンドを実行してプロジェクトを再構築します。

    ```command&nbsp;line
    npm run build
    ```

### <a name="open-the-dialog-from-the-task-pane"></a>作業ウィンドウからダイアログを開く

1. ファイル **./src/taskpane/taskpane.html** を開きます。

2. `freeze-header` ボタンの `<button>` 要素を見つけ、その行の後に次のマークアップを追加します。

    ```html
    <button class="ms-Button" id="open-dialog">Open Dialog</button><br/><br/>
    ```

3. このダイアログでは、ユーザーに名前の入力を求めて、ユーザーの名前を作業ウィンドウに渡します。 作業ウィンドウでは、それがラベルに表示されます。 前の手順で追加した `button` の直後に、次のマークアップを追加します。

    ```html
    <label id="user-name"></label><br/><br/>
    ```

4. ファイル **./src/taskpane/taskpane.js** を開きます。

5. `Office.onReady` メソッドの呼び出し内で、クリック ハンドラーを `freeze-header` ボタンに割り当てる行を見つけ、その行の後に次のコードを追加します。 `openDialog`メソッドは後の手順で作成します。

    ```js
    document.getElementById("open-dialog").onclick = openDialog;
    ```

6. ファイルの最後に次の宣言を追加します。 この変数は、親ページの実行コンテキスト内のオブジェクトを保持するために使用され、ダイアログ ページの実行コンテキストへの仲介者として機能します。

    ```js
    var dialog = null;
    ```

7. (`dialog` の宣言の後で) ファイルの最後に次の関数を追加します。 このコードで注目する重要な点は、そこに `Excel.run` の呼び出しが存在*しない*ことです。 これは、ダイアログを開く API はすべての Office ホストで共有されるため、Excel 固有の API ではなく Office JavaScript 共通 API に含まれているからです。

    ```js
    function openDialog() {
        // TODO1: Call the Office Common API that opens a dialog
    }
    ```

8. `TODO1` を次のコードに置き換えます。次の点に注意してください。

   - `displayDialogAsync` メソッドでは、画面の中央にダイアログを開きます。

   - 最初のパラメーターは、開くページの URL です。

   - 2 番目のパラメーターでオプションを渡します。`height` と `width` は、Office アプリケーションのウィンドウ サイズの比率です。

    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},

        // TODO2: Add callback parameter.
    );
    ```

### <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a>ダイアログからのメッセージを処理してダイアログを閉じる

1. ファイル **./src/taskpane/taskpane.js** の `openDialog` 関数内で、`TODO2` を次のコードに置き換えます。 注:

   - コールバックは、ダイアログが正常に開いた直後、ユーザーがダイアログで操作を行う前に実行されます。

   - `result.value` は、親ページとダイアログ ページの実行コンテキストの間で仲介者のように機能するオブジェクトです。

   - このハンドラーは、`processMessage` 関数の呼び出しによって、ダイアログから送信されるあらゆる値を処理します。

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. `openDialog` 関数の後に次の関数を追加します。

    ```js
    function processMessage(arg) {
        document.getElementById("user-name").innerHTML = arg.message;
        dialog.close();
    }
    ```

3. プロジェクトに行ったすべての変更が保存されていることを確認します。

### <a name="test-the-add-in"></a>アドインをテストする

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. アドイン タスク ウィンドウが Excel でまだ開いていない場合は、[**ホーム**] タブに移動し、リボンの [**作業ウィンドウを表示**] ボタンを選択して開きます。

3. 作業ウィンドウで、**[Open Dialog]** ボタンをクリックします。

4. ダイアログが開いたら、ドラッグしたりサイズ変更したりします。 ワークシートを操作して、作業ウィンドウの他のボタンを押すことはできますが、同じ作業ウィンドウのページから 2 番目のダイアログを起動することはできないことに注意してください。

5. ダイアログで、名前を入力して [**OK**] ボタンを選択します。 作業ウィンドウに名前が表示され、ダイアログが閉じられます。

6. 必要に応じて、`processMessage` 関数の行 `dialog.close();` をコメントにします。このセクションの手順を繰り返します。ダイアログは開いたままで、名前を変更できます。右上隅の **X** ボタンを押して、手動で閉じることができます。

    ![Excel チュートリアル - ダイアログ](../images/excel-tutorial-dialog-open-2.png)

## <a name="next-steps"></a>次の手順

このチュートリアルでは、Excel ブック内のテーブル、グラフ、ワークシート、ダイアログの操作を行う、Excel 作業ウィンドウ アドインを作成しました。 Excel アドインの構築に関する詳細については、次の記事にお進みください。

> [!div class="nextstepaction"]
> [Excel アドインの概要](../excel/excel-add-ins-overview.md)

## <a name="see-also"></a>関連項目

* [Office アドイン プラットフォームの概要](../overview/office-add-ins.md)
* [Office アドインを構築する](../overview/office-add-ins-fundamentals.md)
* [Office アドインを開発する](../develop/develop-overview.md)
* [Excel JavaScript API を使用した基本的なプログラミングの概念](../excel/excel-add-ins-core-concepts.md)
