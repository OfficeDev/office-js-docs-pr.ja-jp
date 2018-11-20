チュートリアルのこの手順では、以前に作成した表をフィルター処理したり並べ替えたりします。

> [!NOTE]
> このページでは、Excel のアドインのチュートリアルの個々 の手順について説明します。 このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[Excel アドインのチュートリアル](../tutorials/excel-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。

## <a name="filter-the-table"></a>表のフィルター処理

1. コード エディターでプロジェクトを開きます。
2. index.html ファイルを開きます。
3. `create-table` ボタンを格納している `div` の直下に、次のマークアップを追加します。

    ```html
    <div class="padding">
        <button class="ms-Button" id="filter-table">Filter Table</button>
    </div>
    ```

4. app.js ファイルを開きます。

5. `create-table` ボタンにクリック ハンドラーを割り当てる行の直下に、次のコードを追加します。

    ```js
    $('#filter-table').click(filterTable);
    ```

6. `createTable` 関数の直下に、次の関数を追加します。

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

7. `TODO1` を次のコードに置き換えます。次の点に注意してください。
   - このコードでは最初に、`getItem` メソッドに列名を渡すことによって、フィルター処理が必要な列への参照を取得します。`createTable` メソッドが行うように、列のインデックスを `getItemAt` メソッドに渡すわけではありません。 ユーザーは表の列を移動させることができるので、表を作成した後、指定したインデックスにある列が変わってしまう可能性があります。 そのため、列名を使用して列への参照を取得するほうが安全です。 前のチュートリアルでは、表を作成するのとまったく同じ方法で `getItemAt` を使用したため、ユーザーが列を移動させた可能性はなく、よって安全に使用できました。
   - `applyValuesFilter` メソッドは、`Filter` オブジェクトのフィルター処理方法の 1 つです。

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);
    ``` 

## <a name="sort-the-table"></a>表の並べ替え

1. index.html ファイルを開きます。
2. `filter-table` ボタンを格納している `div` の下に、次のマークアップを追加します。

    ```html
    <div class="padding">
        <button class="ms-Button" id="sort-table">Sort Table</button>
    </div>
    ```

3. app.js ファイルを開きます。

4. `filter-table` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。

    ```js
    $('#sort-table').click(sortTable);
    ```

5. `filterTable` 関数の下に、次の関数を追加します。

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

7. `TODO1` を次のコードに置き換えます。次の点に注意してください。
   - アドインで並べ替えるのは Merchant 列のみであるため、このコードでは、1 つのメンバーだけを含む `SortField` オブジェクトの配列を作成します。
   - `SortField` オブジェクトの `key` プロパティは、並べ替える対象列の 0 から始まるインデックスです。
   - `Table` の `sort` メンバーは、`TableSort` オブジェクトであり、メソッドではありません。 `TableSort` オブジェクトの `apply` メソッドには、`SortField` が渡されます。

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const sortFields = [
        {
            key: 1,            // Merchant column
            ascending: false,
        }
    ];

    expensesTable.sort.apply(sortFields);
    ```

## <a name="test-the-add-in"></a>アドインをテストする

1. Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、Ctrl-C を 2 回入力して実行中の Web サーバーを停止します。 それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。

     > [!NOTE]
     > ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。 そのためには、ビルド コマンドの入力を求めるプロンプトが表示されるように、サーバー プロセスを強制終了する必要があります。 ビルド後に、サーバーを再起動します。 次の数ステップで、このプロセスを実行します。

1. `npm run build` コマンドを実行して、ES6 ソース コードを Internet Explorer でサポートされている以前のバージョンの JavaScript にトランスパイルします (これは、Excel アドインを実行するために Excel の内部で使用されます)。
2. `npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。
4. 作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。
5. 何らかの理由から開いているワークシートに表が含まれていない場合は、作業ウィンドウの **[Create Table]** (表の作成) ボタンを選択します。
6. **[Filter Table]** (表のフィルター) ボタンと **[Sort Table]** (表の並べ替え) ボタンを任意の順序で選択します。

    ![Excel のチュートリアル - 表のフィルター処理と並べ替え](../images/excel-tutorial-filter-and-sort-table.png)
