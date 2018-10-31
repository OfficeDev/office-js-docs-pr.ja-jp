<span data-ttu-id="88005-101">チュートリアルのこの手順では、以前に作成した表をフィルター処理したり並べ替えたりします。</span><span class="sxs-lookup"><span data-stu-id="88005-101">In this step of the tutorial, you'll filter and sort the table that you created previously.</span></span>

> [!NOTE]
> <span data-ttu-id="88005-102">このページでは、Excel のアドインのチュートリアルの個々 の手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="88005-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="88005-103">このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[Excel アドインのチュートリアル](../tutorials/excel-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。</span><span class="sxs-lookup"><span data-stu-id="88005-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="filter-the-table"></a><span data-ttu-id="88005-104">表のフィルター処理</span><span class="sxs-lookup"><span data-stu-id="88005-104">Filter the table</span></span>

1. <span data-ttu-id="88005-105">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="88005-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="88005-106">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="88005-106">Open the file index.html.</span></span>
3. <span data-ttu-id="88005-107">ボタンを格納している `div` の直下に、次のマークアップを追加します。`create-table`</span><span class="sxs-lookup"><span data-stu-id="88005-107">Just below the `div` that contains the `create-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="filter-table">Filter Table</button>            
    </div>
    ```

4. <span data-ttu-id="88005-108">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="88005-108">Open the app.js file.</span></span>

5. <span data-ttu-id="88005-109">ボタンにクリック ハンドラーを割り当てる行の直下に、次のコードを追加します。`create-table`</span><span class="sxs-lookup"><span data-stu-id="88005-109">Just below the line that assigns a click handler to the `create-table` button, add the following code:</span></span>

    ```js
    $('#filter-table').click(filterTable);
    ```

6. <span data-ttu-id="88005-110">関数の直下に、次の関数を追加します。`createTable`</span><span class="sxs-lookup"><span data-stu-id="88005-110">Just below the `createTable` function, add the following function:</span></span>

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

7. <span data-ttu-id="88005-p102">`TODO1`を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="88005-p102">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="88005-113">このコードでは最初に、`getItem` メソッドに列名を渡すことによって、フィルター処理が必要な列への参照を取得します。`createTable` メソッドが行うように、列のインデックスを `getItemAt` メソッドに渡すわけではありません。</span><span class="sxs-lookup"><span data-stu-id="88005-113">The code first gets a reference to the column that needs filtering by passing the column name to the `getItem` method, instead of passing its index to the `getItemAt` method as the `createTable` method does.</span></span> <span data-ttu-id="88005-114">ユーザーは表の列を移動させることができるので、表を作成した後、指定したインデックスにある列が変わってしまう可能性があります。</span><span class="sxs-lookup"><span data-stu-id="88005-114">Since users can move table columns, the column at a given index might change after the table is created.</span></span> <span data-ttu-id="88005-115">そのため、列名を使用して列への参照を取得するほうが安全です。</span><span class="sxs-lookup"><span data-stu-id="88005-115">Hence, it is safer to use the column name to get a reference to the column.</span></span> <span data-ttu-id="88005-116">前のチュートリアルでは、表を作成するのとまったく同じ方法で `getItemAt` を使用したため、ユーザーが列を移動させた可能性はなく、よって安全に使用できました。</span><span class="sxs-lookup"><span data-stu-id="88005-116">We used `getItemAt` safely in the preceding tutorial, because we used it in the very same method that creates the table, so there is no chance that a user has moved the column.</span></span>
   - <span data-ttu-id="88005-117">メソッドは、`Filter` オブジェクトのフィルター処理方法の 1 つです。`applyValuesFilter`</span><span class="sxs-lookup"><span data-stu-id="88005-117">The `applyValuesFilter` method is one of several filtering methods on the `Filter` object.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);
    ``` 

## <a name="sort-the-table"></a><span data-ttu-id="88005-118">表の並べ替え</span><span class="sxs-lookup"><span data-stu-id="88005-118">Sort the table</span></span>

1. <span data-ttu-id="88005-119">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="88005-119">Open the file index.html.</span></span>
2. <span data-ttu-id="88005-120">ボタンを格納している `div` の下に、次のマークアップを追加します。`filter-table`</span><span class="sxs-lookup"><span data-stu-id="88005-120">Below the `div` that contains the `filter-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="sort-table">Sort Table</button>            
    </div>
    ```

3. <span data-ttu-id="88005-121">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="88005-121">Open the app.js file.</span></span>

4. <span data-ttu-id="88005-122">ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。`filter-table`</span><span class="sxs-lookup"><span data-stu-id="88005-122">Below the line that assigns a click handler to the `filter-table` button, add the following code:</span></span>

    ```js
    $('#sort-table').click(sortTable);
    ```

5. <span data-ttu-id="88005-123">関数の下に、次の関数を追加します。`filterTable`</span><span class="sxs-lookup"><span data-stu-id="88005-123">Below the `filterTable` function add the following function.</span></span>

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

7. <span data-ttu-id="88005-p104">`TODO1`を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="88005-p104">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="88005-126">アドインで並べ替えるのは Merchant 列のみであるため、このコードでは、1 つのメンバーだけを含む `SortField` オブジェクトの配列を作成します。</span><span class="sxs-lookup"><span data-stu-id="88005-126">The code creates an array of `SortField` objects which has just one member since the add-in only sorts on the Merchant column.</span></span>
   - <span data-ttu-id="88005-127">オブジェクトの `key` プロパティは、並べ替える対象列の 0 から始まるインデックスです。`SortField`</span><span class="sxs-lookup"><span data-stu-id="88005-127">The `key` property of a `SortField` object is the zero-based index of the column to sort-on.</span></span>
   - <span data-ttu-id="88005-128">|||UNTRANSLATED_CONTENT_START|||The `sort` member of a `Table` is a `TableSort` object, not a method.|||UNTRANSLATED_CONTENT_END|||</span><span class="sxs-lookup"><span data-stu-id="88005-128">The `sort` member of a `Table` is a `TableSort` object, not a method.</span></span> <span data-ttu-id="88005-129">`TableSort`オブジェクトの `apply`メソッドでは、`SortField` が渡されます。</span><span class="sxs-lookup"><span data-stu-id="88005-129">The `SortField`s are passed the `TableSort` object's `apply` method.</span></span>

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

## <a name="test-the-add-in"></a><span data-ttu-id="88005-130">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="88005-130">Test the add-in</span></span>

1. <span data-ttu-id="88005-131">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、Ctrl-C を 2 回入力して実行中の Web サーバーを停止します。</span><span class="sxs-lookup"><span data-stu-id="88005-131">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="88005-132">それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="88005-132">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="88005-133">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。</span><span class="sxs-lookup"><span data-stu-id="88005-133">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="88005-134">そのためには、ビルド コマンドの入力を求めるプロンプトが表示されるように、サーバー プロセスを強制終了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="88005-134">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="88005-135">ビルド後に、サーバーを再起動します。</span><span class="sxs-lookup"><span data-stu-id="88005-135">After the build, you restart the server.</span></span> <span data-ttu-id="88005-136">次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="88005-136">The next few steps carry out this process.</span></span>

1. <span data-ttu-id="88005-137">コマンドを実行して、ES6 ソース コードを Internet Explorer でサポートされている以前のバージョンの JavaScript にトランスパイルします (これは、Excel アドインを実行するために Excel の内部で使用されます)。`npm run build`</span><span class="sxs-lookup"><span data-stu-id="88005-137">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
2. <span data-ttu-id="88005-138">コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。`npm start`</span><span class="sxs-lookup"><span data-stu-id="88005-138">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="88005-139">作業ウィンドウを再読み込みするために、そのウィンドウを閉じ、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="88005-139">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="88005-140">何らかの理由から開いているワークシートに表が含まれていない場合は、作業ウィンドウの **[Create Table]** (表の作成) ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="88005-140">If for any reason the table is not in the open worksheet, in the taskpane, choose **Create Table**.</span></span> 
6. <span data-ttu-id="88005-141">**[Filter Table]** (表のフィルター) ボタンと **[Sort Table]** (表の並べ替え) ボタンを任意の順序で選択します。</span><span class="sxs-lookup"><span data-stu-id="88005-141">Choose the **Filter Table** and **Sort Table** buttons, in either order.</span></span>

    ![Excel のチュートリアル - 表のフィルター処理と並べ替え](../images/excel-tutorial-filter-and-sort-table.png)
