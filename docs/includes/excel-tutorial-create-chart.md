<span data-ttu-id="ed77c-101">チュートリアルのこの手順では、前の手順で作成したテーブルのデータを使用してグラフを作成して、そのグラフの書式を設定します。</span><span class="sxs-lookup"><span data-stu-id="ed77c-101">In this step of the tutorial, you'll create a chart using data from the table that you created previously, and then format the chart.</span></span>

> [!NOTE]
> <span data-ttu-id="ed77c-102">このページでは、Excel のアドインのチュートリアルの個々 の手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="ed77c-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="ed77c-103">このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[Excel アドインのチュートリアル](../tutorials/excel-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。</span><span class="sxs-lookup"><span data-stu-id="ed77c-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="chart-table-data"></a><span data-ttu-id="ed77c-104">テーブル データのグラフを作成する</span><span class="sxs-lookup"><span data-stu-id="ed77c-104">Chart table data</span></span>

1. <span data-ttu-id="ed77c-105">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="ed77c-105">Open the project in your code editor.</span></span>
2. <span data-ttu-id="ed77c-106">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="ed77c-106">Open the file index.html.</span></span>
3. <span data-ttu-id="ed77c-107">`sort-table` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="ed77c-107">Below the `div` that contains the `sort-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="create-chart">Create Chart</button>
    </div>
    ```

4. <span data-ttu-id="ed77c-108">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="ed77c-108">Open the app.js file.</span></span>

5. <span data-ttu-id="ed77c-109">`sort-chart` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="ed77c-109">Below the line that assigns a click handler to the `sort-chart` button, add the following code:</span></span>

    ```js
    $('#create-chart').click(createChart);
    ```

6. <span data-ttu-id="ed77c-110">`sortTable` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="ed77c-110">Below the `sortTable` function add the following function.</span></span>

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

7. <span data-ttu-id="ed77c-p102">`TODO1` を次のコードに置き換えます。ヘッダー行を除外するために、このコードでは、`getRange` メソッドではなく `Table.getDataBodyRange` メソッドを使用してグラフを作成するデータの範囲を取得しています。</span><span class="sxs-lookup"><span data-stu-id="ed77c-p102">Replace `TODO1` with the following code. Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const dataRange = expensesTable.getDataBodyRange();
    ```

8. <span data-ttu-id="ed77c-113">`TODO2` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="ed77c-113">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="ed77c-114">次のパラメーターに注意してください。</span><span class="sxs-lookup"><span data-stu-id="ed77c-114">Note the following parameters:</span></span>
   - <span data-ttu-id="ed77c-p104">`add` への最初のパラメーターでは、グラフの種類を指定します。数十種類あります。</span><span class="sxs-lookup"><span data-stu-id="ed77c-p104">The first parameter to the `add` method specifies the type of chart. There are several dozen types.</span></span>
   - <span data-ttu-id="ed77c-117">2 番目のパラメーターでは、グラフに含めるデータの範囲を指定します。</span><span class="sxs-lookup"><span data-stu-id="ed77c-117">The second parameter specifies the range of data to include in the chart.</span></span>
   - <span data-ttu-id="ed77c-118">3 番目のパラメーターでは、テーブルからの一連のデータ ポイントを行方向と列方向のどちらでグラフ化する必要があるかを決定します。</span><span class="sxs-lookup"><span data-stu-id="ed77c-118">The third parameter determines whether a series of data points from the table should be charted rowwise or columnwise.</span></span> <span data-ttu-id="ed77c-119">オプション `auto` は、最適な方法を判断するように Excel に指示します。</span><span class="sxs-lookup"><span data-stu-id="ed77c-119">The option `auto` tells Excel to decide the best method.</span></span>

    ```js
    let chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ```

9. <span data-ttu-id="ed77c-120">`TODO3` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="ed77c-120">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="ed77c-121">このコードのほとんどの部分は、わかりやすく説明不要なものです。</span><span class="sxs-lookup"><span data-stu-id="ed77c-121">Most of this code is self-explanatory.</span></span> <span data-ttu-id="ed77c-122">注意:</span><span class="sxs-lookup"><span data-stu-id="ed77c-122">Note:</span></span>
   - <span data-ttu-id="ed77c-123">`setPosition` メソッドへのパラメーターでは、グラフを収容するワークシート領域の左上と右下のセルを指定します。</span><span class="sxs-lookup"><span data-stu-id="ed77c-123">The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart.</span></span> <span data-ttu-id="ed77c-124">Excel では、所定の空間内でグラフの外観を整えるために線幅などを調整できます。</span><span class="sxs-lookup"><span data-stu-id="ed77c-124">Excel can adjust things like line width to make the chart look good in the space it has been given.</span></span>
   - <span data-ttu-id="ed77c-125">"series" は、テーブルに含まれる列からのデータ ポイントのセットです。</span><span class="sxs-lookup"><span data-stu-id="ed77c-125">A "series" is a set of data points from a column of the table.</span></span> <span data-ttu-id="ed77c-126">このテーブルに存在する文字列以外の列は 1 列のみであるため、Excel は、その列がグラフ化するデータ ポイントの唯一の列であることを推測します。</span><span class="sxs-lookup"><span data-stu-id="ed77c-126">Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart.</span></span> <span data-ttu-id="ed77c-127">その他の列は、グラフのラベルとして解釈されます。</span><span class="sxs-lookup"><span data-stu-id="ed77c-127">It interprets the other columns as chart labels.</span></span> <span data-ttu-id="ed77c-128">そのため、グラフの series は 1 つ存在することになり、インデックス 0 を含みます。</span><span class="sxs-lookup"><span data-stu-id="ed77c-128">So there will be just one series in the chart and it will have index 0.</span></span> <span data-ttu-id="ed77c-129">これに、"Value in €" のラベルを付けます。</span><span class="sxs-lookup"><span data-stu-id="ed77c-129">This is the one to label with "Value in €".</span></span>

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in €';
    ```

## <a name="test-the-add-in"></a><span data-ttu-id="ed77c-130">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="ed77c-130">Test the add-in</span></span>


1. <span data-ttu-id="ed77c-131">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、Ctrl-C を 2 回入力して実行中の Web サーバーを停止します。</span><span class="sxs-lookup"><span data-stu-id="ed77c-131">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="ed77c-132">それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="ed77c-132">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="ed77c-133">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。</span><span class="sxs-lookup"><span data-stu-id="ed77c-133">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="ed77c-134">そのためには、ビルド コマンドの入力を求めるプロンプトが表示されるように、サーバー プロセスを強制終了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ed77c-134">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="ed77c-135">ビルド後に、サーバーを再起動します。</span><span class="sxs-lookup"><span data-stu-id="ed77c-135">After the build, you restart the server.</span></span> <span data-ttu-id="ed77c-136">次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="ed77c-136">The next few steps carry out this process.</span></span>

1. <span data-ttu-id="ed77c-137">`npm run build` コマンドを実行して、ES6 ソース コードを Internet Explorer でサポートされている以前のバージョンの JavaScript にトランスパイルします (これは、Excel アドインを実行するために Excel の内部で使用されます)。</span><span class="sxs-lookup"><span data-stu-id="ed77c-137">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
2. <span data-ttu-id="ed77c-138">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="ed77c-138">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="ed77c-139">作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="ed77c-139">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="ed77c-140">何らかの理由から開いているワークシートにテーブルが含まれていない場合は、**[Create Table]** (テーブルの作成) ボタンをクリックしてから、**[Filter Table]** (テーブルのフィルター) ボタンと **[Sort Table]** (テーブルの並べ替え) ボタンを任意の順序でクリックします。</span><span class="sxs-lookup"><span data-stu-id="ed77c-140">If for any reason the table is not in the open worksheet, in the taskpane, choose **Create Table** and then **Filter Table** and **Sort Table** buttons, in either order.</span></span>
6. <span data-ttu-id="ed77c-141">**[グラフの作成]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="ed77c-141">Choose the **Create Chart** button.</span></span> <span data-ttu-id="ed77c-142">グラフが作成され、フィルターが適用された行からのデータのみが含まれます。</span><span class="sxs-lookup"><span data-stu-id="ed77c-142">A chart is created and only the data from the rows that have been filtered are included.</span></span> <span data-ttu-id="ed77c-143">データ ポイントの下側のラベルは、グラフの並べ替え順序になります。つまり、[Merchant] (業者) の名前の逆アルファベット順になります。</span><span class="sxs-lookup"><span data-stu-id="ed77c-143">The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.</span></span>

    ![Excel チュートリアル: グラフの作成](../images/excel-tutorial-create-chart.png)
