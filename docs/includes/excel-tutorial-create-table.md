<span data-ttu-id="0c79d-101">チュートリアルのこの手順では、プログラムによってアドインがユーザーの Excel の現在のバージョンをサポートしているかどうかをテストし、ワークシートにテーブルを追加して、そのテーブルのデータ設定と書式設定を実行します。</span><span class="sxs-lookup"><span data-stu-id="0c79d-101">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Excel, add a table to a worksheet, populate the table with data, and format it.</span></span>

> [!NOTE]
> <span data-ttu-id="0c79d-102">このページでは、Excel のアドインのチュートリアルの個々 の手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="0c79d-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="0c79d-103">このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[Excel アドインのチュートリアル](../tutorials/excel-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。</span><span class="sxs-lookup"><span data-stu-id="0c79d-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="code-the-add-in"></a><span data-ttu-id="0c79d-104">アドインのコードを作成する</span><span class="sxs-lookup"><span data-stu-id="0c79d-104">Code the add-in</span></span>

1. <span data-ttu-id="0c79d-105">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="0c79d-105">Open the project in your code editor.</span></span>
2. <span data-ttu-id="0c79d-106">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="0c79d-106">Open the file index.html.</span></span>
3. <span data-ttu-id="0c79d-107">`TODO1` を次のマークアップに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="0c79d-107">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="create-table">Create Table</button>
    ```

4. <span data-ttu-id="0c79d-108">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="0c79d-108">Open the app.js file.</span></span>
5. <span data-ttu-id="0c79d-109">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="0c79d-109">Replace the `TODO1` with the following code.</span></span> <span data-ttu-id="0c79d-110">このコードでは、ユーザーの Excel のバージョンが、このチュートリアルのシリーズで使用する API をすべて含んでいるバージョンの Excel.js をサポートしているかどうかを調べます。</span><span class="sxs-lookup"><span data-stu-id="0c79d-110">This code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use.</span></span> <span data-ttu-id="0c79d-111">運用アドインでは、未サポートの API を呼び出す UI を非表示または無効化する条件ブロックの本体を使用してください。</span><span class="sxs-lookup"><span data-stu-id="0c79d-111">In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs.</span></span> <span data-ttu-id="0c79d-112">これにより、ユーザーは、そのユーザーの Excel のバージョンでサポートされているアドインの部分を使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="0c79d-112">This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }
    ```

6. <span data-ttu-id="0c79d-113">`TODO2` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="0c79d-113">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#create-table').click(createTable);
    ```

7. <span data-ttu-id="0c79d-114">`TODO3` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="0c79d-114">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="0c79d-115">次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="0c79d-115">Note the following:</span></span>
   - <span data-ttu-id="0c79d-116">Excel.js のビジネス ロジックは、`Excel.run` に渡される関数に追加します。</span><span class="sxs-lookup"><span data-stu-id="0c79d-116">Your Excel.js business logic will be added to the function that is passed to `Excel.run`.</span></span> <span data-ttu-id="0c79d-117">このロジックは、すぐには実行されません。</span><span class="sxs-lookup"><span data-stu-id="0c79d-117">This logic does not execute immediately.</span></span> <span data-ttu-id="0c79d-118">その代わりに、保留中のコマンドのキューに追加されます。</span><span class="sxs-lookup"><span data-stu-id="0c79d-118">Instead, it is added to a queue of pending commands.</span></span>
   - <span data-ttu-id="0c79d-119">`context.sync` メソッドは、キューに登録されたすべてのコマンドを実行するために Excel に送信します。</span><span class="sxs-lookup"><span data-stu-id="0c79d-119">The `context.sync` method sends all queued commands to Excel for execution.</span></span>
   - <span data-ttu-id="0c79d-120">`Excel.run` の後に `catch` ブロックを続けます。</span><span class="sxs-lookup"><span data-stu-id="0c79d-120">The `Excel.run` is followed by a `catch` block.</span></span> <span data-ttu-id="0c79d-121">これは、どのような場合にも当てはまるベスト プラクティスです。</span><span class="sxs-lookup"><span data-stu-id="0c79d-121">This is a best practice that you should always follow.</span></span> 

    ```js
    function createTable() {
        Excel.run(function (context) {

            // TODO4: Queue table creation logic here.

            // TODO5: Queue commands to populate the table with data.

            // TODO6: Queue commands to format the table.

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

8. <span data-ttu-id="0c79d-p106">`TODO4` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="0c79d-p106">Replace `TODO4` with the following code. Note:</span></span>
   - <span data-ttu-id="0c79d-124">このコードでは、ワークシートのテーブル コレクションの `add` メソッドを使用してテーブルを作成します。このコレクションは空であったとしても常に存在します。</span><span class="sxs-lookup"><span data-stu-id="0c79d-124">The code creates a table by using `add` method of a worksheet's table collection, which always exists even if it is empty.</span></span> <span data-ttu-id="0c79d-125">これは、Excel.js オブジェクトの標準的な作成方法です。</span><span class="sxs-lookup"><span data-stu-id="0c79d-125">This is the standard way that Excel.js objects are created.</span></span> <span data-ttu-id="0c79d-126">クラス コンストラクタ API は存在しません。Excel オブジェクトを作成するために、`new` 演算子は使用できません。</span><span class="sxs-lookup"><span data-stu-id="0c79d-126">There are no class constructor APIs, and you never use a `new` operator to create an Excel object.</span></span> <span data-ttu-id="0c79d-127">その代わりに、親コレクションにオブジェクトを追加します。</span><span class="sxs-lookup"><span data-stu-id="0c79d-127">Instead, you add to a parent collection object.</span></span>
   - <span data-ttu-id="0c79d-128">`add` メソッドの最初のパラメーターは、テーブルの先頭行のみの範囲です。そのテーブルで最終的に使用する全体の範囲ではありません。</span><span class="sxs-lookup"><span data-stu-id="0c79d-128">The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use.</span></span> <span data-ttu-id="0c79d-129">これは、アドインでデータ行を設定するときに (この後の手順で実行します)、既存の行のセルに値を書き込むのではなく、新しい行をテーブルに追加するためです。</span><span class="sxs-lookup"><span data-stu-id="0c79d-129">This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows.</span></span> <span data-ttu-id="0c79d-130">多くの場合、テーブルの作成時には、そのテーブルに含める行の数がわからないため、このパターンのほうが一般的になります。</span><span class="sxs-lookup"><span data-stu-id="0c79d-130">This is a more common pattern because the number of rows that a table will have is often not known when the table is created.</span></span>
   - <span data-ttu-id="0c79d-131">テーブルの名前は、ワークシート内だけでなくブック全体で一意にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="0c79d-131">Table names must be unique across the entire workbook, not just the worksheet.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

9. <span data-ttu-id="0c79d-p109">`TODO5` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="0c79d-p109">Replace `TODO5` with the following code. Note:</span></span>
   - <span data-ttu-id="0c79d-134">範囲に含まれるセルの値は、配列の配列で設定します。</span><span class="sxs-lookup"><span data-stu-id="0c79d-134">The cell values of a range are set with an array of arrays.</span></span>
   - <span data-ttu-id="0c79d-135">テーブル内に新しい行を作成するために、そのテーブルの行コレクションの `add` メソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="0c79d-135">New rows are created in a table by calling the `add` method of the table's row collection.</span></span> <span data-ttu-id="0c79d-136">`add` の 1 回の呼び出しで複数の行を追加できるようにするには、2 番目のパラメーターとして渡す親配列に複数のセル値の配列を含めます。</span><span class="sxs-lookup"><span data-stu-id="0c79d-136">You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.</span></span>

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

10. <span data-ttu-id="0c79d-p111">`TODO6` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="0c79d-p111">Replace `TODO6` with the following code. Note:</span></span>
   - <span data-ttu-id="0c79d-139">このコードでは、ゼロから始まるインデックスをテーブルの列コレクションの `getItemAt` メソッドに渡すことで、**Amount** 列への参照を取得します。</span><span class="sxs-lookup"><span data-stu-id="0c79d-139">The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection.</span></span>

     > [!NOTE]
     > <span data-ttu-id="0c79d-140">Excel.js のコレクション オブジェクト (`TableCollection`、`WorksheetCollection`、`TableColumnCollection` など) には、`items` プロパティがあります。このプロパティは、子オブジェクト タイプ (`Table`、`Worksheet`、`TableColumn` など) の配列ですが、`*Collection` オブジェクト自体は配列ではありません。</span><span class="sxs-lookup"><span data-stu-id="0c79d-140">Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.</span></span>

   - <span data-ttu-id="0c79d-141">その次に、コードでは、**Amount** 列の範囲を小数点以下 2 桁までのユーロとして書式設定します。</span><span class="sxs-lookup"><span data-stu-id="0c79d-141">The code then formats the range of the **Amount** column as Euros to the second decimal.</span></span> 
   - <span data-ttu-id="0c79d-142">最後に、列の幅と行の高さが最長 (最高) のデータ アイテムを収めるために十分な大きさになるようにしています。</span><span class="sxs-lookup"><span data-stu-id="0c79d-142">Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item.</span></span> <span data-ttu-id="0c79d-143">このコードでは、書式設定のために `Range` オブジェクトを取得している点に注目してください。</span><span class="sxs-lookup"><span data-stu-id="0c79d-143">Notice that the code must get `Range` objects to format.</span></span> <span data-ttu-id="0c79d-144">`TableColumn` オブジェクトと `TableRow` オブジェクトには、書式設定のプロパティがありません。</span><span class="sxs-lookup"><span data-stu-id="0c79d-144">`TableColumn` and `TableRow` objects do not have format properties.</span></span>

        ```js
        expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();
        ```

## <a name="test-the-add-in"></a><span data-ttu-id="0c79d-145">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="0c79d-145">Test the add-in</span></span>

1. <span data-ttu-id="0c79d-146">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="0c79d-146">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
2. <span data-ttu-id="0c79d-147">`npm run build` コマンドを実行して、ES6 ソース コードを Internet Explorer でサポートされている以前のバージョンの JavaScript にトランスパイルします (これは、Excel アドインを実行するために Excel の内部で使用されます)。</span><span class="sxs-lookup"><span data-stu-id="0c79d-147">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
3. <span data-ttu-id="0c79d-148">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="0c79d-148">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="0c79d-149">次のいずれかの方法を使用して、アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="0c79d-149">Sideload the add-in by using one of the following methods:</span></span>
    - <span data-ttu-id="0c79d-150">Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="0c79d-150">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="0c79d-151">Excel Online:[Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="0c79d-151">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="0c79d-152">iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="0c79d-152">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>
5. <span data-ttu-id="0c79d-153">**[ホーム]** メニューで、**[作業ウィンドウの表示]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="0c79d-153">On the **Home** menu, choose **Show Taskpane**.</span></span>
6. <span data-ttu-id="0c79d-154">作業ウィンドウで、**[Create Table]** (表の作成) を選択します。</span><span class="sxs-lookup"><span data-stu-id="0c79d-154">In the taskpane, choose **Create Table**.</span></span>

    ![Excel チュートリアル: テーブルの作成](../images/excel-tutorial-create-table.png)
