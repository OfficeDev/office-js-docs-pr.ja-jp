---
title: Excel アドインのチュートリアル
description: このチュートリアルでは、Excel アドインを構築します。このアドインでは、テーブルの作成、表示、フィルター処理、並べ替えを行うことができ、グラフの作成、テーブルのヘッダーの固定、ワークシートの保護も可能となります。また、ダイアログを開くこともできます。
ms.date: 06/20/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 9efbd1380587244fae60551fe104f859d22b4aa2
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771885"
---
# <a name="tutorial-create-an-excel-task-pane-add-in"></a><span data-ttu-id="e54db-103">チュートリアル: Excel 作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="e54db-103">Tutorial: Create an Excel task pane add-in</span></span>

<span data-ttu-id="e54db-104">このチュートリアルでは、以下を実行する Excel 作業ウィンドウ アドインを作成します。</span><span class="sxs-lookup"><span data-stu-id="e54db-104">In this tutorial, you'll create an Excel task pane add-in that:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="e54db-105">テーブルの作成</span><span class="sxs-lookup"><span data-stu-id="e54db-105">Creates a table</span></span>
> * <span data-ttu-id="e54db-106">テーブルのフィルター処理と並べ替え</span><span class="sxs-lookup"><span data-stu-id="e54db-106">Filters and sorts a table</span></span>
> * <span data-ttu-id="e54db-107">グラフの作成</span><span class="sxs-lookup"><span data-stu-id="e54db-107">Creates a chart</span></span>
> * <span data-ttu-id="e54db-108">テーブルのヘッダーの固定</span><span class="sxs-lookup"><span data-stu-id="e54db-108">Freezes a table header</span></span>
> * <span data-ttu-id="e54db-109">ワークシートの保護</span><span class="sxs-lookup"><span data-stu-id="e54db-109">Protects a worksheet</span></span>
> * <span data-ttu-id="e54db-110">ダイアログを開く</span><span class="sxs-lookup"><span data-stu-id="e54db-110">Opens a dialog</span></span>

## <a name="prerequisites"></a><span data-ttu-id="e54db-111">前提条件</span><span class="sxs-lookup"><span data-stu-id="e54db-111">Prerequisites</span></span>

<span data-ttu-id="e54db-112">このチュートリアルを使用するには、以下のバージョンがインストールされている必要があります。</span><span class="sxs-lookup"><span data-stu-id="e54db-112">To use this tutorial, you need to have the following installed.</span></span> 

- <span data-ttu-id="e54db-p101">Excel 2016、バージョン 1711 (ビルド 8730.1000 クイック実行) 以降。 このバージョンを入手するには、Office Insider への参加が必要になることがあります。 詳細については、「[Office Insider](https://products.office.com/office-insider?tab=tab-1)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e54db-p101">Excel 2016, version 1711 (Build 8730.1000 Click-to-Run) or later. You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>

- [<span data-ttu-id="e54db-116">ノード</span><span class="sxs-lookup"><span data-stu-id="e54db-116">Node</span></span>](https://nodejs.org/en/) 

- <span data-ttu-id="e54db-117">[Git バッシュ](https://git-scm.com/downloads) (または別の Git クライアント)</span><span class="sxs-lookup"><span data-stu-id="e54db-117">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

- <span data-ttu-id="e54db-118">このチュートリアルでアドインをテストするには、インターネット接続が必要です。</span><span class="sxs-lookup"><span data-stu-id="e54db-118">You need to have an Internet connection to test the add-in in this tutorial.</span></span>

## <a name="create-your-add-in-project"></a><span data-ttu-id="e54db-119">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="e54db-119">Create your add-in project</span></span>

<span data-ttu-id="e54db-120">このチュートリアルの基礎として使用する Excel アドイン プロジェクトを作成するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="e54db-120">Complete the following steps to create the Excel add-in project that you'll use as the basis for this tutorial.</span></span>

1. <span data-ttu-id="e54db-121">「[Excel アドインのチュートリアル](https://github.com/OfficeDev/Excel-Add-in-Tutorial)」で、GitHub リポジトリを複製します。</span><span class="sxs-lookup"><span data-stu-id="e54db-121">Clone the GitHub repository [Excel add-in tutorial](https://github.com/OfficeDev/Excel-Add-in-Tutorial).</span></span>

2. <span data-ttu-id="e54db-122">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="e54db-122">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

3. <span data-ttu-id="e54db-123">`npm install` コマンドを実行して、package.json ファイルに一覧表示されているツールとライブラリをインストールします。</span><span class="sxs-lookup"><span data-stu-id="e54db-123">Run the command `npm install` to install the tools and libraries listed in the package.json file.</span></span> 

4. <span data-ttu-id="e54db-124">開発用のコンピューターのオペレーティングシステムの証明書を信頼するように、[自己署名証明書をインストール](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)する手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="e54db-124">Carry out the steps in [Installing the self-signed certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

## <a name="create-a-table"></a><span data-ttu-id="e54db-125">テーブルの作成</span><span class="sxs-lookup"><span data-stu-id="e54db-125">Create a table</span></span>

<span data-ttu-id="e54db-126">チュートリアルのこの手順では、プログラムによってアドインがユーザーの Excel の現在のバージョンをサポートしているかどうかをテストし、ワークシートにテーブルを追加して、そのテーブルのデータ設定と書式設定を実行します。</span><span class="sxs-lookup"><span data-stu-id="e54db-126">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Excel, add a table to a worksheet, populate the table with data, and format it.</span></span>

### <a name="code-the-add-in"></a><span data-ttu-id="e54db-127">アドインのコードを作成する</span><span class="sxs-lookup"><span data-stu-id="e54db-127">Code the add-in</span></span>

1. <span data-ttu-id="e54db-128">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-128">Open the project in your code editor.</span></span>

2. <span data-ttu-id="e54db-129">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-129">Open the file index.html.</span></span>

3. <span data-ttu-id="e54db-130">`TODO1` を次のマークアップに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e54db-130">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="create-table">Create Table</button>
    ```

4. <span data-ttu-id="e54db-131">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-131">Open the app.js file.</span></span>

5. <span data-ttu-id="e54db-p102">`TODO1` を次のコードに置き換えます。 このコードでは、ユーザーの Excel のバージョンが、このチュートリアルのシリーズで使用する API をすべて含んでいるバージョンの Excel.js をサポートしているかどうかを調べます。 運用アドインでは、未サポートの API を呼び出す UI を非表示または無効化する条件ブロックの本体を使用してください。 これにより、ユーザーは、そのユーザーの Excel のバージョンでサポートされているアドインの部分を使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="e54db-p102">Replace the `TODO1` with the following code. This code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use. In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs. This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }
    ```

6. <span data-ttu-id="e54db-136">`TODO2` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e54db-136">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#create-table').click(createTable);
    ```

7. <span data-ttu-id="e54db-137">`TODO3` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e54db-137">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="e54db-138">注:</span><span class="sxs-lookup"><span data-stu-id="e54db-138">Note:</span></span>

   - <span data-ttu-id="e54db-p104">Excel.js のビジネス ロジックは、`Excel.run` に渡される関数に追加します。 このロジックは、すぐには実行されません。 その代わりに、保留中のコマンドのキューに追加されます。</span><span class="sxs-lookup"><span data-stu-id="e54db-p104">Your Excel.js business logic will be added to the function that is passed to `Excel.run`. This logic does not execute immediately. Instead, it is added to a queue of pending commands.</span></span>

   - <span data-ttu-id="e54db-142">`context.sync` メソッドは、キューに登録されたすべてのコマンドを実行するために Excel に送信します。</span><span class="sxs-lookup"><span data-stu-id="e54db-142">The `context.sync` method sends all queued commands to Excel for execution.</span></span>

   - <span data-ttu-id="e54db-p105">`Excel.run` の後に `catch` ブロックを続けます。 これは、どのような場合にも当てはまるベスト プラクティスです。</span><span class="sxs-lookup"><span data-stu-id="e54db-p105">The `Excel.run` is followed by a `catch` block. This is a best practice that you should always follow.</span></span> 

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

8. <span data-ttu-id="e54db-145">`TODO4` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e54db-145">Replace `TODO4` with the following code.</span></span> <span data-ttu-id="e54db-146">次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e54db-146">Note:</span></span>

   - <span data-ttu-id="e54db-p107">このコードでは、ワークシートのテーブル コレクションの `add` メソッドを使用してテーブルを作成します。このコレクションは空であったとしても常に存在します。 これは、Excel.js オブジェクトの標準的な作成方法です。 クラス コンストラクタ API は存在しません。Excel オブジェクトを作成するために、`new` 演算子は使用できません。 その代わりに、親コレクションにオブジェクトを追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-p107">The code creates a table by using `add` method of a worksheet's table collection, which always exists even if it is empty. This is the standard way that Excel.js objects are created. There are no class constructor APIs, and you never use a `new` operator to create an Excel object. Instead, you add to a parent collection object.</span></span>

   - <span data-ttu-id="e54db-p108">`add` メソッドの最初のパラメーターは、テーブルの先頭行のみの範囲です。そのテーブルで最終的に使用する全体の範囲ではありません。 これは、アドインでデータ行を設定するときに (この後の手順で実行します)、既存の行のセルに値を書き込むのではなく、新しい行をテーブルに追加するためです。 多くの場合、テーブルの作成時には、そのテーブルに含める行の数がわからないため、このパターンのほうが一般的になります。</span><span class="sxs-lookup"><span data-stu-id="e54db-p108">The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use. This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows. This is a more common pattern because the number of rows that a table will have is often not known when the table is created.</span></span>

   - <span data-ttu-id="e54db-154">テーブルの名前は、ワークシート内だけでなくブック全体で一意にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="e54db-154">Table names must be unique across the entire workbook, not just the worksheet.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

9. <span data-ttu-id="e54db-155">`TODO5` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e54db-155">Replace `TODO5` with the following code.</span></span> <span data-ttu-id="e54db-156">次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e54db-156">Note:</span></span>

   - <span data-ttu-id="e54db-157">範囲に含まれるセルの値は、配列の配列で設定します。</span><span class="sxs-lookup"><span data-stu-id="e54db-157">The cell values of a range are set with an array of arrays.</span></span>

   - <span data-ttu-id="e54db-p110">テーブル内に新しい行を作成するために、そのテーブルの行コレクションの `add` メソッドを呼び出します。 `add` の 1 回の呼び出しで複数の行を追加できるようにするには、2 番目のパラメーターとして渡す親配列に複数のセル値の配列を含めます。</span><span class="sxs-lookup"><span data-stu-id="e54db-p110">New rows are created in a table by calling the `add` method of the table's row collection. You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.</span></span>

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

10. <span data-ttu-id="e54db-160">`TODO6` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e54db-160">Replace `TODO6` with the following code.</span></span> <span data-ttu-id="e54db-161">次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e54db-161">Note:</span></span>

   - <span data-ttu-id="e54db-162">このコードでは、ゼロから始まるインデックスをテーブルの列コレクションの \*\*\*\* メソッドに渡すことで、`getItemAt` 列への参照を取得します。</span><span class="sxs-lookup"><span data-stu-id="e54db-162">The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection.</span></span>

     > [!NOTE]
     > <span data-ttu-id="e54db-163">Excel.js のコレクション オブジェクト (`TableCollection`、`WorksheetCollection`、`TableColumnCollection` など) には、`items` プロパティがあります。このプロパティは、子オブジェクト タイプ (`Table`、`Worksheet`、`TableColumn` など) の配列ですが、`*Collection` オブジェクト自体は配列ではありません。</span><span class="sxs-lookup"><span data-stu-id="e54db-163">Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.</span></span>

   - <span data-ttu-id="e54db-164">その次に、コードでは、**Amount** 列の範囲を小数点以下 2 桁までのユーロとして書式設定します。</span><span class="sxs-lookup"><span data-stu-id="e54db-164">The code then formats the range of the **Amount** column as Euros to the second decimal.</span></span> 

   - <span data-ttu-id="e54db-p112">最後に、列の幅と行の高さが最長 (最高) のデータ アイテムを収めるために十分な大きさになるようにしています。 このコードでは、書式設定のために `Range` オブジェクトを取得している点に注目してください。 `TableColumn` オブジェクトと `TableRow` オブジェクトには、書式設定のプロパティがありません。</span><span class="sxs-lookup"><span data-stu-id="e54db-p112">Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item. Notice that the code must get `Range` objects to format. `TableColumn` and `TableRow` objects do not have format properties.</span></span>

        ```js
        expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();
        ```

### <a name="test-the-add-in"></a><span data-ttu-id="e54db-168">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="e54db-168">Test the add-in</span></span>

1. <span data-ttu-id="e54db-169">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="e54db-169">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

2. <span data-ttu-id="e54db-170">コマンド`npm run build`を実行して、Internet Explorer でサポートされている以前のバージョンの JAVASCRIPT に ES6 のソースコードを transpile します (excel の一部のバージョンでは、excel アドインを実行するために使用されます)。</span><span class="sxs-lookup"><span data-stu-id="e54db-170">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used by some versions of Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="e54db-171">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="e54db-171">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="e54db-172">次のいずれかの方法を使用して、アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="e54db-172">Sideload the add-in by using one of the following methods:</span></span>

    - <span data-ttu-id="e54db-173">Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="e54db-173">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>

    - <span data-ttu-id="e54db-174">Web ブラウザー:[サイドロード Office アドイン (web)](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="e54db-174">Web browser: [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span></span>

    - <span data-ttu-id="e54db-175">iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="e54db-175">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

5. <span data-ttu-id="e54db-176">**[ホーム]** メニューで、**[作業ウィンドウの表示]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="e54db-176">On the **Home** menu, choose **Show Taskpane**.</span></span>

6. <span data-ttu-id="e54db-177">作業ウィンドウで、**[Create Table]** (表の作成) を選択します。</span><span class="sxs-lookup"><span data-stu-id="e54db-177">In the task pane, choose **Create Table**.</span></span>

    ![Excel チュートリアル - テーブルの作成](../images/excel-tutorial-create-table.png)

## <a name="filter-and-sort-a-table"></a><span data-ttu-id="e54db-179">テーブルのフィルター処理と並べ替え</span><span class="sxs-lookup"><span data-stu-id="e54db-179">Filter and sort a table</span></span>

<span data-ttu-id="e54db-180">チュートリアルのこの手順では、以前に作成したテーブルをフィルター処理したり並べ替えたりします。</span><span class="sxs-lookup"><span data-stu-id="e54db-180">In this step of the tutorial, you'll filter and sort the table that you created previously.</span></span>

### <a name="filter-the-table"></a><span data-ttu-id="e54db-181">表のフィルター処理</span><span class="sxs-lookup"><span data-stu-id="e54db-181">Filter the table</span></span>

1. <span data-ttu-id="e54db-182">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-182">Open the project in your code editor.</span></span>

2. <span data-ttu-id="e54db-183">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-183">Open the file index.html.</span></span>

3. <span data-ttu-id="e54db-184">`create-table` ボタンを格納している `div` の直下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-184">Just below the `div` that contains the `create-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="filter-table">Filter Table</button>
    </div>
    ```

4. <span data-ttu-id="e54db-185">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-185">Open the app.js file.</span></span>

5. <span data-ttu-id="e54db-186">`create-table` ボタンにクリック ハンドラーを割り当てる行の直下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-186">Just below the line that assigns a click handler to the `create-table` button, add the following code:</span></span>

    ```js
    $('#filter-table').click(filterTable);
    ```

6. <span data-ttu-id="e54db-187">`createTable` 関数の直下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-187">Just below the `createTable` function, add the following function:</span></span>

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

7. <span data-ttu-id="e54db-188">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e54db-188">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="e54db-189">次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e54db-189">Note:</span></span>

   - <span data-ttu-id="e54db-p114">このコードでは最初に、`getItem` メソッドに列名を渡すことによって、フィルター処理が必要な列への参照を取得します。`getItemAt` メソッドが行うように、列のインデックスを `createTable` メソッドに渡すわけではありません。 ユーザーは表の列を移動させることができるので、表を作成した後、指定したインデックスにある列が変わってしまう可能性があります。 そのため、列名を使用して列への参照を取得するほうが安全です。 前のチュートリアルでは、表を作成するのとまったく同じ方法で `getItemAt` を使用したため、ユーザーが列を移動させた可能性はなく、よって安全に使用できました。</span><span class="sxs-lookup"><span data-stu-id="e54db-p114">The code first gets a reference to the column that needs filtering by passing the column name to the `getItem` method, instead of passing its index to the `getItemAt` method as the `createTable` method does. Since users can move table columns, the column at a given index might change after the table is created. Hence, it is safer to use the column name to get a reference to the column. We used `getItemAt` safely in the preceding tutorial, because we used it in the very same method that creates the table, so there is no chance that a user has moved the column.</span></span>

   - <span data-ttu-id="e54db-194">`applyValuesFilter` メソッドは、`Filter` オブジェクトのフィルター処理方法の 1 つです。</span><span class="sxs-lookup"><span data-stu-id="e54db-194">The `applyValuesFilter` method is one of several filtering methods on the `Filter` object.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);
    ``` 

### <a name="sort-the-table"></a><span data-ttu-id="e54db-195">表の並べ替え</span><span class="sxs-lookup"><span data-stu-id="e54db-195">Sort the table</span></span>

1. <span data-ttu-id="e54db-196">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-196">Open the file index.html.</span></span>

2. <span data-ttu-id="e54db-197">`filter-table` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-197">Below the `div` that contains the `filter-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="sort-table">Sort Table</button>
    </div>
    ```

3. <span data-ttu-id="e54db-198">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-198">Open the app.js file.</span></span>

4. <span data-ttu-id="e54db-199">`filter-table` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-199">Below the line that assigns a click handler to the `filter-table` button, add the following code:</span></span>

    ```js
    $('#sort-table').click(sortTable);
    ```

5. <span data-ttu-id="e54db-200">`filterTable` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-200">Below the `filterTable` function add the following function.</span></span>

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

6. <span data-ttu-id="e54db-201">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e54db-201">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="e54db-202">次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e54db-202">Note:</span></span>

   - <span data-ttu-id="e54db-203">アドインで並べ替えるのは Merchant 列のみであるため、このコードでは、1 つのメンバーだけを含む `SortField` オブジェクトの配列を作成します。</span><span class="sxs-lookup"><span data-stu-id="e54db-203">The code creates an array of `SortField` objects which has just one member since the add-in only sorts on the Merchant column.</span></span>

   - <span data-ttu-id="e54db-204">`key` オブジェクトの `SortField` プロパティは、並べ替える対象列の 0 から始まるインデックスです。</span><span class="sxs-lookup"><span data-stu-id="e54db-204">The `key` property of a `SortField` object is the zero-based index of the column to sort-on.</span></span>

   - <span data-ttu-id="e54db-205">`Table` の `sort` メンバーは、`TableSort` オブジェクトであり、メソッドではありません。</span><span class="sxs-lookup"><span data-stu-id="e54db-205">The `sort` member of a `Table` is a `TableSort` object, not a method.</span></span> <span data-ttu-id="e54db-206">`TableSort` オブジェクトの `apply` メソッドには、`SortField` が渡されます。</span><span class="sxs-lookup"><span data-stu-id="e54db-206">The `SortField`s are passed to the `TableSort` object's `apply` method.</span></span>

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

### <a name="test-the-add-in"></a><span data-ttu-id="e54db-207">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="e54db-207">Test the add-in</span></span>

1. <span data-ttu-id="e54db-208">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、**Ctrl + C** を 2 回入力して実行中の Web サーバーを停止します。</span><span class="sxs-lookup"><span data-stu-id="e54db-208">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter **Ctrl+C** twice to stop the running web server.</span></span> <span data-ttu-id="e54db-209">それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="e54db-209">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="e54db-210">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。</span><span class="sxs-lookup"><span data-stu-id="e54db-210">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="e54db-211">そのためには、ビルド コマンドの入力を求めるプロンプトが表示されるように、サーバー プロセスを強制終了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e54db-211">In order to do this, you need to kill the server process so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="e54db-212">ビルド後に、サーバーを再起動します。</span><span class="sxs-lookup"><span data-stu-id="e54db-212">After the build, you restart the server.</span></span> <span data-ttu-id="e54db-213">次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="e54db-213">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="e54db-214">コマンド`npm run build`を実行して、Internet Explorer でサポートされている以前のバージョンの JAVASCRIPT に ES6 のソースコードを transpile します (excel の一部のバージョンでは、excel アドインを実行するために使用されます)。</span><span class="sxs-lookup"><span data-stu-id="e54db-214">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used by some versions of Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="e54db-215">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="e54db-215">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="e54db-216">作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-216">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="e54db-217">何らかの理由から開いているワークシートに表が含まれていない場合は、作業ウィンドウの **[Create Table]** (表の作成) ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="e54db-217">If for any reason the table is not in the open worksheet, in the task pane, choose **Create Table**.</span></span>

6. <span data-ttu-id="e54db-218">**[Filter Table]** (表のフィルター) ボタンと **[Sort Table]** (表の並べ替え) ボタンを任意の順序で選択します。</span><span class="sxs-lookup"><span data-stu-id="e54db-218">Choose the **Filter Table** and **Sort Table** buttons, in either order.</span></span>

    ![Excel のチュートリアル - テーブルのフィルター処理と並べ替え](../images/excel-tutorial-filter-and-sort-table.png)

## <a name="create-a-chart"></a><span data-ttu-id="e54db-220">グラフの作成</span><span class="sxs-lookup"><span data-stu-id="e54db-220">Create a chart</span></span>

<span data-ttu-id="e54db-221">チュートリアルのこの手順では、前の手順で作成したテーブルのデータを使用してグラフを作成して、そのグラフの書式を設定します。</span><span class="sxs-lookup"><span data-stu-id="e54db-221">In this step of the tutorial, you'll create a chart using data from the table that you created previously, and then format the chart.</span></span>

### <a name="chart-a-chart-using-table-data"></a><span data-ttu-id="e54db-222">テーブルのデータを使用してグラフを作成する</span><span class="sxs-lookup"><span data-stu-id="e54db-222">Chart a chart using table data</span></span>

1. <span data-ttu-id="e54db-223">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-223">Open the project in your code editor.</span></span>

2. <span data-ttu-id="e54db-224">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-224">Open the file index.html.</span></span>

3. <span data-ttu-id="e54db-225">`sort-table` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-225">Below the `div` that contains the `sort-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="create-chart">Create Chart</button>
    </div>
    ```

4. <span data-ttu-id="e54db-226">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-226">Open the app.js file.</span></span>

5. <span data-ttu-id="e54db-227">`sort-chart` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-227">Below the line that assigns a click handler to the `sort-chart` button, add the following code:</span></span>

    ```js
    $('#create-chart').click(createChart);
    ```

6. <span data-ttu-id="e54db-228">`sortTable` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-228">Below the `sortTable` function add the following function.</span></span>

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

7. <span data-ttu-id="e54db-p119">`TODO1` を次のコードに置き換えます。ヘッダー行を除外するために、このコードでは、`Table.getDataBodyRange` メソッドではなく `getRange` メソッドを使用してグラフを作成するデータの範囲を取得しています。</span><span class="sxs-lookup"><span data-stu-id="e54db-p119">Replace `TODO1` with the following code. Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var dataRange = expensesTable.getDataBodyRange();
    ```

8. <span data-ttu-id="e54db-p120">`TODO2` を次のコードに置き換えます。 次のパラメーターに注意してください。</span><span class="sxs-lookup"><span data-stu-id="e54db-p120">Replace `TODO2` with the following code. Note the following parameters:</span></span>

   - <span data-ttu-id="e54db-p121">`add` への最初のパラメーターでは、グラフの種類を指定します。数十種類あります。</span><span class="sxs-lookup"><span data-stu-id="e54db-p121">The first parameter to the `add` method specifies the type of chart. There are several dozen types.</span></span>

   - <span data-ttu-id="e54db-235">2 番目のパラメーターでは、グラフに含めるデータの範囲を指定します。</span><span class="sxs-lookup"><span data-stu-id="e54db-235">The second parameter specifies the range of data to include in the chart.</span></span>

   - <span data-ttu-id="e54db-236">3 番目のパラメーターでは、テーブルからの一連のデータ ポイントを行方向と列方向のどちらでグラフ化する必要があるかを決定します。</span><span class="sxs-lookup"><span data-stu-id="e54db-236">The third parameter determines whether a series of data points from the table should be charted row-wise or column-wise.</span></span> <span data-ttu-id="e54db-237">オプション `auto` は、最適な方法を判断するように Excel に指示します。</span><span class="sxs-lookup"><span data-stu-id="e54db-237">The option `auto` tells Excel to decide the best method.</span></span>

    ```js
    var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ```

9. <span data-ttu-id="e54db-p123">`TODO3` を次のコードに置き換えます。 このコードのほとんどの部分は、わかりやすく説明不要なものです。 注意:</span><span class="sxs-lookup"><span data-stu-id="e54db-p123">Replace `TODO3` with the following code. Most of this code is self-explanatory. Note:</span></span>
   
   - <span data-ttu-id="e54db-p124">`setPosition` メソッドへのパラメーターでは、グラフを収容するワークシート領域の左上と右下のセルを指定します。 Excel では、所定の空間内でグラフの外観を整えるために線幅などを調整できます。</span><span class="sxs-lookup"><span data-stu-id="e54db-p124">The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart. Excel can adjust things like line width to make the chart look good in the space it has been given.</span></span>
   
   - <span data-ttu-id="e54db-p125">"series" は、テーブルに含まれる列からのデータ ポイントのセットです。 このテーブルに存在する文字列以外の列は 1 列のみであるため、Excel は、その列がグラフ化するデータ ポイントの唯一の列であることを推測します。 その他の列は、グラフのラベルとして解釈されます。 そのため、グラフの series は 1 つ存在することになり、インデックス 0 を含みます。 これに、"Value in €" のラベルを付けます。</span><span class="sxs-lookup"><span data-stu-id="e54db-p125">A "series" is a set of data points from a column of the table. Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart. It interprets the other columns as chart labels. So there will be just one series in the chart and it will have index 0. This is the one to label with "Value in €".</span></span>

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in €';
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="e54db-248">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="e54db-248">Test the add-in</span></span>

1. <span data-ttu-id="e54db-249">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、**Ctrl + C** を 2 回入力して実行中の Web サーバーを停止します。</span><span class="sxs-lookup"><span data-stu-id="e54db-249">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter **Ctrl+C** twice to stop the running web server.</span></span> <span data-ttu-id="e54db-250">それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="e54db-250">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="e54db-p127">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。 そのためには、ビルド コマンドの入力を求めるプロンプトが表示されるように、サーバー プロセスを強制終了する必要があります。 ビルド後に、サーバーを再起動します。 次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="e54db-p127">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command. After the build, you restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="e54db-255">コマンド`npm run build`を実行して、Internet Explorer でサポートされている以前のバージョンの JAVASCRIPT に ES6 のソースコードを transpile します (excel の一部のバージョンでは、excel アドインを実行するために使用されます)。</span><span class="sxs-lookup"><span data-stu-id="e54db-255">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used by some versions of Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="e54db-256">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="e54db-256">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="e54db-257">作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-257">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="e54db-258">何らかの理由から開いているワークシートにテーブルが含まれていない場合は、**[Create Table]** (テーブルの作成) ボタンをクリックしてから、**[Filter Table]** (テーブルのフィルター) ボタンと **[Sort Table]** (テーブルの並べ替え) ボタンを任意の順序でクリックします。</span><span class="sxs-lookup"><span data-stu-id="e54db-258">If for any reason the table is not in the open worksheet, in the task pane, choose **Create Table** and then **Filter Table** and **Sort Table** buttons, in either order.</span></span>

6. <span data-ttu-id="e54db-p128">**[グラフの作成]** ボタンをクリックします。 グラフが作成され、フィルターが適用された行からのデータのみが含まれます。 データ ポイントの下側のラベルは、グラフの並べ替え順序になります。つまり、[Merchant] (業者) の名前の逆アルファベット順になります。</span><span class="sxs-lookup"><span data-stu-id="e54db-p128">Choose the **Create Chart** button. A chart is created and only the data from the rows that have been filtered are included. The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.</span></span>

    ![Excel チュートリアル - グラフの作成](../images/excel-tutorial-create-chart.png)

## <a name="freeze-a-table-header"></a><span data-ttu-id="e54db-263">テーブルのヘッダーの固定</span><span class="sxs-lookup"><span data-stu-id="e54db-263">Freeze a table header</span></span>

<span data-ttu-id="e54db-p129">表がとても長く、行を参照するためにスクロールしなければならない場合、ヘッダー行が画面の外に移動して見えなくなることがあります。 チュートリアルのこの手順では、以前に作成した表のヘッダー行を固定して、ワークシートを下にスクロールしても表示されるようにします。</span><span class="sxs-lookup"><span data-stu-id="e54db-p129">When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight. In this step of the tutorial, you'll freeze the header row of the table that you created previously, so that it remains visible even as the user scrolls down the worksheet.</span></span>

### <a name="freeze-the-tables-header-row"></a><span data-ttu-id="e54db-266">表のヘッダー行を固定する</span><span class="sxs-lookup"><span data-stu-id="e54db-266">Freeze the table's header row</span></span>

1. <span data-ttu-id="e54db-267">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-267">Open the project in your code editor.</span></span>

2. <span data-ttu-id="e54db-268">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-268">Open the file index.html.</span></span>

3. <span data-ttu-id="e54db-269">`create-chart` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-269">Below the `div` that contains the `create-chart` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="freeze-header">Freeze Header</button>
    </div>
    ```

4. <span data-ttu-id="e54db-270">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-270">Open the app.js file.</span></span>

5. <span data-ttu-id="e54db-271">`create-chart` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-271">Below the line that assigns a click handler to the `create-chart` button, add the following code:</span></span>

    ```js
    $('#freeze-header').click(freezeHeader);
    ```

6. <span data-ttu-id="e54db-272">`createChart` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-272">Below the `createChart` function add the following function:</span></span>

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

7. <span data-ttu-id="e54db-273">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e54db-273">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="e54db-274">次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e54db-274">Note:</span></span>

   - <span data-ttu-id="e54db-275">`Worksheet.freezePanes` コレクションは、ワークシートのスクロール操作時に、ワークシート上でピン留めつまり固定される一式のペインのことです。</span><span class="sxs-lookup"><span data-stu-id="e54db-275">The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.</span></span>

   - <span data-ttu-id="e54db-p131">`freezeRows` メソッドでは、上から数えた行数を、ピン留めする位置のパラメーターとして使用します。`1` を渡して最初の行を適所にピン留めします。</span><span class="sxs-lookup"><span data-stu-id="e54db-p131">The `freezeRows` method takes as a parameter the number of rows, from the top that are to be pinned in place. We pass `1` to pin the first row in place.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="e54db-278">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="e54db-278">Test the add-in</span></span>

1. <span data-ttu-id="e54db-279">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、**Ctrl + C** を 2 回入力して実行中の Web サーバーを停止します。</span><span class="sxs-lookup"><span data-stu-id="e54db-279">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter **Ctrl+C** twice to stop the running web server.</span></span> <span data-ttu-id="e54db-280">それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="e54db-280">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="e54db-p133">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。 そのためには、ビルド コマンドの入力を求めるプロンプトが表示されるように、サーバー プロセスを強制終了する必要があります。 ビルド後に、サーバーを再起動します。 次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="e54db-p133">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command. After the build, you restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="e54db-285">コマンド`npm run build`を実行して、Internet Explorer でサポートされている以前のバージョンの JAVASCRIPT に ES6 のソースコードを transpile します (excel の一部のバージョンでは、excel アドインを実行するために使用されます)。</span><span class="sxs-lookup"><span data-stu-id="e54db-285">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used by some versions of Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="e54db-286">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="e54db-286">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="e54db-287">作業ウィンドウを再読み込みするために、そのウィンドウを閉じ、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-287">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="e54db-288">ワークシート内に表があれば、削除します。</span><span class="sxs-lookup"><span data-stu-id="e54db-288">If the table is in the worksheet, delete it.</span></span>

6. <span data-ttu-id="e54db-289">作業ウィンドウで、**[Create Table]** (表の作成) を選択します。</span><span class="sxs-lookup"><span data-stu-id="e54db-289">In the task pane, choose **Create Table**.</span></span>

7. <span data-ttu-id="e54db-290">**[Freeze Header]** (ヘッダーを固定) ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="e54db-290">Choose the **Freeze Header** button.</span></span>

8. <span data-ttu-id="e54db-291">ヘッダー以降の行が画面の外に出て見えなくなるまでワークシートを下にスクロールしても、表のヘッダーが最上部に表示されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="e54db-291">Scroll down the worksheet enough to to see that the table header remains visible at the top even when the higher rows scroll out of sight.</span></span>

    ![Excel のチュートリアル - ヘッダーの固定](../images/excel-tutorial-freeze-header.png)

## <a name="protect-a-worksheet"></a><span data-ttu-id="e54db-293">ワークシートの保護</span><span class="sxs-lookup"><span data-stu-id="e54db-293">Protect a worksheet</span></span>

<span data-ttu-id="e54db-294">チュートリアルのこの手順では、リボンに別のボタンを追加します。このボタンをクリックすると、ワークシートの保護のオン/オフが切り替わるように定義した関数が実行されるようにします。</span><span class="sxs-lookup"><span data-stu-id="e54db-294">In this step of the tutorial, you'll add another button to the ribbon that, when chosen, executes a function that you'll define to toggle worksheet protection on and off.</span></span>

### <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a><span data-ttu-id="e54db-295">2 つ目のリボン ボタンを追加するようにマニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="e54db-295">Configure the manifest to add a second ribbon button</span></span>

1. <span data-ttu-id="e54db-296">マニフェスト ファイル my-office-add-in-manifest.xml を開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-296">Open the manifest file my-office-add-in-manifest.xml.</span></span>

2. <span data-ttu-id="e54db-p134">`<Control>` 要素を検索します。 この要素では、アドインの起動に使用している **[ホーム]** リボンの **[作業ウィンドウの表示]** ボタンを定義しています。 ここでは、**[ホーム]** リボンの同じグループに 2 つ目のボタンを追加します。 Control 終了タグ (`</Control>`) と Group 終了タグ (`</Group>`) の間に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-p134">Find the `<Control>` element. This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in. We're going to add a second button to the same group on the **Home** ribbon. In between the end Control tag (`</Control>`) and the end Group tag (`</Group>`), add the following markup.</span></span>

    ```xml
    <Control xsi:type="Button" id="<!--TODO1: Unique (in manifest) name for button -->">
        <Label resid="<!--TODO2: Button label -->" />
        <Supertip>            
            <Title resid="<!-- TODO3: Button tool tip title -->" />
            <Description resid="<!-- TODO4: Button tool tip description -->" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="<!-- TODO5: Specify the type of action-->">
            <!-- TODO6: Identify the function.-->
        </Action>
    </Control>
    ```

3. <span data-ttu-id="e54db-301">`TODO1` は文字列に置き換えて、このマニフェスト ファイル内で一意の ID をボタンに割り当てます。</span><span class="sxs-lookup"><span data-stu-id="e54db-301">Replace `TODO1` with a string that gives the button an ID that is unique within this manifest file.</span></span> <span data-ttu-id="e54db-302">このボタンでは、ワークシートの保護のオン/オフを切り替える予定なので、"ToggleProtection" を使用することにします。</span><span class="sxs-lookup"><span data-stu-id="e54db-302">Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection".</span></span> <span data-ttu-id="e54db-303">作業が完了すると、Control 開始タグの全体は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="e54db-303">When you are done, the entire start Control tag should look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. <span data-ttu-id="e54db-p136">その次の 3 つの `TODO` では、"resid" を設定します ("resid" はリソース ID の略号です)。 リソースは文字列です。これら 3 つの文字列は、この後の手順で作成します。 ここでは、そのリソースに ID を割り当てる必要があります。 ボタンのラベルは "Toggle Protection" と表示されるようにしますが、この文字列の *ID* は "ProtectionButtonLabel" にします。そのため、完成した `Label` 要素は次のコードのようになります。</span><span class="sxs-lookup"><span data-stu-id="e54db-p136">The next three `TODO`s set "resid"s, which is short for resource ID. A resource is a string, and you'll create these three strings in a later step. For now, you need to give IDs to the resources. The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the completed `Label` element should look like the following code:</span></span>

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. <span data-ttu-id="e54db-p137">`SuperTip` 要素では、このボタンのツール ヒントを定義します。 ツール ヒントのタイトルはボタンのラベルと同じにする必要があるため、リソース ID にはまったく同じ "ProtectionButtonLabel" を使用することにします。 ツール ヒントの説明は、"Click to turn protection of the worksheet on and off" にする予定です。 ただし、`ID` は "ProtectionButtonToolTip" にします。 作業が完了すると、`SuperTip` マークアップの全体は次のコードのようになります。</span><span class="sxs-lookup"><span data-stu-id="e54db-p137">The `SuperTip` element defines the tool tip for the button. The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel". The tool tip description will be "Click to turn protection of the worksheet on and off". But the `ID` should be "ProtectionButtonToolTip". So, when you are done, the whole `SuperTip` markup should look like the following code:</span></span> 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > <span data-ttu-id="e54db-p138">運用アドインでは、異なる 2 つのボタンに同じアイコンを使用することは避けたいところですが、このチュートリアルでは説明を簡単にするために同じアイコンを使用します。 そのため、この新しい `Icon` の `Control` マークアップは、単に既存の `Icon` から `Control` 要素をコピーします。</span><span class="sxs-lookup"><span data-stu-id="e54db-p138">In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that. So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`.</span></span> 

6. <span data-ttu-id="e54db-p139">既にマニフェストに存在している元の `Action` 要素の内側にある `Control` 要素では、その要素のタイプが `ShowTaskpane` に設定されていますが、新しいボタンで作業ウィンドウを開く予定はありません。このボタンでは、この後の手順で作成するカスタム関数を実行する予定です。 そのため、`TODO5` は、カスタム関数をトリガーするボタンのアクション タイプである `ExecuteFunction` に置き換えます。 `Action` 開始タグは次のようになります。</span><span class="sxs-lookup"><span data-stu-id="e54db-p139">The `Action` element inside the original `Control` element that was already present in the manifest, has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step. So replace `TODO5` with `ExecuteFunction` which is the action type for buttons that trigger custom functions. The start `Action` tag should look like the following code:</span></span>
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. <span data-ttu-id="e54db-p140">元の `Action` 要素には、作業ウィンドウ ID を指定する子要素と、作業ウィンドウで開かれるページの URL を指定する子要素があります。 ただし、`Action` タイプの `ExecuteFunction` 要素には、実行を制御する関数の名前を指定する子要素を 1 つ含めます。 その関数は、`toggleProtection` という名前にして、この後の手順で作成します。 そのために、`TODO6` を次のマークアップに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e54db-p140">The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane. But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes. You'll create that function in a later step, and it will be called `toggleProtection`. So, replace `TODO6` with the following markup:</span></span>
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    <span data-ttu-id="e54db-322">`Control` マークアップの全体は、次のようになりました。</span><span class="sxs-lookup"><span data-stu-id="e54db-322">The entire `Control` markup should now look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
        <Label resid="ProtectionButtonLabel" />
        <Supertip>            
            <Title resid="ProtectionButtonLabel" />
            <Description resid="ProtectionButtonToolTip" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="ExecuteFunction">
           <FunctionName>toggleProtection</FunctionName>
        </Action>
    </Control>
    ```

8. <span data-ttu-id="e54db-323">マニフェストの `Resources` セクションまで下にスクロールします。</span><span class="sxs-lookup"><span data-stu-id="e54db-323">Scroll down to the `Resources` section of the manifest.</span></span>

9. <span data-ttu-id="e54db-324">`bt:ShortStrings` 要素の子として、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-324">Add the following markup as a child of the `bt:ShortStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. <span data-ttu-id="e54db-325">`bt:LongStrings` 要素の子として、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-325">Add the following markup as a child of the `bt:LongStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. <span data-ttu-id="e54db-326">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="e54db-326">Save the file.</span></span>

### <a name="create-the-function-that-protects-the-sheet"></a><span data-ttu-id="e54db-327">シートを保護する関数を作成する</span><span class="sxs-lookup"><span data-stu-id="e54db-327">Create the function that protects the sheet</span></span>

1. <span data-ttu-id="e54db-328">ファイル \function-file\function-file.js を開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-328">Open the file \function-file\function-file.js.</span></span>

2. <span data-ttu-id="e54db-329">このファイルには、即時実行関数式 (IIFE) が既に含まれています。</span><span class="sxs-lookup"><span data-stu-id="e54db-329">The file already has an Immediately Invoked Function Expression (IFFE).</span></span> <span data-ttu-id="e54db-330">*IIFE の外部*に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-330">*Outside of the IIFE*, add the following code.</span></span> <span data-ttu-id="e54db-331">メソッドに `args` パラメーターを指定していることと、メソッドの最後のほうの行で `args.completed` を呼び出していることに注目してください。</span><span class="sxs-lookup"><span data-stu-id="e54db-331">Note that we specify an `args` parameter to the method and the very last line of the method calls `args.completed`.</span></span> <span data-ttu-id="e54db-332">**ExecuteFunction** タイプのすべてのアドイン コマンドでは、これが要件になります。</span><span class="sxs-lookup"><span data-stu-id="e54db-332">This is a requirement for all add-in commands of type **ExecuteFunction**.</span></span> <span data-ttu-id="e54db-333">これにより、関数が終了したことと、UI が再度応答可能になることを Office ホスト アプリケーションに通知します。</span><span class="sxs-lookup"><span data-stu-id="e54db-333">It signals the Office host application that the function has finished and the UI can become responsive again.</span></span>

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

3. <span data-ttu-id="e54db-p142">`TODO1` を次のコードに置き換えます。 このコードでは、標準の切り替えパターンで、ワークシート オブジェクトの protection プロパティを使用します。 `TODO2` については、次のセクションで説明します。</span><span class="sxs-lookup"><span data-stu-id="e54db-p142">Replace `TODO1` with the following code. This code uses the worksheet object's protection property in a standard toggle pattern. The `TODO2` will be explained in the next section.</span></span>

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

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="e54db-337">ドキュメントのプロパティを作業ウィンドウのスクリプト オブジェクトにフェッチするコードを追加する</span><span class="sxs-lookup"><span data-stu-id="e54db-337">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="e54db-p143">このチュートリアルのシリーズで前述したすべての関数では、Office ドキュメントへの*書き込み*コマンドをキューに登録していました。 各関数は、キューに登録されたコマンドを実行対象のドキュメントに送信する `context.sync()` メソッドを呼び出すことで終了しています。 ただし、最後の手順で追加したコードでは、`sheet.protection.protected` プロパティを呼び出しています。このことが、これまでに作成した関数とは大きく異なります。`sheet` オブジェクトは、この作業ウィンドウのスクリプトに存在する単なるプロキシ オブジェクトなので、 ドキュメントの実際の保護の状態を認識できません。そのため、その `protection.protected` プロパティでは実際の値が保持できません。 まず、ドキュメントから保護の状態をフェッチする必要があり、その状態を使用して `sheet.protection.protected` の値を設定します。 そのようにした場合にのみ、例外がスローされることなく `sheet.protection.protected` を呼び出せるようになります。 このフェッチ処理には、3 つの手順があります。</span><span class="sxs-lookup"><span data-stu-id="e54db-p143">In all the earlier functions in this series of tutorials, you queued commands to *write* to the Office document. Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed. But the code you added in the last step calls the `sheet.protection.protected` property, and this is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script. It doesn't know what the actual protection state of the document is, so its `protection.protected` property can't have a real value. It is necessary to first fetch the protection status from the document and use it set the value of `sheet.protection.protected`. Only then can `sheet.protection.protected` be called without causing an exception to be thrown. This fetching process has three steps:</span></span>

   1. <span data-ttu-id="e54db-345">コードで読み取る必要があるプロパティをロードする (つまりフェッチする) コマンドをキューに登録します。</span><span class="sxs-lookup"><span data-stu-id="e54db-345">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>

   2. <span data-ttu-id="e54db-346">コンテキスト オブジェクトの `sync` メソッドを呼び出します。このメソッドは、キューに登録されたコマンドを実行対象のドキュメントに送信して、要求された情報を返します。</span><span class="sxs-lookup"><span data-stu-id="e54db-346">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>

   3. <span data-ttu-id="e54db-347">`sync` メソッドは非同期であるため、フェッチされたプロパティをコードで呼び出す前に、そのメソッドが完了していることを確認します。</span><span class="sxs-lookup"><span data-stu-id="e54db-347">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="e54db-348">こうした手順は、コードで Office ドキュメントから情報を*読み取る*必要がある場合には必ず完了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e54db-348">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="e54db-p144">`toggleProtection` 関数で、`TODO2` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e54db-p144">In the `toggleProtection` function, replace `TODO2` with the following code. Note:</span></span>
   
   - <span data-ttu-id="e54db-p145">すべての Excel オブジェクトに `load` メソッドがあります。 読み取る必要のあるオブジェクトのプロパティは、コンマ区切りの名前の文字列としてパラメーターで指定します。 この場合、読み取る必要のあるプロパティは、`protection` プロパティのサブプロパティです。 サブプロパティはその他のコードの場合とほとんど同じ方法で参照しますが、"." 記号の代わりにスラッシュ ('/') 記号を使用する点が異なります。</span><span class="sxs-lookup"><span data-stu-id="e54db-p145">Every Excel object has a `load` method. You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names. In this case, the property you need to read is a subproperty of the `protection` property. You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.</span></span>

   - <span data-ttu-id="e54db-355">`sheet.protection.protected` が完了してドキュメントからフェッチされた適切な値が `sync` に割り当てられるまで、`sheet.protection.protected` を読み取る切り替えロジックが実行されないようにするために、そのロジックを `then` が完了するまで実行されない `sync` 関数に (この後の手順で) 移動します。</span><span class="sxs-lookup"><span data-stu-id="e54db-355">To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed.</span></span> 

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

2. <span data-ttu-id="e54db-p146">分岐していない同一のコード パスに 2 つの `return` ステートメントを含めることはできないため、`return context.sync();` の最後にある最終行の `Excel.run` を削除します。 この後の手順で、新しい最終の `context.sync` を追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-p146">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`. You will add a new final `context.sync`, in a later step.</span></span>

3. <span data-ttu-id="e54db-358">`if ... else` 関数内の `toggleProtection` 構造を切り取って、`TODO3` の代わりに貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="e54db-358">Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.</span></span>

4. <span data-ttu-id="e54db-p147">`TODO4` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e54db-p147">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="e54db-361">`sync` メソッドを `then` 関数に渡すことで、`sheet.protection.unprotect()` または `sheet.protection.protect()` のどちらかがキューに登録されるまで、そのメソッドが実行されないようにします。</span><span class="sxs-lookup"><span data-stu-id="e54db-361">Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.</span></span>

   - <span data-ttu-id="e54db-362">`then` メソッドは渡された関数を呼び出します。`sync` が 2 回呼び出されないように、`context.sync` の末尾の "()" は省略します。</span><span class="sxs-lookup"><span data-stu-id="e54db-362">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.</span></span>

    ```js
    .then(context.sync);
    ```

   <span data-ttu-id="e54db-363">作業が完了すると、関数の全体は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="e54db-363">When you are done, the entire function should look like the following:</span></span>

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

### <a name="configure-the-script-loading-html-file"></a><span data-ttu-id="e54db-364">スクリプト読み込み HTMl ファイルを構成する</span><span class="sxs-lookup"><span data-stu-id="e54db-364">Configure the script-loading HTML file</span></span>

<span data-ttu-id="e54db-p148">/function-file/function-file.html ファイルを開きます。 これは、ユーザーが **[Toggle Worksheet Protection]** ボタンをクリックしたときに呼び出される UI のない HTML ファイルです。 ボタンがクリックされたときに実行する JavaScript メソッドを読み込むことを目的としています。 このファイルには変更を加えません。 2 番目の `<script>` タグで functionfile.js が読み込まれる点に注目してください。</span><span class="sxs-lookup"><span data-stu-id="e54db-p148">Open the /function-file/function-file.html file. This is a UI-less HTML file that is called when the user presses the **Toggle Worksheet Protection** button. Its purpose is to load the JavaScript method that should run when the button is pushed. You are not going to change this file. Simply note that the second `<script>` tag loads the functionfile.js.</span></span>

   > [!NOTE]
   > <span data-ttu-id="e54db-p149">function-file.html ファイルと、そのファイルが読み込む function-file.js ファイルは、アドインの作業ウィンドウとは完全に別の IE プロセスで実行されます。 function-file.js が app.js ファイルと同じ bundle.js ファイルからトランスパイルされていた場合、アドインでは bundle.js の 2 つのコピーを読み込むことが必要になり、バンドル化の意味がなくなります。 さらに、function-file.js ファイルには IE で未サポートの JavaScript は含まれていません。 これら 2 つの理由から、このアドインでは function-file.js を一切トランスパイルしていません。</span><span class="sxs-lookup"><span data-stu-id="e54db-p149">The function-file.html file and the function-file.js file that it loads run in an entirely separate IE process from the add-in's task pane. If the function-file.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling. In addition, the function-file.js file does not contain any JavaScript that is unsupported by IE. For these two reasons, this add-in does not transpile the function-file.js at all.</span></span> 

### <a name="test-the-add-in"></a><span data-ttu-id="e54db-374">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="e54db-374">Test the add-in</span></span>

1. <span data-ttu-id="e54db-375">Excel も含めて、すべての Office アプリケーションを閉じます。</span><span class="sxs-lookup"><span data-stu-id="e54db-375">Close all Office applications, including Excel.</span></span> 

2. <span data-ttu-id="e54db-p150">キャッシュ フォルダーの内容を削除して、Office キャッシュを削除します。 これは、ホストから古いバージョンのアドインを完全に削除するために必要です。</span><span class="sxs-lookup"><span data-stu-id="e54db-p150">Delete the Office cache by deleting the contents of the cache folder. This is necessary to completely clear the old version of the add-in from the host.</span></span> 

    - <span data-ttu-id="e54db-378">Windows の場合: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`。</span><span class="sxs-lookup"><span data-stu-id="e54db-378">For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

    - <span data-ttu-id="e54db-379">Mac の場合: `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`。</span><span class="sxs-lookup"><span data-stu-id="e54db-379">For Mac: `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span> 
    
        [!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

3. <span data-ttu-id="e54db-p151">何らかの理由で、サーバーが稼働中でない場合は、Git Bash ウィンドウ、または Node.JS 対応のシステム プロンプトで、プロジェクトの **Start** フォルダーに移動して、`npm start` コマンドを実行します。 変更した JavaScript ファイルはビルド済みの bundle.js に含まれていないため、プロジェクトをリビルドする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="e54db-p151">If for any reason, your server is not running, then in a Git Bash window, or Node.JS-enabled system prompt, navigate to the **Start** folder of the project and run the command `npm start`. You do not need to rebuild the project because the only JavaScript file you changed is not part of the built bundle.js.</span></span>

4. <span data-ttu-id="e54db-p152">新しいバージョンの変更済みマニフェスト ファイルを使用して、次のいずれかの方法でサイドローディング プロセスを繰り返します。 *マニフェスト ファイルの以前のコピーを上書きする必要があります。*</span><span class="sxs-lookup"><span data-stu-id="e54db-p152">Using the new version of the changed manifest file, repeat the sideloading process by using one of the following methods. *You should overwrite the previous copy of the manifest file.*</span></span>

    - <span data-ttu-id="e54db-384">Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="e54db-384">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>

    - <span data-ttu-id="e54db-385">Web ブラウザー:[サイドロード Office アドイン (web)](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="e54db-385">Web browser: [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span></span>

    - <span data-ttu-id="e54db-386">iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="e54db-386">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

5. <span data-ttu-id="e54db-387">Excel で任意のワークシートを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-387">Open any worksheet in Excel.</span></span>

6. <span data-ttu-id="e54db-p153">**[ホーム]** リボンで、**[ワークシート保護の切り替え]** を選択します。次のスクリーンショットに示すように、リボンのほとんどのコントロールは、無効化 (淡色表示) されます。</span><span class="sxs-lookup"><span data-stu-id="e54db-p153">On the **Home** ribbon, choose **Toggle Worksheet Protection**. Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in screenshot below.</span></span> 

7. <span data-ttu-id="e54db-p154">セルの内容を変更する場合は、そのセルを選択します。 ワークシートが保護されているというエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="e54db-p154">Choose a cell as you would if you wanted to change its content. You get an error telling you that the worksheet is protected.</span></span>

8. <span data-ttu-id="e54db-392">もう一度 **[Toggle Worksheet Protection]** を選択すると、コントロールが再有効化され、再びセルの値を変更できるようになります。</span><span class="sxs-lookup"><span data-stu-id="e54db-392">Choose **Toggle Worksheet Protection** again, and the controls are reenabled, and you can change cell values again.</span></span>

    ![Excel チュートリアル - 保護がオンになっているリボン](../images/excel-tutorial-ribbon-with-protection-on.png)

## <a name="open-a-dialog"></a><span data-ttu-id="e54db-394">ダイアログを開く</span><span class="sxs-lookup"><span data-stu-id="e54db-394">Open a dialog</span></span>

<span data-ttu-id="e54db-p155">このチュートリアルの最後の手順では、アドインでダイアログを開いて、ダイアログのプロセスから作業ウィンドウのプロセスにメッセージを渡して、ダイアログを閉じます。 Office アドインのダイアログは、*モードレス*です。ユーザーは、ホスト Office アプリケーション内のドキュメントと作業ウィンドウ内のホスト ページの両方の操作を続行できます。</span><span class="sxs-lookup"><span data-stu-id="e54db-p155">In this final step of the tutorial, you'll open a dialog in your add-in, pass a message from the dialog process to the task pane process, and close the dialog. Office Add-in dialogs are *nonmodal*: a user can continue to interact with both the document in the host Office application and with the host page in the task pane.</span></span>

### <a name="create-the-dialog-page"></a><span data-ttu-id="e54db-397">ダイアログ ページを作成する</span><span class="sxs-lookup"><span data-stu-id="e54db-397">Create the dialog page</span></span>

1. <span data-ttu-id="e54db-398">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-398">Open the project in your code editor.</span></span>

2. <span data-ttu-id="e54db-399">プロジェクトのルート (index.html がある場所) で、popup.html というファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="e54db-399">Create a file in the root of the project (where index.html is) called popup.html.</span></span>

3. <span data-ttu-id="e54db-p156">popup.html に、次のコードを追加します。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e54db-p156">Add the following markup to popup.html. Note:</span></span>

   - <span data-ttu-id="e54db-402">このページには、ユーザーが自分の名前を入力する `<input>` と、その名前を作業ウィンドウ内のページ (入力した名前が表示されるページ) に送信するボタンが含まれています。</span><span class="sxs-lookup"><span data-stu-id="e54db-402">The page has a `<input>` where the user will enter their name and a button that will send the name to the page in the task pane where it will be displayed.</span></span>

   - <span data-ttu-id="e54db-403">このマークアップでは、popup.js というスクリプトを読み込みます。このスクリプトは、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="e54db-403">The markup loads a script called popup.js that you will create in a later step.</span></span>

   - <span data-ttu-id="e54db-404">また、popup.js で使用することになる Office.JS ライブラリと jQuery も読み込みます。</span><span class="sxs-lookup"><span data-stu-id="e54db-404">It also loads the Office.JS library and jQuery because they will be used in popup.js.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head lang="en">
            <title>Dialog for My Office Add-in</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">

            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.min.css" />
            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.components.css" />
            <link rel="stylesheet" href="app.css" />

            <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
            <script type="text/javascript" src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.2.1.min.js"></script>
            <script type="text/javascript" src="popup.js"></script>

        </head>
        <body style="display:flex;flex-direction:column;align-items:center;justify-content:center">
            <div class="padding">
                <p class="ms-font-xl">ENTER YOUR NAME</p>
            </div>
            <div class="padding">
                <input id="name-box" type="text"/>
            </div>
            <div class="padding">
                <button id="ok-button" class="ms-Button">OK</button>
            </div>
        </body>
    </html>
    ```

4. <span data-ttu-id="e54db-405">プロジェクトのルートに popup.js というファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="e54db-405">Create a file in the root of the project called popup.js.</span></span>

5. <span data-ttu-id="e54db-406">popup.js に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-406">Add the following code to popup.js.</span></span> <span data-ttu-id="e54db-407">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e54db-407">Note the following about this code:</span></span>

   - <span data-ttu-id="e54db-408">*Office.JS ライブラリ内の API を呼び出すすべてのページでは、まずライブラリが完全に初期化されていることを確認する必要があります。*</span><span class="sxs-lookup"><span data-stu-id="e54db-408">*Every page that calls APIs in the Office.JS library must first ensure that the library is fully initialized.*</span></span> <span data-ttu-id="e54db-409">これを行う最善の方法は `Office.onReady()` メソッドを呼び出すことです。</span><span class="sxs-lookup"><span data-stu-id="e54db-409">The best way to do that is to call the `Office.onReady()` method.</span></span> <span data-ttu-id="e54db-410">アドインに独自の初期化タスクがある場合、コードを `Office.onReady()` の呼び出しにチェーンされている `then()` メソッドに含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="e54db-410">If your add-in has its own initialization tasks, the code should go in a `then()` method that is chained to the call of `Office.onReady()`.</span></span> <span data-ttu-id="e54db-411">たとえば、プロジェクト ルートにある app.js ファイルを確認してください。</span><span class="sxs-lookup"><span data-stu-id="e54db-411">For an example, see the app.js file in the project root.</span></span> <span data-ttu-id="e54db-412">`Office.onReady()` の呼び出しは、Office.JS を呼び出す前に実行する必要があります。そのため、この例で示すように、割り当てはページによって読み込まれるスクリプト ファイル内に入れてあります。</span><span class="sxs-lookup"><span data-stu-id="e54db-412">The call of `Office.onReady()` must run before any calls to Office.JS; hence the assignment is in a script file that is loaded by the page, as it is in this case.</span></span>
   - <span data-ttu-id="e54db-413">jQuery の `ready` 関数は、`then()` メソッドの内側から呼び出します。</span><span class="sxs-lookup"><span data-stu-id="e54db-413">The jQuery `ready` function is called inside the `then()` method.</span></span> <span data-ttu-id="e54db-414">通常は、その他の JavaScript ライブラリの読み込み、初期化、ブートストラップのコードを、`Office.onReady()` の呼び出しにチェーンされている `then()` メソッドの内側に含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="e54db-414">In most cases, the loading, initializing, or bootstrapping code of other JavaScript libraries should be inside the `then()` method that is chained to the call of `Office.onReady()`.</span></span>

    ```js
    (function () {
    "use strict";

        Office.onReady()
            .then(function() {
                $(document).ready(function () {  

                    // TODO1: Assign handler to the OK button.

                });
            });

        // TODO2: Create the OK button handler

    }());
    ```

6. <span data-ttu-id="e54db-p160">`TODO1` を次のコードに置き換えます。 `sendStringToParentPage` 関数は、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="e54db-p160">Replace `TODO1` with the following code. You'll create the `sendStringToParentPage` function in the next step.</span></span>

    ```js
    $('#ok-button').click(sendStringToParentPage);
    ```

7. <span data-ttu-id="e54db-p161">`TODO2` を次のコードに置き換えます。 `messageParent` メソッドは、パラメーターを親ページ (この例では、作業ウィンドウ内のページ) に渡します。 パラメーターには、ブール値または文字列を使用できます (XML や JSON など、文字列としてシリアル化できるすべてのものが含まれます)。</span><span class="sxs-lookup"><span data-stu-id="e54db-p161">Replace `TODO2` with the following code. The `messageParent` method passes its parameter to the parent page, in this case, the page in the task pane. The parameter can be a boolean or a string, which includes anything that can be serialized as a string, such as XML or JSON.</span></span>

    ```js
    function sendStringToParentPage() {
        var userName = $('#name-box').val();
        Office.context.ui.messageParent(userName);
    }
    ```

8. <span data-ttu-id="e54db-420">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="e54db-420">Save the file.</span></span>

   > [!NOTE]
   > <span data-ttu-id="e54db-421">ポップアップファイルとそれによって読み込まれるポップアップファイルは、アドインの作業ウィンドウから完全に独立した Microsoft Edge または Internet Explorer 11 プロセスで実行されます。</span><span class="sxs-lookup"><span data-stu-id="e54db-421">The popup.html file, and the popup.js file that it loads, run in an entirely separate Microsoft Edge or Internet Explorer 11 process from the add-in's task pane.</span></span> <span data-ttu-id="e54db-422">popup.js が app.js ファイルと同じ bundle.js ファイルからトランスパイルされていた場合、アドインでは bundle.js の 2 つのコピーを読み込むことが必要になり、バンドル化の意味がなくなります。</span><span class="sxs-lookup"><span data-stu-id="e54db-422">If the popup.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="e54db-423">さらに、ポップアップ .js ファイルには、Internet Explorer 11 でサポートされていない JavaScript は含まれていません。</span><span class="sxs-lookup"><span data-stu-id="e54db-423">In addition, the popup.js file does not contain any JavaScript that is unsupported by Internet Explorer 11.</span></span> <span data-ttu-id="e54db-424">これら 2 つの理由から、このアドインでは popup.js を一切トランスパイルしていません。</span><span class="sxs-lookup"><span data-stu-id="e54db-424">For these two reasons, this add-in does not transpile the popup.js file at all.</span></span>

### <a name="open-the-dialog-from-the-task-pane"></a><span data-ttu-id="e54db-425">作業ウィンドウからダイアログを開く</span><span class="sxs-lookup"><span data-stu-id="e54db-425">Open the dialog from the task pane</span></span>

1. <span data-ttu-id="e54db-426">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-426">Open the file index.html.</span></span>

2. <span data-ttu-id="e54db-427">`freeze-header` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-427">Below the `div` that contains the `freeze-header` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="open-dialog">Open Dialog</button>
    </div>
    ```

3. <span data-ttu-id="e54db-p163">このダイアログでは、ユーザーに名前の入力を求めて、ユーザーの名前を作業ウィンドウに渡します。 作業ウィンドウでは、それがラベルに表示されます。 前の手順で追加した `div` のすぐ下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-p163">The dialog will prompt the user to enter a name and pass the user's name to the task pane. The task pane will display it in a label. Immediately below the `div` that you just added, add the following markup:</span></span>

    ```html
    <div class="padding">
        <label id="user-name"></label>
    </div>
    ```

4. <span data-ttu-id="e54db-431">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-431">Open the app.js file.</span></span>

5. <span data-ttu-id="e54db-p164">`freeze-header` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。 `openDialog` メソッドは、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="e54db-p164">Below the line that assigns a click handler to the `freeze-header` button, add the following code. You'll create the `openDialog` method in a later step.</span></span>

    ```js
    $('#open-dialog').click(openDialog);
    ```

6. <span data-ttu-id="e54db-p165">`freezeHeader` 関数の下に、次の宣言を追加します。この変数は、親ページの実行コンテキスト内のオブジェクトを保持するために使用され、ダイアログ ページの実行コンテキストへの仲介者として機能します。</span><span class="sxs-lookup"><span data-stu-id="e54db-p165">Below the `freezeHeader` function add the following declaration. This variable is used to hold an object in the parent page's execution context that acts as an intermediator to the dialog page's execution context.</span></span>

    ```js
    var dialog = null;
    ```

7. <span data-ttu-id="e54db-p166">`dialog` の宣言の下に、次の関数を追加します。 このコードで注目する重要な点は、そこに \*\* の呼び出しが存在`Excel.run`ことです。 これは、ダイアログを開く API はすべての Office ホストで共有されるため、Excel 固有の API ではなく Office JavaScript 共通 API に含まれているからです。</span><span class="sxs-lookup"><span data-stu-id="e54db-p166">Below the declaration of `dialog`, add the following function. The important thing to notice about this code is what is *not* there: there is no call of `Excel.run`. This is because the API to open a dialog is shared among all Office hosts, so it is part of the Office JavaScript Common API, not the Excel-specific API.</span></span>

    ```js
    function openDialog() {
        // TODO1: Call the Office Common API that opens a dialog
    }
    ```

8. <span data-ttu-id="e54db-439">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e54db-439">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="e54db-440">次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e54db-440">Note:</span></span>

   - <span data-ttu-id="e54db-441">`displayDialogAsync` メソッドでは、画面の中央にダイアログを開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-441">The `displayDialogAsync` method opens a dialog in the center of the screen.</span></span>

   - <span data-ttu-id="e54db-442">最初のパラメーターは、開くページの URL です。</span><span class="sxs-lookup"><span data-stu-id="e54db-442">The first parameter is the URL of the page to open.</span></span>

   - <span data-ttu-id="e54db-p168">2 番目のパラメーターでオプションを渡します。`height` と `width` は、Office アプリケーションのウィンドウ サイズの比率です。</span><span class="sxs-lookup"><span data-stu-id="e54db-p168">The second parameter passes options. `height` and `width` are percentages of the size of the Office application's window.</span></span>

    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},

        // TODO2: Add callback parameter.
    );
    ```

### <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a><span data-ttu-id="e54db-445">ダイアログからのメッセージを処理してダイアログを閉じる</span><span class="sxs-lookup"><span data-stu-id="e54db-445">Process the message from the dialog and close the dialog</span></span>

1. <span data-ttu-id="e54db-p169">app.js ファイルでの作業を続けます。`TODO2` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e54db-p169">Continue in the app.js file, and replace `TODO2` with the following code. Note:</span></span>

   - <span data-ttu-id="e54db-448">コールバックは、ダイアログが正常に開いた直後、ユーザーがダイアログで操作を行う前に実行されます。</span><span class="sxs-lookup"><span data-stu-id="e54db-448">The callback is executed immediately after the dialog successfully opens and before the user has taken any action in the dialog.</span></span>

   - <span data-ttu-id="e54db-449">`result.value` は、親ページとダイアログ ページの実行コンテキストの間で仲介者のように機能するオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="e54db-449">The `result.value` is the object that acts as a kind of middleman between the execution contexts of the parent and dialog pages.</span></span>

   - <span data-ttu-id="e54db-p170">`processMessage` 関数は、この後の手順で作成します。 このハンドラーは、`messageParent` 関数の呼び出しによって、ダイアログから送信されるあらゆる値を処理します。</span><span class="sxs-lookup"><span data-stu-id="e54db-p170">The `processMessage` function will be created in a later step. This handler will process any values that are sent from the dialog page with calls of the `messageParent` function.</span></span>

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. <span data-ttu-id="e54db-452">`openDialog` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="e54db-452">Below the `openDialog` function, add the following function.</span></span>

    ```js
    function processMessage(arg) {
        $('#user-name').text(arg.message);
        dialog.close();
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="e54db-453">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="e54db-453">Test the add-in</span></span>

1. <span data-ttu-id="e54db-454">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、**Ctrl + C** を 2 回入力して実行中の Web サーバーを停止します。</span><span class="sxs-lookup"><span data-stu-id="e54db-454">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter **Ctrl+C** twice to stop the running web server.</span></span> <span data-ttu-id="e54db-455">それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="e54db-455">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="e54db-p172">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。 そのためには、ビルド コマンドの入力を求めるプロンプトが表示されるように、サーバー プロセスを強制終了する必要があります。 ビルド後に、サーバーを再起動します。 次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="e54db-p172">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command. After the build, you restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="e54db-460">コマンド`npm run build`を実行して、Internet Explorer でサポートされている以前のバージョンの JAVASCRIPT に ES6 のソースコードを transpile します (excel の一部のバージョンでは、excel アドインを実行するために使用されます)。</span><span class="sxs-lookup"><span data-stu-id="e54db-460">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used by some versions of Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="e54db-461">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="e54db-461">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="e54db-462">作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="e54db-462">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="e54db-463">作業ウィンドウで、**[Open Dialog]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="e54db-463">Choose the **Open Dialog** button in the task pane.</span></span>

6. <span data-ttu-id="e54db-464">ダイアログが開いたら、ドラッグしたりサイズ変更したりします。</span><span class="sxs-lookup"><span data-stu-id="e54db-464">While the dialog is open, drag it and resize it.</span></span> <span data-ttu-id="e54db-465">ワークシートを操作して、作業ウィンドウの他のボタンを押すことはできますが、同じ作業ウィンドウのページから 2 番目のダイアログを起動することはできないことに注意してください。</span><span class="sxs-lookup"><span data-stu-id="e54db-465">Note that you can interact with the worksheet and press other buttons on the task pane, but you cannot launch a second dialog from the same task pane page.</span></span>

7. <span data-ttu-id="e54db-p174">ダイアログで、名前を入力して **[OK]** をクリックします。 作業ウィンドウに名前が表示され、ダイアログが閉じられます。</span><span class="sxs-lookup"><span data-stu-id="e54db-p174">In the dialog, enter a name and choose **OK**. The name appears on the task pane and the dialog closes.</span></span>

8. <span data-ttu-id="e54db-p175">オプションとして、`dialog.close();` 関数内の行 `processMessage` をコメントにします。 その後で、このセクションの手順を繰り返します。 ダイアログを開いたまま名前を変更できます。 右上の **[X]** ボタンをクリックすることで、手動で閉じることができます。</span><span class="sxs-lookup"><span data-stu-id="e54db-p175">Optionally, comment out the line `dialog.close();` in the `processMessage` function. Then repeat the steps of this section. The dialog stays open and you can change the name. You can close it manually by pressing the **X** button in the upper right corner.</span></span>

    ![Excel チュートリアル - ダイアログ](../images/excel-tutorial-dialog-open.png)

## <a name="next-steps"></a><span data-ttu-id="e54db-473">次の手順</span><span class="sxs-lookup"><span data-stu-id="e54db-473">Next steps</span></span>

<span data-ttu-id="e54db-474">このチュートリアルでは、Excel ブック内のテーブル、グラフ、ワークシート、ダイアログの操作を行う、Excel 作業ウィンドウ アドインを作成しました。</span><span class="sxs-lookup"><span data-stu-id="e54db-474">In this tutorial, you've created an Excel task pane add-in that interacts with tables, charts, worksheets, and dialogs in an Excel workbook.</span></span> <span data-ttu-id="e54db-475">Excel アドインの構築に関する詳細については、次の記事にお進みください。</span><span class="sxs-lookup"><span data-stu-id="e54db-475">To learn more about building Excel add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="e54db-476">Excel アドインの概要</span><span class="sxs-lookup"><span data-stu-id="e54db-476">Excel add-ins overview</span></span>](../excel/excel-add-ins-overview.md)
