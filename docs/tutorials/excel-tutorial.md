---
title: Excel アドインのチュートリアル
description: このチュートリアルでは、Excel アドインを構築します。このアドインでは、テーブルの作成、表示、フィルター処理、並べ替えを行うことができ、グラフの作成、テーブルのヘッダーの固定、ワークシートの保護も可能となります。また、ダイアログを開くこともできます。
ms.date: 12/31/2018
ms.topic: tutorial
ms.openlocfilehash: fe4350f5f3fdbe34250c1739c7651a1dde1e28ef
ms.sourcegitcommit: 3007bf57515b0811ff98a7e1518ecc6fc9462276
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/04/2019
ms.locfileid: "27724958"
---
# <a name="tutorial-create-an-excel-task-pane-add-in"></a><span data-ttu-id="a7282-103">チュートリアル: Excel 作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="a7282-103">Tutorial: Create an Excel task pane add-in</span></span>

<span data-ttu-id="a7282-104">このチュートリアルでは、以下を実行する Excel 作業ウィンドウ アドインを作成します。</span><span class="sxs-lookup"><span data-stu-id="a7282-104">In this tutorial, you'll create an Excel task pane add-in that:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="a7282-105">テーブルの作成</span><span class="sxs-lookup"><span data-stu-id="a7282-105">Creates a new table.</span></span>
> * <span data-ttu-id="a7282-106">テーブルのフィルター処理と並べ替え</span><span class="sxs-lookup"><span data-stu-id="a7282-106">Filters and sorts a table</span></span>
> * <span data-ttu-id="a7282-107">グラフの作成</span><span class="sxs-lookup"><span data-stu-id="a7282-107">Creates a new chart.</span></span>
> * <span data-ttu-id="a7282-108">テーブルのヘッダーの固定</span><span class="sxs-lookup"><span data-stu-id="a7282-108">Freezes a table header</span></span>
> * <span data-ttu-id="a7282-109">ワークシートの保護</span><span class="sxs-lookup"><span data-stu-id="a7282-109">Protects a worksheet</span></span>
> * <span data-ttu-id="a7282-110">ダイアログを開く</span><span class="sxs-lookup"><span data-stu-id="a7282-110">Opens a dialog</span></span>

## <a name="prerequisites"></a><span data-ttu-id="a7282-111">前提条件</span><span class="sxs-lookup"><span data-stu-id="a7282-111">Prerequisites</span></span>

<span data-ttu-id="a7282-112">このチュートリアルを使用するには、以下のバージョンがインストールされている必要があります。</span><span class="sxs-lookup"><span data-stu-id="a7282-112">To use this tutorial, you need to have the following installed.</span></span> 

- <span data-ttu-id="a7282-113">Excel 2016、バージョン 1711 (ビルド 8730.1000 クイック実行) 以降。</span><span class="sxs-lookup"><span data-stu-id="a7282-113">Excel 2016, version 1711 (Build 8730.1000 Click-to-Run) or later.</span></span> <span data-ttu-id="a7282-114">このバージョンを入手するには、Office Insider への参加が必要になることがあります。</span><span class="sxs-lookup"><span data-stu-id="a7282-114">You might need to be an Office Insider to get this version.</span></span> <span data-ttu-id="a7282-115">詳細については、「[Office Insider になる](https://products.office.com/office-insider?tab=tab-1)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-115">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>

- [<span data-ttu-id="a7282-116">ノード</span><span class="sxs-lookup"><span data-stu-id="a7282-116">Node</span></span>](https://nodejs.org/en/) 

- <span data-ttu-id="a7282-117">[Git バッシュ](https://git-scm.com/downloads) (または別の Git クライアント)</span><span class="sxs-lookup"><span data-stu-id="a7282-117">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

## <a name="create-your-add-in-project"></a><span data-ttu-id="a7282-118">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="a7282-118">Create your add-in project</span></span>

<span data-ttu-id="a7282-119">このチュートリアルの基礎として使用する Excel アドイン プロジェクトを作成するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="a7282-119">Complete the following steps to create the Excel add-in project that you'll use as the basis for this tutorial.</span></span>

1. <span data-ttu-id="a7282-120">「[Excel アドインのチュートリアル](https://github.com/OfficeDev/Excel-Add-in-Tutorial)」で、GitHub リポジトリを複製します。</span><span class="sxs-lookup"><span data-stu-id="a7282-120">Clone the GitHub repository [Excel Add-in Tutorial](https://github.com/OfficeDev/Excel-Add-in-Tutorial).</span></span>

2. <span data-ttu-id="a7282-121">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="a7282-121">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

3. <span data-ttu-id="a7282-122">`npm install` コマンドを実行して、package.json ファイルに一覧表示されているツールとライブラリをインストールします。</span><span class="sxs-lookup"><span data-stu-id="a7282-122">Run the command `npm install` to install the tools and libraries listed in the package.json file.</span></span> 

4. <span data-ttu-id="a7282-123">「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」の手順を実行して、開発用コンピューターのオペレーティング システムの証明書を信頼します。</span><span class="sxs-lookup"><span data-stu-id="a7282-123">Carry out the steps in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

## <a name="create-a-table"></a><span data-ttu-id="a7282-124">テーブルの作成</span><span class="sxs-lookup"><span data-stu-id="a7282-124">Create a table</span></span>

<span data-ttu-id="a7282-125">チュートリアルのこの手順では、プログラムによってアドインがユーザーの Excel の現在のバージョンをサポートしているかどうかをテストし、ワークシートにテーブルを追加して、そのテーブルのデータ設定と書式設定を実行します。</span><span class="sxs-lookup"><span data-stu-id="a7282-125">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Excel, add a table to a worksheet, populate the table with data, and format it.</span></span>

### <a name="code-the-add-in"></a><span data-ttu-id="a7282-126">アドインのコードを作成する</span><span class="sxs-lookup"><span data-stu-id="a7282-126">Code the add-in</span></span>

1. <span data-ttu-id="a7282-127">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-127">Open the project in your code editor.</span></span>

2. <span data-ttu-id="a7282-128">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-128">Open the file index.html.</span></span>

3. <span data-ttu-id="a7282-129">`TODO1` を次のマークアップに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a7282-129">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="create-table">Create Table</button>
    ```

4. <span data-ttu-id="a7282-130">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-130">Open the app.js file.</span></span>

5. <span data-ttu-id="a7282-131">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a7282-131">Replace the `TODO1` with the following code.</span></span> <span data-ttu-id="a7282-132">このコードでは、ユーザーの Excel のバージョンが、このチュートリアルのシリーズで使用する API をすべて含んでいるバージョンの Excel.js をサポートしているかどうかを調べます。</span><span class="sxs-lookup"><span data-stu-id="a7282-132">This code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use.</span></span> <span data-ttu-id="a7282-133">運用アドインでは、未サポートの API を呼び出す UI を非表示または無効化する条件ブロックの本体を使用してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-133">In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs.</span></span> <span data-ttu-id="a7282-134">これにより、ユーザーは、そのユーザーの Excel のバージョンでサポートされているアドインの部分を使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="a7282-134">This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }
    ```

6. <span data-ttu-id="a7282-135">`TODO2` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a7282-135">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#create-table').click(createTable);
    ```

7. <span data-ttu-id="a7282-136">`TODO3` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a7282-136">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="a7282-137">注:</span><span class="sxs-lookup"><span data-stu-id="a7282-137">Note:</span></span>

   - <span data-ttu-id="a7282-138">Excel.js のビジネス ロジックは、`Excel.run` に渡される関数に追加されます。</span><span class="sxs-lookup"><span data-stu-id="a7282-138">Your Excel.js business logic will be added to the function that is passed to `Excel.run`.</span></span> <span data-ttu-id="a7282-139">このロジックは、すぐには実行されません。</span><span class="sxs-lookup"><span data-stu-id="a7282-139">This logic does not execute immediately.</span></span> <span data-ttu-id="a7282-140">その代わりに、保留中のコマンドのキューに追加されます。</span><span class="sxs-lookup"><span data-stu-id="a7282-140">Instead, it is added to a queue of pending commands.</span></span>

   - <span data-ttu-id="a7282-141">`context.sync` メソッドは、キューに登録されたすべてのコマンドを実行するために Excel に送信します。</span><span class="sxs-lookup"><span data-stu-id="a7282-141">The `context.sync` method sends all queued commands to Excel for execution.</span></span>

   - <span data-ttu-id="a7282-142">`Excel.run` の後に `catch` ブロックを続けます。</span><span class="sxs-lookup"><span data-stu-id="a7282-142">The `Excel.run` is followed by a `catch` block.</span></span> <span data-ttu-id="a7282-143">これは、どのような場合にも当てはまるベスト プラクティスです。</span><span class="sxs-lookup"><span data-stu-id="a7282-143">This is a best practice that you should always follow.</span></span> 

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

8. <span data-ttu-id="a7282-p106">`TODO4` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-p106">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="a7282-146">このコードでは、ワークシートのテーブル コレクションの `add` メソッドを使用してテーブルを作成します。このコレクションは空であったとしても常に存在します。</span><span class="sxs-lookup"><span data-stu-id="a7282-146">The code creates a table by using `add` method of a worksheet's table collection, which always exists even if it is empty.</span></span> <span data-ttu-id="a7282-147">これは、Excel.js オブジェクトの標準的な作成方法です。</span><span class="sxs-lookup"><span data-stu-id="a7282-147">This is the standard way that Excel.js objects are created.</span></span> <span data-ttu-id="a7282-148">クラス コンストラクタ API は存在しません。Excel オブジェクトを作成するために、`new` 演算子は使用できません。</span><span class="sxs-lookup"><span data-stu-id="a7282-148">There are no class constructor APIs, and you never use a `new` operator to create an Excel object.</span></span> <span data-ttu-id="a7282-149">その代わりに、親コレクションにオブジェクトを追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-149">Instead, you add to a parent collection object.</span></span>

   - <span data-ttu-id="a7282-150">`add` メソッドの最初のパラメーターは、テーブルの先頭行のみの範囲です。そのテーブルで最終的に使用する全体の範囲ではありません。</span><span class="sxs-lookup"><span data-stu-id="a7282-150">The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use.</span></span> <span data-ttu-id="a7282-151">これは、アドインでデータ行を設定するときに (この後の手順で実行します)、既存の行のセルに値を書き込むのではなく、新しい行をテーブルに追加するためです。</span><span class="sxs-lookup"><span data-stu-id="a7282-151">This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows.</span></span> <span data-ttu-id="a7282-152">多くの場合、テーブルの作成時には、そのテーブルに含める行の数がわからないため、このパターンのほうが一般的になります。</span><span class="sxs-lookup"><span data-stu-id="a7282-152">This is a more common pattern because the number of rows that a table will have is often not known when the table is created.</span></span>

   - <span data-ttu-id="a7282-153">テーブルの名前は、ワークシート内だけでなくブック全体で一意にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="a7282-153">Table names must be unique across the entire workbook, not just the worksheet.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

9. <span data-ttu-id="a7282-p109">`TODO5` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-p109">Replace `TODO5` with the following code. Note:</span></span>

   - <span data-ttu-id="a7282-156">範囲に含まれるセルの値は、配列の配列で設定します。</span><span class="sxs-lookup"><span data-stu-id="a7282-156">The cell values of a range are set with an array of arrays.</span></span>

   - <span data-ttu-id="a7282-157">テーブル内に新しい行を作成するために、そのテーブルの行コレクションの `add` メソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="a7282-157">New rows are created in a table by calling the `add` method of the table's row collection.</span></span> <span data-ttu-id="a7282-158">`add` の 1 回の呼び出しで複数の行を追加できるようにするには、2 番目のパラメーターとして渡す親配列に複数のセル値の配列を含めます。</span><span class="sxs-lookup"><span data-stu-id="a7282-158">You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.</span></span>

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

10. <span data-ttu-id="a7282-p111">`TODO6` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-p111">Replace `TODO6` with the following code. Note:</span></span>

   - <span data-ttu-id="a7282-161">このコードでは、ゼロから始まるインデックスをテーブルの列コレクションの `getItemAt` メソッドに渡すことで、**Amount** 列への参照を取得します。</span><span class="sxs-lookup"><span data-stu-id="a7282-161">The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection.</span></span>

     > [!NOTE]
     > <span data-ttu-id="a7282-162">Excel.js のコレクション オブジェクト (`TableCollection`、`WorksheetCollection`、`TableColumnCollection` など) には、`items` プロパティがあります。このプロパティは、子オブジェクト タイプ (`Table`、`Worksheet`、`TableColumn` など) の配列ですが、`*Collection` オブジェクト自体は配列ではありません。</span><span class="sxs-lookup"><span data-stu-id="a7282-162">Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.</span></span>

   - <span data-ttu-id="a7282-163">その次に、コードでは、**Amount** 列の範囲を小数点以下 2 桁までのユーロとして書式設定します。</span><span class="sxs-lookup"><span data-stu-id="a7282-163">The code then formats the range of the **Amount** column as Euros to the second decimal.</span></span> 

   - <span data-ttu-id="a7282-164">最後に、列の幅と行の高さが最長 (最高) のデータ アイテムを収めるために十分な大きさになるようにしています。</span><span class="sxs-lookup"><span data-stu-id="a7282-164">Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item.</span></span> <span data-ttu-id="a7282-165">このコードでは、書式設定のために `Range` オブジェクトを取得している点に注目してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-165">Notice that the code must get `Range` objects to format.</span></span> <span data-ttu-id="a7282-166">`TableColumn` オブジェクトと `TableRow` オブジェクトには、書式設定のプロパティがありません。</span><span class="sxs-lookup"><span data-stu-id="a7282-166">`TableColumn` and `TableRow` objects do not have format properties.</span></span>

        ```js
        expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();
        ```

### <a name="test-the-add-in"></a><span data-ttu-id="a7282-167">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="a7282-167">Test the add-in</span></span>

1. <span data-ttu-id="a7282-168">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="a7282-168">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

2. <span data-ttu-id="a7282-169">`npm run build` コマンドを実行して、ES6 ソース コードを Internet Explorer でサポートされている以前のバージョンの JavaScript にトランスパイルします (これは、Excel アドインを実行するために Excel の内部で使用されます)。</span><span class="sxs-lookup"><span data-stu-id="a7282-169">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="a7282-170">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="a7282-170">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="a7282-171">次のいずれかの方法を使用して、アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="a7282-171">Sideload the add-in by using one of the following methods:</span></span>

    - <span data-ttu-id="a7282-172">Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="a7282-172">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>

    - <span data-ttu-id="a7282-173">Excel Online:[Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="a7282-173">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>

    - <span data-ttu-id="a7282-174">iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="a7282-174">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

5. <span data-ttu-id="a7282-175">**[ホーム]** メニューで、**[作業ウィンドウの表示]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a7282-175">On the **Home** menu, choose **Show Taskpane**.</span></span>

6. <span data-ttu-id="a7282-176">作業ウィンドウで、**[Create Table]** (表の作成) を選択します。</span><span class="sxs-lookup"><span data-stu-id="a7282-176">In the task pane, choose **Create Table**.</span></span>

    ![Excel チュートリアル - テーブルの作成](../images/excel-tutorial-create-table.png)

## <a name="filter-and-sort-a-table"></a><span data-ttu-id="a7282-178">テーブルのフィルター処理と並べ替え</span><span class="sxs-lookup"><span data-stu-id="a7282-178">Filter and sort a table</span></span>

<span data-ttu-id="a7282-179">チュートリアルのこの手順では、以前に作成したテーブルをフィルター処理したり並べ替えたりします。</span><span class="sxs-lookup"><span data-stu-id="a7282-179">In this step of the tutorial, you'll filter and sort the table that you created previously.</span></span>

### <a name="filter-the-table"></a><span data-ttu-id="a7282-180">表のフィルター処理</span><span class="sxs-lookup"><span data-stu-id="a7282-180">Filter the table</span></span>

1. <span data-ttu-id="a7282-181">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-181">Open the project in your code editor.</span></span>

2. <span data-ttu-id="a7282-182">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-182">Open the file index.html.</span></span>

3. <span data-ttu-id="a7282-183">`create-table` ボタンを格納している `div` の直下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-183">Just below the `div` that contains the `create-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="filter-table">Filter Table</button>
    </div>
    ```

4. <span data-ttu-id="a7282-184">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-184">Open the app.js file.</span></span>

5. <span data-ttu-id="a7282-185">`create-table` ボタンにクリック ハンドラーを割り当てる行の直下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-185">Just below the line that assigns a click handler to the `create-table` button, add the following code:</span></span>

    ```js
    $('#filter-table').click(filterTable);
    ```

6. <span data-ttu-id="a7282-186">`createTable` 関数の直下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-186">Just below the `createTable` function, add the following function:</span></span>

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

7. <span data-ttu-id="a7282-p113">`TODO1` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-p113">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="a7282-189">このコードでは最初に、`getItem` メソッドに列名を渡すことによって、フィルター処理が必要な列への参照を取得します。`createTable` メソッドが行うように、列のインデックスを `getItemAt` メソッドに渡すわけではありません。</span><span class="sxs-lookup"><span data-stu-id="a7282-189">The code first gets a reference to the column that needs filtering by passing the column name to the `getItem` method, instead of passing its index to the `getItemAt` method as the `createTable` method does.</span></span> <span data-ttu-id="a7282-190">ユーザーは表の列を移動させることができるので、表を作成した後、指定したインデックスにある列が変わってしまう可能性があります。</span><span class="sxs-lookup"><span data-stu-id="a7282-190">Since users can move table columns, the column at a given index might change after the table is created.</span></span> <span data-ttu-id="a7282-191">そのため、列名を使用して列への参照を取得するほうが安全です。</span><span class="sxs-lookup"><span data-stu-id="a7282-191">Hence, it is safer to use the column name to get a reference to the column.</span></span> <span data-ttu-id="a7282-192">前のチュートリアルでは、表を作成するのとまったく同じ方法で `getItemAt` を使用したため、ユーザーが列を移動させた可能性はなく、よって安全に使用できました。</span><span class="sxs-lookup"><span data-stu-id="a7282-192">We used `getItemAt` safely in the preceding tutorial, because we used it in the very same method that creates the table, so there is no chance that a user has moved the column.</span></span>

   - <span data-ttu-id="a7282-193">`applyValuesFilter` メソッドは、`Filter` オブジェクトのフィルター処理方法の 1 つです。</span><span class="sxs-lookup"><span data-stu-id="a7282-193">The `applyValuesFilter` method is one of several filtering methods on the `Filter` object.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);
    ``` 

### <a name="sort-the-table"></a><span data-ttu-id="a7282-194">表の並べ替え</span><span class="sxs-lookup"><span data-stu-id="a7282-194">Sort the table</span></span>

1. <span data-ttu-id="a7282-195">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-195">Open the file index.html.</span></span>

2. <span data-ttu-id="a7282-196">`filter-table` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-196">Below the `div` that contains the `filter-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="sort-table">Sort Table</button>
    </div>
    ```

3. <span data-ttu-id="a7282-197">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-197">Open the app.js file.</span></span>

4. <span data-ttu-id="a7282-198">`filter-table` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-198">Below the line that assigns a click handler to the `filter-table` button, add the following code:</span></span>

    ```js
    $('#sort-table').click(sortTable);
    ```

5. <span data-ttu-id="a7282-199">`filterTable` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-199">Below the `filterTable` function add the following function.</span></span>

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

6. <span data-ttu-id="a7282-p115">`TODO1` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-p115">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="a7282-202">アドインで並べ替えるのは Merchant 列のみであるため、このコードでは、1 つのメンバーだけを含む `SortField` オブジェクトの配列を作成します。</span><span class="sxs-lookup"><span data-stu-id="a7282-202">The code creates an array of `SortField` objects which has just one member since the add-in only sorts on the Merchant column.</span></span>

   - <span data-ttu-id="a7282-203">`SortField` オブジェクトの `key` プロパティは、並べ替える対象列の 0 から始まるインデックスです。</span><span class="sxs-lookup"><span data-stu-id="a7282-203">The `key` property of a `SortField` object is the zero-based index of the column to sort-on.</span></span>

   - <span data-ttu-id="a7282-204">`Table` の `sort` メンバーは、`TableSort` オブジェクトであり、メソッドではありません。</span><span class="sxs-lookup"><span data-stu-id="a7282-204">The `sort` member of a `Table` is a `TableSort` object, not a method.</span></span> <span data-ttu-id="a7282-205">`TableSort` オブジェクトの `apply` メソッドには、`SortField` が渡されます。</span><span class="sxs-lookup"><span data-stu-id="a7282-205">The `SortField`s are passed to the `TableSort` object's `apply` method.</span></span>

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

### <a name="test-the-add-in"></a><span data-ttu-id="a7282-206">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="a7282-206">Test the add-in</span></span>

1. <span data-ttu-id="a7282-207">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、**Ctrl + C** を 2 回入力して実行中の Web サーバーを停止します。</span><span class="sxs-lookup"><span data-stu-id="a7282-207">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="a7282-208">それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="a7282-208">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="a7282-209">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。</span><span class="sxs-lookup"><span data-stu-id="a7282-209">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="a7282-210">そのためには、ビルド コマンドの入力を求めるプロンプトが表示されるように、サーバー プロセスを強制終了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a7282-210">In order to do this, you need to kill the server process so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="a7282-211">ビルド後に、サーバーを再起動します。</span><span class="sxs-lookup"><span data-stu-id="a7282-211">After the build, you restart the server.</span></span> <span data-ttu-id="a7282-212">次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="a7282-212">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="a7282-213">`npm run build` コマンドを実行して、ES6 ソース コードを Internet Explorer でサポートされている以前のバージョンの JavaScript にトランスパイルします (これは、Excel アドインを実行するために Excel の内部で使用されます)。</span><span class="sxs-lookup"><span data-stu-id="a7282-213">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="a7282-214">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="a7282-214">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="a7282-215">作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-215">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="a7282-216">何らかの理由から開いているワークシートに表が含まれていない場合は、作業ウィンドウの **[Create Table]** (表の作成) ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="a7282-216">If for any reason the table is not in the open worksheet, in the task pane, choose **Create Table**.</span></span>

6. <span data-ttu-id="a7282-217">**[Filter Table]** (表のフィルター) ボタンと **[Sort Table]** (表の並べ替え) ボタンを任意の順序で選択します。</span><span class="sxs-lookup"><span data-stu-id="a7282-217">Choose the **Filter Table** and **Sort Table** buttons, in either order.</span></span>

    ![Excel のチュートリアル - テーブルのフィルター処理と並べ替え](../images/excel-tutorial-filter-and-sort-table.png)

## <a name="create-a-chart"></a><span data-ttu-id="a7282-219">グラフの作成</span><span class="sxs-lookup"><span data-stu-id="a7282-219">Create a chart</span></span>

<span data-ttu-id="a7282-220">チュートリアルのこの手順では、前の手順で作成したテーブルのデータを使用してグラフを作成して、そのグラフの書式を設定します。</span><span class="sxs-lookup"><span data-stu-id="a7282-220">In this step of the tutorial, you'll create a chart using data from the table that you created previously, and then format the chart.</span></span>

### <a name="chart-a-chart-using-table-data"></a><span data-ttu-id="a7282-221">テーブルのデータを使用してグラフを作成する</span><span class="sxs-lookup"><span data-stu-id="a7282-221">Chart a chart using table data</span></span>

1. <span data-ttu-id="a7282-222">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-222">Open the project in your code editor.</span></span>

2. <span data-ttu-id="a7282-223">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-223">Open the file index.html.</span></span>

3. <span data-ttu-id="a7282-224">`sort-table` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-224">Below the `div` that contains the `sort-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="create-chart">Create Chart</button>
    </div>
    ```

4. <span data-ttu-id="a7282-225">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-225">Open the app.js file.</span></span>

5. <span data-ttu-id="a7282-226">`sort-chart` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-226">Below the line that assigns a click handler to the `sort-chart` button, add the following code:</span></span>

    ```js
    $('#create-chart').click(createChart);
    ```

6. <span data-ttu-id="a7282-227">`sortTable` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-227">Below the `sortTable` function add the following function.</span></span>

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

7. <span data-ttu-id="a7282-p119">`TODO1` を次のコードに置き換えます。ヘッダー行を除外するために、このコードでは、`getRange` メソッドではなく `Table.getDataBodyRange` メソッドを使用してグラフを作成するデータの範囲を取得しています。</span><span class="sxs-lookup"><span data-stu-id="a7282-p119">Replace `TODO1` with the following code. Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var dataRange = expensesTable.getDataBodyRange();
    ```

8. <span data-ttu-id="a7282-230">`TODO2` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a7282-230">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="a7282-231">次のパラメーターに注意してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-231">Note the following parameters:</span></span>

   - <span data-ttu-id="a7282-p121">`add` への最初のパラメーターでは、グラフの種類を指定します。数十種類あります。</span><span class="sxs-lookup"><span data-stu-id="a7282-p121">The first parameter to the `add` method specifies the type of chart. There are several dozen types.</span></span>

   - <span data-ttu-id="a7282-234">2 番目のパラメーターでは、グラフに含めるデータの範囲を指定します。</span><span class="sxs-lookup"><span data-stu-id="a7282-234">The second parameter specifies the range of data to include in the chart.</span></span>

   - <span data-ttu-id="a7282-235">3 番目のパラメーターでは、テーブルからの一連のデータ ポイントを行方向と列方向のどちらでグラフ化する必要があるかを決定します。</span><span class="sxs-lookup"><span data-stu-id="a7282-235">The third parameter determines whether a series of data points from the table should be charted row-wise or column-wise.</span></span> <span data-ttu-id="a7282-236">オプション `auto` は、最適な方法を判断するように Excel に指示します。</span><span class="sxs-lookup"><span data-stu-id="a7282-236">The option `auto` tells Excel to decide the best method.</span></span>

    ```js
    var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ```

9. <span data-ttu-id="a7282-237">`TODO3` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a7282-237">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="a7282-238">このコードのほとんどの部分は、わかりやすく説明不要なものです。</span><span class="sxs-lookup"><span data-stu-id="a7282-238">Most of this code is self-explanatory.</span></span> <span data-ttu-id="a7282-239">注意:</span><span class="sxs-lookup"><span data-stu-id="a7282-239">Note:</span></span>
   
   - <span data-ttu-id="a7282-240">`setPosition` メソッドへのパラメーターでは、グラフを収容するワークシート領域の左上と右下のセルを指定します。</span><span class="sxs-lookup"><span data-stu-id="a7282-240">The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart.</span></span> <span data-ttu-id="a7282-241">Excel では、所定の空間内でグラフの外観を整えるために線幅などを調整できます。</span><span class="sxs-lookup"><span data-stu-id="a7282-241">Excel can adjust things like line width to make the chart look good in the space it has been given.</span></span>
   
   - <span data-ttu-id="a7282-242">"series" は、テーブルに含まれる列からのデータ ポイントのセットです。</span><span class="sxs-lookup"><span data-stu-id="a7282-242">A "series" is a set of data points from a column of the table.</span></span> <span data-ttu-id="a7282-243">このテーブルに存在する文字列以外の列は 1 列のみであるため、Excel は、その列がグラフ化するデータ ポイントの唯一の列であることを推測します。</span><span class="sxs-lookup"><span data-stu-id="a7282-243">Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart.</span></span> <span data-ttu-id="a7282-244">その他の列は、グラフのラベルとして解釈されます。</span><span class="sxs-lookup"><span data-stu-id="a7282-244">It interprets the other columns as chart labels.</span></span> <span data-ttu-id="a7282-245">そのため、グラフの series は 1 つ存在することになり、インデックス 0 を含みます。</span><span class="sxs-lookup"><span data-stu-id="a7282-245">So there will be just one series in the chart and it will have index 0.</span></span> <span data-ttu-id="a7282-246">これに、"Value in €" のラベルを付けます。</span><span class="sxs-lookup"><span data-stu-id="a7282-246">This is the one to label with "Value in €".</span></span>

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in €';
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="a7282-247">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="a7282-247">Test the add-in</span></span>

1. <span data-ttu-id="a7282-248">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、**Ctrl + C** を 2 回入力して実行中の Web サーバーを停止します。</span><span class="sxs-lookup"><span data-stu-id="a7282-248">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="a7282-249">それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="a7282-249">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="a7282-250">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。</span><span class="sxs-lookup"><span data-stu-id="a7282-250">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="a7282-251">そのためには、ビルド コマンドの入力を求めるプロンプトが表示されるように、サーバー プロセスを強制終了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a7282-251">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="a7282-252">ビルド後に、サーバーを再起動します。</span><span class="sxs-lookup"><span data-stu-id="a7282-252">After the build, you restart the server.</span></span> <span data-ttu-id="a7282-253">次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="a7282-253">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="a7282-254">`npm run build` コマンドを実行して、ES6 ソース コードを Internet Explorer でサポートされている以前のバージョンの JavaScript にトランスパイルします (これは、Excel アドインを実行するために Excel の内部で使用されます)。</span><span class="sxs-lookup"><span data-stu-id="a7282-254">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="a7282-255">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="a7282-255">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="a7282-256">作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-256">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="a7282-257">何らかの理由から開いているワークシートにテーブルが含まれていない場合は、**[Create Table]** (テーブルの作成) ボタンをクリックしてから、**[Filter Table]** (テーブルのフィルター) ボタンと **[Sort Table]** (テーブルの並べ替え) ボタンを任意の順序でクリックします。</span><span class="sxs-lookup"><span data-stu-id="a7282-257">If for any reason the table is not in the open worksheet, in the task pane, choose **Create Table** and then **Filter Table** and **Sort Table** buttons, in either order.</span></span>

6. <span data-ttu-id="a7282-258">**[グラフの作成]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="a7282-258">Choose the **Create Chart** button.</span></span> <span data-ttu-id="a7282-259">グラフが作成され、フィルターが適用された行からのデータのみが含まれます。</span><span class="sxs-lookup"><span data-stu-id="a7282-259">A chart is created and only the data from the rows that have been filtered are included.</span></span> <span data-ttu-id="a7282-260">データ ポイントの下側のラベルは、グラフの並べ替え順序になります。つまり、[Merchant] (業者) の名前の逆アルファベット順になります。</span><span class="sxs-lookup"><span data-stu-id="a7282-260">The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.</span></span>

    ![Excel チュートリアル - グラフの作成](../images/excel-tutorial-create-chart.png)

## <a name="freeze-a-table-header"></a><span data-ttu-id="a7282-262">テーブルのヘッダーの固定</span><span class="sxs-lookup"><span data-stu-id="a7282-262">Freeze a table header in place</span></span>

<span data-ttu-id="a7282-263">テーブルがとても長く、行を参照するためにスクロールしなければならない場合、ヘッダー行が画面の外に移動して見えなくなることがあります。</span><span class="sxs-lookup"><span data-stu-id="a7282-263">When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight.</span></span> <span data-ttu-id="a7282-264">チュートリアルのこの手順では、以前に作成した表のヘッダー行を固定して、ワークシートを下にスクロールしても表示されるようにします。</span><span class="sxs-lookup"><span data-stu-id="a7282-264">In this step of the tutorial, you'll freeze the header row of the table that you created previously, so that it remains visible even as the user scrolls down the worksheet.</span></span>

### <a name="freeze-the-tables-header-row"></a><span data-ttu-id="a7282-265">表のヘッダー行を固定する</span><span class="sxs-lookup"><span data-stu-id="a7282-265">Freeze the table's header row</span></span>

1. <span data-ttu-id="a7282-266">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-266">Open the project in your code editor.</span></span>

2. <span data-ttu-id="a7282-267">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-267">Open the file index.html.</span></span>

3. <span data-ttu-id="a7282-268">`create-chart` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-268">Below the `div` that contains the `create-chart` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="freeze-header">Freeze Header</button>
    </div>
    ```

4. <span data-ttu-id="a7282-269">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-269">Open the app.js file.</span></span>

5. <span data-ttu-id="a7282-270">`create-chart` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-270">Below the line that assigns a click handler to the `create-chart` button, add the following code:</span></span>

    ```js
    $('#freeze-header').click(freezeHeader);
    ```

6. <span data-ttu-id="a7282-271">`createChart` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-271">Below the `createChart` function add the following function:</span></span>

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

7. <span data-ttu-id="a7282-p130">`TODO1` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-p130">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="a7282-274">`Worksheet.freezePanes` コレクションは、ワークシートのスクロール操作時に、ワークシート上でピン留めつまり固定される一式のペインのことです。</span><span class="sxs-lookup"><span data-stu-id="a7282-274">The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.</span></span>

   - <span data-ttu-id="a7282-p131">`freezeRows` メソッドでは、上から数えた行数を、ピン留めする位置のパラメーターとして使用します。`1` を渡して最初の行を適所にピン留めします。</span><span class="sxs-lookup"><span data-stu-id="a7282-p131">The `freezeRows` method takes as a parameter the number of rows, from the top that are to be pinned in place. We pass `1` to pin the first row in place.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="a7282-277">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="a7282-277">Test the add-in</span></span>

1. <span data-ttu-id="a7282-278">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、**Ctrl + C** を 2 回入力して実行中の Web サーバーを停止します。</span><span class="sxs-lookup"><span data-stu-id="a7282-278">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="a7282-279">それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="a7282-279">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="a7282-280">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。</span><span class="sxs-lookup"><span data-stu-id="a7282-280">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="a7282-281">そのためには、ビルド コマンドの入力を求めるプロンプトが表示されるように、サーバー プロセスを強制終了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a7282-281">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="a7282-282">ビルド後に、サーバーを再起動します。</span><span class="sxs-lookup"><span data-stu-id="a7282-282">After the build, you restart the server.</span></span> <span data-ttu-id="a7282-283">次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="a7282-283">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="a7282-284">`npm run build` コマンドを実行して、ES6 ソース コードを Internet Explorer でサポートされている以前のバージョンの JavaScript にトランスパイルします (これは、Excel アドインを実行するために Excel の内部で使用されます)。</span><span class="sxs-lookup"><span data-stu-id="a7282-284">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="a7282-285">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="a7282-285">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="a7282-286">作業ウィンドウを再読み込みするために、そのウィンドウを閉じ、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-286">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="a7282-287">ワークシート内に表があれば、削除します。</span><span class="sxs-lookup"><span data-stu-id="a7282-287">If the table is in the worksheet, delete it.</span></span>

6. <span data-ttu-id="a7282-288">作業ウィンドウで、**[Create Table]** (表の作成) を選択します。</span><span class="sxs-lookup"><span data-stu-id="a7282-288">In the task pane, choose **Create Table**.</span></span>

7. <span data-ttu-id="a7282-289">**[Freeze Header]** (ヘッダーを固定) ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="a7282-289">Choose the **Freeze Header** button.</span></span>

8. <span data-ttu-id="a7282-290">ヘッダー以降の行が画面の外に出て見えなくなるまでワークシートを下にスクロールしても、表のヘッダーが最上部に表示されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="a7282-290">Scroll down the worksheet enough to to see that the table header remains visible at the top even when the higher rows scroll out of sight.</span></span>

    ![Excel のチュートリアル - ヘッダーの固定](../images/excel-tutorial-freeze-header.png)

## <a name="protect-a-worksheet"></a><span data-ttu-id="a7282-292">ワークシートの保護</span><span class="sxs-lookup"><span data-stu-id="a7282-292">Protect a worksheet from changes</span></span>

<span data-ttu-id="a7282-293">チュートリアルのこの手順では、リボンに別のボタンを追加します。このボタンをクリックすると、ワークシートの保護のオン/オフが切り替わるように定義した関数が実行されるようにします。</span><span class="sxs-lookup"><span data-stu-id="a7282-293">In this step of the tutorial, you'll add another button to the ribbon that, when chosen, executes a function that you'll define to toggle worksheet protection on and off.</span></span>

### <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a><span data-ttu-id="a7282-294">2 つ目のリボン ボタンを追加するようにマニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="a7282-294">Configure the manifest to add a second ribbon button</span></span>

1. <span data-ttu-id="a7282-295">マニフェスト ファイル my-office-add-in-manifest.xml を開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-295">Open the manifest file my-office-add-in-manifest.xml.</span></span>

2. <span data-ttu-id="a7282-296">`<Control>` 要素を検索します。</span><span class="sxs-lookup"><span data-stu-id="a7282-296">Find the `<Control>` element.</span></span> <span data-ttu-id="a7282-297">この要素では、アドインの起動に使用している **[ホーム]** リボンの **[作業ウィンドウの表示]** ボタンを定義しています。</span><span class="sxs-lookup"><span data-stu-id="a7282-297">This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in.</span></span> <span data-ttu-id="a7282-298">ここでは、**[ホーム]** リボンの同じグループに 2 つ目のボタンを追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-298">We're going to add a second button to the same group on the **Home** ribbon.</span></span> <span data-ttu-id="a7282-299">Control 終了タグ (`</Control>`) と Group 終了タグ (`</Group>`) の間に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-299">In between the end Control tag (`</Control>`) and the end Group tag (`</Group>`), add the following markup.</span></span>

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

3. <span data-ttu-id="a7282-300">`TODO1` は文字列に置き換えて、このマニフェスト ファイル内で一意の ID をボタンに割り当てます。</span><span class="sxs-lookup"><span data-stu-id="a7282-300">Replace `TODO1` with a string that gives the button an ID that is unique within this manifest file.</span></span> <span data-ttu-id="a7282-301">このボタンでは、ワークシートの保護のオン/オフを切り替える予定なので、"ToggleProtection" を使用することにします。</span><span class="sxs-lookup"><span data-stu-id="a7282-301">Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection".</span></span> <span data-ttu-id="a7282-302">作業が完了すると、Control 開始タグの全体は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="a7282-302">When you are done, the entire start Control tag should look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. <span data-ttu-id="a7282-303">その次の 3 つの `TODO` では、"resid" を設定します ("resid" はリソース ID の略号です)。</span><span class="sxs-lookup"><span data-stu-id="a7282-303">The next three `TODO`s set "resid"s, which is short for resource ID.</span></span> <span data-ttu-id="a7282-304">リソースは文字列です。これら 3 つの文字列は、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="a7282-304">A resource is a string, and you'll create these three strings in a later step.</span></span> <span data-ttu-id="a7282-305">ここでは、そのリソースに ID を割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="a7282-305">For now, you need to give IDs to the resources.</span></span> <span data-ttu-id="a7282-306">ボタンのラベルは "Toggle Protection" と表示されるようにしますが、この文字列の *ID* は "ProtectionButtonLabel" にします。そのため、完成した `Label` 要素は次のコードのようになります。</span><span class="sxs-lookup"><span data-stu-id="a7282-306">The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the completed `Label` element should look like the following code:</span></span>

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. <span data-ttu-id="a7282-307">`SuperTip` 要素では、このボタンのツール ヒントを定義します。</span><span class="sxs-lookup"><span data-stu-id="a7282-307">The `SuperTip` element defines the tool tip for the button.</span></span> <span data-ttu-id="a7282-308">ツール ヒントのタイトルはボタンのラベルと同じにする必要があるため、リソース ID にはまったく同じ "ProtectionButtonLabel" を使用することにします。</span><span class="sxs-lookup"><span data-stu-id="a7282-308">The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel".</span></span> <span data-ttu-id="a7282-309">ツール ヒントの説明は、"Click to turn protection of the worksheet on and off" にする予定です。</span><span class="sxs-lookup"><span data-stu-id="a7282-309">The tool tip description will be "Click to turn protection of the worksheet on and off".</span></span> <span data-ttu-id="a7282-310">ただし、`ID` は "ProtectionButtonToolTip" にします。</span><span class="sxs-lookup"><span data-stu-id="a7282-310">But the `ID` should be "ProtectionButtonToolTip".</span></span> <span data-ttu-id="a7282-311">作業が完了すると、`SuperTip` マークアップの全体は次のコードのようになります。</span><span class="sxs-lookup"><span data-stu-id="a7282-311">So, when you are done, the whole `SuperTip` markup should look like the following code:</span></span> 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > <span data-ttu-id="a7282-312">運用アドインでは、異なる 2 つのボタンに同じアイコンを使用することは避けたいところですが、このチュートリアルでは説明を簡単にするために同じアイコンを使用します。</span><span class="sxs-lookup"><span data-stu-id="a7282-312">In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that.</span></span> <span data-ttu-id="a7282-313">そのため、この新しい `Control` の `Icon` マークアップは、単に既存の `Control` から `Icon` 要素をコピーします。</span><span class="sxs-lookup"><span data-stu-id="a7282-313">So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`.</span></span> 

6. <span data-ttu-id="a7282-314">既にマニフェストに存在している元の `Control` 要素の内側にある `Action` 要素では、その要素のタイプが `ShowTaskpane` に設定されていますが、新しいボタンで作業ウィンドウを開く予定はありません。このボタンでは、この後の手順で作成するカスタム関数を実行する予定です。</span><span class="sxs-lookup"><span data-stu-id="a7282-314">The `Action` element inside the original `Control` element that was already present in the manifest, has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step.</span></span> <span data-ttu-id="a7282-315">そのため、`TODO5` は、カスタム関数をトリガーするボタンのアクション タイプである `ExecuteFunction` に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a7282-315">So replace `TODO5` with `ExecuteFunction` which is the action type for buttons that trigger custom functions.</span></span> <span data-ttu-id="a7282-316">`Action` 開始タグは次のようになります。</span><span class="sxs-lookup"><span data-stu-id="a7282-316">The start `Action` tag should look like the following code:</span></span>
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. <span data-ttu-id="a7282-317">元の `Action` 要素には、作業ウィンドウ ID を指定する子要素と、作業ウィンドウで開かれるページの URL を指定する子要素があります。</span><span class="sxs-lookup"><span data-stu-id="a7282-317">The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane.</span></span> <span data-ttu-id="a7282-318">ただし、`ExecuteFunction` タイプの `Action` 要素には、実行を制御する関数の名前を指定する子要素を 1 つ含めます。</span><span class="sxs-lookup"><span data-stu-id="a7282-318">But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes.</span></span> <span data-ttu-id="a7282-319">その関数は、`toggleProtection` という名前にして、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="a7282-319">You'll create that function in a later step, and it will be called `toggleProtection`.</span></span> <span data-ttu-id="a7282-320">そのために、`TODO6` を次のマークアップに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a7282-320">So, replace `TODO6` with the following markup:</span></span>
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    <span data-ttu-id="a7282-321">`Control` マークアップの全体は、次のようになりました。</span><span class="sxs-lookup"><span data-stu-id="a7282-321">The entire `Control` markup should now look like the following:</span></span>

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

8. <span data-ttu-id="a7282-322">マニフェストの `Resources` セクションまで下にスクロールします。</span><span class="sxs-lookup"><span data-stu-id="a7282-322">Scroll down to the `Resources` section of the manifest.</span></span>

9. <span data-ttu-id="a7282-323">`bt:ShortStrings` 要素の子として、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-323">Add the following markup as a child of the `bt:ShortStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. <span data-ttu-id="a7282-324">`bt:LongStrings` 要素の子として、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-324">Add the following markup as a child of the `bt:LongStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. <span data-ttu-id="a7282-325">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="a7282-325">Save the file.</span></span>

### <a name="create-the-function-that-protects-the-sheet"></a><span data-ttu-id="a7282-326">シートを保護する関数を作成する</span><span class="sxs-lookup"><span data-stu-id="a7282-326">Create the function that protects the sheet</span></span>

1. <span data-ttu-id="a7282-327">ファイル \function-file\function-file.js を開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-327">Open the file \function-file\function-file.js.</span></span>

2. <span data-ttu-id="a7282-328">このファイルには、即時実行関数式 (IIFE) が既に含まれています。</span><span class="sxs-lookup"><span data-stu-id="a7282-328">The file already has an Immediately Invoked Function Expression (IFFE).</span></span> <span data-ttu-id="a7282-329">カスタムの初期化ロジックは必要ないため、`Office.initialize` に割り当てられた関数は空のままにしておきます </span><span class="sxs-lookup"><span data-stu-id="a7282-329">No custom initialization logic is needed, so leave the function that is assigned to `Office.initialize` with an empty body.</span></span> <span data-ttu-id="a7282-330">(ただし、削除してはいけません。</span><span class="sxs-lookup"><span data-stu-id="a7282-330">(But do not delete it.</span></span> <span data-ttu-id="a7282-331">`Office.initialize` プロパティは Null や未定義にすることはできません)。*IIFE の外側に*、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-331">The `Office.initialize` property cannot be null or undefined.) *Outside of the IIFE*, add the following code.</span></span> <span data-ttu-id="a7282-332">メソッドに `args` パラメーターを指定していることと、メソッドの最後のほうの行で `args.completed` を呼び出していることに注目してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-332">Note that we specify an `args` parameter to the method and the very last line of the method calls `args.completed`.</span></span> <span data-ttu-id="a7282-333">**ExecuteFunction** タイプのすべてのアドイン コマンドでは、これが要件になります。</span><span class="sxs-lookup"><span data-stu-id="a7282-333">This is a requirement for all add-in commands of type **ExecuteFunction**.</span></span> <span data-ttu-id="a7282-334">これにより、関数が終了したことと、UI が再度応答可能になることを Office ホスト アプリケーションに通知します。</span><span class="sxs-lookup"><span data-stu-id="a7282-334">It signals the Office host application that the function has finished and the UI can become responsive again.</span></span>

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

3. <span data-ttu-id="a7282-335">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a7282-335">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="a7282-336">このコードでは、標準の切り替えパターンで、ワークシート オブジェクトの protection プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="a7282-336">This code uses the worksheet object's protection property in a standard toggle pattern.</span></span> <span data-ttu-id="a7282-337">`TODO2` については、次のセクションで説明します。</span><span class="sxs-lookup"><span data-stu-id="a7282-337">The `TODO2` will be explained in the next section.</span></span>

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

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="a7282-338">ドキュメントのプロパティを作業ウィンドウのスクリプト オブジェクトにフェッチするコードを追加する</span><span class="sxs-lookup"><span data-stu-id="a7282-338">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="a7282-339">このチュートリアルのシリーズで前述したすべての関数では、Office ドキュメントへの*書き込み*コマンドをキューに登録していました。</span><span class="sxs-lookup"><span data-stu-id="a7282-339">In all the earlier functions in this series of tutorials, you queued commands to *write* to the Office document.</span></span> <span data-ttu-id="a7282-340">各関数は、キューに登録されたコマンドを実行対象のドキュメントに送信する `context.sync()` メソッドを呼び出すことで終了しています。</span><span class="sxs-lookup"><span data-stu-id="a7282-340">Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed.</span></span> <span data-ttu-id="a7282-341">ただし、最後の手順で追加したコードでは、`sheet.protection.protected` プロパティを呼び出しています。このことが、これまでに作成した関数とは大きく異なります。`sheet` オブジェクトは、この作業ウィンドウのスクリプトに存在する単なるプロキシ オブジェクトなので、</span><span class="sxs-lookup"><span data-stu-id="a7282-341">But the code you added in the last step calls the `sheet.protection.protected` property, and this is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script.</span></span> <span data-ttu-id="a7282-342">ドキュメントの実際の保護の状態を認識できません。そのため、その `protection.protected` プロパティでは実際の値が保持できません。</span><span class="sxs-lookup"><span data-stu-id="a7282-342">It doesn't know what the actual protection state of the document is, so its `protection.protected` property can't have a real value.</span></span> <span data-ttu-id="a7282-343">まず、ドキュメントから保護の状態をフェッチする必要があり、その状態を使用して `sheet.protection.protected` の値を設定します。</span><span class="sxs-lookup"><span data-stu-id="a7282-343">It is necessary to first fetch the protection status from the document and use it set the value of `sheet.protection.protected`.</span></span> <span data-ttu-id="a7282-344">そのようにした場合にのみ、例外がスローされることなく `sheet.protection.protected` を呼び出せるようになります。</span><span class="sxs-lookup"><span data-stu-id="a7282-344">Only then can `sheet.protection.protected` be called without causing an exception to be thrown.</span></span> <span data-ttu-id="a7282-345">このフェッチ処理には、3 つの手順があります。</span><span class="sxs-lookup"><span data-stu-id="a7282-345">This fetching process has three steps:</span></span>

   1. <span data-ttu-id="a7282-346">コードで読み取る必要があるプロパティをロードする (つまりフェッチする) コマンドをキューに登録します。</span><span class="sxs-lookup"><span data-stu-id="a7282-346">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>

   2. <span data-ttu-id="a7282-347">コンテキスト オブジェクトの `sync` メソッドを呼び出します。このメソッドは、キューに登録されたコマンドを実行対象のドキュメントに送信して、要求された情報を返します。</span><span class="sxs-lookup"><span data-stu-id="a7282-347">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>

   3. <span data-ttu-id="a7282-348">`sync` メソッドは非同期であるため、フェッチされたプロパティをコードで呼び出す前に、そのメソッドが完了していることを確認します。</span><span class="sxs-lookup"><span data-stu-id="a7282-348">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="a7282-349">こうした手順は、コードで Office ドキュメントから情報を*読み取る*必要がある場合には必ず完了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a7282-349">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="a7282-p144">`toggleProtection` 関数で、`TODO2` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-p144">In the `toggleProtection` function, replace `TODO2` with the following code. Note:</span></span>
   
   - <span data-ttu-id="a7282-352">すべての Excel オブジェクトに `load` メソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="a7282-352">Every Excel object has a `load` method.</span></span> <span data-ttu-id="a7282-353">読み取る必要のあるオブジェクトのプロパティは、コンマ区切りの名前の文字列としてパラメーターで指定します。</span><span class="sxs-lookup"><span data-stu-id="a7282-353">You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names.</span></span> <span data-ttu-id="a7282-354">この場合、読み取る必要のあるプロパティは、`protection` プロパティのサブプロパティです。</span><span class="sxs-lookup"><span data-stu-id="a7282-354">In this case, the property you need to read is a subproperty of the `protection` property.</span></span> <span data-ttu-id="a7282-355">サブプロパティはその他のコードの場合とほとんど同じ方法で参照しますが、"." 記号の代わりにスラッシュ ('/') 記号を使用する点が異なります。</span><span class="sxs-lookup"><span data-stu-id="a7282-355">You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.</span></span>

   - <span data-ttu-id="a7282-356">`sync` が完了してドキュメントからフェッチされた適切な値が `sheet.protection.protected` に割り当てられるまで、`sheet.protection.protected` を読み取る切り替えロジックが実行されないようにするために、そのロジックを `sync` が完了するまで実行されない `then` 関数に (この後の手順で) 移動します。</span><span class="sxs-lookup"><span data-stu-id="a7282-356">To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed.</span></span> 

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

2. <span data-ttu-id="a7282-357">分岐していない同一のコード パスに 2 つの `return` ステートメントを含めることはできないため、`Excel.run` の最後にある最終行の `return context.sync();` を削除します。</span><span class="sxs-lookup"><span data-stu-id="a7282-357">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`.</span></span> <span data-ttu-id="a7282-358">この後の手順で、新しい最終の `context.sync` を追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-358">You will add a new final `context.sync`, in a later step.</span></span>

3. <span data-ttu-id="a7282-359">`toggleProtection` 関数内の `if ... else` 構造を切り取って、`TODO3` の代わりに貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="a7282-359">Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.</span></span>

4. <span data-ttu-id="a7282-p147">`TODO4` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-p147">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="a7282-362">`sync` メソッドを `then` 関数に渡すことで、`sheet.protection.unprotect()` または `sheet.protection.protect()` のどちらかがキューに登録されるまで、そのメソッドが実行されないようにします。</span><span class="sxs-lookup"><span data-stu-id="a7282-362">Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.</span></span>

   - <span data-ttu-id="a7282-363">`then` メソッドは渡された関数を呼び出します。`sync` が 2 回呼び出されないように、`context.sync` の末尾の "()" は省略します。</span><span class="sxs-lookup"><span data-stu-id="a7282-363">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.</span></span>

    ```js
    .then(context.sync);
    ```

   <span data-ttu-id="a7282-364">作業が完了すると、関数の全体は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="a7282-364">When you are done, the entire function should look like the following:</span></span>

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

### <a name="configure-the-script-loading-html-file"></a><span data-ttu-id="a7282-365">スクリプト読み込み HTMl ファイルを構成する</span><span class="sxs-lookup"><span data-stu-id="a7282-365">Configure the script-loading HTML file</span></span>

<span data-ttu-id="a7282-366">/function-file/function-file.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-366">Open the /function-file/function-file.html file.</span></span> <span data-ttu-id="a7282-367">これは、ユーザーが **[Toggle Worksheet Protection]** ボタンをクリックしたときに呼び出される UI のない HTML ファイルです。</span><span class="sxs-lookup"><span data-stu-id="a7282-367">This is a UI-less HTML file that is called when the user presses the **Toggle Worksheet Protection** button.</span></span> <span data-ttu-id="a7282-368">ボタンがクリックされたときに実行する JavaScript メソッドを読み込むことを目的としています。</span><span class="sxs-lookup"><span data-stu-id="a7282-368">Its purpose is to load the JavaScript method that should run when the button is pushed.</span></span> <span data-ttu-id="a7282-369">このファイルには変更を加えません。</span><span class="sxs-lookup"><span data-stu-id="a7282-369">You are not going to change this file.</span></span> <span data-ttu-id="a7282-370">2 番目の `<script>` タグで functionfile.js が読み込まれる点に注目してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-370">Simply note that the second `<script>` tag loads the functionfile.js.</span></span>

   > [!NOTE]
   > <span data-ttu-id="a7282-371">function-file.html ファイルと、そのファイルが読み込む function-file.js ファイルは、アドインの作業ウィンドウとは完全に別の IE プロセスで実行されます。</span><span class="sxs-lookup"><span data-stu-id="a7282-371">The function-file.html file and the function-file.js file that it loads run in an entirely separate IE process from the add-in's task pane.</span></span> <span data-ttu-id="a7282-372">function-file.js が app.js ファイルと同じ bundle.js ファイルからトランスパイルされていた場合、アドインでは bundle.js の 2 つのコピーを読み込むことが必要になり、バンドル化の意味がなくなります。</span><span class="sxs-lookup"><span data-stu-id="a7282-372">If the function-file.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="a7282-373">さらに、function-file.js ファイルには IE で未サポートの JavaScript は含まれていません。</span><span class="sxs-lookup"><span data-stu-id="a7282-373">In addition, the function-file.js file does not contain any JavaScript that is unsupported by IE.</span></span> <span data-ttu-id="a7282-374">これら 2 つの理由から、このアドインでは function-file.js を一切トランスパイルしていません。</span><span class="sxs-lookup"><span data-stu-id="a7282-374">For these two reasons, this add-in does not transpile the function-file.js at all.</span></span> 

### <a name="test-the-add-in"></a><span data-ttu-id="a7282-375">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="a7282-375">Test the add-in</span></span>

1. <span data-ttu-id="a7282-376">Excel も含めて、すべての Office アプリケーションを閉じます。</span><span class="sxs-lookup"><span data-stu-id="a7282-376">Close all Office applications, including Excel.</span></span> 

2. <span data-ttu-id="a7282-377">キャッシュ フォルダーの内容を削除して、Office キャッシュを削除します。</span><span class="sxs-lookup"><span data-stu-id="a7282-377">Delete the Office cache by deleting the contents of the cache folder.</span></span> <span data-ttu-id="a7282-378">これは、ホストから古いバージョンのアドインを完全に削除するために必要です。</span><span class="sxs-lookup"><span data-stu-id="a7282-378">This is necessary to completely clear the old version of the add-in from the host.</span></span> 

    - <span data-ttu-id="a7282-379">Windows の場合: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`。</span><span class="sxs-lookup"><span data-stu-id="a7282-379">For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

    - <span data-ttu-id="a7282-380">Mac の場合: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`。</span><span class="sxs-lookup"><span data-stu-id="a7282-380">For Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

3. <span data-ttu-id="a7282-381">何らかの理由で、サーバーが稼働中でない場合は、Git Bash ウィンドウ、または Node.JS 対応のシステム プロンプトで、プロジェクトの **Start** フォルダーに移動して、`npm start` コマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="a7282-381">If for any reason, your server is not running, then in a Git Bash window, or Node.JS-enabled system prompt, navigate to the **Start** folder of the project and run the command `npm start`.</span></span> <span data-ttu-id="a7282-382">変更した JavaScript ファイルはビルド済みの bundle.js に含まれていないため、プロジェクトをリビルドする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="a7282-382">You do not need to rebuild the project because the only JavaScript file you changed is not part of the built bundle.js.</span></span>

4. <span data-ttu-id="a7282-383">新しいバージョンの変更済みマニフェスト ファイルを使用して、次のいずれかの方法でサイドローディング プロセスを繰り返します。</span><span class="sxs-lookup"><span data-stu-id="a7282-383">Using the new version of the changed manifest file, repeat the sideloading process by using one of the following methods.</span></span> <span data-ttu-id="a7282-384">*マニフェスト ファイルの以前のコピーを上書きする必要があります。*</span><span class="sxs-lookup"><span data-stu-id="a7282-384">*You should overwrite the previous copy of the manifest file.*</span></span>

    - <span data-ttu-id="a7282-385">Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="a7282-385">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>

    - <span data-ttu-id="a7282-386">Excel Online:[Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="a7282-386">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>

    - <span data-ttu-id="a7282-387">iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="a7282-387">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

5. <span data-ttu-id="a7282-388">Excel で任意のワークシートを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-388">Open any worksheet in Excel.</span></span>

6. <span data-ttu-id="a7282-p153">**[ホーム]** リボンで、**[ワークシート保護の切り替え]** を選択します。次のスクリーンショットに示すように、リボンのほとんどのコントロールは、無効化 (淡色表示) されます。</span><span class="sxs-lookup"><span data-stu-id="a7282-p153">On the **Home** ribbon, choose **Toggle Worksheet Protection**. Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in screenshot below.</span></span> 

7. <span data-ttu-id="a7282-391">セルの内容を変更する場合は、そのセルを選択します。</span><span class="sxs-lookup"><span data-stu-id="a7282-391">Choose a cell as you would if you wanted to change its content.</span></span> <span data-ttu-id="a7282-392">ワークシートが保護されているというエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="a7282-392">You get an error telling you that the worksheet is protected.</span></span>

8. <span data-ttu-id="a7282-393">もう一度 **[Toggle Worksheet Protection]** を選択すると、コントロールが再有効化され、再びセルの値を変更できるようになります。</span><span class="sxs-lookup"><span data-stu-id="a7282-393">Choose **Toggle Worksheet Protection** again, and the controls are reenabled, and you can change cell values again.</span></span>

    ![Excel チュートリアル - 保護がオンになっているリボン](../images/excel-tutorial-ribbon-with-protection-on.png)

## <a name="open-a-dialog"></a><span data-ttu-id="a7282-395">ダイアログを開く</span><span class="sxs-lookup"><span data-stu-id="a7282-395">Open a dialog box.</span></span>

<span data-ttu-id="a7282-396">このチュートリアルの最後の手順では、アドインでダイアログを開いて、ダイアログのプロセスから作業ウィンドウのプロセスにメッセージを渡して、ダイアログを閉じます。</span><span class="sxs-lookup"><span data-stu-id="a7282-396">In this final step of the tutorial, you'll open a dialog in your add-in, pass a message from the dialog process to the task pane process, and close the dialog.</span></span> <span data-ttu-id="a7282-397">Office アドインのダイアログは、*モードレス*です。ユーザーは、ホスト Office アプリケーション内のドキュメントと作業ウィンドウ内のホスト ページの両方の操作を続行できます。</span><span class="sxs-lookup"><span data-stu-id="a7282-397">Office Add-in dialogs are *nonmodal*: a user can continue to interact with both the document in the host Office application and with the host page in the task pane.</span></span>

### <a name="create-the-dialog-page"></a><span data-ttu-id="a7282-398">ダイアログ ページを作成する</span><span class="sxs-lookup"><span data-stu-id="a7282-398">Create the dialog page</span></span>

1. <span data-ttu-id="a7282-399">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-399">Open the project in your code editor.</span></span>

2. <span data-ttu-id="a7282-400">プロジェクトのルート (index.html がある場所) で、popup.html というファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="a7282-400">Create a file in the root of the project (where index.html is) called popup.html.</span></span>

3. <span data-ttu-id="a7282-p156">popup.html に、次のコードを追加します。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-p156">Add the following markup to popup.html. Note:</span></span>

   - <span data-ttu-id="a7282-403">このページには、ユーザーが自分の名前を入力する `<input>` と、その名前を作業ウィンドウ内のページ (入力した名前が表示されるページ) に送信するボタンが含まれています。</span><span class="sxs-lookup"><span data-stu-id="a7282-403">The page has a `<input>` where the user will enter their name and a button that will send the name to the page in the task pane where it will be displayed.</span></span>

   - <span data-ttu-id="a7282-404">このマークアップでは、popup.js というスクリプトを読み込みます。このスクリプトは、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="a7282-404">The markup loads a script called popup.js that you will create in a later step.</span></span>

   - <span data-ttu-id="a7282-405">また、popup.js で使用することになる Office.JS ライブラリと jQuery も読み込みます。</span><span class="sxs-lookup"><span data-stu-id="a7282-405">It also loads the Office.JS library and jQuery because they will be used in popup.js.</span></span>

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

4. <span data-ttu-id="a7282-406">プロジェクトのルートに popup.js というファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="a7282-406">Create a file in the root of the project called popup.js.</span></span>

5. <span data-ttu-id="a7282-p157">popup.js に、次のコードを追加します。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-p157">Add the following code to popup.js. Note:</span></span>

   - <span data-ttu-id="a7282-409">*Office.JS 内の API を呼び出すページは、どのページでも `Office.initialize` プロパティに関数を割り当てる必要があります。*</span><span class="sxs-lookup"><span data-stu-id="a7282-409">*Every page that calls APIs in the Office.JS library must assign a function to the `Office.initialize` property.*</span></span> <span data-ttu-id="a7282-410">初期化が不要な場合は、関数の本体を空にすることができますが、プロパティを未定義のままにすることや、Null または関数以外の値を割り当てることはできません。</span><span class="sxs-lookup"><span data-stu-id="a7282-410">If no initialization is needed, then the function can have an empty body, but the property must not be left undefined, assigned to null or to a non-function value.</span></span> <span data-ttu-id="a7282-411">たとえば、プロジェクト ルートにある app.js ファイルを確認してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-411">For an example, see the app.js file in the project root.</span></span> <span data-ttu-id="a7282-412">この割り当てを実施するコードは、Office.JS を呼び出す前に実行する必要があります。そのため、この例で示すように、割り当てはページによって読み込まれるスクリプト ファイル内に入れてあります。</span><span class="sxs-lookup"><span data-stu-id="a7282-412">The code that makes the assignment must run before any calls to Office.JS; hence the assignment is in a script file that is loaded by the page, as it is in this case.</span></span>
   
   - <span data-ttu-id="a7282-p159">jQuery の `ready` 関数は、`initialize` メソッド内から呼び出します。別の JavaScript ライブラリのコードの読み込み、初期化、またはブートストラップを `Office.initialize` 関数内に入れることは、ほとんどすべての場合に通用するルールです。</span><span class="sxs-lookup"><span data-stu-id="a7282-p159">The jQuery `ready` function is called inside the `initialize` method. It is an almost universal rule that the loading, initializing, or bootstrapping code of other JavaScript libraries should be inside the `Office.initialize` function.</span></span>

    ```js
    (function () {
    "use strict";

        Office.initialize = function() {
            $(document).ready(function () {  

                // TODO1: Assign handler to the OK button.

            });
        }

        // TODO2: Create the OK button handler

    }());
    ```

6. <span data-ttu-id="a7282-415">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a7282-415">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="a7282-416">`sendStringToParentPage` 関数は、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="a7282-416">You'll create the `sendStringToParentPage` function in the next step.</span></span>

    ```js
    $('#ok-button').click(sendStringToParentPage);
    ```

7. <span data-ttu-id="a7282-417">`TODO2` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a7282-417">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="a7282-418">`messageParent` メソッドは、パラメーターを親ページ (この例では、作業ウィンドウ内のページ) に渡します。</span><span class="sxs-lookup"><span data-stu-id="a7282-418">The `messageParent` method passes its parameter to the parent page, in this case, the page in the task pane.</span></span> <span data-ttu-id="a7282-419">パラメーターには、ブール値または文字列を使用できます (XML や JSON など、文字列としてシリアル化できるすべてのものが含まれます)。</span><span class="sxs-lookup"><span data-stu-id="a7282-419">The parameter can be a boolean or a string, which includes anything that can be serialized as a string, such as XML or JSON.</span></span>

    ```js
    function sendStringToParentPage() {
        var userName = $('#name-box').val();
        Office.context.ui.messageParent(userName);
    }
    ```

8. <span data-ttu-id="a7282-420">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="a7282-420">Save the file.</span></span>

   > [!NOTE]
   > <span data-ttu-id="a7282-421">popup.html ファイルと、そのファイルで読み込む popup.js ファイルは、アドインの作業ウィンドウとは完全に別な Internet Explorer プロセスで実行されます。</span><span class="sxs-lookup"><span data-stu-id="a7282-421">The popup.html file, and the popup.js file that it loads, run in an entirely separate Internet Explorer process from the add-in's task pane.</span></span> <span data-ttu-id="a7282-422">popup.js が app.js ファイルと同じ bundle.js ファイルからトランスパイルされていた場合、アドインでは bundle.js の 2 つのコピーを読み込むことが必要になり、バンドル化の意味がなくなります。</span><span class="sxs-lookup"><span data-stu-id="a7282-422">If the popup.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="a7282-423">さらに、popup.js ファイルには IE で未サポートの JavaScript は含まれていません。</span><span class="sxs-lookup"><span data-stu-id="a7282-423">In addition, the popup.js file does not contain any JavaScript that is unsupported by IE.</span></span> <span data-ttu-id="a7282-424">これら 2 つの理由から、このアドインでは popup.js を一切トランスパイルしていません。</span><span class="sxs-lookup"><span data-stu-id="a7282-424">For these two reasons, this add-in does not transpile the popup.js file at all.</span></span>

### <a name="open-the-dialog-from-the-task-pane"></a><span data-ttu-id="a7282-425">作業ウィンドウからダイアログを開く</span><span class="sxs-lookup"><span data-stu-id="a7282-425">Open the dialog from the task pane</span></span>

1. <span data-ttu-id="a7282-426">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-426">Open the file index.html.</span></span>

2. <span data-ttu-id="a7282-427">`freeze-header` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-427">Below the `div` that contains the `freeze-header` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="open-dialog">Open Dialog</button>
    </div>
    ```

3. <span data-ttu-id="a7282-428">このダイアログでは、ユーザーに名前の入力を求めて、ユーザーの名前を作業ウィンドウに渡します。</span><span class="sxs-lookup"><span data-stu-id="a7282-428">The dialog will prompt the user to enter a name and pass the user's name to the task pane.</span></span> <span data-ttu-id="a7282-429">作業ウィンドウでは、それがラベルに表示されます。</span><span class="sxs-lookup"><span data-stu-id="a7282-429">The task pane will display it in a label.</span></span> <span data-ttu-id="a7282-430">前の手順で追加した `div` のすぐ下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-430">Immediately below the `div` that you just added, add the following markup:</span></span>

    ```html
    <div class="padding">
        <label id="user-name"></label>
    </div>
    ```

4. <span data-ttu-id="a7282-431">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-431">Open the app.js file.</span></span>

5. <span data-ttu-id="a7282-432">`freeze-header` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-432">Below the line that assigns a click handler to the `freeze-header` button, add the following code.</span></span> <span data-ttu-id="a7282-433">`openDialog` メソッドは、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="a7282-433">You'll create the `openDialog` method in a later step.</span></span>

    ```js
    $('#open-dialog').click(openDialog);
    ```

6. <span data-ttu-id="a7282-p165">`freezeHeader` 関数の下に、次の宣言を追加します。この変数は、親ページの実行コンテキスト内のオブジェクトを保持するために使用され、ダイアログ ページの実行コンテキストへの仲介者として機能します。</span><span class="sxs-lookup"><span data-stu-id="a7282-p165">Below the `freezeHeader` function add the following declaration. This variable is used to hold an object in the parent page's execution context that acts as an intermediator to the dialog page's execution context.</span></span>

    ```js
    var dialog = null;
    ```

7. <span data-ttu-id="a7282-436">`dialog` の宣言の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-436">Below the declaration of `dialog`, add the following function.</span></span> <span data-ttu-id="a7282-437">このコードで注目する重要な点は、そこに `Excel.run` の呼び出しが存在*しない*ことです。</span><span class="sxs-lookup"><span data-stu-id="a7282-437">The important thing to notice about this code is what is *not* there: there is no call of `Excel.run`.</span></span> <span data-ttu-id="a7282-438">これは、ダイアログを開く API はすべての Office ホストで共有されるため、Excel 固有の API ではなく Office JavaScript 共通 API に含まれているからです。</span><span class="sxs-lookup"><span data-stu-id="a7282-438">This is because the API to open a dialog is shared among all Office hosts, so it is part of the Office JavaScript Common API, not the Excel-specific API.</span></span>

    ```js
    function openDialog() {
        // TODO1: Call the Office Common API that opens a dialog
    }
    ```

8. <span data-ttu-id="a7282-p167">`TODO1` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-p167">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="a7282-441">`displayDialogAsync` メソッドでは、画面の中央にダイアログを開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-441">The `displayDialogAsync` method opens a dialog in the center of the screen.</span></span>

   - <span data-ttu-id="a7282-442">最初のパラメーターは、開くページの URL です。</span><span class="sxs-lookup"><span data-stu-id="a7282-442">The first parameter is the URL of the page to open.</span></span>

   - <span data-ttu-id="a7282-p168">2 番目のパラメーターでオプションを渡します。`height` と `width` は、Office アプリケーションのウィンドウ サイズの比率です。</span><span class="sxs-lookup"><span data-stu-id="a7282-p168">The second parameter passes options. `height` and `width` are percentages of the size of the Office application's window.</span></span>

    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},

        // TODO2: Add callback parameter.
    );
    ```

### <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a><span data-ttu-id="a7282-445">ダイアログからのメッセージを処理してダイアログを閉じる</span><span class="sxs-lookup"><span data-stu-id="a7282-445">Process the message from the dialog and close the dialog</span></span>

1. <span data-ttu-id="a7282-p169">app.js ファイルでの作業を続けます。`TODO2` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-p169">Continue in the app.js file, and replace `TODO2` with the following code. Note:</span></span>

   - <span data-ttu-id="a7282-448">コールバックは、ダイアログが正常に開いた直後、ユーザーがダイアログで操作を行う前に実行されます。</span><span class="sxs-lookup"><span data-stu-id="a7282-448">The callback is executed immediately after the dialog successfully opens and before the user has taken any action in the dialog.</span></span>

   - <span data-ttu-id="a7282-449">`result.value` は、親ページとダイアログ ページの実行コンテキストの間で仲介者のように機能するオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="a7282-449">The `result.value` is the object that acts as a kind of middleman between the execution contexts of the parent and dialog pages.</span></span>

   - <span data-ttu-id="a7282-450">`processMessage` 関数は、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="a7282-450">The `processMessage` function will be created in a later step.</span></span> <span data-ttu-id="a7282-451">このハンドラーは、`messageParent` 関数の呼び出しによって、ダイアログから送信されるあらゆる値を処理します。</span><span class="sxs-lookup"><span data-stu-id="a7282-451">This handler will process any values that are sent from the dialog page with calls of the `messageParent` function.</span></span>

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. <span data-ttu-id="a7282-452">`openDialog` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="a7282-452">Below the `openDialog` function, add the following function.</span></span>

    ```js
    function processMessage(arg) {
        $('#user-name').text(arg.message);
        dialog.close();
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="a7282-453">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="a7282-453">Test the add-in</span></span>

1. <span data-ttu-id="a7282-454">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、**Ctrl + C** を 2 回入力して実行中の Web サーバーを停止します。</span><span class="sxs-lookup"><span data-stu-id="a7282-454">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="a7282-455">それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="a7282-455">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="a7282-456">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。</span><span class="sxs-lookup"><span data-stu-id="a7282-456">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="a7282-457">そのためには、ビルド コマンドの入力を求めるプロンプトが表示されるように、サーバー プロセスを強制終了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a7282-457">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="a7282-458">ビルド後に、サーバーを再起動します。</span><span class="sxs-lookup"><span data-stu-id="a7282-458">After the build, you restart the server.</span></span> <span data-ttu-id="a7282-459">次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="a7282-459">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="a7282-460">`npm run build` コマンドを実行して、ES6 ソース コードを Internet Explorer でサポートされている以前のバージョンの JavaScript にトランスパイルします (これは、Excel アドインを実行するために Excel の内部で使用されます)。</span><span class="sxs-lookup"><span data-stu-id="a7282-460">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="a7282-461">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="a7282-461">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="a7282-462">作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="a7282-462">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="a7282-463">作業ウィンドウで、**[Open Dialog]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="a7282-463">Choose the **Open Dialog** button in the task pane.</span></span>

6. <span data-ttu-id="a7282-464">ダイアログが開いたら、ドラッグしたりサイズ変更したりします。</span><span class="sxs-lookup"><span data-stu-id="a7282-464">While the dialog is open, drag it and resize it.</span></span> <span data-ttu-id="a7282-465">ワークシートを操作して、作業ウィンドウの他のボタンを押すことはできますが、同じ作業ウィンドウのページから 2 番目のダイアログを起動することはできないことに注意してください。</span><span class="sxs-lookup"><span data-stu-id="a7282-465">Note that you can interact with the worksheet and press other buttons on the task pane, but you cannot launch a second dialog from the same task pane page.</span></span>

7. <span data-ttu-id="a7282-466">ダイアログで、名前を入力して **[OK]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="a7282-466">In the dialog, enter a name and choose **OK**.</span></span> <span data-ttu-id="a7282-467">作業ウィンドウに名前が表示され、ダイアログが閉じられます。</span><span class="sxs-lookup"><span data-stu-id="a7282-467">The name appears on the task pane and the dialog closes.</span></span>

8. <span data-ttu-id="a7282-468">オプションとして、`processMessage` 関数内の行 `dialog.close();` をコメントにします。</span><span class="sxs-lookup"><span data-stu-id="a7282-468">Optionally, comment out the line `dialog.close();` in the `processMessage` function.</span></span> <span data-ttu-id="a7282-469">その後で、このセクションの手順を繰り返します。</span><span class="sxs-lookup"><span data-stu-id="a7282-469">Then repeat the steps of this section.</span></span> <span data-ttu-id="a7282-470">ダイアログを開いたまま名前を変更できます。</span><span class="sxs-lookup"><span data-stu-id="a7282-470">The dialog stays open and you can change the name.</span></span> <span data-ttu-id="a7282-471">右上の **[X]** ボタンをクリックすることで、手動で閉じることができます。</span><span class="sxs-lookup"><span data-stu-id="a7282-471">You can close it manually by pressing the **X** button in the upper right corner.</span></span>

    ![Excel チュートリアル - ダイアログ](../images/excel-tutorial-dialog-open.png)

## <a name="next-steps"></a><span data-ttu-id="a7282-473">次の手順</span><span class="sxs-lookup"><span data-stu-id="a7282-473">Next steps</span></span>

<span data-ttu-id="a7282-474">このチュートリアルでは、Excel ブック内のテーブル、グラフ、ワークシート、ダイアログの操作を行う、Excel 作業ウィンドウ アドインを作成しました。</span><span class="sxs-lookup"><span data-stu-id="a7282-474">In this tutorial, you've created an Excel task pane add-in that interacts with tables, charts, worksheets, and dialogs in an Excel workbook.</span></span> <span data-ttu-id="a7282-475">Excel アドインの構築に関する詳細については、次の記事にお進みください。</span><span class="sxs-lookup"><span data-stu-id="a7282-475">To learn more about developing Outlook add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="a7282-476">Excel アドインの概要</span><span class="sxs-lookup"><span data-stu-id="a7282-476">Excel add-ins overview</span></span>](../excel/excel-add-ins-overview.md)
