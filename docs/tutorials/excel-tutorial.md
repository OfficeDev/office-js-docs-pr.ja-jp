---
title: Excel アドインのチュートリアル
description: このチュートリアルでは、Excel アドインを構築します。このアドインでは、テーブルの作成、表示、フィルター処理、並べ替えを行うことができ、グラフの作成、テーブルのヘッダーの固定、ワークシートの保護も可能となります。また、ダイアログを開くこともできます。
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 3ad286e248b60afa16d4c18c9090e54e9c44cc39
ms.sourcegitcommit: f4fa1a0187466ea136009d1fe48ec67e4312c934
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/25/2020
ms.locfileid: "49408863"
---
# <a name="tutorial-create-an-excel-task-pane-add-in"></a><span data-ttu-id="2ee18-103">チュートリアル: Excel 作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="2ee18-103">Tutorial: Create an Excel task pane add-in</span></span>

<span data-ttu-id="2ee18-104">このチュートリアルでは、以下を実行する Excel 作業ウィンドウ アドインを作成します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-104">In this tutorial, you'll create an Excel task pane add-in that:</span></span>

> [!div class="checklist"]
>
> - <span data-ttu-id="2ee18-105">テーブルの作成</span><span class="sxs-lookup"><span data-stu-id="2ee18-105">Creates a table</span></span>
> - <span data-ttu-id="2ee18-106">テーブルのフィルター処理と並べ替え</span><span class="sxs-lookup"><span data-stu-id="2ee18-106">Filters and sorts a table</span></span>
> - <span data-ttu-id="2ee18-107">グラフの作成</span><span class="sxs-lookup"><span data-stu-id="2ee18-107">Creates a chart</span></span>
> - <span data-ttu-id="2ee18-108">テーブルのヘッダーの固定</span><span class="sxs-lookup"><span data-stu-id="2ee18-108">Freezes a table header</span></span>
> - <span data-ttu-id="2ee18-109">ワークシートの保護</span><span class="sxs-lookup"><span data-stu-id="2ee18-109">Protects a worksheet</span></span>
> - <span data-ttu-id="2ee18-110">ダイアログを開く</span><span class="sxs-lookup"><span data-stu-id="2ee18-110">Opens a dialog</span></span>

> [!TIP]
> <span data-ttu-id="2ee18-111">既に Yeoman ジェネレーターを使用した [[Excel タスク ウィンドウ アドインのビルド](../quickstarts/excel-quickstart-jquery.md)] の クイックスタートを完​​了しており、このチュートリアルの出発点としてそのプロジェクトを使用する場合は、[[テーブルの作成](#create-a-table)] セクションに直接移動します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-111">If you've already completed the [Build an Excel task pane add-in](../quickstarts/excel-quickstart-jquery.md) quick start using the Yeoman generator, and want to use that project as a starting point for this tutorial, go directly to the [Create a table](#create-a-table) section to start this tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="2ee18-112">前提条件</span><span class="sxs-lookup"><span data-stu-id="2ee18-112">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-your-add-in-project"></a><span data-ttu-id="2ee18-113">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="2ee18-113">Create your add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="2ee18-114">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="2ee18-114">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="2ee18-115">**Choose a script type: (スクリプトの種類を選択)** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="2ee18-115">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="2ee18-116">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="2ee18-116">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="2ee18-117">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)**</span><span class="sxs-lookup"><span data-stu-id="2ee18-117">**Which Office client application would you like to support?**</span></span> `Excel`

![Yeoman Office アドイン ジェネレーター コマンドライン インターフェイスのスクリーンショット](../images/yo-office-excel.png)

<span data-ttu-id="2ee18-119">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-119">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="create-a-table"></a><span data-ttu-id="2ee18-120">表を作成する</span><span class="sxs-lookup"><span data-stu-id="2ee18-120">Create a table</span></span>

<span data-ttu-id="2ee18-121">チュートリアルのこの手順では、プログラムによってアドインがユーザーの Excel の現在のバージョンをサポートしているかどうかをテストし、ワークシートにテーブルを追加して、そのテーブルのデータ設定と書式設定を実行します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-121">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Excel, add a table to a worksheet, populate the table with data, and format it.</span></span>

### <a name="code-the-add-in"></a><span data-ttu-id="2ee18-122">アドインのコードを作成する</span><span class="sxs-lookup"><span data-stu-id="2ee18-122">Code the add-in</span></span>

1. <span data-ttu-id="2ee18-123">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-123">Open the project in your code editor.</span></span>

2. <span data-ttu-id="2ee18-p101">ファイル **./src/taskpane/taskpane.html** を開きます。このファイルには、作業ウィンドウ用の HTML マークアップが含まれています。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p101">Open the file **./src/taskpane/taskpane.html**.  This file contains the HTML markup for the task pane.</span></span>

3. <span data-ttu-id="2ee18-126">`<main>` 要素を見つけて、開始 `<main>` タグの後、終了 `</main>` タグの前に表示されるすべての行を削除します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-126">Locate the `<main>` element and delete all lines that appear after the opening `<main>` tag and before the closing `</main>` tag.</span></span>

4. <span data-ttu-id="2ee18-127">開始 `<main>` タグの後に次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-127">Add the following markup immediately after the opening `<main>` tag:</span></span>

    ```html
    <button class="ms-Button" id="create-table">Create Table</button><br/><br/>
    ```

5. <span data-ttu-id="2ee18-p102">ファイル **./src/taskpane/taskpane.js** を開きます。このファイルには、作業ウィンドウと Office クライアント アプリケーションの間の相互作用を容易にする Office JavaScript API コードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p102">Open the file **./src/taskpane/taskpane.js**. This file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application.</span></span>

6. <span data-ttu-id="2ee18-130">次の操作を行って、[`run`] ボタンと `run()` 関数へのすべての参照を削除します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-130">Remove all references to the `run` button and the `run()` function by doing the following:</span></span>

    - <span data-ttu-id="2ee18-131">`document.getElementById("run").onclick = run;` 行を見つけて削除します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-131">Locate and delete the line `document.getElementById("run").onclick = run;`.</span></span>

    - <span data-ttu-id="2ee18-132">`run()` 関数全体を見つけて削除します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-132">Locate and delete the entire `run()` function.</span></span>

7. <span data-ttu-id="2ee18-p103">`Office.onReady` メソッドの呼び出しで、`if (info.host === Office.HostType.Excel) {` 行を見つけ、その行の直後に次のコードを追加します。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p103">Within the `Office.onReady` method call, locate the line `if (info.host === Office.HostType.Excel) {` and add the following code immediately after that line. Note:</span></span>

    - <span data-ttu-id="2ee18-p104">このコードの最初の部分では、ユーザーの Excel のバージョンが、このチュートリアルのシリーズで使用する API をすべて含んでいるバージョンの Excel.js をサポートしているかどうかを調べます。運用アドインでは、未サポートの API を呼び出す UI を非表示または無効化する条件ブロックの本体を使用してください。これにより、ユーザーは、そのユーザーの Excel のバージョンでサポートされているアドインの部分を使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p104">The first part of this code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use. In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs. This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.</span></span>

    - <span data-ttu-id="2ee18-138">このコードの 2 番目の部分では、[`create-table`] ボタンのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-138">The second part of this code adds an event handler for the `create-table` button.</span></span>

    ```js
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = createTable;
    ```

8. <span data-ttu-id="2ee18-p105">次の関数をファイルの最後に追加します。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p105">Add the following function to the end of the file. Note:</span></span>

    - <span data-ttu-id="2ee18-p106">Excel.js のビジネス ロジックが `Excel.run` に渡される関数に追加されます。このロジックは直ちには実行されません。代わりに、保留中のコマンドのキューに追加されます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p106">Your Excel.js business logic will be added to the function that is passed to `Excel.run`. This logic does not execute immediately. Instead, it is added to a queue of pending commands.</span></span>

    - <span data-ttu-id="2ee18-144">`context.sync` メソッドは、キューに登録されたすべてのコマンドを実行するために Excel に送信します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-144">The `context.sync` method sends all queued commands to Excel for execution.</span></span>

    - <span data-ttu-id="2ee18-p107">`Excel.run` の後に `catch` ブロックが続きます。これは、常に従う必要のあるベスト プラクティスです。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p107">The `Excel.run` is followed by a `catch` block. This is a best practice that you should always follow.</span></span> 

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

9. <span data-ttu-id="2ee18-p108">`createTable()` 関数内で、`TODO1` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p108">Within the `createTable()` function, replace `TODO1` with the following code. Note:</span></span>

    - <span data-ttu-id="2ee18-p109">コードは、ワークシートのテーブル コレクションの `add` メソッドを使用してテーブルを作成します。これは空の場合でも常に存在します。これは、Excel.js オブジェクトが作成される標準的な方法です。クラス コンストラクター API はありません。`new` 演算子を使用して Excel オブジェクトを作成することはできません。代わりに、親コレクション オブジェクトに追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p109">The code creates a table by using the `add` method of a worksheet's table collection, which always exists even if it is empty. This is the standard way that Excel.js objects are created. There are no class constructor APIs, and you never use a `new` operator to create an Excel object. Instead, you add to a parent collection object.</span></span>

    - <span data-ttu-id="2ee18-p110">`add` の最初のパラメーターは、テーブルの一番上の行のみの範囲で、テーブルが最終的に使用する範囲全体ではありません。これは、アドインがデータ行にデータを入力するときに、既存の行のセルに値を入力する代わりに、テーブルに新しい行を追加するためです。テーブルを作成するときに、テーブルに含まれる行の数がわからないことが多いため、これはより一般的なパターンです。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p110">The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use. This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows. This is a common pattern, because the number of rows a table will have is often unknown when the table is created.</span></span>

    - <span data-ttu-id="2ee18-156">テーブルの名前は、ワークシートだけでなくブック全体で一意にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="2ee18-156">Table names must be unique across the entire workbook, not just the worksheet.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

10. <span data-ttu-id="2ee18-p111">`createTable()` 関数内で、`TODO2` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p111">Within the `createTable()` function, replace `TODO2` with the following code. Note:</span></span>

    - <span data-ttu-id="2ee18-159">範囲に含まれるセルの値は、配列の配列で設定します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-159">The cell values of a range are set with an array of arrays.</span></span>

    - <span data-ttu-id="2ee18-p112">テーブル内に新しい行を作成するために、そのテーブルの行コレクションの `add` メソッドを呼び出します。2番目のパラメーターとして渡される親の配列に複数のセル値の配列を含めることにより、1 つの `add` 呼び出しに複数の行を追加できます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p112">New rows are created in a table by calling the `add` method of the table's row collection. You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.</span></span>

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

11. <span data-ttu-id="2ee18-p113">`createTable()` 関数内で、`TODO3` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p113">Within the `createTable()` function, replace `TODO3` with the following code. Note:</span></span>

    - <span data-ttu-id="2ee18-164">このコードでは、ゼロから始まるインデックスをテーブルの列コレクションの `getItemAt` メソッドに渡すことで、**Amount** 列への参照を取得します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-164">The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection.</span></span>

        > [!NOTE]
        > <span data-ttu-id="2ee18-165">Excel.js のコレクション オブジェクト (`TableCollection`、`WorksheetCollection`、`TableColumnCollection` など) には、`items` プロパティがあります。このプロパティは、子オブジェクト タイプ (`Table`、`Worksheet`、`TableColumn` など) の配列ですが、`*Collection` オブジェクト自体は配列ではありません。</span><span class="sxs-lookup"><span data-stu-id="2ee18-165">Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.</span></span>

    - <span data-ttu-id="2ee18-166">その次に、コードでは、**Amount** 列の範囲を小数点以下 2 桁までのユーロとして書式設定します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-166">The code then formats the range of the **Amount** column as Euros to the second decimal.</span></span>

    - <span data-ttu-id="2ee18-p114">最後に、列の幅と行の高さが、最も長い (または最も高い) データ項目の幅になるようにします。コードを書式設定するには `Range` オブジェクトを取得する必要があります。`TableColumn` と `TableRow` オブジェクトには、書式プロパティがありません。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p114">Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item. Notice that the code must get `Range` objects to format. `TableColumn` and `TableRow` objects do not have format properties.</span></span>

    ```js
    expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();
    ```

12. <span data-ttu-id="2ee18-170">プロジェクトに行ったすべての変更が保存されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-170">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="2ee18-171">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="2ee18-171">Test the add-in</span></span>

1. <span data-ttu-id="2ee18-172">以下の手順を実行し、ローカル Web サーバーを起動してアドインのサイドロードを行います。</span><span class="sxs-lookup"><span data-stu-id="2ee18-172">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="2ee18-p115">開発の最中でも、Office アドインは HTTP ではなく HTTPS を使用する必要があります。次のいずれかのコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p115">Office Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="2ee18-p116">Mac でアドインをテストする場合は、先に進む前にプロジェクトのルート ディレクトリで次のコマンドを実行します。このコマンドを実行すると、ローカル Web サーバーが起動します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p116">If you're testing your add-in on Mac, run the following command in the root directory of your project before proceeding. When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="2ee18-p117">Excel でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。ローカル Web サーバーが (まだ実行されていない場合) 起動し、アドインが読み込まれた Excel が開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p117">To test your add-in in Excel, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="2ee18-p118">Excel on the web でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p118">To test your add-in in Excel on the web, run the following command in the root directory of your project. When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="2ee18-181">アドインを使用するには、Excel on the web で新しいドキュメントを開き、「[Office on the web で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)」の手順に従ってアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="2ee18-181">To use your add-in, open a new document in Excel on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

2. <span data-ttu-id="2ee18-182">Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-182">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![[作業ウィンドウの表示] ボタンが強調表示されている Excel ホームメニューのスクリーンショット](../images/excel-quickstart-addin-3b.png)

3. <span data-ttu-id="2ee18-184">作業ウィンドウで、[**テーブルの作成**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-184">In the task pane, choose the **Create Table** button.</span></span>

    ![Excelのスクリーンショット。[テーブルの作成] ボタンが付いたアドイン作業ウィンドウと、日付、販売者、カテゴリ、および金額のデータが入力されたワーク シートのテーブルが表示されます](../images/excel-tutorial-create-table-2.png)

## <a name="filter-and-sort-a-table"></a><span data-ttu-id="2ee18-186">テーブルのフィルター処理と並べ替え</span><span class="sxs-lookup"><span data-stu-id="2ee18-186">Filter and sort a table</span></span>

<span data-ttu-id="2ee18-187">チュートリアルのこの手順では、以前に作成したテーブルをフィルター処理したり並べ替えたりします。</span><span class="sxs-lookup"><span data-stu-id="2ee18-187">In this step of the tutorial, you'll filter and sort the table that you created previously.</span></span>

### <a name="filter-the-table"></a><span data-ttu-id="2ee18-188">表のフィルター処理</span><span class="sxs-lookup"><span data-stu-id="2ee18-188">Filter the table</span></span>

1. <span data-ttu-id="2ee18-189">ファイル **./src/taskpane/taskpane.html** を開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-189">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="2ee18-190">`create-table` ボタンの `<button>` 要素を見つけ、その行の後に次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-190">Locate the `<button>` element for the `create-table` button, and add the following markup after that line:</span></span>

    ```html
    <button class="ms-Button" id="filter-table">Filter Table</button><br/><br/>
    ```

3. <span data-ttu-id="2ee18-191">ファイル **./src/taskpane/taskpane.js** を開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-191">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="2ee18-192">`Office.onReady` メソッドの呼び出し内で、クリック ハンドラーを `create-table` ボタンに割り当てる行を見つけ、その行の後に次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-192">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `create-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("filter-table").onclick = filterTable;
    ```

5. <span data-ttu-id="2ee18-193">次の関数をファイルの最後に追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-193">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="2ee18-p119">`filterTable()` 関数内で、`TODO1` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p119">Within the `filterTable()` function, replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="2ee18-p120">このコードでは最初に、`getItem` メソッドに列名を渡すことで、フィルター処理が必要な列への参照を取得します。`createTable` メソッドが行うように、列のインデックスを `getItemAt` メソッドに渡すわけではありません。ユーザーはテーブルの列を移動させることができるので、テーブルを作成した後、指定したインデックスにある列が変わってしまう可能性があります。そのため、列名を使用して列への参照を取得するほうが安全です。前のチュートリアルでは `getItemAt` を安全に使用しました。これは、テーブルを作成するのとまったく同じ方法で使用したため、ユーザーが列を移動する可能性がないためです。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p120">The code first gets a reference to the column that needs filtering by passing the column name to the `getItem` method, instead of passing its index to the `getItemAt` method as the `createTable` method does. Since users can move table columns, the column at a given index might change after the table is created. Hence, it is safer to use the column name to get a reference to the column. We used `getItemAt` safely in the preceding tutorial, because we used it in the very same method that creates the table, so there is no chance that a user has moved the column.</span></span>

   - <span data-ttu-id="2ee18-200">`applyValuesFilter` メソッドは、`Filter` オブジェクトのフィルター処理方法の 1 つです。</span><span class="sxs-lookup"><span data-stu-id="2ee18-200">The `applyValuesFilter` method is one of several filtering methods on the `Filter` object.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(['Education', 'Groceries']);
    ```

### <a name="sort-the-table"></a><span data-ttu-id="2ee18-201">表の並べ替え</span><span class="sxs-lookup"><span data-stu-id="2ee18-201">Sort the table</span></span>

1. <span data-ttu-id="2ee18-202">ファイル **./src/taskpane/taskpane.html** を開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-202">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="2ee18-203">`filter-table` ボタンの `<button>` 要素を見つけ、その行の後に次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-203">Locate the `<button>` element for the `filter-table` button, and add the following markup after that line:</span></span>

    ```html
    <button class="ms-Button" id="sort-table">Sort Table</button><br/><br/>
    ```

3. <span data-ttu-id="2ee18-204">ファイル **./src/taskpane/taskpane.js** を開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-204">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="2ee18-205">`Office.onReady` メソッドの呼び出し内で、クリック ハンドラーを `filter-table` ボタンに割り当てる行を見つけ、その行の後に次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-205">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `filter-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("sort-table").onclick = sortTable;
    ```

5. <span data-ttu-id="2ee18-206">次の関数をファイルの最後に追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-206">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="2ee18-p121">`sortTable()` 関数内で、`TODO1` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p121">Within the `sortTable()` function, replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="2ee18-209">アドインで並べ替えるのは Merchant 列のみであるため、このコードでは、1 つのメンバーだけを含む `SortField` オブジェクトの配列を作成します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-209">The code creates an array of `SortField` objects, which has just one member since the add-in only sorts on the Merchant column.</span></span>

   - <span data-ttu-id="2ee18-p122">`SortField` オブジェクトの `key` プロパティは、並べ替えに使用される対象列のゼロから始まるインデックスです。テーブルの行は、参照する列の値に基づいて並べ替えられます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p122">The `key` property of a `SortField` object is the zero-based index of the column used for sorting. The rows of the table are sorted based on the values in the referenced column.</span></span>

   - <span data-ttu-id="2ee18-p123">`Table` の `sort` メンバーは、`TableSort` オブジェクトであり、メソッドではありません。`SortField` は、`TableSort` オブジェクトの `apply` メソッドに渡されます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p123">The `sort` member of a `Table` is a `TableSort` object, not a method. The `SortField`s are passed to the `TableSort` object's `apply` method.</span></span>

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

7. <span data-ttu-id="2ee18-214">プロジェクトに行ったすべての変更が保存されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-214">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="2ee18-215">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="2ee18-215">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="2ee18-216">アドイン タスク ウィンドウが Excel でまだ開いていない場合は、[**ホーム**] タブに移動し、リボンの [**作業ウィンドウを表示**] ボタンを選択して開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-216">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="2ee18-217">このチュートリアルで以前に追加したテーブルが、開いているワークシートにない場合は、タスク ウィンドウの [**テーブルの作成**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-217">If the table you added previously in this tutorial is not present in the open worksheet, choose the **Create Table** button in the task pane.</span></span>

4. <span data-ttu-id="2ee18-218">[**テーブルのフィルター**] ボタンと [**テーブルの並べ替え**] ボタンを任意の順序で選択します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-218">Choose the **Filter Table** button and the **Sort Table** button, in either order.</span></span>

    ![Excel のスクリーンショット。アドインの作業ウィンドウに [フィルター テーブル] ボタンと [テーブルの並べ替え] ボタンが表示されています。](../images/excel-tutorial-filter-and-sort-table-2.png)

## <a name="create-a-chart"></a><span data-ttu-id="2ee18-220">グラフの作成</span><span class="sxs-lookup"><span data-stu-id="2ee18-220">Create a chart</span></span>

<span data-ttu-id="2ee18-221">チュートリアルのこの手順では、前の手順で作成したテーブルのデータを使用してグラフを作成して、そのグラフの書式を設定します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-221">In this step of the tutorial, you'll create a chart using data from the table that you created previously, and then format the chart.</span></span>

### <a name="chart-a-chart-using-table-data"></a><span data-ttu-id="2ee18-222">テーブルのデータを使用してグラフを作成する</span><span class="sxs-lookup"><span data-stu-id="2ee18-222">Chart a chart using table data</span></span>

1. <span data-ttu-id="2ee18-223">ファイル **./src/taskpane/taskpane.html** を開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-223">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="2ee18-224">`sort-table` ボタンの `<button>` 要素を見つけ、その行の後に次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-224">Locate the `<button>` element for the `sort-table` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="create-chart">Create Chart</button><br/><br/>
    ```

3. <span data-ttu-id="2ee18-225">ファイル **./src/taskpane/taskpane.js** を開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-225">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="2ee18-226">`Office.onReady` メソッドの呼び出し内で、クリック ハンドラーを `sort-table` ボタンに割り当てる行を見つけ、その行の後に次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-226">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `sort-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("create-chart").onclick = createChart;
    ```

5. <span data-ttu-id="2ee18-227">次の関数をファイルの最後に追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-227">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="2ee18-p124">`createChart()` 関数で、`TODO1` を次のコードに置き換えます。ヘッダー行を除外するために、このコードでは、`getRange` メソッドではなく `Table.getDataBodyRange` メソッドを使用してグラフを作成するデータの範囲を取得しています。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p124">Within the `createChart()` function, replace `TODO1` with the following code. Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var dataRange = expensesTable.getDataBodyRange();
    ```

7. <span data-ttu-id="2ee18-p125">`createChart()` 関数内で、`TODO2` を次のコードに置き換えます。次のパラメーターに注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p125">Within the `createChart()` function, replace `TODO2` with the following code. Note the following parameters:</span></span>

   - <span data-ttu-id="2ee18-p126">`add` への最初のパラメーターでは、グラフの種類を指定します。数十種類あります。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p126">The first parameter to the `add` method specifies the type of chart. There are several dozen types.</span></span>

   - <span data-ttu-id="2ee18-234">2 番目のパラメーターでは、グラフに含めるデータの範囲を指定します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-234">The second parameter specifies the range of data to include in the chart.</span></span>

   - <span data-ttu-id="2ee18-p127">3 番目のパラメーターでは、テーブルからの一連のデータ ポイントを行方向と列方向のどちらでグラフ化する必要があるかを決定します。オプション `auto` は、最適な方法を判断するように Excel に指示します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p127">The third parameter determines whether a series of data points from the table should be charted row-wise or column-wise. The option `auto` tells Excel to decide the best method.</span></span>

    ```js
    var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ```

8. <span data-ttu-id="2ee18-p128">`createChart()` 関数で、`TODO3` を次のコードに置き換えます。このコードのほとんどの部分は、わかりやすく説明不要なものです。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p128">Within the `createChart()` function, replace `TODO3` with the following code. Most of this code is self-explanatory. Note:</span></span>

   - <span data-ttu-id="2ee18-p129">`setPosition` メソッドへのパラメーターでは、グラフを挿入するワークシート領域の左上と右下のセルを指定します。Excel では、指定した空間でグラフを見栄えよくするために、線の太さなどの調整ができます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p129">The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart. Excel can adjust things like line width to make the chart look good in the space it has been given.</span></span>

   - <span data-ttu-id="2ee18-p130">"系列" とは、テーブルの 1 つの列にある一連のデータ ポイントのことです。このテーブルには文字列以外の列は 1 列しか含まれていないため、Excel は、グラフ化するデータ ポイントの列は、この列のみであると推測します。その他の列はグラフのラベルであると解釈されます。従って、グラフに含まれる系列は 1 つのみとなり、この系列のインデックスは 0 となります。を含みます。"&euro; での値" というラベルを付ける系列は、 この系列です。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p130">A "series" is a set of data points from a column of the table. Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart. It interprets the other columns as chart labels. So there will be just one series in the chart and it will have index 0. This is the one to label with "Value in &euro;".</span></span>

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in &euro;';
    ```

9. <span data-ttu-id="2ee18-247">プロジェクトに行ったすべての変更が保存されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-247">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="2ee18-248">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="2ee18-248">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="2ee18-249">アドイン タスク ウィンドウが Excel でまだ開いていない場合は、[**ホーム**] タブに移動し、リボンの [**作業ウィンドウを表示**] ボタンを選択して開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-249">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="2ee18-250">このチュートリアルで以前に追加したテーブルが、開いているワークシートにない場合は、タスク ウィンドウの [**テーブルの作成**] ボタンを選択します。次に、[**テーブルのフィルター処理**] ボタン、および [**テーブルの並べ替え**] ボタンのいずれかを選択します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-250">If the table you added previously in this tutorial is not present in the open worksheet, choose the **Create Table** button, and then the **Filter Table** button and the **Sort Table** button, in either order.</span></span>

4. <span data-ttu-id="2ee18-p131">**[グラフの作成]** ボタンを選択します。グラフが作成され、フィルター処理された行のデータのみが含まれます。一番下にあるデータポイントのラベルは、グラフの並べ替え順序になります。つまり、名前の逆アルファベット順での商社の名前です。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p131">Choose the **Create Chart** button. A chart is created and only the data from the rows that have been filtered are included. The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.</span></span>

    ![Excelのスクリーンショット。アドインの作業ウィンドウに [グラフの作成] ボタンが表示され、ワークシートに食料品と教育費のデータを表示するグラフが表示されます。](../images/excel-tutorial-create-chart-2.png)

## <a name="freeze-a-table-header"></a><span data-ttu-id="2ee18-255">テーブルのヘッダーの固定</span><span class="sxs-lookup"><span data-stu-id="2ee18-255">Freeze a table header</span></span>

<span data-ttu-id="2ee18-p132">ユーザーがスクロールして一部の行を見る必要があるほど長いテーブルがあると、見出し行がスクロールして見えなくなります。チュートリアルのこの手順では、ユーザーがワークシートを下にスクロールしても表示されるように、前に作成したテーブルの見出し行を固定します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p132">When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight. In this step of the tutorial, you'll freeze the header row of the table that you created previously, so that it remains visible even as the user scrolls down the worksheet.</span></span>

### <a name="freeze-the-tables-header-row"></a><span data-ttu-id="2ee18-258">表のヘッダー行を固定する</span><span class="sxs-lookup"><span data-stu-id="2ee18-258">Freeze the table's header row</span></span>

1. <span data-ttu-id="2ee18-259">ファイル **./src/taskpane/taskpane.html** を開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-259">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="2ee18-260">`create-chart` ボタンの `<button>` 要素を見つけ、その行の後に次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-260">Locate the `<button>` element for the `create-chart` button, and add the following markup after that line:</span></span>

    ```html
    <button class="ms-Button" id="freeze-header">Freeze Header</button><br/><br/>
    ```

3. <span data-ttu-id="2ee18-261">ファイル **./src/taskpane/taskpane.js** を開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-261">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="2ee18-262">`Office.onReady` メソッドの呼び出し内で、クリック ハンドラーを `create-chart` ボタンに割り当てる行を見つけ、その行の後に次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-262">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `create-chart` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("freeze-header").onclick = freezeHeader;
    ```

5. <span data-ttu-id="2ee18-263">次の関数をファイルの最後に追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-263">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="2ee18-p133">`freezeHeader()` 関数内で、`TODO1` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p133">Within the `freezeHeader()` function, replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="2ee18-266">`Worksheet.freezePanes` コレクションは、ワークシートのスクロール操作時に、ワークシート上でピン留めつまり固定される一式のウィンドウのことです。</span><span class="sxs-lookup"><span data-stu-id="2ee18-266">The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.</span></span>

   - <span data-ttu-id="2ee18-p134">`freezeRows` メソッドでは、上から数えた行数を、ピン留めする位置のパラメーターとして使用します。`1` を渡して最初の行を適所にピン留めします。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p134">The `freezeRows` method takes as a parameter the number of rows, from the top, that are to be pinned in place. We pass `1` to pin the first row in place.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

7. <span data-ttu-id="2ee18-269">プロジェクトに行ったすべての変更が保存されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-269">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="2ee18-270">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="2ee18-270">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="2ee18-271">アドイン タスク ウィンドウが Excel でまだ開いていない場合は、[**ホーム**] タブに移動し、リボンの [**作業ウィンドウを表示**] ボタンを選択して開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-271">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="2ee18-272">このチュートリアルに以前追加したテーブルがワークシートに存在する場合は、それを削除します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-272">If the table you added previously in this tutorial is present in the worksheet, delete it.</span></span>

4. <span data-ttu-id="2ee18-273">作業ウィンドウで、[**テーブルの作成**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-273">In the task pane, choose the **Create Table** button.</span></span>

5. <span data-ttu-id="2ee18-274">作業ウィンドウで、[**ヘッダーを固定**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-274">In the task pane, choose the **Freeze Header** button.</span></span>

6. <span data-ttu-id="2ee18-275">ヘッダー以降の行が画面の外に出て見えなくなるまでワークシートを十分下にスクロールしても、表のヘッダーが最上部に表示されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-275">Scroll down the worksheet far enough to see that the table header remains visible at the top even when the higher rows scroll out of sight.</span></span>

    ![固定テーブル ヘッダーがある Excel ワーク シートを表示するスクリーンショット](../images/excel-tutorial-freeze-header-2.png)

## <a name="protect-a-worksheet"></a><span data-ttu-id="2ee18-277">ワークシートの保護</span><span class="sxs-lookup"><span data-stu-id="2ee18-277">Protect a worksheet</span></span>

<span data-ttu-id="2ee18-278">チュートリアルのこの手順では、ワークシートの保護のオンとオフを切り替えるボタンをリボンに追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-278">In this step of the tutorial, you'll add a button to the ribbon that toggles worksheet protection on and off.</span></span>

### <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a><span data-ttu-id="2ee18-279">2 つ目のリボン ボタンを追加するようにマニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="2ee18-279">Configure the manifest to add a second ribbon button</span></span>

1. <span data-ttu-id="2ee18-280">マニフェスト ファイル **./manifest.xml** を開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-280">Open the manifest file **./manifest.xml**.</span></span>

2. <span data-ttu-id="2ee18-p135">`<Control>` 要素を見つけます。この要素は、アドインの起動に使用する **[ホーム]** リボンの **[作業ウィンドウを表示]** ボタンを定義します。**[ホーム]** リボンの同じグループに 2 番目のボタンを追加します。終了 `</Control>` タグと終了 `</Group>` タグの間に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p135">Locate the `<Control>` element. This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in. We're going to add a second button to the same group on the **Home** ribbon. In between the closing `</Control>` tag and the closing `</Group>` tag, add the following markup.</span></span>

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

3. <span data-ttu-id="2ee18-p136">マニフェスト ファイルに追加した XML 内で `TODO1` を文字列に置き換えて、このマニフェスト ファイル内で一意の ID をボタンに割り当てます。このボタンでは、ワークシートの保護のオン/オフを切り替えるので、「ToggleProtection」を使用することにします。完了すると、`Control` 要素の開始タグは次のようになります。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p136">Within the XML you just added to the manifest file, replace `TODO1` with a string that gives the button an ID that is unique within this manifest file. Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection". When you are done, the opening tag for the `Control` element should look like this:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. <span data-ttu-id="2ee18-p137">次の 3 つの `TODO` は、リソース ID または `resid` を設定します。リソースは文字列です。これら 3 つの文字列は、この後の手順で作成します。ここでは、そのリソースに ID を割り当てる必要があります。ボタンのラベルは「保護を切り換える」と表示されますが、この文字列の *ID* は「ProtectionButtonLabel」である必要があるため、`Label` 要素は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p137">The next three `TODO`s set resource IDs, or `resid`s. A resource is a string, and you'll create these three strings in a later step. For now, you need to give IDs to the resources. The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the `Label` element should look like this:</span></span>

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. <span data-ttu-id="2ee18-p138">`SuperTip` 要素では、このボタンのツール ヒントを定義します。ツール ヒントのタイトルはボタンのラベルと同じにする必要があるため、リソース ID にはまったく同じ「ProtectionButtonLabel」を使用することにします。ツール ヒントの説明は、「クリックして、ワークシートの保護をオンまたはオフにします」にする予定です。ただし、`resid` は「ProtectionButtonToolTip」である必要があります。したがって、完了すると、`SuperTip` 要素は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p138">The `SuperTip` element defines the tool tip for the button. The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel". The tool tip description will be "Click to turn protection of the worksheet on and off". But the `resid` should be "ProtectionButtonToolTip". So, when you are done, the `SuperTip` element should look like this:</span></span>

    ```xml
    <Supertip>
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE]
   > <span data-ttu-id="2ee18-p139">運用アドインでは、異なる 2 つのボタンに同じアイコンを使用することは避けたいところですが、このチュートリアルでは説明を簡単にするために同じアイコンを使用します。そのため、この新しい `Control` の `Icon` マークアップは、単に既存の `Control` から `Icon` 要素をコピーします。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p139">In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that. So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`.</span></span>

6. <span data-ttu-id="2ee18-p140">元の `Control` 要素内の `Action` 要素の種類は `ShowTaskpane` に設定されていますが、新しいボタンでは作業ウィンドウが開きません。後の手順で作成するカスタム関数を実行します。したがって、`TODO5` を `ExecuteFunction` に置き換えます。これは、カスタム関数をトリガーするボタンの操作の種類です。`Action` 要素の開始タグは次のようになります。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p140">The `Action` element inside the original `Control` element has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step. So, replace `TODO5` with `ExecuteFunction`, which is the action type for buttons that trigger custom functions. The opening tag for the `Action` element should look like this:</span></span>

    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. <span data-ttu-id="2ee18-p141">元の `Action` 要素には、作業ウィンドウの ID と、作業ウィンドウで開くページの URL を指定する子要素があります。ただし、`ExecuteFunction` の種類の `Action` の要素には、そのコントロールが実行している関数に名前を付けた単一の子要素があります。この関数は、後の手順で作成し、`toggleProtection` と呼ばれます。`TODO6` は次のマークアップに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p141">The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane. But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes. You'll create that function in a later step, and it will be called `toggleProtection`. So, replace `TODO6` with the following markup:</span></span>

    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    <span data-ttu-id="2ee18-306">`Control` マークアップの全体は、次のようになりました。</span><span class="sxs-lookup"><span data-stu-id="2ee18-306">The entire `Control` markup should now look like the following:</span></span>

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

8. <span data-ttu-id="2ee18-307">マニフェストの `Resources` セクションまで下にスクロールします。</span><span class="sxs-lookup"><span data-stu-id="2ee18-307">Scroll down to the `Resources` section of the manifest.</span></span>

9. <span data-ttu-id="2ee18-308">`bt:ShortStrings` 要素の子として、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-308">Add the following markup as a child of the `bt:ShortStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. <span data-ttu-id="2ee18-309">`bt:LongStrings` 要素の子として、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-309">Add the following markup as a child of the `bt:LongStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. <span data-ttu-id="2ee18-310">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-310">Save the file.</span></span>

### <a name="create-the-function-that-protects-the-sheet"></a><span data-ttu-id="2ee18-311">シートを保護する関数を作成する</span><span class="sxs-lookup"><span data-stu-id="2ee18-311">Create the function that protects the sheet</span></span>

1. <span data-ttu-id="2ee18-312">**.\commands\commands.js** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-312">Open the file **.\commands\commands.js**.</span></span>

2. <span data-ttu-id="2ee18-p142">`action` 関数の直後に次の関数を追加します。関数に `args` パラメーターを指定し、関数の最後の行で `args.completed` を呼び出すことに注意してください。これは、種類 **ExecuteFunction** のすべてのアドイン コマンドの要件です。これにより、関数が終了したことと、UI が再度応答可能になることを Office クライアント アプリケーションに通知します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p142">Add the following function immediately after the `action` function. Note that we specify an `args` parameter to the function and the very last line of the function calls `args.completed`. This is a requirement for all add-in commands of type **ExecuteFunction**. It signals the Office client application that the function has finished and the UI can become responsive again.</span></span>

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

3. <span data-ttu-id="2ee18-317">次の行を、ファイルの最後に追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-317">Add the following line to the end of the file:</span></span>

    ```js
    g.toggleProtection = toggleProtection;
    ```

4. <span data-ttu-id="2ee18-p143">`toggleProtection` 関数で、`TODO1` を次のコードに置き換えます。このコードでは、標準の切り替えパターンで、ワークシート オブジェクトの protection プロパティを使用します。`TODO2` については次のセクションで説明します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p143">Within the `toggleProtection` function, replace `TODO1` with the following code. This code uses the worksheet object's protection property in a standard toggle pattern. The `TODO2` will be explained in the next section.</span></span>

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

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="2ee18-321">ドキュメントのプロパティを作業ウィンドウのスクリプト オブジェクトに取得するコードを追加する</span><span class="sxs-lookup"><span data-stu-id="2ee18-321">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="2ee18-p144">これまでこのチュートリアルで作成した各関数では、Office ドキュメントに *書き込み* するコマンドをキューに登録します。各関数は `context.sync()` メソッドの呼び出しで終了しました。このメソッドは、キューに登録されたコマンドを実行するドキュメントに送信します。ただし、最後の手順で追加したコードは `sheet.protection.protected property` を呼び出します。`sheet` オブジェクトは、作業ウィンドウのスクリプトに存在するプロキシ オブジェクトにすぎないため、これは以前に書き込んだ関数と大きく違います。プロキシ オブジェクトはドキュメントの実際の保護状態を認識していないため、その `protection.protected` プロパティに実際の値を設定することはできません。例外エラーを回避するには、最初にドキュメントから保護状態を取得し、それを使用して `sheet.protection.protected` の値を設定する必要があります。この取得プロセスには、次の 3 つの手順があります。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p144">In each function that you've created in this tutorial until now, you queued commands to *write* to the Office document. Each function ended with a call to the `context.sync()` method, which sends the queued commands to the document to be executed. However, the code you added in the last step calls the `sheet.protection.protected property`. This is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script. The proxy object doesn't know the actual protection state of the document, so its `protection.protected` property can't have a real value. To avoid an exception error, you must first fetch the protection status from the document and use it set the value of `sheet.protection.protected`. This fetching process has three steps:</span></span>

   1. <span data-ttu-id="2ee18-329">コードで読み取る必要があるプロパティを読み込む (つまり取得する) コマンドをキューに登録します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-329">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>

   2. <span data-ttu-id="2ee18-330">コンテキスト オブジェクトの `sync` メソッドを呼び出します。このメソッドは、キューに登録されたコマンドを実行対象のドキュメントに送信して、要求された情報を返します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-330">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>

   3. <span data-ttu-id="2ee18-331">`sync` メソッドは非同期であるため、フェッチされたプロパティをコードで呼び出す前に、そのメソッドが完了していることを確認します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-331">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="2ee18-332">こうした手順は、コードで Office ドキュメントから情報を *読み取る* 必要がある場合には必ず完了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2ee18-332">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="2ee18-p145">`toggleProtection` 関数内で、`TODO2` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p145">Within the `toggleProtection` function, replace `TODO2` with the following code. Note:</span></span>

   - <span data-ttu-id="2ee18-p146">すべての Excel オブジェクトに `load` メソッドがあります。パラメーターで読み取るオブジェクトのプロパティをコンマ区切り名前の文字列として指定します。この場合、読み取る必要のあるプロパティは、`protection` プロパティのサブプロパティです。サブプロパティは、コード内の他の場所とほぼ同じ方法で参照します。ただし、「.」文字の代わりにスラッシュ ('/') を使用します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p146">Every Excel object has a `load` method. You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names. In this case, the property you need to read is a subproperty of the `protection` property. You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.</span></span>

   - <span data-ttu-id="2ee18-339">`sync` が完了してドキュメントからフェッチされた適切な値が `sheet.protection.protected` に割り当てられるまで、`sheet.protection.protected` を読み取る切り替えロジックが実行されないようにするために、そのロジックを `sync` が完了するまで実行されない `then` 関数に (この後の手順で) 移動します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-339">To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed.</span></span>

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

2. <span data-ttu-id="2ee18-p147">分岐していない同一のコード パスに 2 つの `return` ステートメントを含めることはできないため、`Excel.run` の最後にある最終行の `return context.sync();` を削除します。新しい最後の `context.sync` は、このチュートリアルの後の方で追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p147">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`. You will add a new final `context.sync`, in a later step.</span></span>

3. <span data-ttu-id="2ee18-342">`toggleProtection` 関数内の `if ... else` 構造を切り取って、`TODO3` の代わりに貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-342">Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.</span></span>

4. <span data-ttu-id="2ee18-p148">`TODO4` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p148">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="2ee18-345">`sync` メソッドを `then` 関数に渡すことで、`sheet.protection.unprotect()` または `sheet.protection.protect()` のどちらかがキューに登録されるまで、そのメソッドが実行されないようにします。</span><span class="sxs-lookup"><span data-stu-id="2ee18-345">Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.</span></span>

   - <span data-ttu-id="2ee18-346">`then` メソッドは渡された関数を呼び出します。`sync` が 2 回呼び出されないように、`context.sync` の末尾の "()" は省略します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-346">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.</span></span>

    ```js
    .then(context.sync);
    ```

   <span data-ttu-id="2ee18-347">作業が完了すると、関数の全体は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="2ee18-347">When you are done, the entire function should look like the following:</span></span>

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

5. <span data-ttu-id="2ee18-348">プロジェクトに行ったすべての変更が保存されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-348">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="2ee18-349">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="2ee18-349">Test the add-in</span></span>

1. <span data-ttu-id="2ee18-350">Excel も含めて、すべての Office アプリケーションを閉じます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-350">Close all Office applications, including Excel.</span></span>

2. <span data-ttu-id="2ee18-p149">キャッシュ フォルダーの内容 (すべてのファイルとサブフォルダー) を削除して、Office キャッシュを削除します。これは、クライアント アプリケーションから以前のバージョンのアドインを完全に削除するために必要です。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p149">Delete the Office cache by deleting the contents (all the files and subfolders) of the cache folder. This is necessary to completely clear the old version of the add-in from the client application.</span></span>

    - <span data-ttu-id="2ee18-353">Windows の場合: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`。</span><span class="sxs-lookup"><span data-stu-id="2ee18-353">For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

    - <span data-ttu-id="2ee18-354">Mac の場合: `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`。</span><span class="sxs-lookup"><span data-stu-id="2ee18-354">For Mac: `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

      > [!NOTE]
      > <span data-ttu-id="2ee18-355">そのフォルダーが存在しない場合は、次のフォルダーを確認し、見つかった場合はフォルダーの内容を削除します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-355">If that folder doesn't exist, check for the following folders and if found, delete the contents of the folder:</span></span>
      >  - <span data-ttu-id="2ee18-356">`{host}` が Office アプリケーション (例: `Excel`) である `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/`</span><span class="sxs-lookup"><span data-stu-id="2ee18-356">`~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/` where `{host}` is the Office application (e.g., `Excel`)</span></span>
      >  - <span data-ttu-id="2ee18-357">`{host}` が Office アプリケーション (例: `Excel`) である `~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`</span><span class="sxs-lookup"><span data-stu-id="2ee18-357">`~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/` where `{host}` is the Office application (e.g., `Excel`)</span></span>
      >  - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
      >  - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2/`

3. <span data-ttu-id="2ee18-358">ローカル Web サーバーが既に実行中の場合は、ノード コマンド ウィンドウを閉じて終了します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-358">If the local web server is already running, stop it by closing the node command window.</span></span>

4. <span data-ttu-id="2ee18-p150">マニフェスト ファイルが更新されているため、更新されたマニフェスト ファイルを使用してアドインを再度サイドロードする必要があります。ローカル Web サーバーを起動し、アドインのサイドロードを行います。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p150">Because your manifest file has been updated, you must sideload your add-in again, using the updated manifest file. Start the local web server and sideload your add-in:</span></span>

    - <span data-ttu-id="2ee18-p151">Excel でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。ローカル Web サーバーが (まだ実行されていない場合) 起動し、アドインが読み込まれた Excel が開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p151">To test your add-in in Excel, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="2ee18-p152">Excel on the web でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p152">To test your add-in in Excel on the web, run the following command in the root directory of your project. When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="2ee18-365">アドインを使用するには、Excel on the web で新しいドキュメントを開き、「[Office on the web で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)」の手順に従ってアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="2ee18-365">To use your add-in, open a new document in Excel on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

5. <span data-ttu-id="2ee18-p153">Excel の [**ホーム**] タブで、[**ワークシート保護を切り換える**] ボタンを選択します。次のスクリーンショットに示すように、リボンのほとんどのコントロールは無効化 (淡色表示) されることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p153">On the **Home** tab in Excel, choose the **Toggle Worksheet Protection** button. Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in the following screenshot.</span></span>

    ![[ワークシート保護の切り替え] ボタンが強調表示され、有効になっている Excel リボンのスクリーンショットです。他のほとんどのボタンは灰色表示され、無効になります。](../images/excel-tutorial-ribbon-with-protection-on-2.png)

6. <span data-ttu-id="2ee18-p155">セルの内容を変更する場合は、そのセルを選択します。Excel にワークシートが保護されていることを示すエラー メッセージが表示されます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p155">Choose a cell as you would if you wanted to change its content. Excel displays an error message indicating that the worksheet is protected.</span></span>

7. <span data-ttu-id="2ee18-372">もう一度 [**ワークシート保護を切り換える**] ボタンを選択すると、コントロールが再有効化され、再びセルの値を変更できるようになります。</span><span class="sxs-lookup"><span data-stu-id="2ee18-372">Choose the **Toggle Worksheet Protection** button again, and the controls are reenabled, and you can change cell values again.</span></span>

## <a name="open-a-dialog"></a><span data-ttu-id="2ee18-373">ダイアログを開く</span><span class="sxs-lookup"><span data-stu-id="2ee18-373">Open a dialog</span></span>

<span data-ttu-id="2ee18-p156">チュートリアルの最後の手順では、アドインのダイアログを開いて、ダイアログ プロセスのメッセージを作業ウィンドウのプロセスに渡し、ダイアログボックスを閉じます。Office アドインのダイアログには *非モーダル* があります。ユーザーは、Office アプリケーションでも、作業ウィンドウのホスト ページでも、ドキュメントの操作を続行できます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p156">In this final step of the tutorial, you'll open a dialog in your add-in, pass a message from the dialog process to the task pane process, and close the dialog. Office Add-in dialogs are *nonmodal*: a user can continue to interact with both the document in the Office application and with the host page in the task pane.</span></span>

### <a name="create-the-dialog-page"></a><span data-ttu-id="2ee18-376">ダイアログ ページを作成する</span><span class="sxs-lookup"><span data-stu-id="2ee18-376">Create the dialog page</span></span>

1. <span data-ttu-id="2ee18-377">プロジェクトのルートにある **./src** フォルダーで、**ダイアログ** という名前の新しいフォルダーを作成します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-377">In the **./src** folder that's located at the root of the project, create a new folder named **dialogs**.</span></span>

2. <span data-ttu-id="2ee18-378">**./src/dialogs** フォルダー に **popup.html** という名前の新しいファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-378">In the **./src/dialogs** folder, create new file named **popup.html**.</span></span>

3. <span data-ttu-id="2ee18-p157">**popup.html** に、次のマークアップを追加します。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p157">Add the following markup to **popup.html**. Note:</span></span>

   - <span data-ttu-id="2ee18-381">このページには、ユーザーが自分の名前を入力する `<input>` フィールドと、その名前が表示される作業ウィンドウ内のページに送信するボタンが含まれています。</span><span class="sxs-lookup"><span data-stu-id="2ee18-381">The page has an `<input>` field where the user will enter their name, and a button that will send this name to the task pane where it will display.</span></span>

   - <span data-ttu-id="2ee18-382">このマークアップでは、**popup.js** という名前のスクリプトを読み込みます。このスクリプトは、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-382">The markup loads a script named **popup.js** that you will create in a later step.</span></span>

   - <span data-ttu-id="2ee18-383">また、**popup.js** で使用することになる Office.JS ライブラリも読み込みます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-383">It also loads the Office.js library because it will be used in **popup.js**.</span></span>

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

4. <span data-ttu-id="2ee18-384">**./src/dialogs** フォルダーで、**popup.js** という名前の新しいファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-384">In the **./src/dialogs** folder, create new file named **popup.js**.</span></span>

5. <span data-ttu-id="2ee18-p158">次のコードを **popup.js** に追加します。このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p158">Add the following code to **popup.js**. Note the following about this code:</span></span>

   - <span data-ttu-id="2ee18-p159">*Office.js ライブラリ内の API を呼び出すすべてのページでは、まずライブラリが完全に初期化されていることを確認する必要があります。* そのための最良の方法は、`Office.onReady()` メソッドを呼び出すことです。アドインに独自の初期化作業がある場合、コードは `Office.onReady()` の呼び出しにチェーンされた `then()` メソッドに配置する必要があります。`Office.onReady()` の呼び出しは、Office.js を呼び出す前に実行する必要があります。したがって、この場合のように、割り当てはページによって読み込まれるスクリプト ファイルにあります。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p159">*Every page that calls APIs in the Office.js library must first ensure that the library is fully initialized.* The best way to do that is to call the `Office.onReady()` method. If your add-in has its own initialization tasks, the code should go in a `then()` method that is chained to the call of `Office.onReady()`. The call of `Office.onReady()` must run before any calls to Office.js; hence the assignment is in a script file that is loaded by the page, as it is in this case.</span></span>

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

6. <span data-ttu-id="2ee18-p160">`TODO1` は次のコードで置き換えます。`sendStringToParentPage` 関数は、次の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p160">Replace `TODO1` with the following code. You'll create the `sendStringToParentPage` function in the next step.</span></span>

    ```js
    document.getElementById("ok-button").onclick = sendStringToParentPage;
    ```

7. <span data-ttu-id="2ee18-p161">`TODO2` は次のコードで置き換えます。`messageParent` メソッドは、パラメーターを親ページ (この場合は作業ウィンドウのページ) に渡します。パラメーターには、ブール値または文字列を指定できます。これには、XML や JSON など、文字列としてシリアル化できるものが含まれます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p161">Replace `TODO2` with the following code. The `messageParent` method passes its parameter to the parent page, in this case, the page in the task pane. The parameter can be a boolean or a string, which includes anything that can be serialized as a string, such as XML or JSON.</span></span>

    ```js
    function sendStringToParentPage() {
        var userName = document.getElementById("name-box").value;
        Office.context.ui.messageParent(userName);
    }
    ```

> [!NOTE]
> <span data-ttu-id="2ee18-p162">**popup.html** ファイルと、そのファイルで読み込む **popup.js** ファイルは、アドインの作業ウィンドウとは完全に別な Microsoft Edge または Internet Explorer 11 プロセスで実行されます。**popup.js** が **app.js** ファイルと同じ **bundle.js** ファイルにトランスパイルされた場合、アドインは 2 つの **bundle.js** ファイルのコピーを読み込む必要があります。これはバンドルの目的に違反します。したがって、このアドインは **popup.js** ファイルをまったくトランスパイルしません。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p162">The **popup.html** file, and the **popup.js** file that it loads, run in an entirely separate Microsoft Edge or Internet Explorer 11 process from the add-in's task pane. If **popup.js** was transpiled into the same **bundle.js** file as the **app.js** file, then the add-in would have to load two copies of the **bundle.js** file, which defeats the purpose of bundling. Therefore, this add-in does not transpile the **popup.js** file at all.</span></span>

### <a name="update-webpack-config-settings"></a><span data-ttu-id="2ee18-399">Webpack の構成設定を更新する</span><span class="sxs-lookup"><span data-stu-id="2ee18-399">Update webpack config settings</span></span>

<span data-ttu-id="2ee18-400">プロジェクトのルートディレクトリにあるファイル **webpack.config.js** を開き、以下の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-400">Open the file **webpack.config.js** in the root directory of the project and complete the following steps.</span></span>

1. <span data-ttu-id="2ee18-401">`config`オブジェクト内で`entry`オブジェクトを探し、`popup`の新しいエントリーを追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-401">Locate the `entry` object within the `config` object and add a new entry for `popup`.</span></span>

    ```js
    popup: "./src/dialogs/popup.js"
    ```

    <span data-ttu-id="2ee18-402">これを実行すると、新しい`entry`オブジェクトは次のようになります。</span><span class="sxs-lookup"><span data-stu-id="2ee18-402">After you've done this, the new `entry` object will look like this:</span></span>

    ```js
    entry: {
      polyfill: "@babel/polyfill",
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
      popup: "./src/dialogs/popup.js"
    },
    ```
  
2. <span data-ttu-id="2ee18-403">`config`オブジェクト内で`plugins`配列を探し、次の新しいオブジェクトをその配列の末尾に追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-403">Locate the `plugins` array within the `config` object and add the following object to the end of that array.</span></span>

    ```js
    new HtmlWebpackPlugin({
      filename: "popup.html",
      template: "./src/dialogs/popup.html",
      chunks: ["polyfill", "popup"]
    })
    ```

    <span data-ttu-id="2ee18-404">これを実行すると、新しい`plugins`配列は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="2ee18-404">After you've done this, the new `plugins` array will look like this:</span></span>

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

3. <span data-ttu-id="2ee18-405">ローカル Web サーバーが実行中の場合は、ノード コマンド ウィンドウを閉じて終了します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-405">If the local web server is running, stop it by closing the node command window.</span></span>

4. <span data-ttu-id="2ee18-406">次のコマンドを実行してプロジェクトを再構築します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-406">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

### <a name="open-the-dialog-from-the-task-pane"></a><span data-ttu-id="2ee18-407">作業ウィンドウからダイアログを開く</span><span class="sxs-lookup"><span data-stu-id="2ee18-407">Open the dialog from the task pane</span></span>

1. <span data-ttu-id="2ee18-408">ファイル **./src/taskpane/taskpane.html** を開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-408">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="2ee18-409">`freeze-header` ボタンの `<button>` 要素を見つけ、その行の後に次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-409">Locate the `<button>` element for the `freeze-header` button, and add the following markup after that line:</span></span>

    ```html
    <button class="ms-Button" id="open-dialog">Open Dialog</button><br/><br/>
    ```

3. <span data-ttu-id="2ee18-p163">このダイアログでは、ユーザーに名前の入力を求めて、ユーザーの名前を作業ウィンドウに渡します。作業ウィンドウでは、それがラベルに表示されます。前の手順で追加した `button` の直後に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p163">The dialog will prompt the user to enter a name and pass the user's name to the task pane. The task pane will display it in a label. Immediately after the `button` that you just added, add the following markup:</span></span>

    ```html
    <label id="user-name"></label><br/><br/>
    ```

4. <span data-ttu-id="2ee18-413">ファイル **./src/taskpane/taskpane.js** を開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-413">Open the file **./src/taskpane/taskpane.js**.</span></span>

5. <span data-ttu-id="2ee18-p164">`Office.onReady` メソッドの呼び出し内で、クリック ハンドラーを `freeze-header` ボタンに割り当てる行を見つけ、その行の後に次のコードを追加します。後の手順で `openDialog` メソッドを作成します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p164">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `freeze-header` button, and add the following code after that line. You'll create the `openDialog` method in a later step.</span></span>

    ```js
    document.getElementById("open-dialog").onclick = openDialog;
    ```

6. <span data-ttu-id="2ee18-p165">ファイルの最後に次の宣言を追加します。この変数は、親ページの実行コンテキスト内のオブジェクトを保持するために使用され、ダイアログ ページの実行コンテキストへの仲介者として機能します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p165">Add the following declaration to the end of the file. This variable is used to hold an object in the parent page's execution context that acts as an intermediator to the dialog page's execution context.</span></span>

    ```js
    var dialog = null;
    ```

7. <span data-ttu-id="2ee18-p166">(`dialog` の宣言の後で) ファイルの最後に次の関数を追加します。このコードについて注意すべき重要なことは、そこに *ない* ものがあることであり、そのないものとは `Excel.run` の呼び出しです。これは、ダイアログを開く API はすべての Office アプリケーションで共有されるため、Excel 固有の API ではなく Office JavaScript 共通 API に含まれているからです。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p166">Add the following function to the end of the file (after the declaration of `dialog`). The important thing to notice about this code is what is *not* there: there is no call of `Excel.run`. This is because the API to open a dialog is shared among all Office applications, so it is part of the Office JavaScript Common API, not the Excel-specific API.</span></span>

    ```js
    function openDialog() {
        // TODO1: Call the Office Common API that opens a dialog
    }
    ```

8. <span data-ttu-id="2ee18-p167">`TODO1` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p167">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="2ee18-423">`displayDialogAsync` メソッドでは、画面の中央にダイアログを開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-423">The `displayDialogAsync` method opens a dialog in the center of the screen.</span></span>

   - <span data-ttu-id="2ee18-424">最初のパラメーターは、開くページの URL です。</span><span class="sxs-lookup"><span data-stu-id="2ee18-424">The first parameter is the URL of the page to open.</span></span>

   - <span data-ttu-id="2ee18-p168">2 番目のパラメーターでオプションを渡します。`height` と `width` は、Office アプリケーションのウィンドウ サイズの比率です。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p168">The second parameter passes options. `height` and `width` are percentages of the size of the Office application's window.</span></span>

    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},

        // TODO2: Add callback parameter.
    );
    ```

### <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a><span data-ttu-id="2ee18-427">ダイアログからのメッセージを処理してダイアログを閉じる</span><span class="sxs-lookup"><span data-stu-id="2ee18-427">Process the message from the dialog and close the dialog</span></span>

1. <span data-ttu-id="2ee18-p169">ファイル **./src/taskpane/taskpane.js** の `openDialog` 関数内で、`TODO2` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p169">Within the `openDialog` function in the file **./src/taskpane/taskpane.js**, replace `TODO2` with the following code. Note:</span></span>

   - <span data-ttu-id="2ee18-430">コールバックは、ダイアログが正常に開いた直後、ユーザーがダイアログで操作を行う前に実行されます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-430">The callback is executed immediately after the dialog successfully opens and before the user has taken any action in the dialog.</span></span>

   - <span data-ttu-id="2ee18-431">`result.value` は、親ページとダイアログ ページの実行コンテキストの間で仲介者として機能するオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="2ee18-431">The `result.value` is the object that acts as an intermediary between the execution contexts of the parent and dialog pages.</span></span>

   - <span data-ttu-id="2ee18-p170">このハンドラーは、`processMessage` 関数の呼び出しによって、ダイアログから送信されるあらゆる値を処理します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p170">The `processMessage` function will be created in a later step. This handler will process any values that are sent from the dialog page with calls of the `messageParent` function.</span></span>

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. <span data-ttu-id="2ee18-434">`openDialog` 関数の後に次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-434">Add the following function after the `openDialog` function.</span></span>

    ```js
    function processMessage(arg) {
        document.getElementById("user-name").innerHTML = arg.message;
        dialog.close();
    }
    ```

3. <span data-ttu-id="2ee18-435">プロジェクトに行ったすべての変更が保存されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="2ee18-435">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="2ee18-436">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="2ee18-436">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="2ee18-437">アドイン タスク ウィンドウが Excel でまだ開いていない場合は、[**ホーム**] タブに移動し、リボンの [**作業ウィンドウを表示**] ボタンを選択して開きます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-437">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="2ee18-438">作業ウィンドウで、**[ダイアログを開く]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="2ee18-438">Choose the **Open Dialog** button in the task pane.</span></span>

4. <span data-ttu-id="2ee18-p171">ダイアログが開いたら、ドラッグしたりサイズ変更したりします。ワークシートを操作して、作業ウィンドウの他のボタンを押すことはできますが、同じ作業ウィンドウのページから 2 番目のダイアログを起動することはできないことに注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p171">While the dialog is open, drag it and resize it. Note that you can interact with the worksheet and press other buttons on the task pane, but you cannot launch a second dialog from the same task pane page.</span></span>

5. <span data-ttu-id="2ee18-p172">ダイアログで、名前を入力して [**OK**] ボタンを選択します。作業ウィンドウに名前が表示され、ダイアログが閉じられます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p172">In the dialog, enter a name and choose the **OK** button. The name appears on the task pane and the dialog closes.</span></span>

6. <span data-ttu-id="2ee18-p173">必要に応じて、`processMessage` 関数の行 `dialog.close();` をコメントにします。このセクションの手順を繰り返します。ダイアログは開いたままで、名前を変更できます。右上隅の **X** ボタンを押して、手動で閉じることができます。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p173">Optionally, comment out the line `dialog.close();` in the `processMessage` function. Then repeat the steps of this section. The dialog stays open and you can change the name. You can close it manually by pressing the **X** button in the upper right corner.</span></span>

    ![Excelのスクリーンショット。アドインの作業ウィンドウに [ダイアログを開く] ボタンが表示され、ワークシートの上にダイアログ ボックスが表示されます。](../images/excel-tutorial-dialog-open-2.png)

## <a name="next-steps"></a><span data-ttu-id="2ee18-448">次の手順</span><span class="sxs-lookup"><span data-stu-id="2ee18-448">Next steps</span></span>

<span data-ttu-id="2ee18-p174">このチュートリアルでは、Excel ブック内のテーブル、グラフ、ワークシート、ダイアログの操作を行う、Excel 作業ウィンドウ アドインを作成しました。Excel アドインの構築に関する詳細については、次の記事にお進みください。</span><span class="sxs-lookup"><span data-stu-id="2ee18-p174">In this tutorial, you've created an Excel task pane add-in that interacts with tables, charts, worksheets, and dialogs in an Excel workbook. To learn more about building Excel add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="2ee18-451">Excel アドインの概要</span><span class="sxs-lookup"><span data-stu-id="2ee18-451">Excel add-ins overview</span></span>](../excel/excel-add-ins-overview.md)

## <a name="see-also"></a><span data-ttu-id="2ee18-452">関連項目</span><span class="sxs-lookup"><span data-stu-id="2ee18-452">See also</span></span>

- [<span data-ttu-id="2ee18-453">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="2ee18-453">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="2ee18-454">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="2ee18-454">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="2ee18-455">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="2ee18-455">Excel JavaScript object model in Office Add-ins</span></span>](../excel/excel-add-ins-core-concepts.md)
