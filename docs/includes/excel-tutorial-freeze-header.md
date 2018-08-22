<span data-ttu-id="288b2-101">表がとても長く、行を参照するためにスクロールしなければならない場合、ヘッダー行が画面の外に移動して見えなくなることがあります。</span><span class="sxs-lookup"><span data-stu-id="288b2-101">When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight.</span></span> <span data-ttu-id="288b2-102">チュートリアルのこの手順では、以前に作成した表のヘッダー行を固定して、ワークシートを下にスクロールしても表示されるようにします。</span><span class="sxs-lookup"><span data-stu-id="288b2-102">In this step of the tutorial, you'll freeze the header row of the table that you created previously, so that it remains visible even as the user scrolls down the worksheet.</span></span> 

> [!NOTE]
> <span data-ttu-id="288b2-103">このページでは、Excel のアドインのチュートリアルの個々 の手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="288b2-103">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="288b2-104">このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[Excel アドインのチュートリアル](../tutorials/excel-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。</span><span class="sxs-lookup"><span data-stu-id="288b2-104">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="freeze-the-tables-header-row"></a><span data-ttu-id="288b2-105">表のヘッダー行を固定する</span><span class="sxs-lookup"><span data-stu-id="288b2-105">Freeze the table's header row</span></span>

1. <span data-ttu-id="288b2-106">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="288b2-106">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="288b2-107">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="288b2-107">Open the file index.html.</span></span>
3. <span data-ttu-id="288b2-108">ボタンを格納している `div` の下に、次のマークアップを追加します。`create-chart`</span><span class="sxs-lookup"><span data-stu-id="288b2-108">Below the `div` that contains the `create-chart` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="freeze-header">Freeze Header</button>            
    </div>
    ```

4. <span data-ttu-id="288b2-109">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="288b2-109">Open the app.js file.</span></span>

5. <span data-ttu-id="288b2-110">`create-chart` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="288b2-110">Below the line that assigns a click handler to the `create-chart` button, add the following code:</span></span>

    ```js
    $('#freeze-header').click(freezeHeader);
    ```

6. <span data-ttu-id="288b2-111">`createChart` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="288b2-111">Below the `createChart` function add the following function:</span></span>

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

7. <span data-ttu-id="288b2-p103">`TODO1` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="288b2-p103">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="288b2-114">`Worksheet.freezePanes` コレクションは、ワークシートのスクロール操作時に、ワークシート上でピン留めつまり固定される一式のペインのことです。</span><span class="sxs-lookup"><span data-stu-id="288b2-114">The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.</span></span>
   - <span data-ttu-id="288b2-p104">`freezeRows` メソッドでは、上から数えた行数を、ピン留めする位置のパラメーターとして使用します。`1` を渡して最初の行を適所にピン留めします。</span><span class="sxs-lookup"><span data-stu-id="288b2-p104">The `freezeRows` method takes as a parameter the number of rows, from the top that are to be pinned in place. We pass `1` to pin the first row in place.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="288b2-117">アドインのテスト</span><span class="sxs-lookup"><span data-stu-id="288b2-117">Test the add-in</span></span>

1. <span data-ttu-id="288b2-118">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、Ctrl-C を 2 回入力して実行中の Web サーバーを停止します。</span><span class="sxs-lookup"><span data-stu-id="288b2-118">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="288b2-119">それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="288b2-119">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="288b2-120">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。</span><span class="sxs-lookup"><span data-stu-id="288b2-120">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="288b2-121">そのためには、ビルド コマンドの入力を求めるプロンプトが表示されるように、サーバー プロセスを強制終了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="288b2-121">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="288b2-122">ビルド後に、サーバーを再起動します。</span><span class="sxs-lookup"><span data-stu-id="288b2-122">After the build, you restart the server.</span></span> <span data-ttu-id="288b2-123">次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="288b2-123">The next few steps carry out this process.</span></span>

1. <span data-ttu-id="288b2-124">コマンドを実行して、ES6 ソース コードを Internet Explorer でサポートされている以前のバージョンの JavaScript にトランスパイルします (これは、Excel アドインを実行するために Excel の内部で使用されます)。`npm run build`</span><span class="sxs-lookup"><span data-stu-id="288b2-124">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
2. <span data-ttu-id="288b2-125">コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。`npm start`</span><span class="sxs-lookup"><span data-stu-id="288b2-125">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="288b2-126">作業ウィンドウを再読み込みするために、そのウィンドウを閉じ、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="288b2-126">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
6. <span data-ttu-id="288b2-127">ワークシート内に表があれば、削除します。</span><span class="sxs-lookup"><span data-stu-id="288b2-127">If the table is in the worksheet, delete it.</span></span>
7. <span data-ttu-id="288b2-128">作業ウィンドウで、**[Create Table]** (表の作成) を選択します。</span><span class="sxs-lookup"><span data-stu-id="288b2-128">In the taskpane, choose **Create Table**.</span></span> 
8. <span data-ttu-id="288b2-129">**[Freeze Header]** (ヘッダーを固定) ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="288b2-129">Choose the **Freeze Header** button.</span></span>
9. <span data-ttu-id="288b2-130">ヘッダー以降の行が画面の外に出て見えなくなるまでワークシートを下にスクロールしても、表のヘッダーが最上部に表示されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="288b2-130">Scroll down the worksheet enough to to see that the table header remains visible at the top even when the higher rows scroll out of sight.</span></span>

    ![Excel のチュートリアル - ヘッダーの固定](../images/excel-tutorial-freeze-header.png)
