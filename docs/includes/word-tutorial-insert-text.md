<span data-ttu-id="ceec6-101">チュートリアルのこの手順では、ユーザーが現在使用している Word のバージョンをアドインがサポートしているかどうかをプログラムによってテストし、ドキュメントにパラグラフを挿入します。</span><span class="sxs-lookup"><span data-stu-id="ceec6-101">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Word, and then insert a paragraph in the document.</span></span>

> [!NOTE]
> <span data-ttu-id="ceec6-p101">このページでは、Word アドインのチュートリアルの個々の手順について説明します。このページに検索エンジンの結果から、またはその他の直接リンクからアクセスした場合は、「[Word アドインのチュートリアル](../tutorials/word-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。</span><span class="sxs-lookup"><span data-stu-id="ceec6-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="code-the-add-in"></a><span data-ttu-id="ceec6-104">アドインのコードを作成する</span><span class="sxs-lookup"><span data-stu-id="ceec6-104">Code the add-in</span></span>

1. <span data-ttu-id="ceec6-105">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="ceec6-105">Open the project in your code editor.</span></span>
2. <span data-ttu-id="ceec6-106">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="ceec6-106">Open the file index.html.</span></span>
3. <span data-ttu-id="ceec6-107">`TODO1` を次のマークアップに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="ceec6-107">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button>
    ```

4. <span data-ttu-id="ceec6-108">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="ceec6-108">Open the app.js file.</span></span>
5. <span data-ttu-id="ceec6-109">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="ceec6-109">Replace the `TODO1` with the following code.</span></span> <span data-ttu-id="ceec6-110">このコードでは、ユーザーの Word のバージョンが、このチュートリアルのすべての段階で使用するすべての API を含んでいる Word.js のバージョンをサポートしているかどうかを調べます。</span><span class="sxs-lookup"><span data-stu-id="ceec6-110">This code determines whether the user's version of Word supports a version of Word.js that includes all the APIs that are used in all the stages of this tutorial.</span></span> <span data-ttu-id="ceec6-111">運用アドインでは、未サポートの API を呼び出す UI を非表示または無効化する条件ブロックの本体を使用してください。</span><span class="sxs-lookup"><span data-stu-id="ceec6-111">In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs.</span></span> <span data-ttu-id="ceec6-112">これにより、ユーザーは、自分が使用している Word のバージョンでサポートされているアドインの部分を使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="ceec6-112">This will enable the user to still use the parts of the add-in that are supported by their version of Word.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
        console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }
    ```

6. <span data-ttu-id="ceec6-113">`TODO2` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="ceec6-113">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#insert-paragraph').click(insertParagraph);
    ```

7. <span data-ttu-id="ceec6-114">`TODO3` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="ceec6-114">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="ceec6-115">次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="ceec6-115">Note the following:</span></span>
   - <span data-ttu-id="ceec6-116">Word.js のビジネス ロジックは、`Word.run` に渡される関数に追加されます。</span><span class="sxs-lookup"><span data-stu-id="ceec6-116">Your Word.js business logic will be added to the function that is passed to `Word.run`.</span></span> <span data-ttu-id="ceec6-117">このロジックは、すぐには実行されません。</span><span class="sxs-lookup"><span data-stu-id="ceec6-117">This logic does not execute immediately.</span></span> <span data-ttu-id="ceec6-118">その代わりに、保留中のコマンドのキューに追加されます。</span><span class="sxs-lookup"><span data-stu-id="ceec6-118">Instead, it is added to a queue of pending commands.</span></span>
   - <span data-ttu-id="ceec6-119">`context.sync` メソッドは、キューに登録されたすべてのコマンドを、実行するために Word に送信します。</span><span class="sxs-lookup"><span data-stu-id="ceec6-119">The `context.sync` method sends all queued commands to Word for execution.</span></span>
   - <span data-ttu-id="ceec6-120">`Word.run` の後に `catch` ブロックを続けます。</span><span class="sxs-lookup"><span data-stu-id="ceec6-120">The `Word.run` is followed by a `catch` block.</span></span> <span data-ttu-id="ceec6-121">これは、どのような場合にも当てはまるベスト プラクティスです。</span><span class="sxs-lookup"><span data-stu-id="ceec6-121">This is a best practice that you should always follow.</span></span> 

    ```js
    function insertParagraph() {
        Word.run(function (context) {

            // TODO4: Queue commands to insert a paragraph into the document.

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

8. <span data-ttu-id="ceec6-p106">`TODO4` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="ceec6-p106">Replace `TODO4` with the following code. Note:</span></span>
   - <span data-ttu-id="ceec6-124">`insertParagraph` メソッドの最初のパラメーターは、新しい段落のテキストです。</span><span class="sxs-lookup"><span data-stu-id="ceec6-124">The first parameter to the `insertParagraph` method is the text for the new paragraph.</span></span>
   - <span data-ttu-id="ceec6-125">2 番目のパラメーターは、段落を挿入する本文内の場所です。</span><span class="sxs-lookup"><span data-stu-id="ceec6-125">The second parameter is the location within the body where the paragraph will be inserted.</span></span> <span data-ttu-id="ceec6-126">親オブジェクトが本文の場合、段落の挿入に使用できるその他のオプションには、End と Replace があります。</span><span class="sxs-lookup"><span data-stu-id="ceec6-126">Other options for insert paragraph, when the parent object is the body, are "End" and "Replace".</span></span>

    ```js
    const docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office Online.",
                            "Start");
    ```

## <a name="test-the-add-in"></a><span data-ttu-id="ceec6-127">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="ceec6-127">Test the add-in</span></span>

1. <span data-ttu-id="ceec6-128">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="ceec6-128">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
2. <span data-ttu-id="ceec6-129">`npm run build` コマンドを実行し、Office アドインを実行できるすべてのホストでサポートされている以前のバージョンの JavaScript に ES6 ソース コードをトランスパイルします。</span><span class="sxs-lookup"><span data-stu-id="ceec6-129">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="ceec6-130">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="ceec6-130">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="ceec6-131">次のいずれかの方法を使用して、アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="ceec6-131">Sideload the add-in by using one of the following methods:</span></span>
    - <span data-ttu-id="ceec6-132">Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="ceec6-132">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="ceec6-133">Word Online: [Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="ceec6-133">Word Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="ceec6-134">iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="ceec6-134">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>
5. <span data-ttu-id="ceec6-135">Word の **[ホーム]** メニューで、**[作業ウィンドウの表示]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="ceec6-135">On the **Home** menu of Word, select **Show Taskpane**.</span></span>
6. <span data-ttu-id="ceec6-136">作業ウィンドウで、**[段落の挿入]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="ceec6-136">In the taskpane, choose **Insert Paragraph**.</span></span>
7. <span data-ttu-id="ceec6-137">段落に変更を加えます。</span><span class="sxs-lookup"><span data-stu-id="ceec6-137">Make a change in the paragraph.</span></span>
8. <span data-ttu-id="ceec6-138">**[段落の挿入]** をもう一度選択します。</span><span class="sxs-lookup"><span data-stu-id="ceec6-138">Choose **Insert Paragraph** again.</span></span> <span data-ttu-id="ceec6-139">`insertParagraph` メソッドはドキュメントの本文の開始位置に挿入を行うため、新しい段落は前の段落より上に追加されます。</span><span class="sxs-lookup"><span data-stu-id="ceec6-139">Note that the new paragraph is above the previous one because the `insertParagraph` method is inserting at the "start" of the document's body.</span></span>

    ![Word のチュートリアル - 段落の挿入](../images/word-tutorial-insert-paragraph.png)
