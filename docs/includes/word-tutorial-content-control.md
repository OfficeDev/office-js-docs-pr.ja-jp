<span data-ttu-id="c8881-101">このチュートリアルの手順では、ドキュメント内にリッチ テキスト コンテンツ コントロールを作成する方法、およびそのコントロールにコンテンツを挿入したり置き換えたりする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="c8881-101">In this step of the tutorial, you'll learn how to create Rich Text content controls in the document, and then how to insert and replace content in the controls.</span></span> 

> [!NOTE]
> <span data-ttu-id="c8881-p101">このページでは、Word アドインのチュートリアルの個々の手順について説明します。このページに検索エンジンの結果から、またはその他の直接リンクからアクセスした場合は、「[Word アドインのチュートリアル](../tutorials/word-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。</span><span class="sxs-lookup"><span data-stu-id="c8881-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

<span data-ttu-id="c8881-104">チュートリアルのこの手順を開始する前に、Word UI からリッチ テキスト コンテンツ コントロールを作成して操作し、コントロールとそのプロパティを理解しておくことをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="c8881-104">Before you start this step of the tutorial, we recommend that you create and manipulate Rich Text content controls through the Word UI, so you can be familiar with the controls and their properties.</span></span> <span data-ttu-id="c8881-105">詳細については、「[ユーザーが Word 上で記入または印刷するフォームを作成する](https://support.office.com/en-us/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c8881-105">For details, see [Create forms that users complete or print in Word](https://support.office.com/en-us/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span></span>

> [!NOTE]
> <span data-ttu-id="c8881-106">UI から Word 文書に追加できるコンテンツ コントロールにはいくつかの種類がありますが、Word.js では現在のところリッチ テキスト コンテンツ コントロールのみがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="c8881-106">There are several types of content controls that can be added to a Word document through the UI; but currently only Rich Text content controls are supported by Word.js.</span></span>


## <a name="create-a-content-control"></a><span data-ttu-id="c8881-107">コンテンツ コントロールを作成する</span><span class="sxs-lookup"><span data-stu-id="c8881-107">Create a content control</span></span>

1. <span data-ttu-id="c8881-108">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="c8881-108">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="c8881-109">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="c8881-109">Open the file index.html.</span></span>
3. <span data-ttu-id="c8881-110">ボタンを格納している `div` の下に、次のマークアップを追加します。`replace-text`</span><span class="sxs-lookup"><span data-stu-id="c8881-110">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="create-content-control">Create Content Control</button>            
    </div>
    ```

4. <span data-ttu-id="c8881-111">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="c8881-111">Open the app.js file.</span></span>

5. <span data-ttu-id="c8881-112">ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。`insert-table`</span><span class="sxs-lookup"><span data-stu-id="c8881-112">Below the line that assigns a click handler to the `insert-table` button, add the following code:</span></span>

    ```js
    $('#create-content-control').click(createContentControl);
    ```

6. <span data-ttu-id="c8881-113">関数の下に、次の関数を追加します。`insertTable`</span><span class="sxs-lookup"><span data-stu-id="c8881-113">Below the `insertTable` function, add the following function:</span></span>

    ```js
    function createContentControl() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to create a content control.

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

7. <span data-ttu-id="c8881-p103"> `TODO1` を次のコードに置換します。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="c8881-p103">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="c8881-116">このコードの目的は、コンテンツ コントロール内の Office 365 という語句をラップすることです。</span><span class="sxs-lookup"><span data-stu-id="c8881-116">This code is intended to wrap the phrase "Office 365" in a content control.</span></span> <span data-ttu-id="c8881-117">これは前提を単純化し、文字列は存在しており、ユーザーがその文字列を選択したものとしています。</span><span class="sxs-lookup"><span data-stu-id="c8881-117">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>
   - <span data-ttu-id="c8881-118">プロパティは、コンテンツ コントロールの表示タイトルを指定します。`ContentControl.title`</span><span class="sxs-lookup"><span data-stu-id="c8881-118">The `ContentControl.title` property specifies the visible title of the content control.</span></span> 
   - <span data-ttu-id="c8881-119">プロパティは、`ContentControlCollection.getByTag` メソッドを使用してコンテンツ コントロールへの参照を取得するために使用できるタグを指定します。これを後述する関数で使用します。`ContentControl.tag`</span><span class="sxs-lookup"><span data-stu-id="c8881-119">The `ContentControl.tag` property specifies an tag that can be used to get a reference to a content control using the `ContentControlCollection.getByTag` method, which you'll use in a later function.</span></span> 
   - <span data-ttu-id="c8881-120">プロパティは、コントロールの外観を指定します。`ContentControl.appearance`</span><span class="sxs-lookup"><span data-stu-id="c8881-120">The `ContentControl.appearance` property specifies the visual look of the control.</span></span> <span data-ttu-id="c8881-121">Tags という値を使用すると、コントロールは開始タグと終了タグにラップされます。開始タグには、コンテンツ コントロールのタイトルが設定されます。</span><span class="sxs-lookup"><span data-stu-id="c8881-121">Using the value "Tags" means that the control will be wrapped in opening and closing tags, and the opening tag will have the content control's title.</span></span> <span data-ttu-id="c8881-122">その他の値として、BoundingBox と None が使用できます。</span><span class="sxs-lookup"><span data-stu-id="c8881-122">Other possible values are "BoundingBox" and "None".</span></span>
   - <span data-ttu-id="c8881-123">プロパティは、タグまたは境界ボックスの境界線の色を指定します。`ContentControl.color`</span><span class="sxs-lookup"><span data-stu-id="c8881-123">The `ContentControl.color` property specifies the color of the tags or the border of the bounding box.</span></span>

    ```js
    const serviceNameRange = context.document.getSelection();
    const serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ``` 

## <a name="replace-the-content-of-the-content-control"></a><span data-ttu-id="c8881-124">コンテンツ コントロールのコンテンツを置き換える</span><span class="sxs-lookup"><span data-stu-id="c8881-124">Replace the content of the content control</span></span>

1. <span data-ttu-id="c8881-125">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="c8881-125">Open the file index.html.</span></span>
2. <span data-ttu-id="c8881-126">ボタンを格納している `div` の下に、次のマークアップを追加します。`create-content-control`</span><span class="sxs-lookup"><span data-stu-id="c8881-126">Below the `div` that contains the `create-content-control` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="replace-content-in-control">Rename Service</button>            
    </div>
    ```

3. <span data-ttu-id="c8881-127">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="c8881-127">Open the app.js file.</span></span>

4. <span data-ttu-id="c8881-128">ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。`create-content-control`</span><span class="sxs-lookup"><span data-stu-id="c8881-128">Below the line that assigns a click handler to the `create-content-control` button, add the following code:</span></span>

    ```js
    $('#replace-content-in-control').click(replaceContentInControl);
    ```

5. <span data-ttu-id="c8881-129">関数の下に、次の関数を追加します。`createContentControl`</span><span class="sxs-lookup"><span data-stu-id="c8881-129">Below the `createContentControl` function, add the following function:</span></span>

    ```js
    function replaceContentInControl() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to replace the text in the Service Name
            //        content control.

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

7. <span data-ttu-id="c8881-130"> `TODO1` を次のコードに置換します。</span><span class="sxs-lookup"><span data-stu-id="c8881-130">Replace `TODO1` with the following code.</span></span> 
    > [!NOTE]
    > <span data-ttu-id="c8881-131"> `ContentControlCollection.getByTag` メソッドは、特定のタグの全てのコンテンツコントロールの `ContentControlCollection` を返します。</span><span class="sxs-lookup"><span data-stu-id="c8881-131">The `ContentControlCollection.getByTag` method returns a `ContentControlCollection` of all content controls of the specified tag.</span></span> <span data-ttu-id="c8881-132"> `getFirst` を使って、目的のコントロールへの参照を取得します。</span><span class="sxs-lookup"><span data-stu-id="c8881-132">We use `getFirst` to get a reference to the desired control.</span></span>

    ```js
    const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="c8881-133">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="c8881-133">Test the add-in</span></span>

1. <span data-ttu-id="c8881-134">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、Ctrl + C を 2 回入力して実行中の Web サーバーを停止します。</span><span class="sxs-lookup"><span data-stu-id="c8881-134">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="c8881-135">それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="c8881-135">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
     > [!NOTE]
     > <span data-ttu-id="c8881-136">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。</span><span class="sxs-lookup"><span data-stu-id="c8881-136">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="c8881-137">これを行うには、プロンプトが表示されてビルド コマンドを入力できるようにするため、サーバー プロセスを強制終了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c8881-137">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="c8881-138">ビルド後に、サーバーを再起動します。</span><span class="sxs-lookup"><span data-stu-id="c8881-138">After the build, restart the server.</span></span> <span data-ttu-id="c8881-139">次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="c8881-139">The next few steps carry out this process.</span></span>
2. <span data-ttu-id="c8881-140">コマンドを実行し、Office アドインを実行できるすべてのホストでサポートされている以前のバージョンの JavaScript に ES6 ソース コードをトランスパイルします。`npm run build`</span><span class="sxs-lookup"><span data-stu-id="c8881-140">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="c8881-141">コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。`npm start`</span><span class="sxs-lookup"><span data-stu-id="c8881-141">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="c8881-142">作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="c8881-142">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="c8881-143">作業ウィンドウで **[段落の挿入]** を選択し、文書の先頭が Office 365 となっている段落があることを確認します。</span><span class="sxs-lookup"><span data-stu-id="c8881-143">In the taskpane, choose **Insert Paragraph** to ensure that there is a paragraph with "Office 365" at the top of the document.</span></span>
6. <span data-ttu-id="c8881-144">追加した段落の Office 365 という語句を選択し、**[コンテンツ コントロールの作成]** ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="c8881-144">Select the phrase "Office 365" in the paragraph you just added, and then choose the **Create Content Control** button.</span></span> <span data-ttu-id="c8881-145">Service Name というラベルが付いたタグで語句がラップされていることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="c8881-145">Note that the phrase is wrapped in tags labelled "Service Name".</span></span>
7. <span data-ttu-id="c8881-146">**[サービス名の変更]** ボタンを選択し、コンテンツ コントロールのテキストが Fabrikam Online Productivity Suite に変わることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="c8881-146">Choose the **Rename Service** button and note that the text of the content control changes to "Fabrikam Online Productivity Suite".</span></span>

    ![Word のチュートリアル - コンテンツ コントロールの作成とテキストの変更](../images/word-tutorial-content-control.png)
