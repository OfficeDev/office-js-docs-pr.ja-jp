<span data-ttu-id="e5a28-101">チュートリアルのこの手順では、ドキュメントに画像、HTML、テーブルを挿入する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-101">In this step of the tutorial, you'll learn how to insert images, HTML, and tables into the document.</span></span>

> [!NOTE]
> <span data-ttu-id="e5a28-p101">このページでは、Word アドインのチュートリアルの個々の手順について説明します。このページに検索エンジンの結果から、またはその他の直接リンクからアクセスした場合は、「[Word アドインのチュートリアル](../tutorials/word-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。</span><span class="sxs-lookup"><span data-stu-id="e5a28-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="insert-an-image"></a><span data-ttu-id="e5a28-104">画像の挿入</span><span class="sxs-lookup"><span data-stu-id="e5a28-104">Insert an image</span></span>

1. <span data-ttu-id="e5a28-105">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="e5a28-105">Open the project in your code editor.</span></span>
2. <span data-ttu-id="e5a28-106">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e5a28-106">Open the file index.html.</span></span>
3. <span data-ttu-id="e5a28-107">`replace-text` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-107">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-image">Insert Image</button>
    </div>
    ```

4. <span data-ttu-id="e5a28-108">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e5a28-108">Open the app.js file.</span></span>

5. <span data-ttu-id="e5a28-109">ファイルの先頭近くにある、use-strict 行のすぐ下に次の行を追加します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-109">Near the top of the file, just below the use-strict line, add the following line.</span></span> <span data-ttu-id="e5a28-110">この行は、別のファイルから変数をインポートします。</span><span class="sxs-lookup"><span data-stu-id="e5a28-110">This line imports a variable from another file.</span></span> <span data-ttu-id="e5a28-111">この変数は、画像をエンコードする Base 64 文字列です。</span><span class="sxs-lookup"><span data-stu-id="e5a28-111">The variable is a base 64 string that encodes an image.</span></span> <span data-ttu-id="e5a28-112">エンコードされた文字列を表示するには、プロジェクトのルートにある base64Image.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e5a28-112">To see the encoded string, open the base64Image.js file in the root of the project.</span></span>

    ```js
    import { base64Image } from "./base64Image";
    ```

6. <span data-ttu-id="e5a28-113">`replace-text` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-113">Below the line that assigns a click handler to the `replace-text` button, add the following code:</span></span>

    ```js
    $('#insert-image').click(insertImage);
    ```

7. <span data-ttu-id="e5a28-114">`replaceText` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-114">Below the `replaceText` function, add the following function:</span></span>

    ```js
    function insertImage() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert an image.

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

8. <span data-ttu-id="e5a28-115">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e5a28-115">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="e5a28-116">この行により、Base 64 でエンコードされた画像がドキュメントの末尾に挿入されることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5a28-116">Note that this line inserts the base 64 encoded image at the end of the document.</span></span> <span data-ttu-id="e5a28-117">(`Paragraph` オブジェクトにも `insertInlinePictureFromBase64` メソッドやその他の `insert*` メソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="e5a28-117">(The `Paragraph` object also has an `insertInlinePictureFromBase64` method and other `insert*` methods.</span></span> <span data-ttu-id="e5a28-118">例については、次の insertHTML セクションを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="e5a28-118">See the following insertHTML section for an example.)</span></span>

    ```js
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    ```

## <a name="insert-html"></a><span data-ttu-id="e5a28-119">HTML の挿入</span><span class="sxs-lookup"><span data-stu-id="e5a28-119">Insert HTML</span></span>

1. <span data-ttu-id="e5a28-120">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e5a28-120">Open the file index.html.</span></span>
2. <span data-ttu-id="e5a28-121">`insert-image` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-121">Below the `div` that contains the `insert-image` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-html">Insert HTML</button>
    </div>
    ```

3. <span data-ttu-id="e5a28-122">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e5a28-122">Open the app.js file.</span></span>

4. <span data-ttu-id="e5a28-123">`insert-image` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-123">Below the line that assigns a click handler to the `insert-image` button, add the following code:</span></span>

    ```js
    $('#insert-html').click(insertHTML);
    ```

5. <span data-ttu-id="e5a28-124">`insertImage` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-124">Below the `insertImage` function, add the following function:</span></span>

    ```js
    function insertHTML() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert a string of HTML.

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

6. <span data-ttu-id="e5a28-p104">`TODO1` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5a28-p104">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="e5a28-127">最初の行は、ドキュメントの末尾に空白の段落を追加します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-127">The first line adds a blank paragraph to the end of the document.</span></span> 
   - <span data-ttu-id="e5a28-128">2 行目は、その段落の末尾に HTML の文字列を挿入します。具体的には、Verdana フォントで書式設定された段落と、Word 文書の既定のスタイルが設定された段落の 2 つの段落が挿入されます。</span><span class="sxs-lookup"><span data-stu-id="e5a28-128">The second line inserts a string of HTML at the end of the paragraph; specifically two paragraphs, one formatted with Verdana font, the other with the default styling of the Word document.</span></span> <span data-ttu-id="e5a28-129">(`insertImage` メソッドで説明したように、`context.document.body` オブジェクトにも `insert*` メソッドがあります)。</span><span class="sxs-lookup"><span data-stu-id="e5a28-129">(As you saw in the `insertImage` method earlier, the `context.document.body` object also has the `insert*` methods.)</span></span>

    ```js
    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    ```

## <a name="insert-table"></a><span data-ttu-id="e5a28-130">テーブルの挿入</span><span class="sxs-lookup"><span data-stu-id="e5a28-130">Insert Table</span></span>

1. <span data-ttu-id="e5a28-131">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e5a28-131">Open the file index.html.</span></span>
2. <span data-ttu-id="e5a28-132">`insert-html` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-132">Below the `div` that contains the `insert-html` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-table">Insert Table</button>
    </div>
    ```

3. <span data-ttu-id="e5a28-133">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e5a28-133">Open the app.js file.</span></span>

4. <span data-ttu-id="e5a28-134">`insert-html` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-134">Below the line that assigns a click handler to the `insert-html` button, add the following code:</span></span>

    ```js
    $('#insert-table').click(insertTable);
    ```

5. <span data-ttu-id="e5a28-135">`insertHTML` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-135">Below the `insertHTML` function, add the following function:</span></span>

    ```js
    function insertTable() {
        Word.run(function (context) {

            // TODO1: Queue commands to get a reference to the paragraph
            //        that will proceed the table.

            // TODO2: Queue commands to create a table and populate it with data.

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

6. <span data-ttu-id="e5a28-136">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e5a28-136">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="e5a28-137">この行は `ParagraphCollection.getFirst` メソッドを使用して最初の段落への参照を取得し、次に `Paragraph.getNext` メソッドを使用して 2 番目の段落への参照を取得することに注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5a28-137">Note that this line uses the `ParagraphCollection.getFirst` method to get a reference ot the first paragraph and then uses the `Paragraph.getNext` method to get a reference to the second paragraph.</span></span>

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ```

7. <span data-ttu-id="e5a28-p107">`TODO2` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5a28-p107">Replace `TODO2` with the following code. Note:</span></span>
   - <span data-ttu-id="e5a28-140">`insertTable` メソッドの最初の 2 つのパラメーターは、行と列の数を指定します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-140">The first two parameters of the `insertTable` method specify the number of rows and columns.</span></span>
   - <span data-ttu-id="e5a28-141">3 番目のパラメーターは、テーブルを挿入する場所を指定します (この例では段落の後)。</span><span class="sxs-lookup"><span data-stu-id="e5a28-141">The third parameter specifies where to insert the table, in this case after the paragraph.</span></span>
   - <span data-ttu-id="e5a28-142">4 番目のパラメーターは、テーブルのセルの値を設定する 2 次元配列です。</span><span class="sxs-lookup"><span data-stu-id="e5a28-142">The fourth parameter is a two-dimensional array that sets the values of the table cells.</span></span>
   - <span data-ttu-id="e5a28-143">このテーブルには既定のスタイルがそのまま設定されますが、`insertTable` メソッドがさまざまなメンバーを持つ `Table` オブジェクトを返し、その一部がテーブルのスタイル設定に使用されます。</span><span class="sxs-lookup"><span data-stu-id="e5a28-143">The table will have plain default styling, but the `insertTable` method returns a `Table` object with many members, some of which are used to style the table.</span></span>

    ```js
    const tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    ```

## <a name="test-the-add-in"></a><span data-ttu-id="e5a28-144">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="e5a28-144">Test the add-in</span></span>


1. <span data-ttu-id="e5a28-145">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、Ctrl + C を 2 回入力して実行中の Web サーバーを停止します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-145">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="e5a28-146">それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-146">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="e5a28-147">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。</span><span class="sxs-lookup"><span data-stu-id="e5a28-147">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="e5a28-148">これを行うには、プロンプトが表示されてビルド コマンドを入力できるようにするため、サーバー プロセスを強制終了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e5a28-148">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="e5a28-149">ビルド後に、サーバーを再起動します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-149">After the build, restart the server.</span></span> <span data-ttu-id="e5a28-150">次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-150">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="e5a28-151">`npm run build` コマンドを実行し、Office アドインを実行できるすべてのホストでサポートされている以前のバージョンの JavaScript に ES6 ソース コードをトランスパイルします。</span><span class="sxs-lookup"><span data-stu-id="e5a28-151">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="e5a28-152">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-152">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="e5a28-153">作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="e5a28-153">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="e5a28-154">作業ウィンドウで **[段落の挿入]** を少なくとも 3 回選択し、ドキュメントに段落がいくつかあることを確認します。</span><span class="sxs-lookup"><span data-stu-id="e5a28-154">In the taskpane, choose **Insert Paragraph** at least three times to ensure that there are a few paragraphs in the document.</span></span>
6. <span data-ttu-id="e5a28-155">**[画像の挿入]** ボタンをクリックし、ドキュメントの末尾に画像が挿入されることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5a28-155">Choose the **Insert Image** button and note that an image is inserted at the end of the document.</span></span>
7. <span data-ttu-id="e5a28-156">**[HTML の挿入]** ボタンをクリックし、ドキュメントの末尾に 2 つの段落が挿入され、最初の段落に Verdana フォントが設定されていることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5a28-156">Choose the **Insert HTML** button and note that two paragraphs are inserted at the end of the document, and that the first one has Verdana font.</span></span>
8. <span data-ttu-id="e5a28-157">**[テーブルの挿入]** ボタンをクリックし、2 番目の段落の後にテーブルが挿入されることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5a28-157">Choose the **Insert Table** button and note that a table is inserted after the second paragraph.</span></span>

    ![Word のチュートリアル - 画像、HTML、テーブルの挿入](../images/word-tutorial-insert-image-html-table.png)
