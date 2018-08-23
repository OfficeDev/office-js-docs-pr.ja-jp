<span data-ttu-id="2a398-101">チュートリアルのこの手順では、テキストのフォントを変更して、テキストに組み込みのスタイルやカスタム スタイルを使用します。</span><span class="sxs-lookup"><span data-stu-id="2a398-101">In this step of the tutorial, you'll change the font of text, and use both built-in and custom styles on the text.</span></span>

> [!NOTE]
> <span data-ttu-id="2a398-p101">このページでは、Word アドインのチュートリアルの個々の手順について説明します。このページに検索エンジンの結果から、またはその他の直接リンクからアクセスした場合は、「[Word アドインのチュートリアル](../tutorials/word-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。</span><span class="sxs-lookup"><span data-stu-id="2a398-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="apply-a-built-in-style-to-text"></a><span data-ttu-id="2a398-104">組み込みのスタイルをテキストに適用する</span><span class="sxs-lookup"><span data-stu-id="2a398-104">Apply a built-in style to text</span></span>

1. <span data-ttu-id="2a398-105">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="2a398-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="2a398-106">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="2a398-106">Open the file index.html.</span></span>
3. <span data-ttu-id="2a398-107">`insert-paragraph` ボタンを格納している `div` の直下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="2a398-107">Just below the `div` that contains the `insert-paragraph` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-style">Apply Style</button>            
    </div>
    ```

4. <span data-ttu-id="2a398-108">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="2a398-108">Open the app.js file.</span></span>

5. <span data-ttu-id="2a398-109">`insert-paragraph` ボタンにクリック ハンドラーを割り当てる行の直下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="2a398-109">Just below the line that assigns a click handler to the `insert-paragraph` button, add the following code:</span></span>

    ```js
    $('#apply-style').click(applyStyle);
    ```

6. <span data-ttu-id="2a398-110">`insertParagraph` 関数の直下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="2a398-110">Just below the `insertParagraph` function, add the following function:</span></span>

    ```js
    function applyStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to style text.

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

7. <span data-ttu-id="2a398-111">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="2a398-111">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="2a398-112">このコードではスタイルを段落に適用していますが、スタイルはテキストの範囲にも適用できます。</span><span class="sxs-lookup"><span data-stu-id="2a398-112">Note that the code applies a style to a paragraph, but styles can also be applied to ranges of text.</span></span>

    ```js
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ``` 

## <a name="apply-a-custom-style-to-text"></a><span data-ttu-id="2a398-113">カスタム スタイルをテキストに適用する</span><span class="sxs-lookup"><span data-stu-id="2a398-113">Apply a custom style to text</span></span>

1. <span data-ttu-id="2a398-114">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="2a398-114">Open the file index.html.</span></span>
2. <span data-ttu-id="2a398-115">ボタンを格納している `div` の下に、次のマークアップを追加します。`apply-style`</span><span class="sxs-lookup"><span data-stu-id="2a398-115">Below the `div` that contains the `apply-style` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-custom-style">Apply Custom Style</button>            
    </div>
    ```

3. <span data-ttu-id="2a398-116">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="2a398-116">Open the app.js file.</span></span>

4. <span data-ttu-id="2a398-117">ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。`apply-style`</span><span class="sxs-lookup"><span data-stu-id="2a398-117">Below the line that assigns a click handler to the `apply-style` button, add the following code:</span></span>

    ```js
    $('#apply-custom-style').click(applyCustomStyle);
    ```

5. <span data-ttu-id="2a398-118">関数の下に、次の関数を追加します。`applyStyle`</span><span class="sxs-lookup"><span data-stu-id="2a398-118">Below the `applyStyle` function, add the following function:</span></span>

    ```js
    function applyCustomStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply the custom style.

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

7. <span data-ttu-id="2a398-119">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="2a398-119">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="2a398-120">このコードでは、まだ存在していないカスタム スタイルを適用しています。</span><span class="sxs-lookup"><span data-stu-id="2a398-120">Note that the code applies a custom style that does not exist yet.</span></span> <span data-ttu-id="2a398-121">「[アドインをテストする](#test-the-add-in)」の手順で **MyCustomStyle** という名前のスタイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="2a398-121">You'll create a style with the name **MyCustomStyle** in the [Test the add-in](#test-the-add-in) step.</span></span>

    ```js
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ``` 

## <a name="change-the-font-of-text"></a><span data-ttu-id="2a398-122">テキストのフォントを変更する</span><span class="sxs-lookup"><span data-stu-id="2a398-122">Change the font of text</span></span>

1. <span data-ttu-id="2a398-123">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="2a398-123">Open the file index.html.</span></span>
2. <span data-ttu-id="2a398-124">ボタンを格納している `div` の下に、次のマークアップを追加します。`apply-custom-style`</span><span class="sxs-lookup"><span data-stu-id="2a398-124">Below the `div` that contains the `apply-custom-style` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="change-font">Change Font</button>            
    </div>
    ```

3. <span data-ttu-id="2a398-125">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="2a398-125">Open the app.js file.</span></span>

4. <span data-ttu-id="2a398-126">ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。`apply-custom-style`</span><span class="sxs-lookup"><span data-stu-id="2a398-126">Below the line that assigns a click handler to the `apply-custom-style` button, add the following code:</span></span>

    ```js
    $('#change-font').click(changeFont);
    ```

5. <span data-ttu-id="2a398-127">関数の下に、次の関数を追加します。`applyCustomStyle`</span><span class="sxs-lookup"><span data-stu-id="2a398-127">Below the `applyCustomStyle` function, add the following function:</span></span>

    ```js
    function changeFont() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply a different font.

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

7. <span data-ttu-id="2a398-128">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="2a398-128">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="2a398-129">このコードでは、`Paragraph.getNext` メソッドにチェーンされた `ParagraphCollection.getFirst` メソッドを使用して 2 番目の段落への参照を取得することに注意してください。</span><span class="sxs-lookup"><span data-stu-id="2a398-129">Note that the code gets a reference to the second paragraph by using the `ParagraphCollection.getFirst` method chained to the `Paragraph.getNext` method.</span></span>

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="2a398-130">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="2a398-130">Test the add-in</span></span>

1. <span data-ttu-id="2a398-131">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、Ctrl + C を 2 回入力して実行中の Web サーバーを停止します。</span><span class="sxs-lookup"><span data-stu-id="2a398-131">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="2a398-132">それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="2a398-132">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="2a398-133">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。</span><span class="sxs-lookup"><span data-stu-id="2a398-133">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="2a398-134">これを行うには、プロンプトが表示されてビルド コマンドを入力できるようにするため、サーバー プロセスを強制終了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2a398-134">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="2a398-135">ビルド後に、サーバーを再起動します。</span><span class="sxs-lookup"><span data-stu-id="2a398-135">After the build, you restart the server.</span></span> <span data-ttu-id="2a398-136">次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="2a398-136">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="2a398-137">コマンドを実行し、Office アドインを実行できるすべてのホストでサポートされている以前のバージョンの JavaScript に ES6 ソース コードをトランスパイルします。`npm run build`</span><span class="sxs-lookup"><span data-stu-id="2a398-137">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="2a398-138">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="2a398-138">Run the command `npm start` to start a web server running on localhost.</span></span>   
4. <span data-ttu-id="2a398-139">作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="2a398-139">Reload the task pane by closing it, and then on the **Home** menu select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="2a398-140">ドキュメントに 3 つ以上の段落があることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="2a398-140">Be sure there are at least three paragraphs in the document.</span></span> <span data-ttu-id="2a398-141">**[段落の挿入]** を 3 回選択できます。</span><span class="sxs-lookup"><span data-stu-id="2a398-141">You can choose **Insert Paragraph** three times.</span></span> <span data-ttu-id="2a398-142">*ドキュメントの最後に空白の段落がないことを慎重にチェックしてください。空白の段落がある場合は、それを削除します。*</span><span class="sxs-lookup"><span data-stu-id="2a398-142">*Check carefully that there's no blank paragraph at the end of the document. If there is, delete it.*</span></span>
6. <span data-ttu-id="2a398-143">Word で、MyCustomStyle という名前のカスタム スタイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="2a398-143">In Word, create a custom style named "MyCustomStyle".</span></span> <span data-ttu-id="2a398-144">このスタイルには、必要に応じて任意の書式を設定できます。</span><span class="sxs-lookup"><span data-stu-id="2a398-144">It can have any formatting that you want.</span></span>
7. <span data-ttu-id="2a398-145">**[スタイルの適用]** ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="2a398-145">Choose the **Apply Style** button.</span></span> <span data-ttu-id="2a398-146">最初の段落は、組み込みのスタイルである **Intense Reference** でスタイル設定されます。</span><span class="sxs-lookup"><span data-stu-id="2a398-146">The first paragraph will be styled with the built-in style **Intense Reference**.</span></span>
8. <span data-ttu-id="2a398-147">**[カスタム スタイルの適用]** ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="2a398-147">Choose the **Apply Custom Style** button.</span></span> <span data-ttu-id="2a398-148">最後の段落は、選択したカスタム スタイルでスタイル設定されます。</span><span class="sxs-lookup"><span data-stu-id="2a398-148">The last paragraph will be styled with your custom style.</span></span> <span data-ttu-id="2a398-149">(何も起こらないように思える場合、最後の段落が空白である可能性があります。</span><span class="sxs-lookup"><span data-stu-id="2a398-149">(If nothing seems to happen, the last paragraph might be blank.</span></span> <span data-ttu-id="2a398-150">その場合は、最後の段落にテキストを追加します)。</span><span class="sxs-lookup"><span data-stu-id="2a398-150">If so, add some text to it.)</span></span>
9. <span data-ttu-id="2a398-151">**[フォントの変更]** ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="2a398-151">Choose the **Change Font** button.</span></span> <span data-ttu-id="2a398-152">2 番目の段落のフォントを、18 ポイントで太字の Courier New に変更します。</span><span class="sxs-lookup"><span data-stu-id="2a398-152">The font of the second paragraph changes to 18 pt., bold, Courier New.</span></span>

    ![Word のチュートリアル - スタイルとフォントを適用する](../images/word-tutorial-apply-styles-and-font.png)
