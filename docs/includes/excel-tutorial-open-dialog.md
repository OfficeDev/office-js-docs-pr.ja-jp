<span data-ttu-id="b77a9-101">このチュートリアルの最後の手順では、アドインでダイアログを開いて、ダイアログのプロセスから作業ウィンドウのプロセスにメッセージを渡して、ダイアログを閉じます。</span><span class="sxs-lookup"><span data-stu-id="b77a9-101">In this final step of the tutorial, you'll open a dialog in your add-in, pass a message from the dialog process to the task pane process, and close the dialog.</span></span> <span data-ttu-id="b77a9-102">Office アドインのダイアログは、*モードレス*です。ユーザーは、ホスト Office アプリケーション内のドキュメントと作業ウィンドウ内のホスト ページの両方の操作を続行できます。</span><span class="sxs-lookup"><span data-stu-id="b77a9-102">Office Add-in dialogs are *nonmodal*: a user can continue to interact with both the document in the host Office application and with the host page in the task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="b77a9-103">このページでは、Excel のアドインのチュートリアルの個々 の手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="b77a9-103">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="b77a9-104">このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[Excel アドインのチュートリアル](../tutorials/excel-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。</span><span class="sxs-lookup"><span data-stu-id="b77a9-104">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="create-the-dialog-page"></a><span data-ttu-id="b77a9-105">ダイアログ ページを作成する</span><span class="sxs-lookup"><span data-stu-id="b77a9-105">Create the dialog page</span></span>

1. <span data-ttu-id="b77a9-106">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="b77a9-106">Open the project in your code editor.</span></span>
2. <span data-ttu-id="b77a9-107">プロジェクトのルート (index.html がある場所) で、popup.html というファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="b77a9-107">Create a file in the root of the project (where index.html is) called popup.html.</span></span>
3. <span data-ttu-id="b77a9-p103">popup.html に、次のコードを追加します。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="b77a9-p103">Add the following markup to popup.html. Note:</span></span>
   - <span data-ttu-id="b77a9-110">このページには、ユーザーが自分の名前を入力する `<input>` と、その名前を作業ウィンドウ内のページ（入力した名前が表示されるページ）に送信するボタンが含まれています。</span><span class="sxs-lookup"><span data-stu-id="b77a9-110">The page has a `<input>` where the user will enter his or her name and a button that will send the name to the page in the task pane where it will be displayed.</span></span>
   - <span data-ttu-id="b77a9-111">このマークアップでは、popup.js というスクリプトを読み込みます。このスクリプトは、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="b77a9-111">The markup loads a script called popup.js that you will create in a later step.</span></span>
   - <span data-ttu-id="b77a9-112">また、popup.js で使用することになる Office.JS ライブラリと jQuery も読み込みます。</span><span class="sxs-lookup"><span data-stu-id="b77a9-112">It also loads the Office.JS library and jQuery because they will be used in popup.js.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head lang="en">
            <title>Dialog for My Office Add-in</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">
        
            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.min.css" />
            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.components.css" />
            <link rel="stylesheet" href="app.css">
    
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
        <div>
        <div class="padding">
            <button id="ok-button" class="ms-Button">OK</button>
        </div>
    </body>
    </html>
    ```

4. <span data-ttu-id="b77a9-113">プロジェクトのルートに popup.js というファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="b77a9-113">Create a file in the root of the project called popup.js.</span></span>
5. <span data-ttu-id="b77a9-p104">popup.js に、次のコードを追加します。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="b77a9-p104">Add the following code to popup.js. Note:</span></span>
   - <span data-ttu-id="b77a9-116">*Office.JS 内の API を呼び出すページは、どのページでも `Office.initialize` プロパティに関数を割り当てる必要があります。*</span><span class="sxs-lookup"><span data-stu-id="b77a9-116">*Every page that calls APIs in the Office.JS library must assign a function to the `Office.initialize` property.*</span></span> <span data-ttu-id="b77a9-117">初期化が不要な場合は、関数の本体を空にすることができますが、プロパティを未定義のままにすることや、Null または関数以外の値を割り当てることはできません。</span><span class="sxs-lookup"><span data-stu-id="b77a9-117">If no initialization is needed, then the function can have an empty body, but the property must not be left undefined, assigned to null or to a non-function value.</span></span> <span data-ttu-id="b77a9-118">たとえば、プロジェクト ルートにある app.js ファイルを確認してください。</span><span class="sxs-lookup"><span data-stu-id="b77a9-118">For an example, see the app.js file in the project root.</span></span> <span data-ttu-id="b77a9-119">この割り当てを実施するコードは、Office.JS を呼び出す前に実行する必要があります。そのため、この例で示すように、割り当てはページによって読み込まれるスクリプト ファイル内に入れてあります。</span><span class="sxs-lookup"><span data-stu-id="b77a9-119">The code that makes the assignment must run before any calls to Office.JS; hence the assignment is in a script file that is loaded by the page, as it is in this case.</span></span>
   - <span data-ttu-id="b77a9-p106">jQuery の `ready` 関数は、`initialize` メソッド内から呼び出します。別の JavaScript ライブラリのコードの読み込み、初期化、またはブートストラップを `Office.initialize` 関数内に入れることは、ほとんどすべての場合に通用するルールです。</span><span class="sxs-lookup"><span data-stu-id="b77a9-p106">The jQuery `ready` function is called inside the `initialize` method. It is an almost universal rule that the loading, initializing, or bootstrapping code of other JavaScript libraries should be inside the `Office.initialize` function.</span></span>

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

6. <span data-ttu-id="b77a9-122">を次のコードに置き換えます。`TODO1`</span><span class="sxs-lookup"><span data-stu-id="b77a9-122">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="b77a9-123">関数は、この後の手順で作成します。`sendStringToParentPage`</span><span class="sxs-lookup"><span data-stu-id="b77a9-123">You'll create the `sendStringToParentPage` function in the next step.</span></span>

    ```js
    $('#ok-button').click(sendStringToParentPage);
    ```

7. <span data-ttu-id="b77a9-124">を次のコードに置き換えます。`TODO2`</span><span class="sxs-lookup"><span data-stu-id="b77a9-124">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="b77a9-125">メソッドは、パラメーターを親ページ (この例では、作業ウィンドウ内のページ) に渡します。`messageParent`</span><span class="sxs-lookup"><span data-stu-id="b77a9-125">The `messageParent` method passes its parameter to the parent page, in this case, the page in the task pane.</span></span> <span data-ttu-id="b77a9-126">パラメーターには、ブール値または文字列を使用できます (XML や JSON など、文字列としてシリアル化できるすべてのものが含まれます)。</span><span class="sxs-lookup"><span data-stu-id="b77a9-126">The parameter can be a boolean or a string, which includes anything that can be serialized as a string, such as XML or JSON.</span></span> 

    ```js
    function sendStringToParentPage() {
        var userName = $('#name-box').val();
        Office.context.ui.messageParent(userName);
    }
    ```

8. <span data-ttu-id="b77a9-127">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="b77a9-127">Save the file.</span></span>

   > [!NOTE]
   > <span data-ttu-id="b77a9-128">popup.html ファイルと、そのファイルで読み込む popup.js ファイルは、アドインの作業ウィンドウとは完全に別な Internet Explorer プロセスで実行されます。</span><span class="sxs-lookup"><span data-stu-id="b77a9-128">The popup.html file, and the popup.js file that it loads, run in an entirely separate Internet Explorer process from the add-in's task pane.</span></span> <span data-ttu-id="b77a9-129">popup.js が app.js ファイルと同じ bundle.js ファイルからトランスパイルされていた場合、アドインでは bundle.js の 2 つのコピーを読み込むことが必要になり、バンドル化の意味がなくなります。</span><span class="sxs-lookup"><span data-stu-id="b77a9-129">If the popup.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="b77a9-130">さらに、popup.js ファイルには IE で未サポートの JavaScript は含まれていません。</span><span class="sxs-lookup"><span data-stu-id="b77a9-130">In addition, the popup.js file does not contain any JavaScript that is unsupported by IE.</span></span> <span data-ttu-id="b77a9-131">これら 2 つの理由から、このアドインでは popup.js を一切トランスパイルしていません。</span><span class="sxs-lookup"><span data-stu-id="b77a9-131">For these two reasons, this add-in does not transpile the popup.js file at all.</span></span> 


## <a name="open-the-dialog-from-the-task-pane"></a><span data-ttu-id="b77a9-132">作業ウィンドウからダイアログを開く</span><span class="sxs-lookup"><span data-stu-id="b77a9-132">Open the dialog from the task pane</span></span>

1. <span data-ttu-id="b77a9-133">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="b77a9-133">Open the file index.html.</span></span>
2. <span data-ttu-id="b77a9-134">ボタンを格納している `div` の下に、次のマークアップを追加します。`freeze-header`</span><span class="sxs-lookup"><span data-stu-id="b77a9-134">Below the `div` that contains the `freeze-header` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="open-dialog">Open Dialog</button>          
    </div>
    ```

3. <span data-ttu-id="b77a9-135">このダイアログでは、ユーザーに名前の入力を求めて、ユーザーの名前を作業ウィンドウに渡します。</span><span class="sxs-lookup"><span data-stu-id="b77a9-135">The dialog will prompt the user to enter a name and pass the user's name to the task pane.</span></span> <span data-ttu-id="b77a9-136">作業ウィンドウでは、それがラベルに表示されます。</span><span class="sxs-lookup"><span data-stu-id="b77a9-136">The task pane will display it in a label.</span></span> <span data-ttu-id="b77a9-137">前の手順で追加した `div` のすぐ下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="b77a9-137">Immediately below the `div` that you just added, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <label id="user-name"></label>            
    </div>
    ```

4. <span data-ttu-id="b77a9-138">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="b77a9-138">Open the app.js file.</span></span>

5. <span data-ttu-id="b77a9-139">ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。`freeze-header`</span><span class="sxs-lookup"><span data-stu-id="b77a9-139">Below the line that assigns a click handler to the `freeze-header` button, add the following code.</span></span> <span data-ttu-id="b77a9-140">メソッドは、この後の手順で作成します。`openDialog`</span><span class="sxs-lookup"><span data-stu-id="b77a9-140">You'll create the `openDialog` method in a later step.</span></span>

    ```js
    $('#open-dialog').click(openDialog);
    ```

6. <span data-ttu-id="b77a9-p112">関数の下に、次の宣言を追加します。この変数は、親ページの実行コンテキスト内のオブジェクトを保持するために使用され、ダイアログ ページの実行コンテキストへの仲介者として機能します。`freezeHeader`</span><span class="sxs-lookup"><span data-stu-id="b77a9-p112">Below the `freezeHeader` function add the following declaration. This variable is used to hold an object in the parent page's execution context that acts as an intermediator to the dialog page's execution context.</span></span>

    ```js
    let dialog = null;
    ```

7. <span data-ttu-id="b77a9-143">の宣言の下に、次の関数を追加します。`dialog`</span><span class="sxs-lookup"><span data-stu-id="b77a9-143">Below the declaration of `dialog`, add the following function.</span></span> <span data-ttu-id="b77a9-144">このコードで注目する重要な点は、そこに `Excel.run` の呼び出しが存在*しない*ことです。</span><span class="sxs-lookup"><span data-stu-id="b77a9-144">The important thing to notice about this code is what is *not* there: there is no call of `Excel.run`.</span></span> <span data-ttu-id="b77a9-145">これは、ダイアログを開く API はすべての Office ホストで共有されるため、Excel 固有の API ではなく Office JavaScript 共通 API に含まれているからです。</span><span class="sxs-lookup"><span data-stu-id="b77a9-145">This is because the API to open a dialog is shared among all Office hosts, so it is part of the Office JavaScript Common API, not the Excel-specific API.</span></span>

    ```js
    function openDialog() {
        // TODO1: Call the Office Shared API that opens a dialog
    }
    ``` 

8. <span data-ttu-id="b77a9-p114">|||UNTRANSLATED_CONTENT_START|||Replace `TODO1` with the following code. Note:|||UNTRANSLATED_CONTENT_END|||</span><span class="sxs-lookup"><span data-stu-id="b77a9-p114">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="b77a9-148">メソッドでは、画面の中央にダイアログを開きます。`displayDialogAsync`</span><span class="sxs-lookup"><span data-stu-id="b77a9-148">The `displayDialogAsync` method opens a dialog in the center of the screen.</span></span>
   - <span data-ttu-id="b77a9-149">最初のパラメーターは、開くページの URL です。</span><span class="sxs-lookup"><span data-stu-id="b77a9-149">The first parameter is the URL of the page to open.</span></span>
   - <span data-ttu-id="b77a9-p115">2 番目のパラメーターでオプションを渡します。`height` と `width` は、Office アプリケーションのウィンドウ サイズの比率です。</span><span class="sxs-lookup"><span data-stu-id="b77a9-p115">The second parameter passes options. `height` and `width` are percentages of the size of the Office application's window.</span></span> 
   
    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},
        
        // TODO2: Add callback parameter.
    );
    ``` 

## <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a><span data-ttu-id="b77a9-152">ダイアログからのメッセージを処理してダイアログを閉じる</span><span class="sxs-lookup"><span data-stu-id="b77a9-152">Process the message from the dialog and close the dialog</span></span>

1. <span data-ttu-id="b77a9-p116">app.js ファイルでの作業を続けます。`TODO2` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="b77a9-p116">Continue in the app.js file, and replace `TODO2` with the following code. Note:</span></span>
   - <span data-ttu-id="b77a9-155">コールバックは、ダイアログが正常に開いた直後、ユーザーがダイアログで操作を行う前に実行されます。</span><span class="sxs-lookup"><span data-stu-id="b77a9-155">The callback is executed immediately after the dialog successfully opens and before the user has taken any action in the dialog.</span></span>
   - <span data-ttu-id="b77a9-156">は、親ページとダイアログ ページの実行コンテキストの間で仲介者のように機能するオブジェクトです。`result.value`</span><span class="sxs-lookup"><span data-stu-id="b77a9-156">The `result.value` is the object that acts as a kind of middleman between the execution contexts of the parent and dialog pages.</span></span>
   - <span data-ttu-id="b77a9-157">関数は、この後の手順で作成します。`processMessage`</span><span class="sxs-lookup"><span data-stu-id="b77a9-157">The `processMessage` function will be created in a later step.</span></span> <span data-ttu-id="b77a9-158">このハンドラーは、`messageParent` 関数の呼び出しによって、ダイアログから送信されるあらゆる値を処理します。</span><span class="sxs-lookup"><span data-stu-id="b77a9-158">This handler will process any values that are sent from the dialog page with calls of the `messageParent` function.</span></span>

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. <span data-ttu-id="b77a9-159">関数の下に、次の関数を追加します。`openDialog`</span><span class="sxs-lookup"><span data-stu-id="b77a9-159">Below the `openDialog` function, add the following function.</span></span>

    ```js
    function processMessage(arg) {
        $('#user-name').text(arg.message);
        dialog.close();
    }
    ```

## <a name="test-the-add-in"></a><span data-ttu-id="b77a9-160">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="b77a9-160">Test the add-in</span></span>

1. <span data-ttu-id="b77a9-161">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、Ctrl-C を 2 回入力して実行中の Web サーバーを停止します。</span><span class="sxs-lookup"><span data-stu-id="b77a9-161">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="b77a9-162">それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="b77a9-162">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="b77a9-163">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。</span><span class="sxs-lookup"><span data-stu-id="b77a9-163">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="b77a9-164">そのためには、ビルド コマンドの入力を求めるプロンプトが表示されるように、サーバー プロセスを強制終了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b77a9-164">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="b77a9-165">ビルド後に、サーバーを再起動します。</span><span class="sxs-lookup"><span data-stu-id="b77a9-165">After the build, you restart the server.</span></span> <span data-ttu-id="b77a9-166">次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="b77a9-166">The next few steps carry out this process.</span></span>

1. <span data-ttu-id="b77a9-167">コマンドを実行して、ES6 ソース コードを Internet Explorer でサポートされている以前のバージョンの JavaScript にトランスパイルします (これは、Excel アドインを実行するために Excel の内部で使用されます)。`npm run build`</span><span class="sxs-lookup"><span data-stu-id="b77a9-167">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
2. <span data-ttu-id="b77a9-168">コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。`npm start`</span><span class="sxs-lookup"><span data-stu-id="b77a9-168">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="b77a9-169">作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="b77a9-169">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
6. <span data-ttu-id="b77a9-170">作業ウィンドウで、**[Open Dialog]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="b77a9-170">Choose the **Open Dialog** button in the task pane.</span></span> 
7. <span data-ttu-id="b77a9-171">ダイアログが開いたら、ドラッグしたりサイズ変更したりします。</span><span class="sxs-lookup"><span data-stu-id="b77a9-171">While the dialog is open, drag it and resize it.</span></span> <span data-ttu-id="b77a9-172">ワークシートの操作と作業ウィンドウの別のボタンのクリックができるようになっています。</span><span class="sxs-lookup"><span data-stu-id="b77a9-172">Note that you can interact with the worksheet and press other buttons on the taskpane.</span></span> <span data-ttu-id="b77a9-173">ただし、同じ作業ウィンドウ ページから 2 番目のダイアログを起動することはできません。</span><span class="sxs-lookup"><span data-stu-id="b77a9-173">But you cannot launch a second dialog from the same task pane page.</span></span>
8. <span data-ttu-id="b77a9-174">ダイアログで、名前を入力して **[OK]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="b77a9-174">In the dialog, enter a name and choose **OK**.</span></span> <span data-ttu-id="b77a9-175">作業ウィンドウに名前が表示され、ダイアログが閉じられます。</span><span class="sxs-lookup"><span data-stu-id="b77a9-175">The name appears on the task pane and the dialog closes.</span></span>
9. <span data-ttu-id="b77a9-176">オプションとして、`processMessage` 関数内の行 `dialog.close();` をコメントにします。</span><span class="sxs-lookup"><span data-stu-id="b77a9-176">Optionally, comment out the line `dialog.close();` in the `processMessage` function.</span></span> <span data-ttu-id="b77a9-177">その後で、このセクションの手順を繰り返します。</span><span class="sxs-lookup"><span data-stu-id="b77a9-177">Then repeat the steps of this section.</span></span> <span data-ttu-id="b77a9-178">ダイアログを開いたまま名前を変更できます。</span><span class="sxs-lookup"><span data-stu-id="b77a9-178">The dialog stays open and you can change the name.</span></span> <span data-ttu-id="b77a9-179">右上の **[X]** ボタンをクリックすることで、手動で閉じることができます。</span><span class="sxs-lookup"><span data-stu-id="b77a9-179">You can close it manually by pressing the **X** button in the upper right corner.</span></span>

    ![Excel チュートリアル: ダイアログ](../images/excel-tutorial-dialog-open.png)

