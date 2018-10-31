# <a name="visio-javascript-api-overview"></a><span data-ttu-id="f5417-101">Visio JavaScript API の概要</span><span class="sxs-lookup"><span data-stu-id="f5417-101">Word-specific JavaScript API overview</span></span>

<span data-ttu-id="f5417-102">Visio JavaScript API を使ってSharePoint オンライン で Visio のダイアグラムを埋め込むことができます。</span><span class="sxs-lookup"><span data-stu-id="f5417-102">You can use the Visio JavaScript APIs to embed Visio diagrams in SharePoint Online.</span></span> <span data-ttu-id="f5417-103">埋め込まれた Visio のダイアグラムは、SharePoint ドキュメント ライブラリに保存され、SharePoint ページに表示されます。</span><span class="sxs-lookup"><span data-stu-id="f5417-103">An embedded Visio diagram is a diagram that is stored in a SharePoint document library and displayed on a SharePoint page.</span></span> <span data-ttu-id="f5417-104">Visio のダイアグラムを HTML `<iframe>` 要素に表示して埋め込みます。</span><span class="sxs-lookup"><span data-stu-id="f5417-104">To embed a Visio diagram, display it in an HTML `<iframe>`iframe element.</span></span> <span data-ttu-id="f5417-105">そうすると、Visio JavaScript API を使用して、プログラムで埋め込まれたダイアグラムを使った作業ができるようになります。</span><span class="sxs-lookup"><span data-stu-id="f5417-105">Then you can use Visio JavaScript APIs to programmatically work with the embedded diagram.</span></span>

![SharePoint ページの iframe 上にある Visio のダイアグラムとスクリプト エディター Web パーツ](../images/visio-api-block-diagram.png)


<span data-ttu-id="f5417-107">Visio JavaScript API を使用して次のことができるようになります。</span><span class="sxs-lookup"><span data-stu-id="f5417-107">You can use the Visio JavaScript APIs to:</span></span>

* <span data-ttu-id="f5417-108">ページや図形などの Visio ダイアグラムの要素を使って操作する。</span><span class="sxs-lookup"><span data-stu-id="f5417-108">Interact with Visio diagram elements like pages and shapes</span></span>
* <span data-ttu-id="f5417-109">Visioダイアグラムのキャンバスにビジュアル マークアップを作成する。</span><span class="sxs-lookup"><span data-stu-id="f5417-109">Create visual markup on the Visio diagram canvas</span></span>
* <span data-ttu-id="f5417-110">図面の中でのマウス イベント用のカスタム ハンドラーを記述する。</span><span class="sxs-lookup"><span data-stu-id="f5417-110">Write custom handlers for mouse events within the drawing</span></span>
* <span data-ttu-id="f5417-111">図形テキスト、図形データ、およびハイパーリンクなどのダイアグラムのデータをソリューションに公開する。</span><span class="sxs-lookup"><span data-stu-id="f5417-111">Expose diagram data, such as shape text, shape data, and hyperlinks, to your solution.</span></span>

<span data-ttu-id="f5417-p102">ここでは、Visio オンラインで Visio JavaScript API を使って SharePoint オンライン のソリューションをビルドする方法について説明します。また、 **EmbeddedSession**、 **RequestContext**、JavaScript プロキシ オブジェクトなどの API、および **sync()**、 **Visio.run()**、 **load()** のメソッドを使用するための根本的な概念について紹介します。コード例により、これらの概念を適用する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="f5417-p102">This article describes how to use the Visio JavaScript APIs with Visio Online to build your solutions for SharePoint Online. It introduces key concepts that are fundamental to using the APIs, such as **EmbeddedSession**, **RequestContext**, and JavaScript proxy objects, and the **sync()**, **Visio.run()**, and **load()** methods. The code examples show you how to apply these concepts.</span></span>

## <a name="embeddedsession"></a><span data-ttu-id="f5417-115">EmbeddedSession</span><span class="sxs-lookup"><span data-stu-id="f5417-115">EmbeddedSession</span></span>

<span data-ttu-id="f5417-116">EmbeddedSession オブジェクトは、開発者のフレームと Visio Online のフレームの間の通信を初期化します。</span><span class="sxs-lookup"><span data-stu-id="f5417-116">The EmbeddedSession object initializes communication between the developer frame and the Visio Online frame.</span></span>

```js
var session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
session.init().then(function () {
    window.console.log("Session successfully initialized");
});
```

## <a name="visiorunsession-functioncontext--batch-"></a><span data-ttu-id="f5417-117">Visio.run(session, function(context) { batch })</span><span class="sxs-lookup"><span data-stu-id="f5417-117">Visio.run(session, function(context) { batch })</span></span>

<span data-ttu-id="f5417-118">**Visio.run()** は、Visio オブジェクト モデルに対して作用するバッチ スクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="f5417-118">**Visio.run()** executes a batch script that performs actions on the Visio object model.</span></span> <span data-ttu-id="f5417-119">バッチ コマンドには、JavaScript のローカル プロキシ オブジェクトの定義と、ローカル オブジェクトと Visio オブジェクトの間で状態を同期し、promise　レゾリューションの **sync()** メソッドが含まれます。</span><span class="sxs-lookup"><span data-stu-id="f5417-119">The batch commands include definitions of local JavaScript proxy objects and **sync()** methods that synchronize the state between local and Visio objects and promise resolution.</span></span> <span data-ttu-id="f5417-120"> *\*Visio.run()** で要求をバッチ処理する利点は、promiseが返されるときに、実行中に割り当てられた追跡ページ オブジェクトが自動的に解放されることです。</span><span class="sxs-lookup"><span data-stu-id="f5417-120">The advantage of batching requests in **Visio.run()** is that when the promise is resolved, any tracked page objects that were allocated during the execution will be automatically released.</span></span>

<span data-ttu-id="f5417-121">run メソッドはセッションと RequestContext オブジェクトを取り込み、promise（通常は**context.sync()** の結果)を返します。</span><span class="sxs-lookup"><span data-stu-id="f5417-121">The run method takes in RequestContext and returns a promise (typically, just the result of **ctx.sync()**).</span></span> <span data-ttu-id="f5417-122">バッチ操作は **Visio.run()** の外部で実行することができます。</span><span class="sxs-lookup"><span data-stu-id="f5417-122">It is possible to run the batch operation outside of the **Visio.run()**.</span></span> <span data-ttu-id="f5417-123">ただし、この場合、ページ オブジェクトの参照は、手動で追跡および管理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f5417-123">However, in such a scenario, any page object references needs to be manually tracked and managed.</span></span>

## <a name="requestcontext"></a><span data-ttu-id="f5417-124">RequestContext</span><span class="sxs-lookup"><span data-stu-id="f5417-124">RequestContext</span></span>

<span data-ttu-id="f5417-125">RequestContext オブジェクトは、Visio アプリケーションへのリクエストを簡単にするものです。</span><span class="sxs-lookup"><span data-stu-id="f5417-125">Request Context: The RequestContext object facilitates requests to the Excel application.</span></span> <span data-ttu-id="f5417-126">開発者のフレームと Visio Online アプリケーションは、異なる 2 つの iframe で実行されるため、開発者のフレームから Visio およびページや図形などの関連するオブジェクトへのアクセスを取得する RequestContext オブジェクト (次の例の内容を含む) が必要です。</span><span class="sxs-lookup"><span data-stu-id="f5417-126">The RequestContext object facilitates requests to the Visio application. Because the developer frame and the Visio Online application run in two different iframes, request context is required to get access to Visio and related objects such as pages and shapes, from the developer frame. The following example shows how to create a request context.</span></span>

```js
function hideToolbars() {
    Visio.run(session, function(context){
        var app = context.document.application;
        app.showToolbars = false;
        return context.sync().then(function () {
            window.console.log("Toolbars Hidden");
        });
    }).catch(function(error)
    {
        window.console.log("Error: " + error);
    });
};
```

## <a name="proxy-objects"></a><span data-ttu-id="f5417-127">プロキシ オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f5417-127">Proxy objects</span></span>

<span data-ttu-id="f5417-p106">アドインで申告され使用される Visio の JavaScript オブジェクトは、Visio ドキュメント内の実際のオブジェクトのためのプロキシ オブジェクトになります。プロキシ オブジェクトで実行されたすべてのアクションは、Visio では認識されません。また、Visio ドキュメントの状態は、ドキュメントの状態が同期されるまでプロキシ オブジェクトで認識されません。ドキュメントの状態は、 `context.sync()` の実行時に同期されます。</span><span class="sxs-lookup"><span data-stu-id="f5417-p106">The Visio JavaScript objects declared and used in an add-in are proxy objects for the real objects in a Visio document. All actions taken on proxy objects are not realized in Visio, and the state of the Visio document is not realized in the proxy objects until the document state has been synchronized. The document state is synchronized when `context.sync()` is run.</span></span>

<span data-ttu-id="f5417-131">たとえば、ローカルの JavaScript オブジェクトの getActivePage は、選択したページを参照するよう表示されます。</span><span class="sxs-lookup"><span data-stu-id="f5417-131">For example, the local JavaScript object  is declared to reference the selected range.</span></span> <span data-ttu-id="f5417-132">これは、オブジェクトのプロパティの設定およびメソッドの呼び出しをキューに入れるために使用します。</span><span class="sxs-lookup"><span data-stu-id="f5417-132">This can be used to queue the setting of its properties and invoking methods.</span></span> <span data-ttu-id="f5417-133"> *\*sync()** メソッドが実行されるまで、これらのオブジェクトのアクションは認識されません。</span><span class="sxs-lookup"><span data-stu-id="f5417-133">The actions on such objects are not realized until the sync() method is run.</span></span>

```js
var activePage = context.document.getActivePage();
```

## <a name="sync"></a><span data-ttu-id="f5417-134">sync()</span><span class="sxs-lookup"><span data-stu-id="f5417-134">sync()</span></span>

<span data-ttu-id="f5417-135"> *\*sync()** メソッドは、Visio内のJavaScript のプロキシ オブジェクトと 実際のオブジェクトの間で状態を同期させます。これは、コンテキストでキューに入れられた指示の実行と、ユーザーのコードで使用するために読み込まれた Office オブジェクトのプロパティを検索することで同期させます。</span><span class="sxs-lookup"><span data-stu-id="f5417-135">The **sync()** method, available on the request context, synchronizes the state between JavaScript proxy objects and real objects in Visio by executing instructions queued on the context and retrieving properties of loaded Office objects for use in your code.</span></span> <span data-ttu-id="f5417-136">このメソッドは、同期処理が完了したときに解決されるpromiseを返します。</span><span class="sxs-lookup"><span data-stu-id="f5417-136">This method returns a promise, which is resolved when synchronization is complete.</span></span> 

## <a name="load"></a><span data-ttu-id="f5417-137">load()</span><span class="sxs-lookup"><span data-stu-id="f5417-137">load()</span></span>

<span data-ttu-id="f5417-p109"> *\*load()** メソッドは、アドインの JavaScript レイヤーで作成されたプロキシ オブジェクトに埋めるために使用します。ドキュメントなどのオブジェクトを検索する場合、まず JavaScript レイヤーでローカル プロキシ オブジェクトが作成されます。このようなオブジェクトは、そのプロパティの設定とメソッドの呼び出しをキューに登録するために使用されます。ただし、オブジェクトのプロパティや関係を読み取りには、最初に *\*load()** メソッドと*\*sync()** メソッドを呼び出す必要があります。 load() メソッドは、 *\*sync()** メソッドが呼び出されたときに読み込まれる必要があるプロパティと関係を取り込みます。</span><span class="sxs-lookup"><span data-stu-id="f5417-p109">The **load()** method is used to fill in the proxy objects created in the add-in JavaScript layer. When trying to retrieve an object such as a document, a local proxy object is created first in the JavaScript layer. Such an object can be used to queue the setting of its properties and invoking methods. However, for reading object properties or relations, the **load()** and **sync()** methods need to be invoked first. The load() method takes in the properties and relations that need to be loaded when the **sync()** method is called.</span></span>

<span data-ttu-id="f5417-143"> *\*load()** メソッドの構文を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="f5417-143">The following shows the syntax for the **load()** method.</span></span>

```js
object.load(string: properties); //or object.load(array: properties); //or object.load({loadOption});
```

1. <span data-ttu-id="f5417-144">**プロパティ** は、読み込まれるプロパティ名の一覧で、コンマ区切りの文字列または名前の配列として指定されます。</span><span class="sxs-lookup"><span data-stu-id="f5417-144">**properties** is the list of properties and/or relationship names to be loaded, specified as comma-delimited strings or array of names.</span></span> <span data-ttu-id="f5417-145">詳細については、各オブジェクトの下の **.load()** メソッドを参照してください。</span><span class="sxs-lookup"><span data-stu-id="f5417-145">See **.load()** methods under each object for details.</span></span>

2. <span data-ttu-id="f5417-p111">**loadOption** は、選択、拡大、トップ、スキップ の各オプションについて説明するオブジェクトを指定します。詳細については、オブジェクトの読み込みの[オプション](/javascript/api/office/officeextension.loadoption)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f5417-p111">**loadOption** specifies an object that describes the selection, expansion, top, and skip options. See object load [options](/javascript/api/office/officeextension.loadoption) for details.</span></span>

## <a name="example-printing-all-shapes-text-in-active-page"></a><span data-ttu-id="f5417-148">例: アクティブ ページですべての図形テキストを印刷する</span><span class="sxs-lookup"><span data-stu-id="f5417-148">Example: Printing all shapes text in active page</span></span>

<span data-ttu-id="f5417-149">次の例では、図形の配列オブジェクトから図形テキストの値を印刷する方法を説明します。</span><span class="sxs-lookup"><span data-stu-id="f5417-149">The following example shows you how to print shape text value from an array shapes object.</span></span>
<span data-ttu-id="f5417-150"> *\*Visio.run()** メソッドには、指示のバッチが含まれています。</span><span class="sxs-lookup"><span data-stu-id="f5417-150">The **Visio.run()** method contains a batch of instructions.</span></span> <span data-ttu-id="f5417-151">このバッチの一部として、作業中のドキュメントの図形を参照するプロキシ オブジェクトが作成されます。</span><span class="sxs-lookup"><span data-stu-id="f5417-151">As part of this batch, a proxy object is created that references shapes on the active document.</span></span>

<span data-ttu-id="f5417-152">すべてのコマンドはキューに登録され、 **context.sync()** が呼び出されたときに実行されます。</span><span class="sxs-lookup"><span data-stu-id="f5417-152">All these commands are queued and run when **ctx.sync()** is called.</span></span> <span data-ttu-id="f5417-153"> *\*sync()** メソッドが返すpromiseは、このメソッドを他の操作とリンクするために使用することができます。</span><span class="sxs-lookup"><span data-stu-id="f5417-153">The **sync()** method returns a promise that can be used to chain it with other operations.</span></span>

```js
Visio.run(session, function (context) {
    var page = context.document.getActivePage();
    var shapes = page.shapes;
    shapes.load();
    return context.sync().then(function () {
        for(var i=0; i<shapes.items.length;i++) {
            var shape = shapes.items[i];
            window.console.log("Shape Text: " + shape.text );
        }
    });
}).catch(function(error) {
    window.console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        window.console.log ("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="error-messages"></a><span data-ttu-id="f5417-154">エラー メッセージ</span><span class="sxs-lookup"><span data-stu-id="f5417-154">Error messages</span></span>

<span data-ttu-id="f5417-p114">エラーは、コードとメッセージで構成される エラー オブジェクトを使用して返されます。次の表は、発生する可能性があるエラー状態の一覧を示しています。</span><span class="sxs-lookup"><span data-stu-id="f5417-p114">Errors are returned using an error object that consists of a code and a message. The following table provides a list of possible error conditions that can occur.</span></span>

| <span data-ttu-id="f5417-157">error.code</span><span class="sxs-lookup"><span data-stu-id="f5417-157">error.code</span></span>            | <span data-ttu-id="f5417-158">error.message</span><span class="sxs-lookup"><span data-stu-id="f5417-158">error.message</span></span> |
|-----------------------|----------------------------------------------------------------|
| <span data-ttu-id="f5417-159">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="f5417-159">InvalidArgument</span></span>       | <span data-ttu-id="f5417-160">引数が無効であるか、存在しません。または形式が正しくありません。</span><span class="sxs-lookup"><span data-stu-id="f5417-160">The argument is invalid or missing or has an incorrect format.</span></span> |
| <span data-ttu-id="f5417-161">GeneralException</span><span class="sxs-lookup"><span data-stu-id="f5417-161">GeneralException</span></span>      | <span data-ttu-id="f5417-162">リクエストの処理中に内部エラーが発生しました。</span><span class="sxs-lookup"><span data-stu-id="f5417-162">There was an internal error while processing the request.</span></span> |
| <span data-ttu-id="f5417-163">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="f5417-163">NotImplemented</span></span>        | <span data-ttu-id="f5417-164">リクエストされた機能は実装されていません。</span><span class="sxs-lookup"><span data-stu-id="f5417-164">The requested feature isn't implemented.</span></span>  |
| <span data-ttu-id="f5417-165">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="f5417-165">UnsupportedOperation</span></span>  | <span data-ttu-id="f5417-166">試行中の操作はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f5417-166">The operation being attempted is not supported.</span></span> |
| <span data-ttu-id="f5417-167">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="f5417-167">AccessDenied</span></span>          | <span data-ttu-id="f5417-168">リクエストされた操作は実行されません。</span><span class="sxs-lookup"><span data-stu-id="f5417-168">You cannot perform the requested operation.</span></span> |
| <span data-ttu-id="f5417-169">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="f5417-169">ItemNotFound</span></span>          | <span data-ttu-id="f5417-170">要求されたリソースは存在しません。</span><span class="sxs-lookup"><span data-stu-id="f5417-170">The requested resource doesn't exist.</span></span> |

## <a name="get-started"></a><span data-ttu-id="f5417-171">はじめてみよう</span><span class="sxs-lookup"><span data-stu-id="f5417-171">Get started</span></span>

<span data-ttu-id="f5417-172">このセクションの例を使用して、実際に試してみましょう。</span><span class="sxs-lookup"><span data-stu-id="f5417-172">You can use the example in this section to get started.</span></span> <span data-ttu-id="f5417-173">この例では、プログラムを使用して Visio のダイアグラムで選択した形の図形のテキストを表示する方法を表示します。</span><span class="sxs-lookup"><span data-stu-id="f5417-173">This example shows you how to programmatically display the shape text of the selected shape in a Visio diagram.</span></span> <span data-ttu-id="f5417-174">まずは、SharePoint Online で通常のページを作成するか、既存のページを編集します。</span><span class="sxs-lookup"><span data-stu-id="f5417-174">To begin, create a classic page in SharePoint Online or edit an existing page.</span></span> <span data-ttu-id="f5417-175">スクリプト エディターの web パーツをページに追加し、次のコードをコピー＆ペーストします。</span><span class="sxs-lookup"><span data-stu-id="f5417-175">Add a script editor webpart on the page and copy-paste the following code.</span></span>

```js
<script src='https://appsforoffice.microsoft.com/embedded/1.0/visio-web-embedded.js' type='text/javascript'></script>

Enter Visio File Url:<br/>
<script language="javascript">
document.write("<input type='text' id='fileUrl' size='120'/>");
document.write("<input type='button' value='InitEmbeddedFrame' onclick='initEmbeddedFrame()' />");
document.write("<br />");
document.write("<input type='button' value='SelectedShapeText' onclick='getSelectedShapeText()' />");
document.write("<textarea id='ResultOutput' style='width:350px;height:60px'> </textarea>");
document.write("<div id='iframeHost' />");

let session; // Global variable to store the session and pass it afterwards in Visio.run()
var textArea;
// Loads the Visio application and Initializes communication between developer frame and Visio online frame
function initEmbeddedFrame() {
    textArea = document.getElementById('ResultOutput');
    var url = document.getElementById('fileUrl').value;
    if (!url) {
        window.alert("File URL should not be empty");
    }
    // APIs are enabled for EmbedView action only.
    url = url.replace("action=view","action=embedview");
    url = url.replace("action=interactivepreview","action=embedview");
    url = url.replace("action=default","action=embedview");
    url = url.replace("action=edit","action=embedview");
  
    session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
    return session.init().then(function () {
        // Initialization is successful
        textArea.value  = "Initialization is successful";
    });
}

// Code for getting selected Shape Text using the shapes collection object
function getSelectedShapeText() {
    Visio.run(session, function (context) {
        var page = context.document.getActivePage();
        var shapes = page.shapes;
        shapes.load();
        return context.sync().then(function () {
            textArea.value = "Please select a Shape in the Diagram";
            for(var i=0; i<shapes.items.length;i++) {
                var shape = shapes.items[i];
                if ( shape.select == true) {
                    textArea.value = shape.text;
                    return;
                }
            }
        });
    }).catch(function(error) {
        textArea.value = "Error: ";
        if (error instanceof OfficeExtension.Error) {
            textArea.value += "Debug info: " + JSON.stringify(error.debugInfo);
        }
    });
}
</script>
```

<span data-ttu-id="f5417-176">次に、作業する Visio ダイアグラムの URL が必要になります。</span><span class="sxs-lookup"><span data-stu-id="f5417-176">After that, all you need is the URL of a Visio diagram that you want to work with.</span></span> <span data-ttu-id="f5417-177">Visio ダイアグラムを SharePoint オンライン にアップロードし、 Visio Online で開きます。</span><span class="sxs-lookup"><span data-stu-id="f5417-177">Just upload the Visio diagram to SharePoint Online and open it in Visio Online.</span></span> <span data-ttu-id="f5417-178">そこから埋め込みダイアログ ボックスを開き、上の例の埋め込み URL を使用します。</span><span class="sxs-lookup"><span data-stu-id="f5417-178">From there, open the Embed dialog and use the Embed URL in the above example.</span></span>

![埋め込みダイアログから Visio ファイルの URL をコピーする](../images/Visio-embed-url.png)

<span data-ttu-id="f5417-180">Visio Onlineを編集モードで使用している場合は、**[File]** > **[Share]** > **[Embed]** を選択し、埋め込みダイアログを開きます。</span><span class="sxs-lookup"><span data-stu-id="f5417-180">If you are using Visio Online in Edit mode, open the Embed dialog by choosing **File** > **Share** > **Embed**.</span></span> <span data-ttu-id="f5417-181">Visio Onlineをビュー モードで使用している場合は、［...］の後 **［埋め込み］** を選択し、埋め込みダイアログを開きます。</span><span class="sxs-lookup"><span data-stu-id="f5417-181">If you are using Visio Online in View mode, open the Embed dialog by choosing '...' and then **Embed**.</span></span>

## <a name="open-api-specifications"></a><span data-ttu-id="f5417-182">Open API の仕様</span><span class="sxs-lookup"><span data-stu-id="f5417-182">Open API specifications</span></span>

<span data-ttu-id="f5417-p118">新しい API の設計と開発にあたり、 [「Open API の仕様」](../openspec.md) ページでフィードバックが可能になります。パイプラインの新機能をご確認いただき、設計の仕様に関する情報をお寄せください。</span><span class="sxs-lookup"><span data-stu-id="f5417-p118">As we design and develop new APIs, we'll make them available for your feedback on our [Open API specifications](../openspec.md) page. Find out what new features are in the pipeline, and provide your input on our design specifications.</span></span>

## <a name="visio-javascript-api-reference"></a><span data-ttu-id="f5417-185">Visio JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="f5417-185">Visio JavaScript APIs reference</span></span>

<span data-ttu-id="f5417-186">Visio JavaScript API の詳細情報については、 [「Visio JavaScript API リファレンス ドキュメント」](/javascript/api/visio) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f5417-186">For detailed information about the Word JavaScript API, see the [Word JavaScript API reference documentation](/javascript/api/visio).</span></span>
