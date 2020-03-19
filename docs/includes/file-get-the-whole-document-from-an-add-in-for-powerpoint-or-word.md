<span data-ttu-id="5c52d-p101">Word 2013 または PowerPoint 2013 のドキュメントをワンクリックでリモートの場所に送信または発行できるようにする Office アドインを作成できます。この記事では、プレゼンテーション全体をデータ オブジェクトとして取得し、そのデータを HTTP 要求を通じて Web サーバーに送信する、PowerPoint 2013 用の簡単な作業ウィンドウ アドインの作成方法を具体例によって示します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-p101">You can create an Office Add-in to provide one-click sending or publishing of a Word 2013 or PowerPoint 2013 document to a remote location. This article demonstrates how to build a simple task pane add-in for PowerPoint 2013 that gets all of the presentation as a data object and sends that data to a web server via an HTTP request.</span></span>

## <a name="prerequisites-for-creating-an-add-in-for-powerpoint-or-word"></a><span data-ttu-id="5c52d-103">PowerPoint または Word 用アドインを作成するための前提条件</span><span class="sxs-lookup"><span data-stu-id="5c52d-103">Prerequisites for creating an add-in for PowerPoint or Word</span></span>

<span data-ttu-id="5c52d-p102">この記事では、PowerPoint または Word 用の作業ウィンドウ アドインの作成にテキスト エディターを使用することを前提にしています。この作業ウィンドウ アドインを作成するには、以下のファイルを作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5c52d-p102">This article assumes that you are using a text editor to create the task pane add-in for PowerPoint or Word. To create the task pane add-in, you must create the following files:</span></span>

- <span data-ttu-id="5c52d-106">共有ネットワーク フォルダーまたは Web サーバー上に次のファイルが必要です。</span><span class="sxs-lookup"><span data-stu-id="5c52d-106">On a shared network folder or on a web server, you need the following files:</span></span>

    - <span data-ttu-id="5c52d-107">HTML ファイル (GetDoc_App.html)。このファイルには、ユーザー インターフェイスに加えて、JavaScript ファイル (office.js やホスト固有の .js ファイルなど) およびカスケード スタイル シート (CSS) ファイルへのリンクが格納されます。</span><span class="sxs-lookup"><span data-stu-id="5c52d-107">An HTML file (GetDoc_App.html) that contains the user interface plus links to the JavaScript files (including office.js and host-specific .js files) and Cascading Style Sheet (CSS) files.</span></span>

    - <span data-ttu-id="5c52d-108">JavaScript ファイル (GetDoc_App.js)。このファイルにはアドインのプログラミング ロジックが格納されます。</span><span class="sxs-lookup"><span data-stu-id="5c52d-108">A JavaScript file (GetDoc_App.js) to contain the programming logic of the add-in.</span></span>

    - <span data-ttu-id="5c52d-109">CSS ファイル (Program.css)。アドインのスタイルと書式を入れるファイルです。</span><span class="sxs-lookup"><span data-stu-id="5c52d-109">A CSS file (Program.css) to contain the styles and formatting for the add-in.</span></span>

- <span data-ttu-id="5c52d-p103">アドインの XML マニフェスト ファイル (GetDoc_App.xml)。共有ネットワーク フォルダーまたはアドイン カタログで使用できます。このマニフェスト ファイルでは、上述の HTML ファイルの場所を指していることが必要です。</span><span class="sxs-lookup"><span data-stu-id="5c52d-p103">An XML manifest file (GetDoc_App.xml) for the add-in, available on a shared network folder or add-in catalog. The manifest file must point to the location of the HTML file mentioned previously.</span></span>

<span data-ttu-id="5c52d-112">また、 [visual studio を使用し](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio)て、または office アドイン用に[visual studio](../quickstarts/word-quickstart.md?tabs=visualstudio)またはサード[オマーンジェネレーター](../quickstarts/word-quickstart.md?tabs=yeomangenerator)を使用して Word に対し[て、PowerPoint](../quickstarts/powerpoint-quickstart.md?tabs=yeomangenerator)用のアドインを作成することもできます。</span><span class="sxs-lookup"><span data-stu-id="5c52d-112">You can also create an add-in for PowerPoint by using [Visual Studio](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio) or the [Yeoman generator for Office Add-ins](../quickstarts/powerpoint-quickstart.md?tabs=yeomangenerator) or for Word by using [Visual Studio](../quickstarts/word-quickstart.md?tabs=visualstudio) or [Yeoman generator for Office Add-ins](../quickstarts/word-quickstart.md?tabs=yeomangenerator).</span></span>

### <a name="core-concepts-to-know-for-creating-a-task-pane-add-in"></a><span data-ttu-id="5c52d-113">作業ウィンドウ アドインを作成するために知っておくべき主要な概念</span><span class="sxs-lookup"><span data-stu-id="5c52d-113">Core concepts to know for creating a task pane add-in</span></span>

<span data-ttu-id="5c52d-p104">この PowerPoint または Word 用アドインの作成を開始する前に、Office アドインの作成と HTTP 要求の操作についてよく理解しておくことが必要です。この記事では、Web サーバー上の HTTP 要求から Base64 エンコード テキストをデコードする方法については説明しません。</span><span class="sxs-lookup"><span data-stu-id="5c52d-p104">Before you begin creating this add-in for PowerPoint or Word, you should be familiar with building Office Add-ins and working with HTTP requests. This article does not discuss how to decode Base64-encoded text from an HTTP request on a web server.</span></span> 

## <a name="create-the-manifest-for-the-add-in"></a><span data-ttu-id="5c52d-116">アドインのマニフェストを作成する</span><span class="sxs-lookup"><span data-stu-id="5c52d-116">Create the manifest for the add-in</span></span>

<span data-ttu-id="5c52d-117">PowerPoint 用アドインの XML マニフェスト ファイルは、アドインをホストできるアプリケーション、HTML ファイルの場所、アドインのタイトルと説明、およびその他の多くの特性に関する重要な情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-117">The XML manifest file for the add-in for PowerPoint provides important information about the add-in: what applications can host it, the location of the HTML file, the add-in title and description, and many other characteristics.</span></span>

1. <span data-ttu-id="5c52d-118">テキスト エディターで、次のコードをマニフェスト ファイルに追加します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-118">In a text editor, add the following code to the manifest file.</span></span>

    ```xml  
    <?xml version="1.0" encoding="utf-8" ?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xsi:type="TaskPaneApp">
        <Id>[Replace_With_Your_GUID]</Id>
        <Version>1.0</Version>
        <ProviderName>[Provider Name]</ProviderName>
        <DefaultLocale>EN-US</DefaultLocale>
        <DisplayName DefaultValue="Get Doc add-in" />
        <Description DefaultValue="My get PowerPoint or Word document add-in." />
        <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" />
        <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
        <Hosts>
        <Host Name="Document" />
        <Host Name="Presentation" />
        </Hosts>
        <DefaultSettings>
        <SourceLocation DefaultValue="[Network location of app]/GetDoc_App.html" />
        </DefaultSettings>
        <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

2. <span data-ttu-id="5c52d-119">このファイルを、UTF-8 エンコードを使用して、GetDoc_App.xml としてネットワークの場所またはアドイン カタログに保存します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-119">Save the file as GetDoc_App.xml using UTF-8 encoding to a network location or to an add-in catalog.</span></span>

## <a name="create-the-user-interface-for-the-add-in"></a><span data-ttu-id="5c52d-120">アドインのユーザー インターフェイスを作成する</span><span class="sxs-lookup"><span data-stu-id="5c52d-120">Create the user interface for the add-in</span></span>

<span data-ttu-id="5c52d-p105">アドインのユーザー インターフェイスとしては、GetDoc_App.html ファイルに直接書き込んだ HTML を使用できます。このアドインのプログラミング ロジックと機能は、JavaScript ファイル (例: GetDoc_App.js) に入れる必要があります。</span><span class="sxs-lookup"><span data-stu-id="5c52d-p105">For the user interface of the add-in, you can use HTML, written directly into the GetDoc_App.html file. The programming logic and functionality of the add-in must be contained in a JavaScript file (for example, GetDoc_App.js).</span></span>

<span data-ttu-id="5c52d-123">以下の手順を使用して、見出しと 1 つのボタンを含むアドインの簡単なユーザー インターフェイスを作成します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-123">Use the following procedure to create a simple user interface for the add-in that includes a heading and a single button.</span></span>

1. <span data-ttu-id="5c52d-124">テキスト エディターで、新しいファイルに次の HTML を追加します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-124">In a new file in the text editor, add the following HTML.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
            <title>Publish presentation</title>
            <link rel="stylesheet" type="text/css" href="Program.css" />
            <script src="https://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js" type="text/javascript"></script>
            <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
            <script src="GetDoc_App.js"></script>
        </head>
        <body>
        <form>
            <h1>Publish presentation</h1>
            <br />
            <div><input id='submit' type="button" value="Submit" /></div>
            <br />
            <div><h2>Status</h2> 
                <div id="status"></div>
            </div>
        </form>
        </body>
    </html>
    ```

2. <span data-ttu-id="5c52d-125">このファイルを、UTF-8 エンコードを使用して GetDoc_App.html としてネットワークの場所または Web サーバーに保存します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-125">Save the file as GetDoc_App.html using UTF-8 encoding to a network location or to a web server.</span></span>

    > [!NOTE]
    > <span data-ttu-id="5c52d-126">必ずアドインの **head** タグに **script** タグと office.js ファイルへの有効なリンクを入れてください。</span><span class="sxs-lookup"><span data-stu-id="5c52d-126">Be sure that the **head** tags of the add-in contains a **script** tag with a valid link to the office.js file.</span></span> 

    <span data-ttu-id="5c52d-p106">CSS を使用して、アドインに単純ながら最新の本格的な外観を与えます。次の CSS を使用して、アドインのスタイルを定義します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-p106">We'll use some CSS to give the add-in a simple, yet modern and professional appearance. Use the following CSS to define the style of the add-in.</span></span>

3. <span data-ttu-id="5c52d-129">テキスト エディターで、新しいファイルに次の CSS を追加します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-129">In a new file in the text editor, add the following CSS.</span></span>

    ```css  
    body
    {
        font-family: "Segoe UI Light","Segoe UI",Tahoma,sans-serif;
    }
    h1,h2
    {
        text-decoration-color:#4ec724;
    }
    input [type="submit"], input[type="button"]
    {
        height:24px;
        padding-left:1em;
        padding-right:1em;
        background-color:white;
        border:1px solid grey;
        border-color: #dedfe0 #b9b9b9 #b9b9b9 #dedfe0;
        cursor:pointer;
    }
    ```

4. <span data-ttu-id="5c52d-130">このファイルを、UTF-8 エンコードを使用して、Program.css としてネットワークの場所または Web サーバー (GetDoc_App.html を保存した場所) に保存します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-130">Save the file as Program.css using UTF-8 encoding to the network location or to the web server where the GetDoc_App.html file is located.</span></span>

## <a name="add-the-javascript-to-get-the-document"></a><span data-ttu-id="5c52d-131">ドキュメントを取得するための JavaScript を追加する</span><span class="sxs-lookup"><span data-stu-id="5c52d-131">Add the JavaScript to get the document</span></span>

<span data-ttu-id="5c52d-132">アドインのコードでは、[Office.initialize](/javascript/api/office) イベントのハンドラーが、フォーム上の **[送信]** ボタンのクリック イベントのハンドラーを追加し、アドインの準備ができたことをユーザーに知らせます。</span><span class="sxs-lookup"><span data-stu-id="5c52d-132">In the code for the add-in, a handler to the [Office.initialize](/javascript/api/office) event adds a handler to the click event of the **Submit** button on the form and informs the user that the add-in is ready.</span></span>

<span data-ttu-id="5c52d-133">次のコード例は、ステータス div に書き込む`Office.initialize`ための`updateStatus`、イベントのイベントハンドラーをヘルパー関数と共に示しています。</span><span class="sxs-lookup"><span data-stu-id="5c52d-133">The following code example shows the event handler for the `Office.initialize` event along with a helper function, `updateStatus`, for writing to the status div.</span></span>

```js
// The initialize function is required for all add-ins.
Office.initialize = function (reason) {

    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {

        // Execute sendFile when submit is clicked
        $('#submit').click(function () {
            sendFile();
        });

        // Update status
        updateStatus("Ready to send file.");
    });
}

// Create a function for writing to the status div.
function updateStatus(message) {
    var statusInfo = $('#status');
    statusInfo.innerHTML += message + "<br/>";
}
```

<span data-ttu-id="5c52d-134">UI で [**送信**] ボタンを選択すると、アドインは、 [getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)メソッド`sendFile`への呼び出しを含む関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-134">When you choose the **Submit** button in the UI, the add-in calls the `sendFile` function, which contains a call to the [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-) method.</span></span> <span data-ttu-id="5c52d-135">メソッド`getFileAsync`は、JavaScript API for Office の他のメソッドと同様に、非同期パターンを使用します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-135">The `getFileAsync` method uses the asynchronous pattern, similar to other methods in the JavaScript API for Office.</span></span> <span data-ttu-id="5c52d-136">このメソッドには、_fileType_ という 1 つの必須パラメーターと、_options_ と _callback_ という 2 つの省略可能なパラメーターがあります。</span><span class="sxs-lookup"><span data-stu-id="5c52d-136">It has one required parameter, _fileType_, and two optional parameters,  _options_ and _callback_.</span></span> 


<span data-ttu-id="5c52d-137">_Filetype_パラメーターには、 [filetype](/javascript/api/office/office.filetype)列挙型`Office.FileType.Compressed` ("圧縮")、" **.pdf** " ("pdf") **、または "テキスト (** " テキスト ") の3つの定数のいずれかが想定されています。</span><span class="sxs-lookup"><span data-stu-id="5c52d-137">The  _fileType_ parameter expects one of three constants from the [FileType](/javascript/api/office/office.filetype) enumeration: `Office.FileType.Compressed` ("compressed"), **Office.FileType.PDF** ("pdf"), or **Office.FileType.Text** ("text").</span></span> <span data-ttu-id="5c52d-138">PowerPoint supports only **Compressed** as an argument; Word supports all three.</span><span class="sxs-lookup"><span data-stu-id="5c52d-138">PowerPoint supports only **Compressed** as an argument; Word supports all three.</span></span> <span data-ttu-id="5c52d-139">_FileType_パラメーターに対して**圧縮**を渡すと、この`getFileAsync`メソッドは、ローカルコンピューターにファイルの一時コピーを作成することによって、PowerPoint 2013 プレゼンテーションファイル (*.pptx) または Word 2013 ドキュメントファイル (*.docx) としてドキュメントを返します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-139">When you pass in **Compressed** for the _fileType_ parameter, the `getFileAsync` method returns the document as a PowerPoint 2013 presentation file (*.pptx) or Word 2013 document file (*.docx) by creating a temporary copy of the file on the local computer.</span></span>

<span data-ttu-id="5c52d-140">この`getFileAsync`メソッドは、[ファイルオブジェクトとしてファイルへ](/javascript/api/office/office.file)の参照を返します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-140">The `getFileAsync` method returns a reference to the file as a [File](/javascript/api/office/office.file) object.</span></span> <span data-ttu-id="5c52d-141">この`File`オブジェクトは、 [Size](/javascript/api/office/office.file#size)プロパティ、[スライス ecount](/javascript/api/office/office.file#slicecount)プロパティ、 [getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-)メソッド、および[closeasync](/javascript/api/office/office.file#closeasync-callback-)メソッドの4つのメンバーを公開します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-141">The `File` object exposes four members: the [size](/javascript/api/office/office.file#size) property, [sliceCount](/javascript/api/office/office.file#slicecount) property, [getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-) method, and [closeAsync](/javascript/api/office/office.file#closeasync-callback-) method.</span></span> <span data-ttu-id="5c52d-142">プロパティ`size`は、ファイル内のバイト数を返します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-142">The `size` property returns the number of bytes in the file.</span></span> <span data-ttu-id="5c52d-143">は`sliceCount` 、この記事で後述する[Slice](/javascript/api/office/office.slice)オブジェクトの数をファイルに返します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-143">The `sliceCount` returns the number of [Slice](/javascript/api/office/office.slice) objects (discussed later in this article) in the file.</span></span>

<span data-ttu-id="5c52d-144">次のコードを使用して、 `File` `Document.getFileAsync`メソッドを使用して PowerPoint または Word 文書をオブジェクトとして取得し、 `getSlice`ローカルに定義された関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-144">Use the following code to get the PowerPoint or Word document as a `File` object using the `Document.getFileAsync` method and then makes a call to the locally defined `getSlice` function.</span></span> <span data-ttu-id="5c52d-145">`File`オブジェクト、カウンタ変数、およびファイル内のスライスの合計数は、匿名オブジェクト`getSlice`での呼び出しに沿って渡されることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="5c52d-145">Note that the `File` object, a counter variable, and the total number of slices in the file are passed along in the call to `getSlice` in an anonymous object.</span></span>

```js
// Get all of the content from a PowerPoint or Word document in 100-KB chunks of text.
function sendFile() {
    Office.context.document.getFileAsync("compressed",
        { sliceSize: 100000 },
        function (result) {

            if (result.status == Office.AsyncResultStatus.Succeeded) {

                // Get the File object from the result.
                var myFile = result.value;
                var state = {
                    file: myFile,
                    counter: 0,
                    sliceCount: myFile.sliceCount
                };

                updateStatus("Getting file of " + myFile.size + " bytes");
                getSlice(state);
            }
            else {
                updateStatus(result.status);
            }
        });
}
```

<span data-ttu-id="5c52d-146">ローカル関数`getSlice`は、 `File.getSliceAsync` `File`オブジェクトからスライスを取得するメソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-146">The local function `getSlice` makes a call to the `File.getSliceAsync` method to retrieve a slice from the `File` object.</span></span> <span data-ttu-id="5c52d-147">メソッド`getSliceAsync`は、スライス`Slice`のコレクションからオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-147">The `getSliceAsync` method returns a `Slice` object from the collection of slices.</span></span> <span data-ttu-id="5c52d-148">このメソッドには、 _sliceIndex_ と _callback_ という 2 つの必須パラメーターがあります。</span><span class="sxs-lookup"><span data-stu-id="5c52d-148">It has two required parameters, _sliceIndex_ and _callback_.</span></span> <span data-ttu-id="5c52d-149">_sliceIndex_ パラメーターは、スライスのコレクションへのインデクサーとして整数を取ります。</span><span class="sxs-lookup"><span data-stu-id="5c52d-149">The  _sliceIndex_ parameter takes an integer as an indexer into the collection of slices.</span></span> <span data-ttu-id="5c52d-150">JavaScript API for Office の他の関数と同様に`getSliceAsync` 、メソッドはコールバック関数をメソッド呼び出しの結果を処理するパラメーターとしても受け取ります。</span><span class="sxs-lookup"><span data-stu-id="5c52d-150">Like other functions in the JavaScript API for Office, the `getSliceAsync` method also takes a callback function as a parameter to handle the results from the method call.</span></span>
<span data-ttu-id="5c52d-151">イオン`getSlice` **ファイル**オブジェクトからスライスを取得するために、getSliceAsync メソッドを呼び出し**ます。**</span><span class="sxs-lookup"><span data-stu-id="5c52d-151">ion `getSlice` makes a call to the **File.getSliceAsync** method to retrieve a slice from the **File** object.</span></span> <span data-ttu-id="5c52d-152">**getSliceAsync** メソッドは、スライスのコレクションから **Slice** オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-152">The **getSliceAsync** method returns a **Slice** object from the collection of slices.</span></span> <span data-ttu-id="5c52d-153">このメソッドには、 _sliceIndex_ と _callback_ という 2 つの必須パラメーターがあります。</span><span class="sxs-lookup"><span data-stu-id="5c52d-153">It has two required parameters, _sliceIndex_ and _callback_.</span></span> <span data-ttu-id="5c52d-154">_sliceIndex_ パラメーターは、スライスのコレクションへのインデクサーとして整数を取ります。</span><span class="sxs-lookup"><span data-stu-id="5c52d-154">The  _sliceIndex_ parameter takes an integer as an indexer into the collection of slices.</span></span> <span data-ttu-id="5c52d-155">Office JavaScript API の他の関数と同様に、 **getSliceAsync**メソッドも、メソッド呼び出しの結果を処理するパラメーターとしてコールバック関数を受け取ります。</span><span class="sxs-lookup"><span data-stu-id="5c52d-155">Like other functions in the Office JavaScript API, the **getSliceAsync** method also takes a callback function as a parameter to handle the results from the method call.</span></span>

<span data-ttu-id="5c52d-156">この`Slice`オブジェクトを使用すると、ファイルに格納されているデータにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="5c52d-156">The `Slice` object gives you access to the data contained in the file.</span></span> <span data-ttu-id="5c52d-157">メソッドの_options_パラメーターで特に指定されてい`Slice`ない限り、オブジェクトのサイズは 4 MB です。 `getFileAsync`</span><span class="sxs-lookup"><span data-stu-id="5c52d-157">Unless otherwise specified in the _options_ parameter of the `getFileAsync` method, the `Slice` object is 4 MB in size.</span></span> <span data-ttu-id="5c52d-158">この`Slice`オブジェクトは、 [size](/javascript/api/office/office.slice#size)、 [data](/javascript/api/office/office.slice#data)、および[index](/javascript/api/office/office.slice#index)という3つのプロパティを公開します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-158">The `Slice` object exposes three properties: [size](/javascript/api/office/office.slice#size), [data](/javascript/api/office/office.slice#data), and [index](/javascript/api/office/office.slice#index).</span></span> <span data-ttu-id="5c52d-159">プロパティ`size`は、スライスのサイズ (バイト単位) を取得します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-159">The `size` property gets the size, in bytes, of the slice.</span></span> <span data-ttu-id="5c52d-160">プロパティ`index`は、スライスのコレクション内のスライスの位置を表す整数を取得します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-160">The `index` property gets an integer that represents the slice's position in the collection of slices.</span></span>

```js
// Get a slice from the file and then call sendSlice.
function getSlice(state) {
    state.file.getSliceAsync(state.counter, function (result) {
        if (result.status == Office.AsyncResultStatus.Succeeded) {
            updateStatus("Sending piece " + (state.counter + 1) + " of " + state.sliceCount);
            sendSlice(result.value, state);
        }
        else {
            updateStatus(result.status);
        }
    });
}
```

<span data-ttu-id="5c52d-161">この`Slice.data`プロパティは、ファイルの生データをバイト配列として返します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-161">The `Slice.data` property returns the raw data of the file as a byte array.</span></span> <span data-ttu-id="5c52d-162">データがテキスト形式 (つまり、XML かプレーン テキスト) の場合、スライスには生テキストが含まれています。</span><span class="sxs-lookup"><span data-stu-id="5c52d-162">If the data is in text format (that is, XML or plain text), the slice contains the raw text.</span></span> <span data-ttu-id="5c52d-163">を指定した場合、ファイルの_filetype_パラメーター**に対して**は`Document.getFileAsync`、バイト配列としてファイルのバイナリデータが格納されます。</span><span class="sxs-lookup"><span data-stu-id="5c52d-163">If you pass in **Office.FileType.Compressed** for the _fileType_ parameter of `Document.getFileAsync`, the slice contains the binary data of the file as a byte array.</span></span> <span data-ttu-id="5c52d-164">In the case of a PowerPoint or Word file, the slices contain byte arrays.</span><span class="sxs-lookup"><span data-stu-id="5c52d-164">In the case of a PowerPoint or Word file, the slices contain byte arrays.</span></span>

<span data-ttu-id="5c52d-p114">バイト配列のデータを Base64 でエンコードされた文字列に変換するには、独自の関数を実装 (または利用可能なライブラリを使用) する必要があります。JavaScript での Base64 エンコーティングについては、「 [Base64 エンコードとデコード](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5c52d-p114">You must implement your own function (or use an available library) to convert byte array data to a Base64-encoded string. For information about Base64 encoding with JavaScript, see [Base64 encoding and decoding](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding).</span></span>

<span data-ttu-id="5c52d-167">データを Base64 に変換した後は、HTTP POST 要求の本体として送信するなど、そのデータをいろいろな方法で Web サーバーに送信できます。</span><span class="sxs-lookup"><span data-stu-id="5c52d-167">Once you have converted the data to Base64, you can then transmit it to a web server in several ways -- including as the body of an HTTP POST request.</span></span>

<span data-ttu-id="5c52d-168">スライスを Web サービスに送信するために次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-168">Add the following code to send a slice to a web service.</span></span>

> [!NOTE]
> <span data-ttu-id="5c52d-169">このコードでは、PowerPoint または Word ファイルを複数のスライスの web サーバーに送信します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-169">This code sends a PowerPoint or Word file to the web server in multiple slices.</span></span> <span data-ttu-id="5c52d-170">Web サーバーまたはサービスでは、個々のスライスを1つのファイルに追加してから、.pptx ファイルまたは .docx ファイルとして保存してから、操作を実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5c52d-170">The web server or service must append each individual slice into a single file, and then save it as a .pptx or .docx file, before you can perform any manipulations on it.</span></span>

```js
function sendSlice(slice, state) {
    var data = slice.data;

    // If the slice contains data, create an HTTP request.
    if (data) {

        // Encode the slice data, a byte array, as a Base64 string.
        // NOTE: The implementation of myEncodeBase64(input) function isn't
        // included with this example. For information about Base64 encoding with
        // JavaScript, see https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding.
        var fileData = myEncodeBase64(data);

        // Create a new HTTP request. You need to send the request
        // to a webpage that can receive a post.
        var request = new XMLHttpRequest();

        // Create a handler function to update the status
        // when the request has been sent.
        request.onreadystatechange = function () {
            if (request.readyState == 4) {

                updateStatus("Sent " + slice.size + " bytes.");
                state.counter++;

                if (state.counter < state.sliceCount) {
                    getSlice(state);
                }
                else {
                    closeFile(state);
                }
            }
        }

        request.open("POST", "[Your receiving page or service]");
        request.setRequestHeader("Slice-Number", slice.index);

        // Send the file as the body of an HTTP POST
        // request to the web server.
        request.send(fileData);
    }
}
```

<span data-ttu-id="5c52d-171">その名前が示すように`File.closeAsync` 、メソッドはドキュメントへの接続を閉じ、リソースを解放します。</span><span class="sxs-lookup"><span data-stu-id="5c52d-171">As the name implies, the `File.closeAsync` method closes the connection to the document and frees up resources.</span></span> <span data-ttu-id="5c52d-172">Office アドインのサンドボックス ガベージはファイルへのスコープ外参照を収集しますが、コードでファイルを使い終わった後、それらのファイルを明示的に閉じるためのベスト プラクティスであることに変わりありません。</span><span class="sxs-lookup"><span data-stu-id="5c52d-172">Although the Office Add-ins sandbox garbage collects out-of-scope references to files, it is still a best practice to explicitly close files once your code is done with them.</span></span> <span data-ttu-id="5c52d-173">`closeAsync`メソッドには、呼び出しの完了時に呼び出す関数を指定する、1つのパラメーターの_callback_があります。</span><span class="sxs-lookup"><span data-stu-id="5c52d-173">The `closeAsync` method has a single parameter, _callback_, that specifies the function to call on the completion of the call.</span></span>

```js
function closeFile(state) {
    // Close the file when you're done with it.
    state.file.closeAsync(function (result) {

        // If the result returns as a success, the
        // file has been successfully closed.
        if (result.status == "succeeded") {
            updateStatus("File closed.");
        }
        else {
            updateStatus("File couldn't be closed.");
        }
    });
}
```