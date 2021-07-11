Word 2013 または PowerPoint 2013 のドキュメントをワンクリックでリモートの場所に送信または発行できるようにする Office アドインを作成できます。この記事では、プレゼンテーション全体をデータ オブジェクトとして取得し、そのデータを HTTP 要求を通じて Web サーバーに送信する、PowerPoint 2013 用の簡単な作業ウィンドウ アドインの作成方法を具体例によって示します。

## <a name="prerequisites-for-creating-an-add-in-for-powerpoint-or-word"></a>PowerPoint または Word 用アドインを作成するための前提条件

この記事では、PowerPoint または Word 用の作業ウィンドウ アドインの作成にテキスト エディターを使用することを前提にしています。 作業ウィンドウ アドインを作成するには、次のファイルを作成する必要があります。

- 共有ネットワーク フォルダーまたは Web サーバーでは、次のファイルが必要です。

  - ユーザー インターフェイスと JavaScript ファイル (office.js およびアプリケーション固有の .js ファイルを含む) およびカスケード スタイル シート (CSS) ファイルへのリンクを含む HTML ファイル (GetDoc_App.html)。

  - JavaScript ファイル (GetDoc_App.js)。このファイルにはアドインのプログラミング ロジックが格納されます。

  - CSS ファイル (Program.css)。アドインのスタイルと書式を入れるファイルです。

- アドインの XML マニフェスト ファイル (GetDoc_App.xml)。共有ネットワーク フォルダーまたはアドイン カタログで使用できます。このマニフェスト ファイルでは、上述の HTML ファイルの場所を指していることが必要です。

Visual Studio または[yeoman](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio)ジェネレーターを Office アドインに使用するか、Office アドインに[Visual Studio](../quickstarts/word-quickstart.md?tabs=visualstudio)または[Yeoman](../quickstarts/powerpoint-quickstart.md?tabs=yeomangenerator)ジェネレーターを使用して Word 用に PowerPoint 用のアドイン[を](../quickstarts/word-quickstart.md?tabs=yeomangenerator)作成できます。

### <a name="core-concepts-to-know-for-creating-a-task-pane-add-in"></a>作業ウィンドウ アドインを作成するために知っておくべき主要な概念

この PowerPoint または Word 用アドインの作成を開始する前に、Office アドインの作成と HTTP 要求の操作についてよく理解しておくことが必要です。この記事では、Web サーバー上の HTTP 要求から Base64 エンコード テキストをデコードする方法については説明しません。

## <a name="create-the-manifest-for-the-add-in"></a>アドインのマニフェストを作成する

PowerPoint 用アドインの XML マニフェスト ファイルは、アドインをホストできるアプリケーション、HTML ファイルの場所、アドインのタイトルと説明、およびその他の多くの特性に関する重要な情報を提供します。

1. テキスト エディターで、次のコードをマニフェスト ファイルに追加します。

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

2. このファイルを、UTF-8 エンコードを使用して、GetDoc_App.xml としてネットワークの場所またはアドイン カタログに保存します。

## <a name="create-the-user-interface-for-the-add-in"></a>アドインのユーザー インターフェイスを作成する

アドインのユーザー インターフェイスとしては、GetDoc_App.html ファイルに直接書き込んだ HTML を使用できます。このアドインのプログラミング ロジックと機能は、JavaScript ファイル (例: GetDoc_App.js) に入れる必要があります。

以下の手順を使用して、見出しと 1 つのボタンを含むアドインの簡単なユーザー インターフェイスを作成します。

1. テキスト エディターで、新しいファイルに次の HTML を追加します。

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

2. このファイルを、UTF-8 エンコードを使用して GetDoc_App.html としてネットワークの場所または Web サーバーに保存します。

    > [!NOTE]
    > 必ずアドインの **head** タグに **script** タグと office.js ファイルへの有効なリンクを入れてください。

    CSS を使用して、アドインに単純ながら最新の本格的な外観を与えます。次の CSS を使用して、アドインのスタイルを定義します。

3. テキスト エディターで、新しいファイルに次の CSS を追加します。

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

4. このファイルを、UTF-8 エンコードを使用して、Program.css としてネットワークの場所または Web サーバー (GetDoc_App.html を保存した場所) に保存します。

## <a name="add-the-javascript-to-get-the-document"></a>ドキュメントを取得するための JavaScript を追加する

アドインのコードでは、[Office.initialize](/javascript/api/office) イベントのハンドラーが、フォーム上の **[送信]** ボタンのクリック イベントのハンドラーを追加し、アドインの準備ができたことをユーザーに知らせます。

次のコード例は、status div に書き込む場合のヘルパー関数と共に、イベントの `Office.initialize` `updateStatus` イベント ハンドラーを示しています。

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
    statusInfo[0].innerHTML += message + "<br/>";
}
```

UI で [ **送信** ] ボタンを選択すると、アドインは `sendFile` [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-) メソッドの呼び出しを含む関数を呼び出します。 このメソッドは、JavaScript API の他のメソッドと同様に、非同期パターン `getFileAsync` を使用Office。 このメソッドには、_fileType_ という 1 つの必須パラメーターと、_options_ と _callback_ という 2 つの省略可能なパラメーターがあります。

_fileType パラメーターは_[、FileType](/javascript/api/office/office.filetype)列挙 `Office.FileType.Compressed` ("compressed")、Office.FileType.PDF("pdf")、または Office の 3 つの定数 **のいずれかを必要とします。FileType.Text** ("text")。 各プラットフォームの現在のファイルの種類のサポートは [、Document.getFileType の備考の下に一覧表示](/javascript/api/office/office.document#getFileAsync_fileType__callback_) されます。 _fileType_ パラメーターに対して **Compressed** を渡す場合、ローカル コンピューターにファイルの一時的なコピーを作成して、ドキュメントを PowerPoint 2013 プレゼンテーション ファイル (.pptx) または `getFileAsync` Word *2013* ドキュメント ファイル (.docx) として返します。

この `getFileAsync` メソッドは、ファイルへの参照を File オブジェクトとして [返](/javascript/api/office/office.file) します。 オブジェクト `File` は[、size プロパティ、sliceCount](/javascript/api/office/office.file#size)プロパティ[](/javascript/api/office/office.file#slicecount)[、getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-)メソッド、[および closeAsync](/javascript/api/office/office.file#closeasync-callback-)メソッドの 4 つのメンバーを公開します。 この `size` プロパティは、ファイル内のバイト数を返します。 ファイル `sliceCount` 内の [Slice](/javascript/api/office/office.slice) オブジェクトの数 (この記事で後で説明します) を返します。

次のコードを使用して、メソッドPowerPoint Word ドキュメントをオブジェクトとして取得し、ローカルで定義された関数 `File` `Document.getFileAsync` を呼び出 `getSlice` します。 オブジェクト、カウンター変数、およびファイル内のスライスの総数は、匿名オブジェクトの呼び出しで `File` `getSlice` 渡されます。

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

ローカル関数は `getSlice` 、オブジェクトからスライスを `File.getSliceAsync` 取得するメソッドを呼び出 `File` します。 この `getSliceAsync` メソッドは、スライス `Slice` のコレクションからオブジェクトを返します。 このメソッドには、 _sliceIndex_ と _callback_ という 2 つの必須パラメーターがあります。 _sliceIndex_ パラメーターは、スライスのコレクションへのインデクサーとして整数を取ります。 JavaScript API for Officeの他の関数と同様に、メソッドはコールバック関数もパラメーターとして受け取り、メソッド呼び出しの結果 `getSliceAsync` を処理します。
ion `getSlice` は **File.getSliceAsync** メソッドを呼び出して **、File** オブジェクトからスライスを取得します。 **getSliceAsync** メソッドは、スライスのコレクションから **Slice** オブジェクトを返します。 このメソッドには、 _sliceIndex_ と _callback_ という 2 つの必須パラメーターがあります。 _sliceIndex_ パラメーターは、スライスのコレクションへのインデクサーとして整数を取ります。 Office JavaScript API の他の関数と同様に **、getSliceAsync** メソッドもコールバック関数をパラメーターとして受け取り、メソッド呼び出しの結果を処理します。

オブジェクト `Slice` を使用すると、ファイルに含まれるデータにアクセスできます。 メソッドの options パラメーターで _特に_ 指定しない限り、 `getFileAsync` `Slice` オブジェクトのサイズは 4 MB です。 オブジェクト `Slice` は、サイズ、データ、 [インデックス](/javascript/api/office/office.slice#size)の 3 [つの](/javascript/api/office/office.slice#data)プロパティを [公開します](/javascript/api/office/office.slice#index)。 プロパティ `size` は、スライスのサイズ (バイト単位) を取得します。 プロパティ `index` は、スライスのコレクション内のスライスの位置を表す整数を取得します。

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

この `Slice.data` プロパティは、ファイルの生データをバイト配列として返します。 データがテキスト形式 (つまり、XML かプレーン テキスト) の場合、スライスには生テキストが含まれています。 fileType パラメーター **の Office.FileType.Compressed** を渡す場合、スライスにはファイルのバイナリ データがバイト配列 `Document.getFileAsync` として含まれる。 In the case of a PowerPoint or Word file, the slices contain byte arrays.

バイト配列のデータを Base64 でエンコードされた文字列に変換するには、独自の関数を実装 (または利用可能なライブラリを使用) する必要があります。JavaScript での Base64 エンコーティングについては、「 [Base64 エンコードとデコード](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding)」を参照してください。

データを Base64 に変換した後は、HTTP POST 要求の本体として送信するなど、そのデータをいろいろな方法で Web サーバーに送信できます。

スライスを Web サービスに送信するために次のコードを追加します。

> [!NOTE]
> このコードは、PowerPointまたは Word ファイルを複数のスライスで Web サーバーに送信します。 Web サーバーまたはサービスは、各スライスを 1 つのファイルに追加し、.pptx または .docx ファイルとして保存してから、操作を実行する必要があります。

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

名前が示すように、メソッドはドキュメントへの接続を閉じ `File.closeAsync` 、リソースを解放します。 Office アドインのサンドボックス ガベージはファイルへのスコープ外参照を収集しますが、コードでファイルを使い終わった後、それらのファイルを明示的に閉じるためのベスト プラクティスであることに変わりありません。 メソッド `closeAsync` には、呼び出しの完了時に呼び出す関数を指定する 1 つのパラメーターコールバックがあります。

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