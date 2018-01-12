
# <a name="get-the-whole-document-from-an-add-in-for-powerpoint-or-word"></a>PowerPoint または Word 用アドインからドキュメント全体を取得する

Word 2013 または PowerPoint 2013 のドキュメントをワンクリックでリモートの場所に送信または発行できるようにする Office アドインを作成できます。この記事では、プレゼンテーション全体をデータ オブジェクトとして取得し、そのデータを HTTP 要求を通じて Web サーバーに送信する、PowerPoint 2013 用の簡単な作業ウィンドウ アドインの作成方法を具体例によって示します。

## <a name="prerequisites-for-creating-an-add-in-for-powerpoint-or-word"></a>PowerPoint または Word 用アドインを作成するための前提条件


この記事では、PowerPoint または Word 用の作業ウィンドウ アドインの作成にテキスト エディターを使用することを前提にしています。この作業ウィンドウ アドインを作成するには、以下のファイルを作成する必要があります。


- 共有ネットワーク フォルダーまたは Web サーバー上に次のファイルが必要です。
    
      - HTML ファイル (GetDoc_App.html)。このファイルには、ユーザー インターフェイスに加えて、JavaScript ファイル (office.js やホスト固有の .js ファイルなど) およびカスケード スタイル シート (CSS) ファイルへのリンクが格納されます。
    
  - JavaScript ファイル (GetDoc_App.js)。このファイルにはアドインのプログラミング ロジックが格納されます。
    
  - CSS ファイル (Program.css)。アドインのスタイルと書式を入れるファイルです。
    
- アドインの XML マニフェスト ファイル (GetDoc_App.xml)。共有ネットワーク フォルダーまたはアドイン カタログで使用できます。このマニフェスト ファイルでは、上述の HTML ファイルの場所を指していることが必要です。
    
また、Visual Studio 2015 を使用して、PowerPoint または Word 用アドインを作成できます。 


### <a name="core-concepts-to-know-for-creating-a-task-pane-add-in"></a>作業ウィンドウ アドインを作成するために知っておくべき主要な概念

この PowerPoint または Word 用アドインの作成を開始する前に、Office アドインの作成と HTTP 要求の操作についてよく理解しておくことが必要です。この記事では、Web サーバー上の HTTP 要求から Base64 エンコード テキストをデコードする方法については説明しません。 


## <a name="create-the-manifest-for-the-add-in"></a>アドインのマニフェストを作成する


PowerPoint 用アドインの XML マニフェスト ファイルは、アドインをホストできるアプリケーション、HTML ファイルの場所、アドインのタイトルと説明、およびその他の多くの特性に関する重要な情報を提供します。


- テキスト エディターで、次のコードをマニフェスト ファイルに追加します。
    
```XML
  
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

- このファイルを、UTF-8 エンコードを使用して、GetDoc_App.xml としてネットワークの場所またはアドイン カタログに保存します。
    

## <a name="create-the-user-interface-for-the-add-in"></a>アドインのユーザー インターフェイスを作成する


アドインのユーザー インターフェイスとしては、GetDoc_App.html ファイルに直接書き込んだ HTML を使用できます。このアドインのプログラミング ロジックと機能は、JavaScript ファイル (例: GetDoc_App.js) に入れる必要があります。

以下の手順を使用して、見出しと 1 つのボタンを含むアドインの簡単なユーザー インターフェイスを作成します。


- テキスト エディターで、新しいファイルに次の HTML を追加します。
    
```html  
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
        <title>Publish presentation</title>
        <link rel="stylesheet" type="text/css" href="Program.css" />
        <script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script>
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

- このファイルを、UTF-8 エンコードを使用して GetDoc_App.html としてネットワークの場所または Web サーバーに保存します。
    

 >**メモ:**  必ずアドインの **head** タグに **script** タグと office.js ファイルへの有効なリンクを入れてください。 

CSS を使用して、アドインに単純ながら最新の本格的な外観を与えます。次の CSS を使用して、アドインのスタイルを定義します。


- テキスト エディターで、新しいファイルに次の CSS を追加します。
    
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

- このファイルを、UTF-8 エンコードを使用して、Program.css としてネットワークの場所または Web サーバー (GetDoc_App.html を保存した場所) に保存します。
    

## <a name="add-the-javascript-to-get-the-document"></a>ドキュメントを取得するための JavaScript を追加する


アドインのコードでは、[Office.initialize](../../reference/shared/office.initialize.md) イベントのハンドラーが、フォーム上の **[送信]** ボタンのクリック イベントのハンドラーを追加し、アドインの準備ができたことをユーザーに知らせます。

次のコード例は、 **Office.initialize** イベントのイベント ハンドラーと、status div に書き込むためのヘルパー関数 `updateStatus` を示しています。




```js
// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {

      // After the DOM is loaded, add-in-specific code can run.
      document.getElementById('submit').addEventListener("click",
          function () {
              sendFile();
          });}
      updateStatus("Ready to send file.");
    });
}

// Create a function for writing to the status div. 
function updateStatus(message) {
    var statusInfo = document.getElementById("status");
    statusInfo.innerHTML += message + "<br/>";
}
```



UI の **[送信]** ボタンをクリックすると、アドインは `sendFile` 関数を呼び出します。この関数には、[Document.getFileAsync](../../reference/shared/document.getfileasync.md) メソッドの呼び出しが含まれています。**getFileAsync** メソッドでは、JavaScript API for Office の他のメソッドと同様に、非同期パターンを使用しています。このメソッドには、_fileType_ という 1 つの必須パラメーターと、_options_ と _callback_ という 2 つの省略可能なパラメーターがあります。 

_fileType_ パラメーターは、[FileType](../../reference/shared/filetype-enumeration.md) 列挙子の 3 つの定数 (**Office.FileType.Compressed** ("compressed")、**Office.FileType.PDF** ("pdf")、または **Office.FileType.Text** ("text")) のうちいずれかを受け付けます。PowerPoint は引数として **Compressed** のみをサポートし、Word は 3 つすべてをサポートしています。**fileType** パラメーターとして _Compressed_ を渡した場合、**getFileAsync** メソッドは、ローカル コンピューター上にファイルの一時コピーを作成することにより、ドキュメントを PowerPoint 2013 のプレゼンテーション ファイル (*.pptx) または Word 2013 のドキュメント ファイル (*.docx) として返します。

**getFileAsync** メソッドは、このファイルへの参照を [File](../../reference/shared/file.md) オブジェクトとして返します。**File** オブジェクトは、[size](../../reference/shared/file.size.md) プロパティ、[sliceCount](../../reference/shared/file.slicecount.md) プロパティ、[getSliceAsync](../../reference/shared/file.getsliceasync.md) メソッド、[closeAsync](../../reference/shared/file.closeasync.md) メソッドという 4 つのメンバーを公開します。**size** プロパティはファイルのバイト数を返します。**sliceCount** は、ファイル内の [Slice](../../reference/shared/document.md) オブジェクト (この記事で後述) の数を返します。

次のコードは、**document.getFileAsync()** メソッドを使用して、PowerPoint または Word ドキュメントを **File** オブジェクトとして取得します。次に、結果の **File** オブジェクト、ゼロになったカウンター、および [sliceCount](../../reference/shared/file.slicecount.md) を匿名オブジェクトにパッケージ化します。次に、このオブジェクトはローカルで定義された `getSlice` 関数に渡されます。 

```js
// Get all the content from a PowerPoint or Word document in 100-KB chunks of text.
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

                updateStatus("Getting file of " + myFile.size +
                    " bytes");

                getSlice(state);
            }
            else {
                updateStatus(result.status);
            }
    });
}
```

ローカル関数  `getSlice` は、 **File** オブジェクトからスライスを取得するために **File.getSliceAsync** メソッドの呼び出しを行います。 **getSliceAsync** メソッドは、スライスのコレクションから **Slice** オブジェクトを返します。このメソッドには、 _sliceIndex_ と _callback_ という 2 つの必須パラメーターがあります。 _sliceIndex_ パラメーターは、スライスのコレクションへのインデクサーとして整数を取ります。JavaScript API for Office の他の関数と同様、 **getSliceAsync** メソッドもメソッド呼び出しからの結果を処理するためにパラメーターとしてコールバック関数を取ります。

**Slice** オブジェクトにより、ファイルに含まれているデータにアクセスできます。_getFileAsync_ メソッドの **options** パラメーターで特に指定しない限り、**Slice** のサイズは 4 MB になります。**Slice** オブジェクトは、[size](../../reference/shared/slice.size.md)、[data](../../reference/shared/slice.data.md)、[index](../../reference/shared/slice.index.md) という 3 つのプロパティを公開します。**size** プロパティは、スライスのサイズ (バイト数) を取得します。**index** プロパティは、スライスのコレクション内でのそのスライスの位置を表す整数を取得します。




```js

// Get a slice from the file and then call sendSlice.
function getSlice(state) {

    state.file.getSliceAsync(state.counter, function (result) {
        if (result.status == Office.AsyncResultStatus.Succeeded) {

            updateStatus("Sending piece " + (state.counter + 1) +
                " of " + state.sliceCount);

            sendSlice(result.value, state);
        }
        else {
            updateStatus(result.status);
        }
    });
}
```

**Slice.data** プロパティは、ファイルの生データをバイト配列として返します。データがテキスト形式 (つまり、XML かプレーン テキスト) の場合、スライスには生テキストが含まれています。**Document.getFileAsync** の _fileType_ パラメーターとして **Office.FileType.Compressed** を渡した場合、スライスにはファイルのバイナリ データがバイト配列として含まれます。PowerPoint ファイルまたは Word ファイルの場合、スライスにはバイト配列が含まれます。

バイト配列のデータを Base64 でエンコードされた文字列に変換するには、独自の関数を実装 (または利用可能なライブラリを使用) する必要があります。JavaScript での Base64 エンコーティングについては、「 [Base64 エンコードとデコード](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding)」を参照してください。

データを Base64 に変換した後は、HTTP POST 要求の本体として送信するなど、そのデータをいろいろな方法で Web サーバーに送信できます。

スライスを Web サービスに送信するために次のコードを追加します。


 >**メモ**  このコードは、PowerPoint ファイルまたは Word ファイルを複数のスライスで Web サーバーに送信します。このファイルに対して何らかの操作を行うためには、Web サーバーまたは Web サービスによって、個々のスライスがコンパイルされて 1 つの .pptx ファイルに変換される必要があります。




```js

function sendSlice(slice, state) {
    var data = slice.data;

    // If the slice contains data, create an HTTP request.
    if (data) {

        // Encode the slice data, a byte array, as a Base64 string.
        // NOTE: The implementation of myEncodeBase64(input) function isn't 
        // included with this example. For information about Base64 encoding with
        // JavaScript, see https://developer.mozilla.org/en-US/docs/Web/JavaScript/Base64_encoding_and_decoding.
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



名前からもわかるように、 **File.closeAsync** メソッドはドキュメントへの接続を閉じて、リソースを解放します。Office アドインのサンドボックス ガベージはファイルへのスコープ外参照を収集しますが、コードでファイルを使い終わった後、それらのファイルを明示的に閉じるためのベスト プラクティスであることに変わりありません。 **closeAsync** メソッドには、 _callback_ という 1 つのパラメーターがあり、これによって呼び出しの完了時に呼び出す関数を指定します。




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
