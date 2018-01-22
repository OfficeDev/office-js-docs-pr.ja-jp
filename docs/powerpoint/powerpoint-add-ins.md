# <a name="powerpoint-add-ins"></a>PowerPoint アドイン

PowerPoint のアドインを使って、Windows、iOS、Office Online、Mac などのプラットフォームでユーザーのプレゼンテーションのための魅力的なソリューションをビルドすることができます。アドインの 2 種類のうちいずれかを作成できます:

- **コンテンツ アドイン**を使うと、プレゼンテーションに HTML5 の動的コンテンツが追加されます。たとえば [PowerPoint のための LucidChart ダイアグラム](https://store.office.com/en-us/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) アドインでは、これを使って LucidChart からデッキにインタラクティブな図を挿入することができます。
- **作業ウィンドウ アドイン**を使えば、サービスを介して、参照情報を取り込んだり、スライドにデータを挿入したりすることができます。たとえば [Shutterstock イメージ](https://store.office.com/en-us/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) アドインでは、これを使ってプロの写真をプレゼンテーションに追加することができます。 

>
  **注:**アドインをビルドするとき、アドインを Office ストアに[発行](../publish/publish.md)する予定であれば、[Office ストア検証ポリシー](https://msdn.microsoft.com/en-us/library/jj220035.aspx)に準拠していることを確認してください。たとえば、検証に合格するには、アドインは、定義したメソッドをサポートするすべてのプラットフォーム全体で機能する必要があります (詳細については、[セクション 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3) と「[Office アドインを使用できるホストおよびプラットフォーム](https://dev.office.com/add-in-availability)」のページを参照してください)。

## <a name="powerpoint-add-in-scenarios"></a>PowerPoint アドインのシナリオ

この記事で紹介するコード例では、PowerPoint のコンテンツ アドインの開発のための基本的なタスクをいくつか示します。 

情報を表示するために、これらの例は `app.showNotification` 関数に依存しています。これは、Visual Studio の Office アドイン プロジェクト テンプレートに含まれています。アドインを開発するのに Visual Studio を使用していない場合は、`showNotification` 関数を独自のコードに置き換える必要があります。これらの例のうちいくつかは、これらの関数 `var globals = {activeViewHandler:0, firstSlideId:0};` の範囲外で宣言されたこの `globals` オブジェクトにも依存しています。

これらのコード例では、プロジェクトが [Office.js v1.1 以降のライブラリを参照](../../docs/develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)している必要があります。


## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>プレゼンテーションのアクティブ ビューの検出と ActiveViewChanged イベントの処理を行う

コンテンツ アドインをビルドする場合は、プレゼンテーションのアクティブ ビューを取得して、Office.Initialize ハンドラーの一部として、ActiveViewChanged イベントを処理する必要があります。


- `getActiveFileView` 関数は [Document.getActiveViewAsync](http://dev.office.com/reference/add-ins/shared/document.getactiveviewasync) メソッドを呼び出して、プレゼンテーションの現在のビューが "編集" ビュー (**[標準]** や **[アウトライン表示]** などの、スライドを編集できるビュー) なのか "読み取り" ビュー (**[スライド ショー]** や **[閲覧表示]**) なのかを返します。


- `registerActiveViewChanged` 関数は、[Document.ActiveViewChanged](http://dev.office.com/reference/add-ins/shared/document.activeviewchanged) イベントのハンドラーを登録するための [addHandlerAsync](http://dev.office.com/reference/add-ins/shared/document.addhandlerasync) メソッドを呼び出します。 
> 注:PowerPoint Online では [Document.ActiveViewChanged](http://dev.office.com/reference/add-ins/shared/document.activeviewchanged) イベントは、スライド ショー モードが新しいセッションとして扱われるようには起動しません。この場合、下に示すように、アドインで読み込むアクティブ ビューをフェッチしなければなりません。



```js

//general Office.initialize function. Fires on load of the add-in.
Office.initialize = function(){

    //Gets whether the current view is edit or read.
    var currentView = getActiveFileView();

    //register for the active view changed handler
    registerActiveViewChanged();

    //render the content based off of the currentView
    //....
}

function getActiveFileView()
{
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification(asyncResult.value);
        }
    });

}


function registerActiveViewChanged() {
    Globals.activeViewHandler = function (args) {
        app.showNotification(JSON.stringify(args));
    }

    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, Globals.activeViewHandler, 
        function (asyncResult) {
            if (asyncResult.status == "failed") {
           app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
            else {
            app.showNotification(asyncResult.status);
            }
        });
}
```
    

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a>プレゼンテーションの特定のスライドに移動する

`getSelectedRange` 関数は [Document.getSelectedDataAsync](http://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) メソッドを呼び出して、`asyncResult.value` から返される JSON オブジェクトを取得します。そのオブジェクトには、選択範囲のスライド (または現在のスライドのみ) の ID、タイトル、インデックスが入った "slides" という名前の配列が含まれています。この関数はまた、選択範囲の最初のスライドの ID をグローバル変数に保存します。


```js
function getSelectedRange() {
    // Get the id, title, and index of the current slide (or selected slides) and store the first slide id */
    Globals.firstSlideId = 0;

    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            Globals.firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
}
```

`goToFirstSlide` 関数は [Document.goToByIdAsync](http://dev.office.com/reference/add-ins/shared/document.gotobyidasync) メソッドを呼び出して、上記の `getSelectedRange` 関数が格納した最初のスライドの ID に移動します。




```js
function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```


## <a name="navigate-between-slides-in-the-presentation"></a>プレゼンテーション内のスライド間を移動する

`goToSlideByIndex` 関数は **Document.goToByIdAsync** メソッドを呼び出して、プレゼンテーションの次のスライドに移動します。


```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```

## <a name="get-the-url-of-the-presentation"></a>プレゼンテーションの URL を取得する

`getFileUrl` 関数は [Document.getFileProperties](http://dev.office.com/reference/add-ins/shared/document.getfilepropertiesasync) メソッドを呼び出して、プレゼンテーション ファイルの URL を取得します。


```js
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            app.showNotification("The file hasn't been saved yet. Save the file and try again");
        }
        else {
            app.showNotification(fileUrl);
        }
    });
}
```



## <a name="additional-resources"></a>追加リソース
- [PowerPoint のコード サンプル](https://dev.office.com/code-samples#?filters=powerpoint)

- [コンテンツ アドインおよび作業ウィンドウ アドインで、ドキュメントごとにアドインの状態と設定を保存する方法](../../docs/develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)

- [ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りと書き込み](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
- [PowerPoint または Word 用のアドインからドキュメント全体を取得する](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)
    
- [PowerPoint アドインでドキュメントのテーマを使用する](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md)
    
