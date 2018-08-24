---
title: PowerPoint アドイン
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: e5c605410601d711e28ca04ff6e26387019cbb41
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925319"
---
# <a name="powerpoint-add-ins"></a><span data-ttu-id="120b3-102">PowerPoint アドイン</span><span class="sxs-lookup"><span data-stu-id="120b3-102">PowerPoint add-ins</span></span>

<span data-ttu-id="120b3-p101">PowerPoint のアドインを使って、Windows、iOS、Office Online、Mac などのプラットフォームでユーザーのプレゼンテーションのための魅力的なソリューションをビルドすることができます。アドインの 2 種類のうちいずれかを作成できます:</span><span class="sxs-lookup"><span data-stu-id="120b3-p101">You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iOS, Office Online, and Mac. You can create one of two types of add-ins:</span></span>

- <span data-ttu-id="120b3-p102">**コンテンツ アドイン**を使うと、プレゼンテーションに HTML5 の動的コンテンツが追加されます。たとえば [PowerPoint のための LucidChart ダイアグラム](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) アドインでは、これを使って LucidChart からデッキにインタラクティブな図を挿入することができます。</span><span class="sxs-lookup"><span data-stu-id="120b3-p102">Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) add-in, which you can use to inject an interactive diagram from LucidChart into your deck.</span></span>
- <span data-ttu-id="120b3-p103">**作業ウィンドウ アドイン**を使えば、サービスを介して、参照情報を取り込んだり、スライドにデータを挿入したりすることができます。たとえば [Shutterstock イメージ](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) アドインでは、これを使ってプロの写真をプレゼンテーションに追加することができます。</span><span class="sxs-lookup"><span data-stu-id="120b3-p103">Use **task pane add-ins** to bring in reference information or insert data into the slide via a service. For example, see the [Shutterstock Images](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) add-in, which you can use to add professional photos to your presentation.</span></span> 

## <a name="powerpoint-add-in-scenarios"></a><span data-ttu-id="120b3-109">PowerPoint アドインのシナリオ</span><span class="sxs-lookup"><span data-stu-id="120b3-109">PowerPoint add-in scenarios</span></span>

<span data-ttu-id="120b3-110">この記事で紹介するコード例では、PowerPoint のコンテンツ アドインの開発のための基本的なタスクをいくつか示します。</span><span class="sxs-lookup"><span data-stu-id="120b3-110">The code examples in the article show you some basic tasks for developing content add-ins for PowerPoint.</span></span> 

<span data-ttu-id="120b3-p104">情報を表示するために、これらの例は `app.showNotification` 関数に依存しています。これは、Visual Studio の Office アドイン プロジェクト テンプレートに含まれています。アドインを開発するのに Visual Studio を使用していない場合は、`showNotification` 関数を独自のコードに置き換える必要があります。これらの例のうちいくつかは、これらの関数 `var globals = {activeViewHandler:0, firstSlideId:0};` の範囲外で宣言されたこの `globals` オブジェクトにも依存しています。</span><span class="sxs-lookup"><span data-stu-id="120b3-p104">To display information, these examples depend on the `app.showNotification` function, which is included in the Visual Studio Office Add-ins project templates. If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code. Several of these examples also depend on this `globals` object that is declared outside of the scope of these functions: `var globals = {activeViewHandler:0, firstSlideId:0};`</span></span>

<span data-ttu-id="120b3-114">これらのコード例では、プロジェクトが [Office.js v1.1 以降のライブラリを参照](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)している必要があります。</span><span class="sxs-lookup"><span data-stu-id="120b3-114">These code examples require your project to [reference Office.js v1.1 library or later](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>


## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a><span data-ttu-id="120b3-115">プレゼンテーションのアクティブ ビューの検出と ActiveViewChanged イベントの処理を行う</span><span class="sxs-lookup"><span data-stu-id="120b3-115">Detect the presentation's active view and handle the ActiveViewChanged event</span></span>

<span data-ttu-id="120b3-116">コンテンツ アドインをビルドする場合は、プレゼンテーションのアクティブ ビューを取得して、Office.Initialize ハンドラーの一部として、ActiveViewChanged イベントを処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="120b3-116">If you are building a content add-in, you will need to get the presentation's active view and handle the ActiveViewChanged event, as part of your Office.Initialize handler.</span></span>


- <span data-ttu-id="120b3-117">`getActiveFileView` 関数は [Document.getActiveViewAsync](https://dev.office.com/reference/add-ins/shared/document.getactiveviewasync) メソッドを呼び出して、プレゼンテーションの現在のビューが "編集" ビュー (**[標準]** や **[アウトライン表示]** などの、スライドを編集できるビュー) なのか "読み取り" ビュー (**[スライド ショー]** や **[閲覧表示]**) なのかを返します。</span><span class="sxs-lookup"><span data-stu-id="120b3-117">The  `getActiveFileView` function calls the [Document.getActiveViewAsync](https://dev.office.com/reference/add-ins/shared/document.getactiveviewasync) method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as **Normal** or **Outline View**) or "read" ( **Slide Show** or **Reading View**) view.</span></span>


- <span data-ttu-id="120b3-118">`registerActiveViewChanged` 関数は、[Document.ActiveViewChanged](https://dev.office.com/reference/add-ins/shared/document.activeviewchanged) イベントのハンドラーを登録するための [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/document.addhandlerasync) メソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="120b3-118">The  `registerActiveViewChanged` function calls the [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/document.addhandlerasync) method to register a handler for the [Document.ActiveViewChanged](https://dev.office.com/reference/add-ins/shared/document.activeviewchanged) event.</span></span> 

> [!NOTE]
> <span data-ttu-id="120b3-p105">PowerPoint Online では [Document.ActiveViewChanged](https://dev.office.com/reference/add-ins/shared/document.activeviewchanged) イベントは、スライド ショー モードが新しいセッションとして扱われるようには起動しません。この場合、下に示すように、アドインで読み込むアクティブ ビューをフェッチしなければなりません。</span><span class="sxs-lookup"><span data-stu-id="120b3-p105">In PowerPoint Online, the [Document.ActiveViewChanged](https://dev.office.com/reference/add-ins/shared/document.activeviewchanged) event will never fire as Slide Show mode is treated as a new session. In this case, the add-in must fetch the active view on load, as noted below.</span></span>

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
    

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a><span data-ttu-id="120b3-121">プレゼンテーションの特定のスライドに移動する</span><span class="sxs-lookup"><span data-stu-id="120b3-121">Navigate to a particular slide in the presentation</span></span>

<span data-ttu-id="120b3-p106">`getSelectedRange` 関数は [Document.getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) メソッドを呼び出して、`asyncResult.value` から返される JSON オブジェクトを取得します。そのオブジェクトには、選択範囲のスライド (または現在のスライドのみ) の ID、タイトル、インデックスが入った "slides" という名前の配列が含まれています。この関数はまた、選択範囲の最初のスライドの ID をグローバル変数に保存します。</span><span class="sxs-lookup"><span data-stu-id="120b3-p106">The  `getSelectedRange` function calls the [Document.getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) method to get a JSON object returned by `asyncResult.value`, which contains an array named "slides" that contains the ids, titles, and indexes of selected range of slides (or just the current slide). It also saves the id of the first slide in the selected range to a global variable.</span></span>


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

<span data-ttu-id="120b3-124">`goToFirstSlide` 関数は [Document.goToByIdAsync](https://dev.office.com/reference/add-ins/shared/document.gotobyidasync) メソッドを呼び出して、上記の `getSelectedRange` 関数が格納した最初のスライドの ID に移動します。</span><span class="sxs-lookup"><span data-stu-id="120b3-124">The  `goToFirstSlide` function calls the [Document.goToByIdAsync](https://dev.office.com/reference/add-ins/shared/document.gotobyidasync) method to go to the id of the first slide stored by the `getSelectedRange` function above.</span></span>




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


## <a name="navigate-between-slides-in-the-presentation"></a><span data-ttu-id="120b3-125">プレゼンテーション内のスライド間を移動する</span><span class="sxs-lookup"><span data-stu-id="120b3-125">Navigate between slides in the presentation</span></span>

<span data-ttu-id="120b3-126">`goToSlideByIndex` 関数は **Document.goToByIdAsync** メソッドを呼び出して、プレゼンテーションの次のスライドに移動します。</span><span class="sxs-lookup"><span data-stu-id="120b3-126">The  `goToSlideByIndex` function calls the **Document.goToByIdAsync** method to navigate to the next slide in the presentation.</span></span>


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

## <a name="get-the-url-of-the-presentation"></a><span data-ttu-id="120b3-127">プレゼンテーションの URL を取得する</span><span class="sxs-lookup"><span data-stu-id="120b3-127">Get the URL of the presentation</span></span>

<span data-ttu-id="120b3-128">`getFileUrl` 関数は [Document.getFileProperties](https://dev.office.com/reference/add-ins/shared/document.getfilepropertiesasync) メソッドを呼び出して、プレゼンテーション ファイルの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="120b3-128">The  `getFileUrl` function calls the [Document.getFileProperties](https://dev.office.com/reference/add-ins/shared/document.getfilepropertiesasync) method to get the URL of the presentation file.</span></span>


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



## <a name="see-also"></a><span data-ttu-id="120b3-129">関連項目</span><span class="sxs-lookup"><span data-stu-id="120b3-129">See also</span></span>
- [<span data-ttu-id="120b3-130">PowerPoint のコード サンプル</span><span class="sxs-lookup"><span data-stu-id="120b3-130">PowerPoint Code Samples</span></span>](https://dev.office.com/code-samples#?filters=powerpoint)
- [<span data-ttu-id="120b3-131">コンテンツ アドインおよび作業ウィンドウ アドインで、ドキュメントごとにアドインの状態と設定を保存する方法</span><span class="sxs-lookup"><span data-stu-id="120b3-131">How to save add-in state and settings per document for content and task pane add-ins</span></span>](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [<span data-ttu-id="120b3-132">ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み</span><span class="sxs-lookup"><span data-stu-id="120b3-132">Read and write data to the active selection in a document or spreadsheet</span></span>](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [<span data-ttu-id="120b3-133">PowerPoint や Word 用のアドインからドキュメント全体を取得する</span><span class="sxs-lookup"><span data-stu-id="120b3-133">Get the whole document from an add-in for PowerPoint or Word</span></span>](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [<span data-ttu-id="120b3-134">PowerPoint アドインでドキュメントのテーマを使用する</span><span class="sxs-lookup"><span data-stu-id="120b3-134">Use document themes in your PowerPoint add-ins</span></span>](use-document-themes-in-your-powerpoint-add-ins.md)
    
