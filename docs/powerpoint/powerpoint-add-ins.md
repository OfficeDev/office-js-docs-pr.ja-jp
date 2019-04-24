---
title: PowerPoint アドイン
description: ''
ms.date: 04/15/2019
localization_priority: Priority
ms.openlocfilehash: 6e518d0bfd37291e39ee17e96ded8debb183c19f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450913"
---
# <a name="powerpoint-add-ins"></a><span data-ttu-id="c865f-102">PowerPoint アドイン</span><span class="sxs-lookup"><span data-stu-id="c865f-102">PowerPoint add-ins</span></span>

<span data-ttu-id="c865f-103">PowerPoint のアドインを使って、Windows、iOS、Office Online、Mac などのプラットフォームでユーザーのプレゼンテーションのための魅力的なソリューションを構築することができます。</span><span class="sxs-lookup"><span data-stu-id="c865f-103">You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iOS, Office Online, and Mac.</span></span> <span data-ttu-id="c865f-104">次の 2 種類の PowerPoint アドインを作成できます。</span><span class="sxs-lookup"><span data-stu-id="c865f-104">You can create two types of PowerPoint add-ins:</span></span>

- <span data-ttu-id="c865f-p102">**コンテンツ アドイン**を使うと、プレゼンテーションに HTML5 の動的コンテンツが追加されます。たとえば [PowerPoint のための LucidChart ダイアグラム](https://appsource.microsoft.com/product/office/WA104380117) アドインでは、これを使って LucidChart からデッキにインタラクティブな図を挿入することができます。</span><span class="sxs-lookup"><span data-stu-id="c865f-p102">Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://appsource.microsoft.com/product/office/WA104380117) add-in, which you can use to inject an interactive diagram from LucidChart into your deck.</span></span>

- <span data-ttu-id="c865f-107">**作業ウィンドウ アドイン**を使えば、サービスを介して、参照情報を取り込んだり、プレゼンテーションにデータを挿入したりすることができます。</span><span class="sxs-lookup"><span data-stu-id="c865f-107">Use **task pane add-ins** to bring in reference information or insert data into the presentation via a service.</span></span> <span data-ttu-id="c865f-108">たとえば [Pixton コミック キャラクター](https://appsource.microsoft.com/product/office/WA104380907) アドインでは、これを使ってプロの写真をプレゼンテーションに追加することができます。</span><span class="sxs-lookup"><span data-stu-id="c865f-108">For example, see the [Pixton Comic Characters](https://appsource.microsoft.com/product/office/WA104380907) add-in, which you can use to add professional photos to your presentation.</span></span> 

## <a name="powerpoint-add-in-scenarios"></a><span data-ttu-id="c865f-109">PowerPoint アドインのシナリオ</span><span class="sxs-lookup"><span data-stu-id="c865f-109">PowerPoint add-in scenarios</span></span>

<span data-ttu-id="c865f-110">この記事で紹介するコード例では、PowerPoint のアドインの開発のための基本的なタスクをいくつか示します。</span><span class="sxs-lookup"><span data-stu-id="c865f-110">The code examples in this article demonstrate some basic tasks for developing add-ins for PowerPoint.</span></span> <span data-ttu-id="c865f-111">以下のことに注意してください。</span><span class="sxs-lookup"><span data-stu-id="c865f-111">Please note the following:</span></span>

- <span data-ttu-id="c865f-112">情報を表示するために、これらの例は `app.showNotification` 関数を使用します。これは、Visual Studio の Office アドイン プロジェクト テンプレートに含まれています。</span><span class="sxs-lookup"><span data-stu-id="c865f-112">To display information, these examples use the `app.showNotification` function, which is included in the Visual Studio Office Add-ins project templates.</span></span> <span data-ttu-id="c865f-113">アドインを開発するのに Visual Studio を使用していない場合は、`showNotification` 関数を独自のコードに置き換える必要があります。</span><span class="sxs-lookup"><span data-stu-id="c865f-113">If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code.</span></span> 

- <span data-ttu-id="c865f-114">これらの例のうちいくつかでは、これらの関数 `var Globals = {activeViewHandler:0, firstSlideId:0};` の範囲を超えて宣言されている `Globals` オブジェクトも使用しています。</span><span class="sxs-lookup"><span data-stu-id="c865f-114">Several of these examples also use a `Globals` object that is declared beyond the scope of these functions as:   `var Globals = {activeViewHandler:0, firstSlideId:0};`</span></span>

- <span data-ttu-id="c865f-115">これらの例を使用するには、アドイン プロジェクトで [Office.js v1.1 ライブラリ以降を参照](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c865f-115">To use these examples, your add-in project must [reference Office.js v1.1 library or later](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a><span data-ttu-id="c865f-116">プレゼンテーションのアクティブ ビューの検出と ActiveViewChanged イベントの処理を行う</span><span class="sxs-lookup"><span data-stu-id="c865f-116">Detect the presentation's active view and handle the ActiveViewChanged event</span></span>

<span data-ttu-id="c865f-117">コンテンツ アドインをビルドする場合は、プレゼンテーションのアクティブ ビューを取得して、`Office.Initialize` ハンドラーの一部として、`ActiveViewChanged` イベントを処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c865f-117">If you are building a content add-in, you will need to get the presentation's active view and handle the `ActiveViewChanged` event, as part of your `Office.Initialize` handler.</span></span>

> [!NOTE]
> <span data-ttu-id="c865f-118">PowerPoint Online では [Document.ActiveViewChanged](/javascript/api/office/office.document) イベントは、スライド ショー モードが新しいセッションとして扱われるようには起動しません。</span><span class="sxs-lookup"><span data-stu-id="c865f-118">In PowerPoint Online, the [Document.ActiveViewChanged](/javascript/api/office/office.document) event will never fire as Slide Show mode is treated as a new session.</span></span> <span data-ttu-id="c865f-119">この場合、次のコード サンプルに示すように、アドインで読み込むアクティブ ビューをフェッチする必要があります。</span><span class="sxs-lookup"><span data-stu-id="c865f-119">In this case, the add-in must fetch the active view on load, as shown in the following code sample.</span></span>

<span data-ttu-id="c865f-120">コード サンプルは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c865f-120">In the following code sample:</span></span>

- <span data-ttu-id="c865f-121">`getActiveFileView` 関数は [Document.getActiveViewAsync](/javascript/api/office/office.document#getactiveviewasync-options--callback-) メソッドを呼び出して、プレゼンテーションの現在のビューが "編集" ビュー (**[標準]** や **[アウトライン表示]** などの、スライドを編集できるビュー) なのか "読み取り" ビュー (**[スライド ショー]** や **[閲覧表示]**) なのかを返します。</span><span class="sxs-lookup"><span data-stu-id="c865f-121">The  `getActiveFileView` function calls the [Document.getActiveViewAsync](/javascript/api/office/office.document#getactiveviewasync-options--callback-) method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as **Normal** or **Outline View**) or "read" (**Slide Show** or **Reading View**).</span></span>

- <span data-ttu-id="c865f-122">`registerActiveViewChanged` 関数は、[Document.ActiveViewChanged](/javascript/api/office/office.document) イベントのハンドラーを登録するための [addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) メソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="c865f-122">The  `registerActiveViewChanged` function calls the [addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) method to register a handler for the [Document.ActiveViewChanged](/javascript/api/office/office.document) event.</span></span>


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

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a><span data-ttu-id="c865f-123">プレゼンテーションの特定のスライドに移動する</span><span class="sxs-lookup"><span data-stu-id="c865f-123">Navigate to a particular slide in the presentation</span></span>

<span data-ttu-id="c865f-124">次のコード サンプルでは、`getSelectedRange` 関数は [Document.getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) メソッドを呼び出して、`asyncResult.value` によって返される JSON オブジェクトを取得します。このオブジェクトには、**slides** という名前の配列が含まれます。</span><span class="sxs-lookup"><span data-stu-id="c865f-124">In the following code sample, the `getSelectedRange` function calls the [Document.getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) method to get the JSON object returned by `asyncResult.value`, which contains an array named **slides**.</span></span> <span data-ttu-id="c865f-125">**slides** 配列には、選択した範囲のスライド (複数のスライドが選択されていない場合は現在のスライド) の ID、タイトル、およびインデックスが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c865f-125">The **slides** array contains the ids, titles, and indexes of selected range of slides (or of the current slide, if multiple slides are not selected).</span></span> <span data-ttu-id="c865f-126">また、選択範囲内の最初のスライド ID をグローバル変数に保存します。</span><span class="sxs-lookup"><span data-stu-id="c865f-126">It also saves the id of the first slide in the selected range to a global variable.</span></span>

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

<span data-ttu-id="c865f-127">次のコード サンプルでは、`goToFirstSlide` 関数は [Document.goToByIdAsync](/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-) メソッドを呼び出して、前に示した `getSelectedRange` 関数で識別された最初のスライドに移動します。</span><span class="sxs-lookup"><span data-stu-id="c865f-127">In the following code sample, the `goToFirstSlide` function calls the [Document.goToByIdAsync](/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-) method to navigate to the first slide that was identified by the `getSelectedRange` function shown previously.</span></span>

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

## <a name="navigate-between-slides-in-the-presentation"></a><span data-ttu-id="c865f-128">プレゼンテーション内のスライド間を移動する</span><span class="sxs-lookup"><span data-stu-id="c865f-128">Navigate between slides in the presentation</span></span>

<span data-ttu-id="c865f-129">次のコード サンプルでは、`goToSlideByIndex` 関数は **Document.goToByIdAsync** メソッドを呼び出して、プレゼンテーションの次のスライドに移動します。</span><span class="sxs-lookup"><span data-stu-id="c865f-129">In the following code sample, the `goToSlideByIndex` function calls the **Document.goToByIdAsync** method to navigate to the next slide in the presentation.</span></span>

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

## <a name="get-the-url-of-the-presentation"></a><span data-ttu-id="c865f-130">プレゼンテーションの URL を取得する</span><span class="sxs-lookup"><span data-stu-id="c865f-130">Get the URL of the presentation</span></span>

<span data-ttu-id="c865f-131">次のコード サンプルでは、`getFileUrl` 関数は [Document.getFileProperties](/javascript/api/office/office.document#getfilepropertiesasync-options--callback-) メソッドを呼び出して、プレゼンテーション ファイルの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="c865f-131">In the following code sample, the  `getFileUrl` function calls the [Document.getFileProperties](/javascript/api/office/office.document#getfilepropertiesasync-options--callback-) method to get the URL of the presentation file.</span></span>

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

## <a name="create-a-presentation"></a><span data-ttu-id="c865f-132">プレゼンテーションの作成</span><span class="sxs-lookup"><span data-stu-id="c865f-132">Create a presentation</span></span>

<span data-ttu-id="c865f-133">アドインでは、アドインが現在実行されている PowerPoint のインスタンスとは異なる新しいプレゼンテーションを作成できます。</span><span class="sxs-lookup"><span data-stu-id="c865f-133">Your add-in can create a new presentation, separate from the PowerPoint instance in which the add-in is currently running.</span></span> <span data-ttu-id="c865f-134">PowerPoint の名前空間には、この目的のための `createPresentation` メソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="c865f-134">The PowerPoint namespace has the `createPresentation` method for this purpose.</span></span> <span data-ttu-id="c865f-135">このメソッドが呼び出されると、新しいプレゼンテーションが PowerPoint の新しいインスタンスですぐに開いて表示されます。</span><span class="sxs-lookup"><span data-stu-id="c865f-135">When this method is called, the new presentation is immediately opened and displayed in a new instance of PowerPoint.</span></span> <span data-ttu-id="c865f-136">アドインは前のプレゼンテーションで開いて実行されたままになります。</span><span class="sxs-lookup"><span data-stu-id="c865f-136">Your add-in remains open and running with the previous presentation.</span></span>

```js
PowerPoint.createPresentation();
```

<span data-ttu-id="c865f-137">`createPresentation` メソッドでは既存のプレゼンテーションのコピーの作成もできます。</span><span class="sxs-lookup"><span data-stu-id="c865f-137">The `createPresentation` method can also create a copy of an existing presentation.</span></span> <span data-ttu-id="c865f-138">このメソッドは、オプションのパラメーターとして .pptx ファイルの base64 エンコード文字列表現を受け取ります。</span><span class="sxs-lookup"><span data-stu-id="c865f-138">The method accepts a base64-encoded string representation of an .pptx file as an optional parameter.</span></span> <span data-ttu-id="c865f-139">文字列の引数は有効な .pptx ファイルと見なされ、作成されるプレゼンテーションはそのファイルのコピーになります。</span><span class="sxs-lookup"><span data-stu-id="c865f-139">The resulting presentation will be a copy of that file, assuming the string argument is a valid .pptx file.</span></span> <span data-ttu-id="c865f-140">次の例に示すように、[FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) クラスを使用して、ファイルを必要な base64 エンコード文字列に変換できます。</span><span class="sxs-lookup"><span data-stu-id="c865f-140">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = function (event) {
    // strip off the metadata before the base64-encoded string
    var startIndex = event.target.result.indexOf("base64,");
    var copyBase64 = event.target.result.substr(startIndex + 7);

    PowerPoint.createPresentation(copyBase64);
};

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="see-also"></a><span data-ttu-id="c865f-141">関連項目</span><span class="sxs-lookup"><span data-stu-id="c865f-141">See also</span></span>

- [<span data-ttu-id="c865f-142">PowerPoint のコード サンプル</span><span class="sxs-lookup"><span data-stu-id="c865f-142">PowerPoint Code Samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,PowerPoint)
- [<span data-ttu-id="c865f-143">コンテンツ アドインおよび作業ウィンドウ アドインで、ドキュメントごとにアドインの状態と設定を保存する方法</span><span class="sxs-lookup"><span data-stu-id="c865f-143">How to save add-in state and settings per document for content and task pane add-ins</span></span>](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [<span data-ttu-id="c865f-144">ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み</span><span class="sxs-lookup"><span data-stu-id="c865f-144">Read and write data to the active selection in a document or spreadsheet</span></span>](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [<span data-ttu-id="c865f-145">PowerPoint や Word 用のアドインからドキュメント全体を取得する</span><span class="sxs-lookup"><span data-stu-id="c865f-145">Get the whole document from an add-in for PowerPoint or Word</span></span>](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [<span data-ttu-id="c865f-146">PowerPoint アドインでドキュメントのテーマを使用する</span><span class="sxs-lookup"><span data-stu-id="c865f-146">Use document themes in your PowerPoint add-ins</span></span>](use-document-themes-in-your-powerpoint-add-ins.md)