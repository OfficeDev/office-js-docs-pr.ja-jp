---
title: PowerPoint アドイン
description: PowerPoint アドインを使用して Windows、iPad、Mac、ブラウザー上など、複数のプラットフォームでプレゼンテーション用の魅力的なソリューションを構築する方法を説明します。
ms.date: 10/14/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 476f8f34bc47d85842d5b31f8a0298bf2d5d7b18
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/23/2020
ms.locfileid: "48740841"
---
# <a name="powerpoint-add-ins"></a><span data-ttu-id="c1082-103">PowerPoint アドイン</span><span class="sxs-lookup"><span data-stu-id="c1082-103">PowerPoint add-ins</span></span>

<span data-ttu-id="c1082-104">PowerPoint アドインを使用して Windows、iPad、Mac、およびブラウザー上など、複数のプラットフォームでのユーザーのプレゼンテーション用に魅力的なソリューションを構築できます。</span><span class="sxs-lookup"><span data-stu-id="c1082-104">You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iPad, Mac, and in a browser.</span></span> <span data-ttu-id="c1082-105">次の 2 種類の PowerPoint アドインを作成できます。</span><span class="sxs-lookup"><span data-stu-id="c1082-105">You can create two types of PowerPoint add-ins:</span></span>

- <span data-ttu-id="c1082-p102">**コンテンツ アドイン**を使うと、プレゼンテーションに HTML5 の動的コンテンツが追加されます。たとえば [PowerPoint のための LucidChart ダイアグラム](https://appsource.microsoft.com/product/office/wa104380117) アドインでは、これを使って LucidChart からデッキにインタラクティブな図を挿入することができます。</span><span class="sxs-lookup"><span data-stu-id="c1082-p102">Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://appsource.microsoft.com/product/office/wa104380117) add-in, which you can use to inject an interactive diagram from LucidChart into your deck.</span></span>

- <span data-ttu-id="c1082-108">**作業ウィンドウ アドイン**を使えば、サービスを介して、参照情報を取り込んだり、プレゼンテーションにデータを挿入したりすることができます。</span><span class="sxs-lookup"><span data-stu-id="c1082-108">Use **task pane add-ins** to bring in reference information or insert data into the presentation via a service.</span></span> <span data-ttu-id="c1082-109">たとえば [Pexels - 無料ストックフォト](https://appsource.microsoft.com/product/office/wa104379997) アドインでは、これを使ってプロの写真をプレゼンテーションに追加することができます。</span><span class="sxs-lookup"><span data-stu-id="c1082-109">For example, see the [Pexels - Free Stock Photos](https://appsource.microsoft.com/product/office/wa104379997) add-in, which you can use to add professional photos to your presentation.</span></span>

## <a name="powerpoint-add-in-scenarios"></a><span data-ttu-id="c1082-110">PowerPoint アドインのシナリオ</span><span class="sxs-lookup"><span data-stu-id="c1082-110">PowerPoint add-in scenarios</span></span>

<span data-ttu-id="c1082-111">この記事で紹介するコード例では、PowerPoint のアドインの開発のための基本的なタスクをいくつか示します。</span><span class="sxs-lookup"><span data-stu-id="c1082-111">The code examples in this article demonstrate some basic tasks for developing add-ins for PowerPoint.</span></span> <span data-ttu-id="c1082-112">以下のことに注意してください。</span><span class="sxs-lookup"><span data-stu-id="c1082-112">Please note the following:</span></span>

- <span data-ttu-id="c1082-113">情報を表示するために、これらの例は `app.showNotification` 関数を使用します。これは、Visual Studio の Office アドイン プロジェクト テンプレートに含まれています。</span><span class="sxs-lookup"><span data-stu-id="c1082-113">To display information, these examples use the `app.showNotification` function, which is included in the Visual Studio Office Add-ins project templates.</span></span> <span data-ttu-id="c1082-114">アドインを開発するのに Visual Studio を使用していない場合は、`showNotification` 関数を独自のコードに置き換える必要があります。</span><span class="sxs-lookup"><span data-stu-id="c1082-114">If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code.</span></span>

- <span data-ttu-id="c1082-115">これらの例のうちいくつかでは、これらの関数 `var Globals = {activeViewHandler:0, firstSlideId:0};` の範囲を超えて宣言されている `Globals` オブジェクトも使用しています。</span><span class="sxs-lookup"><span data-stu-id="c1082-115">Several of these examples also use a `Globals` object that is declared beyond the scope of these functions as:   `var Globals = {activeViewHandler:0, firstSlideId:0};`</span></span>

- <span data-ttu-id="c1082-116">これらの例を使用するには、アドイン プロジェクトで [Office.js v1.1 ライブラリ以降を参照](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c1082-116">To use these examples, your add-in project must [reference Office.js v1.1 library or later](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a><span data-ttu-id="c1082-117">プレゼンテーションのアクティブ ビューの検出と ActiveViewChanged イベントの処理を行う</span><span class="sxs-lookup"><span data-stu-id="c1082-117">Detect the presentation's active view and handle the ActiveViewChanged event</span></span>

<span data-ttu-id="c1082-118">コンテンツ アドインをビルドする場合は、プレゼンテーションのアクティブ ビューを取得して、`Office.Initialize` ハンドラーの一部として、`ActiveViewChanged` イベントを処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c1082-118">If you are building a content add-in, you will need to get the presentation's active view and handle the `ActiveViewChanged` event, as part of your `Office.Initialize` handler.</span></span>

> [!NOTE]
> <span data-ttu-id="c1082-119">PowerPoint on the web では [Document.ActiveViewChanged](/javascript/api/office/office.document) イベントは、スライド ショー モードが新しいセッションとして扱われるようには起動しません。</span><span class="sxs-lookup"><span data-stu-id="c1082-119">In PowerPoint on the web, the [Document.ActiveViewChanged](/javascript/api/office/office.document) event will never fire as Slide Show mode is treated as a new session.</span></span> <span data-ttu-id="c1082-120">この場合、次のコード サンプルに示すように、アドインで読み込むアクティブ ビューをフェッチする必要があります。</span><span class="sxs-lookup"><span data-stu-id="c1082-120">In this case, the add-in must fetch the active view on load, as shown in the following code sample.</span></span>

<span data-ttu-id="c1082-121">コード サンプルは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c1082-121">In the following code sample:</span></span>

- <span data-ttu-id="c1082-122">`getActiveFileView` 関数は [Document.getActiveViewAsync](/javascript/api/office/office.document#getactiveviewasync-options--callback-) メソッドを呼び出して、プレゼンテーションの現在のビューが "編集" ビュー (**[標準]** や **[アウトライン表示]** などの、スライドを編集できるビュー) なのか "読み取り" ビュー (**[スライド ショー]** や **[閲覧表示]**) なのかを返します。</span><span class="sxs-lookup"><span data-stu-id="c1082-122">The  `getActiveFileView` function calls the [Document.getActiveViewAsync](/javascript/api/office/office.document#getactiveviewasync-options--callback-) method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as **Normal** or **Outline View**) or "read" (**Slide Show** or **Reading View**).</span></span>

- <span data-ttu-id="c1082-123">`registerActiveViewChanged` 関数は、[Document.ActiveViewChanged](/javascript/api/office/office.document) イベントのハンドラーを登録するための [addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) メソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="c1082-123">The  `registerActiveViewChanged` function calls the [addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) method to register a handler for the [Document.ActiveViewChanged](/javascript/api/office/office.document) event.</span></span>


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

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a><span data-ttu-id="c1082-124">プレゼンテーションの特定のスライドに移動する</span><span class="sxs-lookup"><span data-stu-id="c1082-124">Navigate to a particular slide in the presentation</span></span>

<span data-ttu-id="c1082-125">次のコード サンプルでは、`getSelectedRange` 関数は [Document.getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) メソッドを呼び出して、`asyncResult.value` によって返される JSON オブジェクトを取得します。このオブジェクトには、`slides` という名前の配列が含まれます。</span><span class="sxs-lookup"><span data-stu-id="c1082-125">In the following code sample, the `getSelectedRange` function calls the [Document.getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) method to get the JSON object returned by `asyncResult.value`, which contains an array named `slides`.</span></span> <span data-ttu-id="c1082-126">`slides`slides 配列には、選択した範囲のスライド (複数のスライドが選択されていない場合は現在のスライド) の ID、タイトル、およびインデックスが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c1082-126">The `slides` array contains the ids, titles, and indexes of selected range of slides (or of the current slide, if multiple slides are not selected).</span></span> <span data-ttu-id="c1082-127">また、選択範囲内の最初のスライド ID をグローバル変数に保存します。</span><span class="sxs-lookup"><span data-stu-id="c1082-127">It also saves the id of the first slide in the selected range to a global variable.</span></span>

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

<span data-ttu-id="c1082-128">次のコード サンプルでは、`goToFirstSlide` 関数は [Document.goToByIdAsync](/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-) メソッドを呼び出して、前に示した `getSelectedRange` 関数で識別された最初のスライドに移動します。</span><span class="sxs-lookup"><span data-stu-id="c1082-128">In the following code sample, the `goToFirstSlide` function calls the [Document.goToByIdAsync](/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-) method to navigate to the first slide that was identified by the `getSelectedRange` function shown previously.</span></span>

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

## <a name="navigate-between-slides-in-the-presentation"></a><span data-ttu-id="c1082-129">プレゼンテーション内のスライド間を移動する</span><span class="sxs-lookup"><span data-stu-id="c1082-129">Navigate between slides in the presentation</span></span>

<span data-ttu-id="c1082-130">次のコード サンプルでは、`goToSlideByIndex` 関数は `Document.goToByIdAsync` メソッドを呼び出して、プレゼンテーションの次のスライドに移動します。</span><span class="sxs-lookup"><span data-stu-id="c1082-130">In the following code sample, the `goToSlideByIndex` function calls the `Document.goToByIdAsync` method to navigate to the next slide in the presentation.</span></span>

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

## <a name="get-the-url-of-the-presentation"></a><span data-ttu-id="c1082-131">プレゼンテーションの URL を取得する</span><span class="sxs-lookup"><span data-stu-id="c1082-131">Get the URL of the presentation</span></span>

<span data-ttu-id="c1082-132">次のコード サンプルでは、`getFileUrl` 関数は [Document.getFileProperties](/javascript/api/office/office.document#getfilepropertiesasync-options--callback-) メソッドを呼び出して、プレゼンテーション ファイルの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="c1082-132">In the following code sample, the  `getFileUrl` function calls the [Document.getFileProperties](/javascript/api/office/office.document#getfilepropertiesasync-options--callback-) method to get the URL of the presentation file.</span></span>

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

## <a name="create-a-presentation"></a><span data-ttu-id="c1082-133">プレゼンテーションの作成</span><span class="sxs-lookup"><span data-stu-id="c1082-133">Create a presentation</span></span>

<span data-ttu-id="c1082-134">アドインでは、アドインが現在実行されている PowerPoint のインスタンスとは異なる新しいプレゼンテーションを作成できます。</span><span class="sxs-lookup"><span data-stu-id="c1082-134">Your add-in can create a new presentation, separate from the PowerPoint instance in which the add-in is currently running.</span></span> <span data-ttu-id="c1082-135">PowerPoint の名前空間には、この目的のための `createPresentation` メソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="c1082-135">The PowerPoint namespace has the `createPresentation` method for this purpose.</span></span> <span data-ttu-id="c1082-136">このメソッドが呼び出されると、新しいプレゼンテーションが PowerPoint の新しいインスタンスですぐに開いて表示されます。</span><span class="sxs-lookup"><span data-stu-id="c1082-136">When this method is called, the new presentation is immediately opened and displayed in a new instance of PowerPoint.</span></span> <span data-ttu-id="c1082-137">アドインは前のプレゼンテーションで開いて実行されたままになります。</span><span class="sxs-lookup"><span data-stu-id="c1082-137">Your add-in remains open and running with the previous presentation.</span></span>

```js
PowerPoint.createPresentation();
```

<span data-ttu-id="c1082-138">`createPresentation` メソッドでは既存のプレゼンテーションのコピーの作成もできます。</span><span class="sxs-lookup"><span data-stu-id="c1082-138">The `createPresentation` method can also create a copy of an existing presentation.</span></span> <span data-ttu-id="c1082-139">このメソッドは、オプションのパラメーターとして .pptx ファイルの base64 エンコード文字列表現を受け取ります。</span><span class="sxs-lookup"><span data-stu-id="c1082-139">The method accepts a base64-encoded string representation of an .pptx file as an optional parameter.</span></span> <span data-ttu-id="c1082-140">文字列の引数は有効な .pptx ファイルと見なされ、作成されるプレゼンテーションはそのファイルのコピーになります。</span><span class="sxs-lookup"><span data-stu-id="c1082-140">The resulting presentation will be a copy of that file, assuming the string argument is a valid .pptx file.</span></span> <span data-ttu-id="c1082-141">次の例に示すように、[FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) クラスを使用して、ファイルを必要な base64 エンコード文字列に変換できます。</span><span class="sxs-lookup"><span data-stu-id="c1082-141">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = function (event) {
    // strip off the metadata before the base64-encoded string
    var startIndex = reader.result.toString().indexOf("base64,");
    var copyBase64 = reader.result.toString().substr(startIndex + 7);

    PowerPoint.createPresentation(copyBase64);
};

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="see-also"></a><span data-ttu-id="c1082-142">関連項目</span><span class="sxs-lookup"><span data-stu-id="c1082-142">See also</span></span>

- [<span data-ttu-id="c1082-143">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="c1082-143">Developing Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="c1082-144">Microsoft 365 開発者プログラムについて</span><span class="sxs-lookup"><span data-stu-id="c1082-144">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
- [<span data-ttu-id="c1082-145">PowerPoint のコード サンプル</span><span class="sxs-lookup"><span data-stu-id="c1082-145">PowerPoint Code Samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,PowerPoint)
- [<span data-ttu-id="c1082-146">コンテンツ アドインおよび作業ウィンドウ アドインで、ドキュメントごとにアドインの状態と設定を保存する方法</span><span class="sxs-lookup"><span data-stu-id="c1082-146">How to save add-in state and settings per document for content and task pane add-ins</span></span>](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [<span data-ttu-id="c1082-147">ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み</span><span class="sxs-lookup"><span data-stu-id="c1082-147">Read and write data to the active selection in a document or spreadsheet</span></span>](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [<span data-ttu-id="c1082-148">PowerPoint や Word 用のアドインからドキュメント全体を取得する</span><span class="sxs-lookup"><span data-stu-id="c1082-148">Get the whole document from an add-in for PowerPoint or Word</span></span>](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [<span data-ttu-id="c1082-149">PowerPoint アドインでドキュメントのテーマを使用する</span><span class="sxs-lookup"><span data-stu-id="c1082-149">Use document themes in your PowerPoint add-ins</span></span>](use-document-themes-in-your-powerpoint-add-ins.md)
