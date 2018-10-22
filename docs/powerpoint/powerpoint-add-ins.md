---
title: PowerPoint アドイン
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: 390497e74d4dc52b9d400f242850ab72bdb0eabc
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640079"
---
# <a name="powerpoint-add-ins"></a><span data-ttu-id="3c6dc-102">PowerPoint アドイン</span><span class="sxs-lookup"><span data-stu-id="3c6dc-102">PowerPoint add-ins</span></span>

<span data-ttu-id="3c6dc-103">PowerPoint のアドインを使って、Windows、iOS、Office Online、Mac などのプラットフォームでユーザーのプレゼンテーションのための魅力的なソリューションをビルドすることができます。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-103">You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iOS, Office Online, and Mac. You can create one of two types of add-ins:</span></span> <span data-ttu-id="3c6dc-104">PowerPoint のアドインの 2 つの種類を作成することができます。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-104">You can create two types of PowerPoint add-ins:</span></span>

- <span data-ttu-id="3c6dc-p102">**コンテンツ アドイン**を使うと、プレゼンテーションに HTML5 の動的コンテンツが追加されます。たとえば [PowerPoint のための LucidChart ダイアグラム](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) アドインでは、これを使って LucidChart からデッキにインタラクティブな図を挿入することができます。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-p102">Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) add-in, which you can use to inject an interactive diagram from LucidChart into your deck.</span></span>

- <span data-ttu-id="3c6dc-107">**作業ウィンドウ アドイン** を使えば、サービスを介して、参照情報を取り込んだり、プレゼンテーションにデータを挿入したりすることができます。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-107">Use **task pane add-ins** to bring in reference information or insert data into the slide via a service. For example, see the Shutterstock Images add-in, which you can use to add professional photos to your presentation.</span></span> <span data-ttu-id="3c6dc-108">たとえば [Shutterstock イメージ 像](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) アドインでは、これを使ってプロの写真をプレゼンテーションに追加することができます。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-108">Use task pane add-ins to bring in reference information or insert data into the slide via a service. For example, see the [Shutterstock Images](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) add-in, which you can use to add professional photos to your presentation.</span></span> 

## <a name="powerpoint-add-in-scenarios"></a><span data-ttu-id="3c6dc-109">PowerPoint アドインのシナリオ</span><span class="sxs-lookup"><span data-stu-id="3c6dc-109">PowerPoint add-in scenarios</span></span>

<span data-ttu-id="3c6dc-110">この記事でデモンストレーションするコード例では、PowerPoint のアドインの開発のための基本的なタスクをいくつか示します。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-110">The code examples in the article show you some basic tasks for developing content add-ins for PowerPoint.</span></span> <span data-ttu-id="3c6dc-111">次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-111">Please note the following in 2nd_ProjServ_12 Beta 2:</span></span>

- <span data-ttu-id="3c6dc-112">情報を表示するために、これらの例では Visual Studio の Office アドイン プロジェクト テンプレートに含まれる `app.showNotification` 関数を使用します。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-112">To display information, these examples use the `app.showNotification` function, which is included in the Visual Studio Office Add-ins project templates.</span></span> <span data-ttu-id="3c6dc-113">アドインの開発に Visual Studio を使用している場合、独自のコードで`showNotification` を置き換える必要があります。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-113">If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code.</span></span> 

- <span data-ttu-id="3c6dc-114">これらの例のうちいくつかでは、これらの関数の範囲を超えて宣言されている `Globals` も使用しています。   `var Globals = {activeViewHandler:0, firstSlideId:0};`</span><span class="sxs-lookup"><span data-stu-id="3c6dc-114">Several of these examples also use a `Globals` object that is declared beyond the scope of these functions as:   `var Globals = {activeViewHandler:0, firstSlideId:0};`</span></span>

- <span data-ttu-id="3c6dc-115">これらの例を使用するには、アドインプロジェクトが [ Office.js v1.1 以降のライブラリを参照する必要がありますしている必要があります](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-115">To use these examples, your add-in project must [reference Office.js v1.1 library or later](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a><span data-ttu-id="3c6dc-116">プレゼンテーションのアクティブ ビューを検出し、ActiveViewChanged イベントを処理します。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-116">Detect the presentation's active view and handle the ActiveViewChanged event</span></span>

<span data-ttu-id="3c6dc-117">コンテンツ アドインをビルドする場合は、プレゼンテーションのアクティブ ビューを取得して、 `Office.Initialize`  ハンドラーの一部として、`ActiveViewChanged`  イベントを処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-117">If you are building a content add-in, you will need to get the presentation's active view and handle the ActiveViewChanged event, as part of your Office.Initialize handler.</span></span> 

> [!NOTE]
> <span data-ttu-id="3c6dc-118">PowerPoint Online では、スライドショーモードが新しいセッションとして扱われるため、 [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) イベントは発生しません。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-118">In PowerPoint Online, the [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) event will never fire as Slide Show mode is treated as a new session. In this case, the add-in must fetch the active view on load, as noted below.</span></span> <span data-ttu-id="3c6dc-119">この場合、次のコードサンプルに示すように、アドインでロード時にアクティブ ビューをフェッチする必要があります。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-119">In this case, the add-in must fetch the active view on load, as shown in the following code sample.</span></span>

<span data-ttu-id="3c6dc-120">コードサンプルは次の通りです。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-120">In the following code sample:</span></span>

- <span data-ttu-id="3c6dc-121">`getActiveFileView` 関数は、プレゼンテーションの現在のビューが "編集" ビュー (**[標準]** や **[アウトライン表示]** などの、スライドを編集できるビュー) なのか "読み取り" ビュー (\*\* [スライド ショー] \*\* や \*\* [閲覧表示]\*\*) なのかを返す [ Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getactiveviewasync-options--callback-)  メソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-121">The getFileView function calls the Document.getActiveViewAsync method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as Normal or Outline View) or "read" (Slide Show or Reading View) view.</span></span>

- <span data-ttu-id="3c6dc-122"> `registerActiveViewChanged` 関数は、 [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) メソッドを呼び出して[Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) イベントのハンドラを登録します。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-122">The  `registerActiveViewChanged` function calls the [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) method to register a handler for the [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) event.</span></span> 


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

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a><span data-ttu-id="3c6dc-123">プレゼンテーション内の特定のスライドに移動します。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-123">Navigate to a particular slide in the presentation</span></span>

<span data-ttu-id="3c6dc-124">次のコードサンプルでは、​​`getSelectedRange`関数は [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) メソッドを呼び出して、`asyncResult.value` から返される JSON オブジェクトを取得します。そのオブジェクトには、**slides** という名前の配列が含まれています。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-124">In the following code sample, the `getSelectedRange` function calls the [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) method to get the JSON object returned by `asyncResult.value`, which contains an array named **slides**.</span></span> <span data-ttu-id="3c6dc-125"> *\*スライド** 配列には、選択した範囲のスライド（複数のスライドが選択されていない場合は現在のスライド）の ID、タイトル、およびインデックスが含まれます。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-125">The **slides** array contains the ids, titles, and indexes of selected range of slides (or of the current slide, if multiple slides are not selected).</span></span> <span data-ttu-id="3c6dc-126">また選択範囲内の最初のスライド ID をグローバル変数に保存します。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-126">It also saves the id of the first slide in the selected range to a global variable.</span></span>

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

<span data-ttu-id="3c6dc-127">次のコードサンプルでは、`goToFirstSlide` 関数は、 [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-) メソッドを呼び出して、先で示した`getSelectedRange` 関数で識別された最初のスライドに移動します。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-127">In the following code sample, the `goToFirstSlide` function calls the [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-) method to navigate to the first slide that was identified by the `getSelectedRange` function shown previously.</span></span>

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

## <a name="navigate-between-slides-in-the-presentation"></a><span data-ttu-id="3c6dc-128">プレゼンテーション内のスライド間を移動します。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-128">Navigate between slides in the presentation</span></span>

<span data-ttu-id="3c6dc-129">次のコードサンプルでは、`goToSlideByIndex` 関数は**Document.goToByIdAsync** メソッドを呼び出して、プレゼンテーション内の次のスライドに移動します。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-129">The  `goToSlideByIndex` function calls the **Document.goToByIdAsync** method to navigate to the next slide in the presentation.</span></span>

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

## <a name="get-the-url-of-the-presentation"></a><span data-ttu-id="3c6dc-130">プレゼンテーションの URL を取得する</span><span class="sxs-lookup"><span data-stu-id="3c6dc-130">Get the URL of the presentation</span></span>

<span data-ttu-id="3c6dc-131">次のコードサンプルでは、`getFileUrl`関数は [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-) メソッドを呼び出してプレゼンテーションファイルの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="3c6dc-131">The  `getFileUrl` function calls the [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-) method to get the URL of the presentation file.</span></span>

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



## <a name="see-also"></a><span data-ttu-id="3c6dc-132">関連項目</span><span class="sxs-lookup"><span data-stu-id="3c6dc-132">See also</span></span>
- [<span data-ttu-id="3c6dc-133">PowerPoint のコード サンプル</span><span class="sxs-lookup"><span data-stu-id="3c6dc-133">PowerPoint Code Samples</span></span>](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,PowerPoint)
- [<span data-ttu-id="3c6dc-134">コンテンツ アドインおよび作業ウィンドウ アドインでドキュメントごとにアドインの状態と設定を保存する方法</span><span class="sxs-lookup"><span data-stu-id="3c6dc-134">How to save add-in state and settings per document for content and task pane add-ins</span></span>](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [<span data-ttu-id="3c6dc-135">現在の選択肢にデータを読み書きする</span><span class="sxs-lookup"><span data-stu-id="3c6dc-135">Read and write data to the active selection in a document or spreadsheet</span></span>](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [<span data-ttu-id="3c6dc-136">PowerPoint や Word 用のアドインからドキュメント全体を取得する</span><span class="sxs-lookup"><span data-stu-id="3c6dc-136">Get the whole document from an add-in for PowerPoint or Word</span></span>](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [<span data-ttu-id="3c6dc-137">PowerPoint アドインでドキュメントのテーマを使用する</span><span class="sxs-lookup"><span data-stu-id="3c6dc-137">Use document themes in your PowerPoint add-ins</span></span>](use-document-themes-in-your-powerpoint-add-ins.md)
    
