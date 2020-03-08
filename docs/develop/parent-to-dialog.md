---
title: ホストページからダイアログボックスにデータとメッセージを渡す
description: MessageChild および DialogParentMessageReceived Api を使用してホストページからダイアログにデータを渡す方法について説明します。
ms.date: 03/11/2020
localization_priority: Normal
ms.openlocfilehash: 03d89a2e5ffb9060edb25dd8e0c3c71c0dd274eb
ms.sourcegitcommit: 153576b1efd0234c6252433e22db213238573534
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/07/2020
ms.locfileid: "42561874"
---
# <a name="passing-data-and-messages-to-a-dialog-box-from-its-host-page-preview"></a><span data-ttu-id="c73ef-103">ホストページからダイアログボックスにデータとメッセージを渡す (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="c73ef-103">Passing data and messages to a dialog box from its host page (preview)</span></span>

<span data-ttu-id="c73ef-104">アドインでは、 [dialog](/javascript/api/office/office.dialog)オブジェクトの[messageChild](/javascript/api/office/office.dialog#messagechild-message-)メソッドを使用して、[ホストページ](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)からダイアログボックスにメッセージを送信できます。</span><span class="sxs-lookup"><span data-stu-id="c73ef-104">Your add-in can send messages from the [host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) to a dialog box using the [messageChild](/javascript/api/office/office.dialog#messagechild-message-) method of the [Dialog](/javascript/api/office/office.dialog) object.</span></span>

> [!Important]
>
> - <span data-ttu-id="c73ef-105">この記事で説明する Api はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="c73ef-105">The APIs described in this article are in preview.</span></span> <span data-ttu-id="c73ef-106">開発者は実験を行うことができます。ただし、運用アドインでは使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="c73ef-106">They are available to developers for experimentation; but should not be used in a production add-in.</span></span> <span data-ttu-id="c73ef-107">この API がリリースされるまでは、「次の操作を実行するには」で説明されている方法を使用して、運用アドインの[ダイアログボックスに情報を渡し](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)ます。</span><span class="sxs-lookup"><span data-stu-id="c73ef-107">Until this API is released, use the techniques described in [Pass information to the dialog box](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) for production add-ins.</span></span>
> - <span data-ttu-id="c73ef-108">この記事に記載されている Api には、Office 365 (サブスクリプション版の Office) が必要です。</span><span class="sxs-lookup"><span data-stu-id="c73ef-108">The APIs described in this article require Office 365 (the subscription version of Office).</span></span> <span data-ttu-id="c73ef-109">Insider チャネルからの最新の月次バージョンとビルドを使ってください。</span><span class="sxs-lookup"><span data-stu-id="c73ef-109">You should use the latest monthly version and build from the Insiders channel.</span></span> <span data-ttu-id="c73ef-110">このバージョンを入手するには、Office Insider への参加が必要です。</span><span class="sxs-lookup"><span data-stu-id="c73ef-110">You need to be an Office Insider to get this version.</span></span> <span data-ttu-id="c73ef-111">詳細については、「[Office Insider になる](https://products.office.com/office-insider?tab=tab-1)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c73ef-111">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span> <span data-ttu-id="c73ef-112">ビルドが生産半期チャネルに graduates されている場合、そのビルドではプレビュー機能のサポートがオフになっていることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="c73ef-112">Please note that when a build graduates to the production semi-annual channel, support for preview features is turned off for that build.</span></span>
> - <span data-ttu-id="c73ef-113">プレビューの初期段階では、Excel、PowerPoint、Word で Api がサポートされています。ただし、Outlook には含まれません。</span><span class="sxs-lookup"><span data-stu-id="c73ef-113">In the initial stage of the preview, the APIs are supported in Excel, PowerPoint, and Word; but not in Outlook.</span></span>
>
> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

## <a name="use-messagechild-from-the-host-page"></a><span data-ttu-id="c73ef-114">ホスト`messageChild()`ページからの使用</span><span class="sxs-lookup"><span data-stu-id="c73ef-114">Use `messageChild()` from the host page</span></span>

<span data-ttu-id="c73ef-115">Office ダイアログ API を呼び出してダイアログボックスを開くと、 [dialog](/javascript/api/office/office.dialog)オブジェクトが返されます。</span><span class="sxs-lookup"><span data-stu-id="c73ef-115">When you call the Office dialog API to open a dialog box, a [Dialog](/javascript/api/office/office.dialog) object is returned.</span></span> <span data-ttu-id="c73ef-116">オブジェクトは他のメソッドによって参照されるため、通常は[Displaydialogasync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-)メソッドよりも広いスコープがある変数に割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="c73ef-116">It should be assigned to a variable, which typically has greater scope than the [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) method because the object will be referenced by other methods.</span></span> <span data-ttu-id="c73ef-117">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="c73ef-117">The following is an example:</span></span>

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

<span data-ttu-id="c73ef-118">この`Dialog`オブジェクトには、すべての文字列または文字列データをダイアログボックスに送信する[messageChild](/javascript/api/office/office.dialog#messagechild-message-)メソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="c73ef-118">This `Dialog` object has a [messageChild](/javascript/api/office/office.dialog#messagechild-message-) method that sends any string, or stringified data, to the dialog box.</span></span> <span data-ttu-id="c73ef-119">これにより`DialogParentMessageReceived` 、ダイアログボックスでイベントが発生します。</span><span class="sxs-lookup"><span data-stu-id="c73ef-119">This raises a `DialogParentMessageReceived` event in the dialog box.</span></span> <span data-ttu-id="c73ef-120">コードでは、次のセクションに示すように、このイベントを処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c73ef-120">Your code should handle this event, as shown in the next section.</span></span>

<span data-ttu-id="c73ef-121">ダイアログの UI が現在アクティブなワークシートと関連付けられ、他のワークシートを基準としたワークシートの相対位置になるシナリオを考えてみます。</span><span class="sxs-lookup"><span data-stu-id="c73ef-121">Consider a scenario in which the UI of the dialog should correlate with the currently active worksheet and that worksheet's position relative to the other worksheets.</span></span> <span data-ttu-id="c73ef-122">次の例では`sheetPropertiesChanged` 、Excel ワークシートのプロパティをダイアログボックスに送信します。</span><span class="sxs-lookup"><span data-stu-id="c73ef-122">In the following example, `sheetPropertiesChanged` sends Excel worksheet properties to the dialog box.</span></span> <span data-ttu-id="c73ef-123">この例では、現在のワークシートの名前は "My Sheet" で、ブックの2番目のシートです。</span><span class="sxs-lookup"><span data-stu-id="c73ef-123">In this case the current worksheet is named "My Sheet" and it is the 2nd sheet in the workbook.</span></span> <span data-ttu-id="c73ef-124">データは、文字列のオブジェクトにカプセル化されるので、に`messageChild`渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="c73ef-124">The data is encapsulated in an object which is stringified so that it can be passed to `messageChild`.</span></span>

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

## <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a><span data-ttu-id="c73ef-125">ダイアログボックスで DialogParentMessageReceived を処理する</span><span class="sxs-lookup"><span data-stu-id="c73ef-125">Handle DialogParentMessageReceived in the dialog box</span></span>

<span data-ttu-id="c73ef-126">ダイアログボックスの JavaScript で、 `DialogParentMessageReceived`イベントのハンドラーを[UI. addhandler async](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-)メソッドに登録します。</span><span class="sxs-lookup"><span data-stu-id="c73ef-126">In the dialog box's JavaScript, register a handler for the `DialogParentMessageReceived` event with the [UI.addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) method.</span></span> <span data-ttu-id="c73ef-127">これは、通常、 [office. onReady または office の initialize メソッド](initialize-add-in.md)で行われます。</span><span class="sxs-lookup"><span data-stu-id="c73ef-127">This is typically done in the [Office.onReady or Office.initialize methods](initialize-add-in.md).</span></span> <span data-ttu-id="c73ef-128">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="c73ef-128">The following is an example:</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

<span data-ttu-id="c73ef-129">その後、 `onMessageFromParent`ハンドラーを定義します。</span><span class="sxs-lookup"><span data-stu-id="c73ef-129">Then, define the `onMessageFromParent` handler.</span></span> <span data-ttu-id="c73ef-130">次のコードでは、前のセクションの例を続行します。</span><span class="sxs-lookup"><span data-stu-id="c73ef-130">The following code continues the example from the preceding section.</span></span> <span data-ttu-id="c73ef-131">Office によってハンドラーに引数が渡され、引数`message`オブジェクトのプロパティにホストページの文字列が含まれていることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="c73ef-131">Note that Office passes an argument to the handler and that the `message` property of argument object contains the string from the host page.</span></span> <span data-ttu-id="c73ef-132">この例では、メッセージはオブジェクトに再変換、jQuery を使用して、新しいワークシート名に一致するダイアログのトップの見出しを設定しています。</span><span class="sxs-lookup"><span data-stu-id="c73ef-132">In this example, the message is reconverted to an object and jQuery is used to set the top heading of the dialog to match the new worksheet name.</span></span>

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

<span data-ttu-id="c73ef-133">ハンドラーが適切に登録されていることを確認することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="c73ef-133">It is a best practice to verify that your handler is properly registered.</span></span> <span data-ttu-id="c73ef-134">これを行うには、ハンドラーの登録が`addHandlerAsync`完了したときに実行されるメソッドにコールバックを渡します。</span><span class="sxs-lookup"><span data-stu-id="c73ef-134">You can do this by passing a callback to the `addHandlerAsync` method that runs when the attempt to register the handler completes.</span></span> <span data-ttu-id="c73ef-135">ハンドラーが正常に登録されなかった場合は、ハンドラーを使用して、エラーを記録または表示します。</span><span class="sxs-lookup"><span data-stu-id="c73ef-135">Use the handler to log or show an error if the handler was not successfully registered.</span></span> <span data-ttu-id="c73ef-136">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="c73ef-136">The following is an example.</span></span> <span data-ttu-id="c73ef-137">ここで`reportError`は、エラーを記録または表示する関数であることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="c73ef-137">Note that `reportError` is a function, not defined here, that logs or displays the error.</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

## <a name="conditional-messaging"></a><span data-ttu-id="c73ef-138">条件付きのメッセージング</span><span class="sxs-lookup"><span data-stu-id="c73ef-138">Conditional messaging</span></span>

<span data-ttu-id="c73ef-139">ホストページから複数`messageChild`の呼び出しを行うことはできますが、 `DialogParentMessageReceived`イベントのダイアログボックスにはハンドラーが1つしかないため、ハンドラーは異なるメッセージを区別するために条件付きロジックを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c73ef-139">Because you can make multiple `messageChild` calls from the host page, but you have only one handler in the dialog box for the `DialogParentMessageReceived` event, the handler must use conditional logic to distinguish different messages.</span></span> <span data-ttu-id="c73ef-140">[条件付き](dialog-api-in-office-add-ins.md#conditional-messaging)メッセージの説明に従って、ダイアログボックスがホストページにメッセージを送信しているときに、条件付きメッセージを構造化する方法で、これを正確に行うことができます。</span><span class="sxs-lookup"><span data-stu-id="c73ef-140">You can do this in a way that is precisely parallel to how you would structure conditional messaging when the dialog box is sending a message to the host page as described in [Conditional messaging](dialog-api-in-office-add-ins.md#conditional-messaging).</span></span>
