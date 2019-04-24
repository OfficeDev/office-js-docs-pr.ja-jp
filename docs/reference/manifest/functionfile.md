---
title: マニフェスト ファイルの FunctionFile 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 5f87d10428b58adfb89f1119ba5741599079afba
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450584"
---
# <a name="functionfile-element"></a><span data-ttu-id="7110f-102">FunctionFile 要素</span><span class="sxs-lookup"><span data-stu-id="7110f-102">FunctionFile element</span></span>

<span data-ttu-id="7110f-p101">UI を表示する代わりに JavaScript 関数を実行するアドイン コマンドによってアドインが公開する操作の、ソース コード ファイルを指定します。**FunctionFile** 要素は、[DesktopFormFactor](desktopformfactor.md) または [MobileFormFactor](mobileformfactor.md) の子要素です。**FunctionFile** 要素の **resid** 属性は、HTML ファイルの URL を含む **Resources** 要素内の **Url** 要素の **id** 属性値に設定されます。この HTML ファイルには、[Control 要素](control.md)の定義に従い、UI なしのアドイン コマンド ボタンに使用されるすべての JavaScript 関数が含まれるか、読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="7110f-p101">Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI. The  **FunctionFile** element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md). The **resid** attribute of the **FunctionFile** element is set to the value of the **id** attribute of a **Url** element in the **Resources** element that contains the URL to an HTML file that contains or loads all  the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="7110f-106">**FunctionFile** 要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="7110f-106">The following is an example of the  **FunctionFile** element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="7110f-107">**FunctionFile** 要素で示される HTML ファイルの JavaScript は、`Office.initialize` を呼び出し、1 つのパラメーター `event` を取る名前付き関数を定義する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7110f-107">The JavaScript in the HTML file indicated by the  **FunctionFile** element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="7110f-108">ユーザーに進捗状況や、成功か失敗かを通知するには、この関数で `item.notificationMessages` API を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7110f-108">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="7110f-109">実行が終了したときに、`event.completed` を呼び出す必要もあります。</span><span class="sxs-lookup"><span data-stu-id="7110f-109">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="7110f-110">関数の名前は、UI なしボタンの **FunctionName** 要素で使用されます。</span><span class="sxs-lookup"><span data-stu-id="7110f-110">The name of the functions are used in the **FunctionName** element for UI-less buttons.</span></span>

<span data-ttu-id="7110f-111">**trackMessage** 関数を定義する HTML ファイルの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="7110f-111">The following is an example of an HTML file defining a **trackMessage** function.</span></span>

```js
Office.initialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

<span data-ttu-id="7110f-112">次のコードは、**FunctionName** で使用される関数の実装方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="7110f-112">The following code shows how to implement the function used by  **FunctionName**.</span></span>

```js
// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

// Your function must be in the global namespace.
function writeText(event) {

    // Implement your custom code here. The following code is a simple example.

    Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed") {
                // Show error message.
            }
            else {
                // Show success message.
            }
        });
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
```

> [!IMPORTANT]
> <span data-ttu-id="7110f-113">**event.completed** シグナルに対する呼び出しにより、イベントが正常に処理されたことが通知されます。</span><span class="sxs-lookup"><span data-stu-id="7110f-113">The call to  **event.completed** signals that you have successfully handled the event.</span></span> <span data-ttu-id="7110f-114">同一のアドイン コマンドを複数回クリックするなどの方法で関数を複数回呼び出すと、すべてのイベントは自動的にキューに入れられます。</span><span class="sxs-lookup"><span data-stu-id="7110f-114">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="7110f-115">最初のイベントが自動的に実行され、その他のイベントはキューに残ります。</span><span class="sxs-lookup"><span data-stu-id="7110f-115">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="7110f-116">関数により **event.completed** が呼び出されると、キューに入れられている、その関数に対する次の呼び出しが実行されます。</span><span class="sxs-lookup"><span data-stu-id="7110f-116">When your function calls **event.completed**, the next queued call to that function runs.</span></span> <span data-ttu-id="7110f-117">そうしないと、関数は実行されません。</span><span class="sxs-lookup"><span data-stu-id="7110f-117">You must call **event.completed**; otherwise your function will not run.</span></span>
