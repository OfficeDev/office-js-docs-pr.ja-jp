---
title: マニフェスト ファイルの FunctionFile 要素
description: UI を表示するのではなく、JavaScript 関数を実行するアドインコマンドを使用して、アドインが公開する操作のソースコードファイルを指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: db447a904c04d07d51119f1eac2556af536a647c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611842"
---
# <a name="functionfile-element"></a><span data-ttu-id="64772-103">FunctionFile 要素</span><span class="sxs-lookup"><span data-stu-id="64772-103">FunctionFile element</span></span>

<span data-ttu-id="64772-104">UI を表示するのではなく、JavaScript 関数を実行するアドインコマンドを使用して、アドインが公開する操作のソースコードファイルを指定します。</span><span class="sxs-lookup"><span data-stu-id="64772-104">Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI.</span></span> <span data-ttu-id="64772-105">`FunctionFile`要素は[Desktopformfactor](desktopformfactor.md)または[MobileFormFactor](mobileformfactor.md)の子要素です。</span><span class="sxs-lookup"><span data-stu-id="64772-105">The `FunctionFile` element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> <span data-ttu-id="64772-106">要素 `resid` の属性は、要素内の要素 `FunctionFile` の属性の値に設定されています。この要素には、 `id` `Url` `Resources` [Control 要素](control.md)で定義されているように、UI に含まれないアドインコマンドボタンによって使用されるすべての JavaScript 関数を含む HTML ファイルへの URL が含まれています。</span><span class="sxs-lookup"><span data-stu-id="64772-106">The `resid` attribute of the `FunctionFile` element is set to the value of the `id` attribute of a `Url` element in the `Resources` element that contains the URL to an HTML file that contains or loads all  the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="64772-107">要素の例を次に示し `FunctionFile` ます。</span><span class="sxs-lookup"><span data-stu-id="64772-107">The following is an example of the `FunctionFile` element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="64772-108">要素によって示される HTML ファイル内の JavaScript は、 `FunctionFile` `Office.initialize` 1 つのパラメーターを取る名前付き関数を呼び出して定義する必要があり `event` ます。</span><span class="sxs-lookup"><span data-stu-id="64772-108">The JavaScript in the HTML file indicated by the `FunctionFile` element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="64772-109">ユーザーに進捗状況や、成功か失敗かを通知するには、この関数で `item.notificationMessages` API を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="64772-109">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="64772-110">実行が終了したときに、`event.completed` を呼び出す必要もあります。</span><span class="sxs-lookup"><span data-stu-id="64772-110">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="64772-111">関数の名前は、UI のないボタンの要素として使用され `FunctionName` ます。</span><span class="sxs-lookup"><span data-stu-id="64772-111">The name of the functions are used in the `FunctionName` element for UI-less buttons.</span></span>

<span data-ttu-id="64772-112">関数を定義する HTML ファイルの例を次に示し `trackMessage` ます。</span><span class="sxs-lookup"><span data-stu-id="64772-112">The following is an example of an HTML file defining a `trackMessage` function.</span></span>

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

<span data-ttu-id="64772-113">次のコードは、で使用される関数を実装する方法を示して `FunctionName` います。</span><span class="sxs-lookup"><span data-stu-id="64772-113">The following code shows how to implement the function used by `FunctionName`.</span></span>

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
> <span data-ttu-id="64772-114">この呼び出しは、 `event.completed` イベントが正常に処理されたことを通知します。</span><span class="sxs-lookup"><span data-stu-id="64772-114">The call to `event.completed` signals that you have successfully handled the event.</span></span> <span data-ttu-id="64772-115">同一のアドイン コマンドを複数回クリックするなど、関数を複数回呼び出すと、すべてのイベントが自動的にキューに入れられます。</span><span class="sxs-lookup"><span data-stu-id="64772-115">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="64772-116">最初のイベントが自動的に実行され、その他のイベントはキューに残ります。</span><span class="sxs-lookup"><span data-stu-id="64772-116">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="64772-117">関数が呼び出されると `event.completed` 、次にキューに入れられた関数の呼び出しが実行されます。</span><span class="sxs-lookup"><span data-stu-id="64772-117">When your function calls `event.completed`, the next queued call to that function runs.</span></span> <span data-ttu-id="64772-118">を呼び出す必要があり `event.completed` ます。それ以外の場合、関数は実行されません。</span><span class="sxs-lookup"><span data-stu-id="64772-118">You must call `event.completed`; otherwise your function will not run.</span></span>
