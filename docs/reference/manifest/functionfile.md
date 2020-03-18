---
title: マニフェスト ファイルの FunctionFile 要素
description: UI を表示するのではなく、JavaScript 関数を実行するアドインコマンドを使用して、アドインが公開する操作のソースコードファイルを指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 376ea82f48360d502ea9be05dc5d6b02f9294add
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718196"
---
# <a name="functionfile-element"></a><span data-ttu-id="f1e54-103">FunctionFile 要素</span><span class="sxs-lookup"><span data-stu-id="f1e54-103">FunctionFile element</span></span>

<span data-ttu-id="f1e54-104">UI を表示するのではなく、JavaScript 関数を実行するアドインコマンドを使用して、アドインが公開する操作のソースコードファイルを指定します。</span><span class="sxs-lookup"><span data-stu-id="f1e54-104">Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI.</span></span> <span data-ttu-id="f1e54-105">要素は[Desktopformfactor](desktopformfactor.md)または MobileFormFactor の子要素です。 [MobileFormFactor](mobileformfactor.md) `FunctionFile`</span><span class="sxs-lookup"><span data-stu-id="f1e54-105">The `FunctionFile` element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> <span data-ttu-id="f1e54-106">要素`resid`の`FunctionFile`属性は、要素内の要素`id`の`Url`属性の値に設定されて`Resources`います。この要素には、 [CONTROL 要素](control.md)で定義されているように、UI に含まれないアドインコマンドボタンによって使用されるすべての JavaScript 関数を含む HTML ファイルへの URL が含まれています。</span><span class="sxs-lookup"><span data-stu-id="f1e54-106">The `resid` attribute of the `FunctionFile` element is set to the value of the `id` attribute of a `Url` element in the `Resources` element that contains the URL to an HTML file that contains or loads all  the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="f1e54-107">`FunctionFile`要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="f1e54-107">The following is an example of the `FunctionFile` element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="f1e54-108">`FunctionFile`要素によって示される HTML ファイル内の JavaScript は`Office.initialize` 、1つのパラメーターを取る名前付き関数`event`を呼び出して定義する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f1e54-108">The JavaScript in the HTML file indicated by the `FunctionFile` element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="f1e54-109">ユーザーに進捗状況や、成功か失敗かを通知するには、この関数で `item.notificationMessages` API を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f1e54-109">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="f1e54-110">実行が終了したときに、`event.completed` を呼び出す必要もあります。</span><span class="sxs-lookup"><span data-stu-id="f1e54-110">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="f1e54-111">関数の名前は、UI のないボタン`FunctionName`の要素として使用されます。</span><span class="sxs-lookup"><span data-stu-id="f1e54-111">The name of the functions are used in the `FunctionName` element for UI-less buttons.</span></span>

<span data-ttu-id="f1e54-112">関数を`trackMessage`定義する HTML ファイルの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="f1e54-112">The following is an example of an HTML file defining a `trackMessage` function.</span></span>

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

<span data-ttu-id="f1e54-113">次のコードは、で`FunctionName`使用される関数を実装する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="f1e54-113">The following code shows how to implement the function used by `FunctionName`.</span></span>

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
> <span data-ttu-id="f1e54-114">この呼び出しは`event.completed` 、イベントが正常に処理されたことを通知します。</span><span class="sxs-lookup"><span data-stu-id="f1e54-114">The call to `event.completed` signals that you have successfully handled the event.</span></span> <span data-ttu-id="f1e54-115">同一のアドイン コマンドを複数回クリックするなど、関数を複数回呼び出すと、すべてのイベントが自動的にキューに入れられます。</span><span class="sxs-lookup"><span data-stu-id="f1e54-115">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="f1e54-116">最初のイベントが自動的に実行され、その他のイベントはキューに残ります。</span><span class="sxs-lookup"><span data-stu-id="f1e54-116">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="f1e54-117">関数が呼び出さ`event.completed`れると、次にキューに入れられた関数の呼び出しが実行されます。</span><span class="sxs-lookup"><span data-stu-id="f1e54-117">When your function calls `event.completed`, the next queued call to that function runs.</span></span> <span data-ttu-id="f1e54-118">を呼び出す`event.completed`必要があります。それ以外の場合、関数は実行されません。</span><span class="sxs-lookup"><span data-stu-id="f1e54-118">You must call `event.completed`; otherwise your function will not run.</span></span>
