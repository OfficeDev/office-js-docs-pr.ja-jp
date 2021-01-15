---
title: マニフェスト ファイルの FunctionFile 要素
description: UI を表示する代わりに JavaScript 関数を実行するアドイン コマンドを使用してアドインが公開する操作のソース コード ファイルを指定します。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 44bfd514025b8a23f4f6acdf3fec004485ca4c5a
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771401"
---
# <a name="functionfile-element"></a><span data-ttu-id="8e00c-103">FunctionFile 要素</span><span class="sxs-lookup"><span data-stu-id="8e00c-103">FunctionFile element</span></span>

<span data-ttu-id="8e00c-104">次のいずれかの方法でアドインが公開する操作のソース コード ファイルを指定します。</span><span class="sxs-lookup"><span data-stu-id="8e00c-104">Specifies the source code file for operations that an add-in exposes in one of the following ways:</span></span>

* <span data-ttu-id="8e00c-105">UI を表示する代わりに JavaScript 関数を実行するアドイン コマンド。</span><span class="sxs-lookup"><span data-stu-id="8e00c-105">Add-in commands that execute a JavaScript function instead of displaying UI.</span></span>
* <span data-ttu-id="8e00c-106">JavaScript 関数を実行するキーボード ショートカット。</span><span class="sxs-lookup"><span data-stu-id="8e00c-106">Keyboard shortcuts that execute a JavaScript function.</span></span>

<span data-ttu-id="8e00c-107">要素 `FunctionFile` は [DesktopFormFactor](desktopformfactor.md) または [MobileFormFactor の子要素です](mobileformfactor.md)。</span><span class="sxs-lookup"><span data-stu-id="8e00c-107">The `FunctionFile` element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> <span data-ttu-id="8e00c-108">要素の属性は 32 文字以内で、Control 要素で定義されている UI なしアドイン コマンド ボタンで使用される `resid` `FunctionFile` `id` `Url` `Resources` JavaScript[](control.md)関数を含む、または読み込む HTML ファイルへの URL を含む要素内の要素の属性の値に設定されます。</span><span class="sxs-lookup"><span data-stu-id="8e00c-108">The `resid` attribute of the `FunctionFile` element can be no more than 32 characters and is set to the value of the `id` attribute of a `Url` element in the `Resources` element that contains the URL to an HTML file that contains or loads all the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="8e00c-109">要素の例を次に示 `FunctionFile` します。</span><span class="sxs-lookup"><span data-stu-id="8e00c-109">The following is an example of the `FunctionFile` element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="8e00c-110">要素によって示される HTML ファイル内の JavaScript は、次の 1 つのパラメーターを受け取る名前付き関数を呼び出して `FunctionFile` `Office.initialize` 定義する必要があります `event` 。</span><span class="sxs-lookup"><span data-stu-id="8e00c-110">The JavaScript in the HTML file indicated by the `FunctionFile` element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="8e00c-111">ユーザーに進捗状況や、成功か失敗かを通知するには、この関数で `item.notificationMessages` API を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8e00c-111">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="8e00c-112">実行が終了したときに、`event.completed` を呼び出す必要もあります。</span><span class="sxs-lookup"><span data-stu-id="8e00c-112">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="8e00c-113">関数の名前は、UI を使用する `FunctionName` ボタンの要素で使用されます。</span><span class="sxs-lookup"><span data-stu-id="8e00c-113">The name of the functions are used in the `FunctionName` element for UI-less buttons.</span></span>

<span data-ttu-id="8e00c-114">関数を定義する HTML ファイルの例を次に示 `trackMessage` します。</span><span class="sxs-lookup"><span data-stu-id="8e00c-114">The following is an example of an HTML file defining a `trackMessage` function.</span></span>

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

<span data-ttu-id="8e00c-115">次のコードは、使用される関数を実装する方法を示しています `FunctionName` 。</span><span class="sxs-lookup"><span data-stu-id="8e00c-115">The following code shows how to implement the function used by `FunctionName`.</span></span>

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
> <span data-ttu-id="8e00c-116">イベントが正常 `event.completed` に処理されたことを示す呼び出し。</span><span class="sxs-lookup"><span data-stu-id="8e00c-116">The call to `event.completed` signals that you have successfully handled the event.</span></span> <span data-ttu-id="8e00c-117">同一のアドイン コマンドを複数回クリックするなど、関数を複数回呼び出すと、すべてのイベントが自動的にキューに入れられます。</span><span class="sxs-lookup"><span data-stu-id="8e00c-117">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="8e00c-118">最初のイベントが自動的に実行され、その他のイベントはキューに残ります。</span><span class="sxs-lookup"><span data-stu-id="8e00c-118">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="8e00c-119">関数が呼び出されると `event.completed` 、その関数の次にキューに入っている呼び出しが実行されます。</span><span class="sxs-lookup"><span data-stu-id="8e00c-119">When your function calls `event.completed`, the next queued call to that function runs.</span></span> <span data-ttu-id="8e00c-120">呼び出す `event.completed` 必要があります。呼び出さなければ、関数は実行されません。</span><span class="sxs-lookup"><span data-stu-id="8e00c-120">You must call `event.completed`; otherwise your function will not run.</span></span>
