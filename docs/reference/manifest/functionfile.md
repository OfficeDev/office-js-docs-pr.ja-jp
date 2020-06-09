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
# <a name="functionfile-element"></a>FunctionFile 要素

UI を表示するのではなく、JavaScript 関数を実行するアドインコマンドを使用して、アドインが公開する操作のソースコードファイルを指定します。 `FunctionFile`要素は[Desktopformfactor](desktopformfactor.md)または[MobileFormFactor](mobileformfactor.md)の子要素です。 要素 `resid` の属性は、要素内の要素 `FunctionFile` の属性の値に設定されています。この要素には、 `id` `Url` `Resources` [Control 要素](control.md)で定義されているように、UI に含まれないアドインコマンドボタンによって使用されるすべての JavaScript 関数を含む HTML ファイルへの URL が含まれています。

要素の例を次に示し `FunctionFile` ます。

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

要素によって示される HTML ファイル内の JavaScript は、 `FunctionFile` `Office.initialize` 1 つのパラメーターを取る名前付き関数を呼び出して定義する必要があり `event` ます。 ユーザーに進捗状況や、成功か失敗かを通知するには、この関数で `item.notificationMessages` API を使用する必要があります。 実行が終了したときに、`event.completed` を呼び出す必要もあります。 関数の名前は、UI のないボタンの要素として使用され `FunctionName` ます。

関数を定義する HTML ファイルの例を次に示し `trackMessage` ます。

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

次のコードは、で使用される関数を実装する方法を示して `FunctionName` います。

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
> この呼び出しは、 `event.completed` イベントが正常に処理されたことを通知します。 同一のアドイン コマンドを複数回クリックするなど、関数を複数回呼び出すと、すべてのイベントが自動的にキューに入れられます。 最初のイベントが自動的に実行され、その他のイベントはキューに残ります。 関数が呼び出されると `event.completed` 、次にキューに入れられた関数の呼び出しが実行されます。 を呼び出す必要があり `event.completed` ます。それ以外の場合、関数は実行されません。
