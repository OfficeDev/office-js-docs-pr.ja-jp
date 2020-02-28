---
title: マニフェスト ファイルの FunctionFile 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: eec1dc8eb2e099670469af6ef300592fc4a31e64
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324870"
---
# <a name="functionfile-element"></a>FunctionFile 要素

UI を表示するのではなく、JavaScript 関数を実行するアドインコマンドを使用して、アドインが公開する操作のソースコードファイルを指定します。 要素は[Desktopformfactor](desktopformfactor.md)または MobileFormFactor の子要素です。 [](mobileformfactor.md) `FunctionFile` 要素`resid`の`FunctionFile`属性は、要素内の要素`id`の`Url`属性の値に設定されて`Resources`います。この要素には、 [CONTROL 要素](control.md)で定義されているように、UI に含まれないアドインコマンドボタンによって使用されるすべての JavaScript 関数を含む HTML ファイルへの URL が含まれています。

`FunctionFile`要素の例を次に示します。

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

`FunctionFile`要素によって示される HTML ファイル内の JavaScript は`Office.initialize` 、1つのパラメーターを取る名前付き関数`event`を呼び出して定義する必要があります。 ユーザーに進捗状況や、成功か失敗かを通知するには、この関数で `item.notificationMessages` API を使用する必要があります。 実行が終了したときに、`event.completed` を呼び出す必要もあります。 関数の名前は、UI のないボタン`FunctionName`の要素として使用されます。

関数を`trackMessage`定義する HTML ファイルの例を次に示します。

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

次のコードは、で`FunctionName`使用される関数を実装する方法を示しています。

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
> この呼び出しは`event.completed` 、イベントが正常に処理されたことを通知します。 同一のアドイン コマンドを複数回クリックするなど、関数を複数回呼び出すと、すべてのイベントが自動的にキューに入れられます。 最初のイベントが自動的に実行され、その他のイベントはキューに残ります。 関数が呼び出さ`event.completed`れると、次にキューに入れられた関数の呼び出しが実行されます。 を呼び出す`event.completed`必要があります。それ以外の場合、関数は実行されません。
