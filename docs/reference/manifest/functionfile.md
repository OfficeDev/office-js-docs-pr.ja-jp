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
# <a name="functionfile-element"></a>FunctionFile 要素

次のいずれかの方法でアドインが公開する操作のソース コード ファイルを指定します。

* UI を表示する代わりに JavaScript 関数を実行するアドイン コマンド。
* JavaScript 関数を実行するキーボード ショートカット。

要素 `FunctionFile` は [DesktopFormFactor](desktopformfactor.md) または [MobileFormFactor の子要素です](mobileformfactor.md)。 要素の属性は 32 文字以内で、Control 要素で定義されている UI なしアドイン コマンド ボタンで使用される `resid` `FunctionFile` `id` `Url` `Resources` JavaScript[](control.md)関数を含む、または読み込む HTML ファイルへの URL を含む要素内の要素の属性の値に設定されます。

要素の例を次に示 `FunctionFile` します。

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

要素によって示される HTML ファイル内の JavaScript は、次の 1 つのパラメーターを受け取る名前付き関数を呼び出して `FunctionFile` `Office.initialize` 定義する必要があります `event` 。 ユーザーに進捗状況や、成功か失敗かを通知するには、この関数で `item.notificationMessages` API を使用する必要があります。 実行が終了したときに、`event.completed` を呼び出す必要もあります。 関数の名前は、UI を使用する `FunctionName` ボタンの要素で使用されます。

関数を定義する HTML ファイルの例を次に示 `trackMessage` します。

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

次のコードは、使用される関数を実装する方法を示しています `FunctionName` 。

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
> イベントが正常 `event.completed` に処理されたことを示す呼び出し。 同一のアドイン コマンドを複数回クリックするなど、関数を複数回呼び出すと、すべてのイベントが自動的にキューに入れられます。 最初のイベントが自動的に実行され、その他のイベントはキューに残ります。 関数が呼び出されると `event.completed` 、その関数の次にキューに入っている呼び出しが実行されます。 呼び出す `event.completed` 必要があります。呼び出さなければ、関数は実行されません。
