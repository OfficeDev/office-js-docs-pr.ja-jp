---
title: Word JavaScript API を使用した基本的なプログラミングの概念
description: Word JavaScript API を使用して、Word 用アドインを構築します。
ms.date: 07/28/2020
localization_priority: Priority
ms.openlocfilehash: 1e7a90d4be378ed9b2c1f30ebebd4a0beec45a11
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293094"
---
# <a name="fundamental-programming-concepts-with-the-word-javascript-api"></a>Word JavaScript API を使用した基本的なプログラミングの概念

この記事では、[Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) を使用して Word 2016 以降のアドインを構築する場合の基本的な概念について説明します。

## <a name="referencing-officejs"></a>Office.js を参照する

Office.js は、次の場所から参照できます。

- `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` - 運用環境のアドインには、このリソースを使用します。

- `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - このリソースを使用してプレビュー機能を試します。

## <a name="word-javascript-api-requirement-sets"></a>Word JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。 Office アドインでは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判断します。 Word JavaScript API 要件セットの詳細については、「[Word JavaScript API の要件セット](../reference/requirement-sets/word-api-requirement-sets.md)」を参照してください。

## <a name="running-word-add-ins"></a>Word アドインを実行する

アドインを実行するには、`Office.initialize` イベント ハンドラーを使用します。 アドインの初期化の詳細については、「[API について](../develop/understanding-the-javascript-api-for-office.md)」を参照してください。

Word 2016 以降を対象とするアドインは、Word 固有の API を使用することができます。 これらは、Word の相互作用ロジックを関数として `Word.run()` メソッドに渡します。 このプログラミング モデルの Word 文書を操作する方法については、「[アプリケーション固有の API モデルの使用](../develop/application-specific-api-model.md)」を参照してください。

次の例では、`Word.run()` メソッドを使用して、Word アドインを初期化および実行する方法を示します。

```js
(function () {
    "use strict";

    // The initialize event handler must be run on each page to initialize Office JS.
    // You can add optional custom initialization code that will run after OfficeJS
    // has initialized.
    Office.initialize = function (reason) {
        // The reason object tells how the add-in was initialized. The values can be:
        // inserted - the add-in was inserted to an open document.
        // documentOpened - the add-in was already inserted in to the document and the document was opened.

        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // Set your optional initialization code.
            // You can also load saved settings from the Office object.
        });
    };

    // Run a batch operation against the Word JavaScript API object model.
    // Use the context argument to get access to the Word document.
    Word.run(function (context) {

        // Create a proxy object for the document.
        var thisDocument = context.document;
        // ...
    })
})();
```

## <a name="see-also"></a>関連項目

- [Word JavaScript API の概要](../reference/overview/word-add-ins-reference-overview.md)
- [最初の Word アドインをビルドする](../quickstarts/word-quickstart.md)
- [Word アドインのチュートリアル](../tutorials/word-tutorial.md)
- [Word JavaScript API リファレンス](/javascript/api/word)
