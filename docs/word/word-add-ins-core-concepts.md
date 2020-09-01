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
# <a name="fundamental-programming-concepts-with-the-word-javascript-api"></a><span data-ttu-id="a6e69-103">Word JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="a6e69-103">Fundamental programming concepts with the Word JavaScript API</span></span>

<span data-ttu-id="a6e69-104">この記事では、[Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) を使用して Word 2016 以降のアドインを構築する場合の基本的な概念について説明します。</span><span class="sxs-lookup"><span data-stu-id="a6e69-104">This article describes concepts that are fundamental to using the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) to build add-ins for Word 2016 or later.</span></span>

## <a name="referencing-officejs"></a><span data-ttu-id="a6e69-105">Office.js を参照する</span><span class="sxs-lookup"><span data-stu-id="a6e69-105">Referencing Office.js</span></span>

<span data-ttu-id="a6e69-106">Office.js は、次の場所から参照できます。</span><span class="sxs-lookup"><span data-stu-id="a6e69-106">You can reference Office.js from the following locations:</span></span>

- <span data-ttu-id="a6e69-107">`https://appsforoffice.microsoft.com/lib/1/hosted/office.js` - 運用環境のアドインには、このリソースを使用します。</span><span class="sxs-lookup"><span data-stu-id="a6e69-107">`https://appsforoffice.microsoft.com/lib/1/hosted/office.js` - use this resource for production add-ins.</span></span>

- <span data-ttu-id="a6e69-108">`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - このリソースを使用してプレビュー機能を試します。</span><span class="sxs-lookup"><span data-stu-id="a6e69-108">`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - use this resource to try out preview features.</span></span>

## <a name="word-javascript-api-requirement-sets"></a><span data-ttu-id="a6e69-109">Word JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="a6e69-109">Word JavaScript API requirement sets</span></span>

<span data-ttu-id="a6e69-110">要件セットは、API メンバーの名前付きグループです。</span><span class="sxs-lookup"><span data-stu-id="a6e69-110">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="a6e69-111">Office アドインでは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判断します。</span><span class="sxs-lookup"><span data-stu-id="a6e69-111">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs.</span></span> <span data-ttu-id="a6e69-112">Word JavaScript API 要件セットの詳細については、「[Word JavaScript API の要件セット](../reference/requirement-sets/word-api-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a6e69-112">For detailed information about Word JavaScript API requirement sets, see [Word JavaScript API requirement sets](../reference/requirement-sets/word-api-requirement-sets.md).</span></span>

## <a name="running-word-add-ins"></a><span data-ttu-id="a6e69-113">Word アドインを実行する</span><span class="sxs-lookup"><span data-stu-id="a6e69-113">Running Word add-ins</span></span>

<span data-ttu-id="a6e69-114">アドインを実行するには、`Office.initialize` イベント ハンドラーを使用します。</span><span class="sxs-lookup"><span data-stu-id="a6e69-114">To run your add-in, use an `Office.initialize` event handler.</span></span> <span data-ttu-id="a6e69-115">アドインの初期化の詳細については、「[API について](../develop/understanding-the-javascript-api-for-office.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a6e69-115">For more information about add-in initialization, see [Understanding the API](../develop/understanding-the-javascript-api-for-office.md).</span></span>

<span data-ttu-id="a6e69-116">Word 2016 以降を対象とするアドインは、Word 固有の API を使用することができます。</span><span class="sxs-lookup"><span data-stu-id="a6e69-116">Add-ins that target Word 2016 or later can use the Word-specific APIs.</span></span> <span data-ttu-id="a6e69-117">これらは、Word の相互作用ロジックを関数として `Word.run()` メソッドに渡します。</span><span class="sxs-lookup"><span data-stu-id="a6e69-117">They pass the Word-interaction logic as a function into the `Word.run()` method.</span></span> <span data-ttu-id="a6e69-118">このプログラミング モデルの Word 文書を操作する方法については、「[アプリケーション固有の API モデルの使用](../develop/application-specific-api-model.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a6e69-118">See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about how to interact with the Word document in this programming model.</span></span>

<span data-ttu-id="a6e69-119">次の例では、`Word.run()` メソッドを使用して、Word アドインを初期化および実行する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="a6e69-119">The following example shows how to initialize and run a Word add-in by using the `Word.run()` method.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="a6e69-120">関連項目</span><span class="sxs-lookup"><span data-stu-id="a6e69-120">See also</span></span>

- [<span data-ttu-id="a6e69-121">Word JavaScript API の概要</span><span class="sxs-lookup"><span data-stu-id="a6e69-121">Word JavaScript API overview</span></span>](../reference/overview/word-add-ins-reference-overview.md)
- [<span data-ttu-id="a6e69-122">最初の Word アドインをビルドする</span><span class="sxs-lookup"><span data-stu-id="a6e69-122">Build your first Word add-in</span></span>](../quickstarts/word-quickstart.md)
- [<span data-ttu-id="a6e69-123">Word アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="a6e69-123">Word add-in tutorial</span></span>](../tutorials/word-tutorial.md)
- [<span data-ttu-id="a6e69-124">Word JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="a6e69-124">Word JavaScript API reference</span></span>](/javascript/api/word)
