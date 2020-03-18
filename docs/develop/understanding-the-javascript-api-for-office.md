---
title: Office JavaScript API について
description: Office JavaScript API の概要
ms.date: 02/27/2020
localization_priority: Priority
ms.openlocfilehash: 67ee9459aab3065466ac8f52f2f835ca1e94bfc3
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718791"
---
# <a name="understanding-the-office-javascript-api"></a><span data-ttu-id="e1e09-103">Office JavaScript API について</span><span class="sxs-lookup"><span data-stu-id="e1e09-103">Understanding the Office JavaScript API</span></span>

<span data-ttu-id="e1e09-104">Office アドインでは、Office JavaScript API を使用することで、アドインが実行されている Office ドキュメント内のコンテンツを操作できます。</span><span class="sxs-lookup"><span data-stu-id="e1e09-104">An Office Add-in can use the Office JavaScript APIs to interact with content in the Office document where the add-in is running.</span></span>

## <a name="accessing-the-office-javascript-api-library"></a><span data-ttu-id="e1e09-105">Office JavaScript API ライブラリへのアクセス</span><span class="sxs-lookup"><span data-stu-id="e1e09-105">Accessing the Office JavaScript API library</span></span>

[!include[information about accessing the Office JS API library](../includes/office-js-access-library.md)]

## <a name="api-models"></a><span data-ttu-id="e1e09-106">API モデル</span><span class="sxs-lookup"><span data-stu-id="e1e09-106">API models</span></span>

[!include[information about the Office JS API models](../includes/office-js-api-models.md)]

## <a name="api-requirement-sets"></a><span data-ttu-id="e1e09-107">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e1e09-107">API requirement sets</span></span>

[!include[information about the Office JS API requirement sets](../includes/office-js-requirement-sets.md)]

> [!NOTE]
> <span data-ttu-id="e1e09-p101">AppSource にアドインを[公開](../publish/publish.md)し、Office エクスペリエンスで利用できるようにする予定がある場合は、[AppSource の検証ポリシー](/office/dev/store/validation-policies)に準拠していることを確認してください。たとえば、検証に合格するには、定義したメソッドをサポートするすべてのプラットフォームでアドインが動作する必要があります (詳細については、[セクション 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) と [Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)のページを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="e1e09-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span> 

## <a name="see-also"></a><span data-ttu-id="e1e09-110">関連項目</span><span class="sxs-lookup"><span data-stu-id="e1e09-110">See also</span></span>

- [<span data-ttu-id="e1e09-111">Office JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="e1e09-111">Office JavaScript API reference</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="e1e09-112">DOM とランタイム環境を読み込む</span><span class="sxs-lookup"><span data-stu-id="e1e09-112">Loading the DOM and runtime environment</span></span>](loading-the-dom-and-runtime-environment.md)
- [<span data-ttu-id="e1e09-113">Office JavaScript API ライブラリの参照</span><span class="sxs-lookup"><span data-stu-id="e1e09-113">Referencing the Office JavaScript API library</span></span>](referencing-the-javascript-api-for-office-library-from-its-cdn.md)
- [<span data-ttu-id="e1e09-114">Office アドインを初期化する</span><span class="sxs-lookup"><span data-stu-id="e1e09-114">Initialize your Office Add-in</span></span>](initialize-add-in.md)
