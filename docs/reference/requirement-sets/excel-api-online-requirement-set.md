---
title: Excel JavaScript API のオンラインのみの要件セット
description: ExcelApiOnline の要件セットの詳細
ms.date: 11/19/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: e583c9832f04e17dc1c82d38d056fe2749888a77
ms.sourcegitcommit: e56bd8f1260c73daf33272a30dc5af242452594f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/21/2019
ms.locfileid: "38757493"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a><span data-ttu-id="2221c-103">Excel JavaScript API のオンラインのみの要件セット</span><span class="sxs-lookup"><span data-stu-id="2221c-103">Excel JavaScript API online-only requirement set</span></span>

<span data-ttu-id="2221c-104">`ExcelApiOnline`要件セットは、web 上の Excel でのみ使用可能な機能を含む特別な要件セットです。</span><span class="sxs-lookup"><span data-stu-id="2221c-104">The `ExcelApiOnline` requirement set is a special requirement set that includes features that are only available for Excel on the web.</span></span> <span data-ttu-id="2221c-105">この要件セットの Api は、web ホスト上の Excel の運用 Api (未提出の行動または構造上の変更による影響を受けない) と見なされます。</span><span class="sxs-lookup"><span data-stu-id="2221c-105">APIs in this requirement set are considered to be production APIs (not subject to undocumented behavioral or structural changes) for the Excel on the web host.</span></span> <span data-ttu-id="2221c-106">`ExcelApiOnline`他のプラットフォーム (Windows、Mac、iOS) の場合は "preview" Api と見なされますが、これらのプラットフォームではサポートされていない場合があります。</span><span class="sxs-lookup"><span data-stu-id="2221c-106">`ExcelApiOnline` are considered to be "preview" APIs for other platforms (Windows, Mac, iOS) and may not be supported by any of those platforms.</span></span>

<span data-ttu-id="2221c-107">`ExcelApiOnline`要件セットの api がすべてのプラットフォームでサポートされている場合は、次にリリースされる`ExcelApi 1.[NEXT]`要件セット () に追加されます。</span><span class="sxs-lookup"><span data-stu-id="2221c-107">When APIs in the `ExcelApiOnline` requirement set are supported across all platforms, they will added to the next released requirement set (`ExcelApi 1.[NEXT]`).</span></span> <span data-ttu-id="2221c-108">新しい要件が公開されると、これらの Api はから`ExcelApiOnline`削除されます。</span><span class="sxs-lookup"><span data-stu-id="2221c-108">Once that new requirement is public, those APIs will be removed from `ExcelApiOnline`.</span></span> <span data-ttu-id="2221c-109">この点は、プレビューからリリースに移行する API と同様に、昇格プロセスと考えることができます。</span><span class="sxs-lookup"><span data-stu-id="2221c-109">Think of this as a similar promotion process as an API moving from preview to release.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2221c-110">`ExcelApiOnline`は、最新の番号付き要件セットのスーパーセットです。</span><span class="sxs-lookup"><span data-stu-id="2221c-110">`ExcelApiOnline` is superset of the latest numbered requirement set.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2221c-111">`ExcelApiOnline 1.1`は、オンライン専用 Api の唯一のバージョンです。</span><span class="sxs-lookup"><span data-stu-id="2221c-111">`ExcelApiOnline 1.1` is the only version of the online-only APIs.</span></span> <span data-ttu-id="2221c-112">これは、web 上の Excel では、最新バージョンのユーザーが常に1つのバージョンを使用できるためです。</span><span class="sxs-lookup"><span data-stu-id="2221c-112">This is because Excel on the web will always have a single version available to users that is the latest version.</span></span>

## <a name="recommended-usage"></a><span data-ttu-id="2221c-113">推奨される使用法</span><span class="sxs-lookup"><span data-stu-id="2221c-113">Recommended usage</span></span>

<span data-ttu-id="2221c-114">Api `ExcelApiOnline`は web 上の Excel でのみサポートされているため、アドインでは、これらの api を呼び出す前に要件セットがサポートされているかどうかを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2221c-114">Because `ExcelApiOnline` APIs are only supported by Excel on the web, your add-in should check if the requirement set is supported before calling these APIs.</span></span> <span data-ttu-id="2221c-115">これにより、別のプラットフォームでオンラインのみの API を呼び出すことを回避できます。</span><span class="sxs-lookup"><span data-stu-id="2221c-115">This avoids calling an online-only API on a different platform.</span></span>

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

<span data-ttu-id="2221c-116">クロスプラットフォームの要件セットに含まれる API は、 `isSetSupported`チェックを削除または編集する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2221c-116">Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check.</span></span> <span data-ttu-id="2221c-117">これにより、他のプラットフォームでアドインの機能が有効になります。</span><span class="sxs-lookup"><span data-stu-id="2221c-117">This will enable your add-in's feature on other platforms.</span></span> <span data-ttu-id="2221c-118">この変更を行うときは、これらのプラットフォームで機能をテストしてください。</span><span class="sxs-lookup"><span data-stu-id="2221c-118">Be sure to test the feature on those platforms when making this change.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2221c-119">マニフェストでライセンス認証`ExcelApiOnline 1.1`の要件として指定することはできません。</span><span class="sxs-lookup"><span data-stu-id="2221c-119">Your manifest cannot specify `ExcelApiOnline 1.1` as an activation requirement.</span></span> <span data-ttu-id="2221c-120">[Set 要素](../manifest/set.md)で使用する有効な値ではありません。</span><span class="sxs-lookup"><span data-stu-id="2221c-120">It is not a valid value to use in the [Set element](../manifest/set.md).</span></span>

## <a name="api-list"></a><span data-ttu-id="2221c-121">API リスト</span><span class="sxs-lookup"><span data-stu-id="2221c-121">API list</span></span>

<span data-ttu-id="2221c-122">現時点では、オンライン専用の Api はありません。</span><span class="sxs-lookup"><span data-stu-id="2221c-122">There are currently no online-only APIs.</span></span> <span data-ttu-id="2221c-123">新機能が web 上の Excel に追加され、Office JavaScript Api によってサポートされるようになると、もう一度確認してください。</span><span class="sxs-lookup"><span data-stu-id="2221c-123">Check back as new features are added to Excel on the web and supported by the Office JavaScript APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="2221c-124">関連項目</span><span class="sxs-lookup"><span data-stu-id="2221c-124">See also</span></span>

- [<span data-ttu-id="2221c-125">Excel JavaScript API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="2221c-125">Excel JavaScript API Reference Documentation</span></span>](/javascript/api/excel?view=excel-js-online)
- [<span data-ttu-id="2221c-126">Excel JavaScript プレビュー API</span><span class="sxs-lookup"><span data-stu-id="2221c-126">Excel JavaScript preview APIs</span></span>](./excel-preview-apis.md)
- [<span data-ttu-id="2221c-127">Excel JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="2221c-127">Excel JavaScript API requirement sets</span></span>](./excel-api-requirement-sets.md)