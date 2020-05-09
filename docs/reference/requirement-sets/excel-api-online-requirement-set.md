---
title: Excel JavaScript API のオンラインのみの要件セット
description: ExcelApiOnline の要件セットの詳細
ms.date: 05/06/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f177e0107de7172c350f94c3a022cb3e0db5c6f5
ms.sourcegitcommit: 735bf94ac3c838f580a992e7ef074dbc8be2b0ea
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/08/2020
ms.locfileid: "44170787"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a><span data-ttu-id="bfc4f-103">Excel JavaScript API のオンラインのみの要件セット</span><span class="sxs-lookup"><span data-stu-id="bfc4f-103">Excel JavaScript API online-only requirement set</span></span>

<span data-ttu-id="bfc4f-104">`ExcelApiOnline`要件セットは、web 上の Excel でのみ使用可能な機能を含む特別な要件セットです。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-104">The `ExcelApiOnline` requirement set is a special requirement set that includes features that are only available for Excel on the web.</span></span> <span data-ttu-id="bfc4f-105">この要件セットの Api は、web ホスト上の Excel の運用 Api (未提出の行動または構造上の変更による影響を受けない) と見なされます。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-105">APIs in this requirement set are considered to be production APIs (not subject to undocumented behavioral or structural changes) for the Excel on the web host.</span></span> <span data-ttu-id="bfc4f-106">`ExcelApiOnline`他のプラットフォーム (Windows、Mac、iOS) の場合は "preview" Api と見なされますが、これらのプラットフォームではサポートされていない場合があります。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-106">`ExcelApiOnline` are considered to be "preview" APIs for other platforms (Windows, Mac, iOS) and may not be supported by any of those platforms.</span></span>

<span data-ttu-id="bfc4f-107">`ExcelApiOnline`要件セットの api がすべてのプラットフォームでサポートされている場合は、次にリリースされる`ExcelApi 1.[NEXT]`要件セット () に追加されます。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-107">When APIs in the `ExcelApiOnline` requirement set are supported across all platforms, they will added to the next released requirement set (`ExcelApi 1.[NEXT]`).</span></span> <span data-ttu-id="bfc4f-108">新しい要件が公開されると、これらの Api はから`ExcelApiOnline`削除されます。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-108">Once that new requirement is public, those APIs will be removed from `ExcelApiOnline`.</span></span> <span data-ttu-id="bfc4f-109">この点は、プレビューからリリースに移行する API と同様に、昇格プロセスと考えることができます。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-109">Think of this as a similar promotion process as an API moving from preview to release.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bfc4f-110">`ExcelApiOnline`は、最新の番号付き要件セットのスーパーセットです。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-110">`ExcelApiOnline` is superset of the latest numbered requirement set.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bfc4f-111">`ExcelApiOnline 1.1`は、オンライン専用 Api の唯一のバージョンです。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-111">`ExcelApiOnline 1.1` is the only version of the online-only APIs.</span></span> <span data-ttu-id="bfc4f-112">これは、web 上の Excel では、最新バージョンのユーザーが常に1つのバージョンを使用できるためです。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-112">This is because Excel on the web will always have a single version available to users that is the latest version.</span></span>

## <a name="recommended-usage"></a><span data-ttu-id="bfc4f-113">推奨される使用法</span><span class="sxs-lookup"><span data-stu-id="bfc4f-113">Recommended usage</span></span>

<span data-ttu-id="bfc4f-114">Api `ExcelApiOnline`は web 上の Excel でのみサポートされているため、アドインでは、これらの api を呼び出す前に要件セットがサポートされているかどうかを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-114">Because `ExcelApiOnline` APIs are only supported by Excel on the web, your add-in should check if the requirement set is supported before calling these APIs.</span></span> <span data-ttu-id="bfc4f-115">これにより、別のプラットフォームでオンラインのみの API を呼び出すことを回避できます。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-115">This avoids calling an online-only API on a different platform.</span></span>

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

<span data-ttu-id="bfc4f-116">クロスプラットフォームの要件セットに含まれる API は、 `isSetSupported`チェックを削除または編集する必要があります。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-116">Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check.</span></span> <span data-ttu-id="bfc4f-117">これにより、他のプラットフォームでアドインの機能が有効になります。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-117">This will enable your add-in's feature on other platforms.</span></span> <span data-ttu-id="bfc4f-118">この変更を行うときは、これらのプラットフォームで機能をテストしてください。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-118">Be sure to test the feature on those platforms when making this change.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bfc4f-119">マニフェストでライセンス認証`ExcelApiOnline 1.1`の要件として指定することはできません。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-119">Your manifest cannot specify `ExcelApiOnline 1.1` as an activation requirement.</span></span> <span data-ttu-id="bfc4f-120">[Set 要素](../manifest/set.md)で使用する有効な値ではありません。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-120">It is not a valid value to use in the [Set element](../manifest/set.md).</span></span>

## <a name="api-list"></a><span data-ttu-id="bfc4f-121">API リスト</span><span class="sxs-lookup"><span data-stu-id="bfc4f-121">API list</span></span>

<span data-ttu-id="bfc4f-122">次の Api は、現在、 `ExcelApiOnline 1.1`要件セットの一部として web 上の Excel で使用できます。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-122">The following APIs are currently available for Excel on the web as part of the `ExcelApiOnline 1.1` requirement set.</span></span>

| <span data-ttu-id="bfc4f-123">クラス</span><span class="sxs-lookup"><span data-stu-id="bfc4f-123">Class</span></span> | <span data-ttu-id="bfc4f-124">フィールド</span><span class="sxs-lookup"><span data-stu-id="bfc4f-124">Fields</span></span> | <span data-ttu-id="bfc4f-125">説明</span><span class="sxs-lookup"><span data-stu-id="bfc4f-125">Description</span></span> |
|:---|:---|:---|
|[<span data-ttu-id="bfc4f-126">ChartAxisTitle</span><span class="sxs-lookup"><span data-stu-id="bfc4f-126">ChartAxisTitle</span></span>](/javascript/api/excel/excel.chartaxistitle)|[<span data-ttu-id="bfc4f-127">textOrientation</span><span class="sxs-lookup"><span data-stu-id="bfc4f-127">textOrientation</span></span>](/javascript/api/excel/excel.chartaxistitle#textorientation)|<span data-ttu-id="bfc4f-128">グラフ軸のタイトルに対して、テキストの方向を指定する角度を指定します。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-128">Specifies the angle to which the text is oriented for the chart axis title.</span></span> <span data-ttu-id="bfc4f-129">この値は、-90 ~ 90 の整数、または垂直方向のテキストの整数の180のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-129">The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.</span></span>|
|[<span data-ttu-id="bfc4f-130">PivotTableScopedCollection</span><span class="sxs-lookup"><span data-stu-id="bfc4f-130">PivotTableScopedCollection</span></span>](/javascript/api/excel/excel.pivottablescopedcollection)|[<span data-ttu-id="bfc4f-131">getCount()</span><span class="sxs-lookup"><span data-stu-id="bfc4f-131">getCount()</span></span>](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|<span data-ttu-id="bfc4f-132">コレクション内のピボットテーブルの数を取得します。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-132">Gets the number of PivotTables in the collection.</span></span>|
||[<span data-ttu-id="bfc4f-133">getFirst()</span><span class="sxs-lookup"><span data-stu-id="bfc4f-133">getFirst()</span></span>](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|<span data-ttu-id="bfc4f-134">コレクション内の最初のピボットテーブルを取得します。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-134">Gets the first PivotTable in the collection.</span></span> <span data-ttu-id="bfc4f-135">コレクション内のピボットテーブルは、上から下、左から右に並べ替えられます。この場合、左上のテーブルはコレクションの最初のピボットテーブルになります。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-135">The PivotTables in the collection are sorted top to bottom and left to right, such that top-left table is the first PivotTable in the collection.</span></span>|
||[<span data-ttu-id="bfc4f-136">getItem(key: string)</span><span class="sxs-lookup"><span data-stu-id="bfc4f-136">getItem(key: string)</span></span>](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|<span data-ttu-id="bfc4f-137">名前に基づいてピボットテーブルを取得します。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-137">Gets a PivotTable by name.</span></span>|
||[<span data-ttu-id="bfc4f-138">getItemOrNullObject(name: string)</span><span class="sxs-lookup"><span data-stu-id="bfc4f-138">getItemOrNullObject(name: string)</span></span>](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|<span data-ttu-id="bfc4f-139">名前を使用してピボットテーブルを取得します。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-139">Gets a PivotTable by name.</span></span> <span data-ttu-id="bfc4f-140">PivotTable が存在しない場合は null オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-140">If the PivotTable does not exist, will return a null object.</span></span>|
||[<span data-ttu-id="bfc4f-141">items</span><span class="sxs-lookup"><span data-stu-id="bfc4f-141">items</span></span>](/javascript/api/excel/excel.pivottablescopedcollection#items)|<span data-ttu-id="bfc4f-142">このコレクション内に読み込まれた子アイテムを取得します。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-142">Gets the loaded child items in this collection.</span></span>|
|[<span data-ttu-id="bfc4f-143">Range</span><span class="sxs-lookup"><span data-stu-id="bfc4f-143">Range</span></span>](/javascript/api/excel/excel.range)|[<span data-ttu-id="bfc4f-144">getPivotTables テーブル (fullyContained?: boolean)</span><span class="sxs-lookup"><span data-stu-id="bfc4f-144">getPivotTables(fullyContained?: boolean)</span></span>](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|<span data-ttu-id="bfc4f-145">範囲に重なっているピボットテーブルのスコープ設定されたコレクションを取得します。</span><span class="sxs-lookup"><span data-stu-id="bfc4f-145">Gets a scoped collection of PivotTables that overlap with the range.</span></span>|

## <a name="see-also"></a><span data-ttu-id="bfc4f-146">関連項目</span><span class="sxs-lookup"><span data-stu-id="bfc4f-146">See also</span></span>

- [<span data-ttu-id="bfc4f-147">Excel JavaScript API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="bfc4f-147">Excel JavaScript API Reference Documentation</span></span>](/javascript/api/excel?view=excel-js-online)
- [<span data-ttu-id="bfc4f-148">Excel JavaScript プレビュー API</span><span class="sxs-lookup"><span data-stu-id="bfc4f-148">Excel JavaScript preview APIs</span></span>](./excel-preview-apis.md)
- [<span data-ttu-id="bfc4f-149">Excel JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="bfc4f-149">Excel JavaScript API requirement sets</span></span>](./excel-api-requirement-sets.md)