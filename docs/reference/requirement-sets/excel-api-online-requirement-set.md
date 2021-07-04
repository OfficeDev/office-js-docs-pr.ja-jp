---
title: ExcelJavaScript API のオンライン専用要件セット
description: ExcelApiOnline 要件セットの詳細。
ms.date: 07/01/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ef4831cf6a6f9be1a5413c89ae0f971bef51a9b1
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290804"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a><span data-ttu-id="5e9c2-103">ExcelJavaScript API のオンライン専用要件セット</span><span class="sxs-lookup"><span data-stu-id="5e9c2-103">Excel JavaScript API online-only requirement set</span></span>

<span data-ttu-id="5e9c2-104">要件セットは、ユーザーが使用できる機能のみを含む特別な要件 `ExcelApiOnline` セットExcel on the web。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-104">The `ExcelApiOnline` requirement set is a special requirement set that includes features that are only available for Excel on the web.</span></span> <span data-ttu-id="5e9c2-105">この要件セットの API は、アプリケーションの実稼働 API (文書化されていない動作や構造上の変更の対象ではない) とExcel on the webされます。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-105">APIs in this requirement set are considered to be production APIs (not subject to undocumented behavioral or structural changes) for the Excel on the web application.</span></span> <span data-ttu-id="5e9c2-106">`ExcelApiOnline`API は、他のプラットフォーム (Windows、Mac、iOS) の 「プレビュー」 API と見なされ、これらのプラットフォームではサポートされない場合があります。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-106">`ExcelApiOnline` APIs are considered to be "preview" APIs for other platforms (Windows, Mac, iOS) and may not be supported by any of those platforms.</span></span>

<span data-ttu-id="5e9c2-107">要件セット内の API がすべてのプラットフォームでサポートされている場合は、次にリリースされた要件セット ( ) に `ExcelApiOnline` 追加されます `ExcelApi 1.[NEXT]` 。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-107">When APIs in the `ExcelApiOnline` requirement set are supported across all platforms, they will added to the next released requirement set (`ExcelApi 1.[NEXT]`).</span></span> <span data-ttu-id="5e9c2-108">その新しい要件が公開されると、これらの API はから削除されます `ExcelApiOnline` 。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-108">Once that new requirement is public, those APIs will be removed from `ExcelApiOnline`.</span></span> <span data-ttu-id="5e9c2-109">これは、プレビューからリリースに移行する API と同様のプロモーション プロセスと考えて下さい。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-109">Think of this as a similar promotion process to an API moving from preview to release.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5e9c2-110">`ExcelApiOnline` は、最新の番号付き要件セットのスーパーセットです。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-110">`ExcelApiOnline` is a superset of the latest numbered requirement set.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5e9c2-111">`ExcelApiOnline 1.1` は、オンライン専用 API の唯一のバージョンです。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-111">`ExcelApiOnline 1.1` is the only version of the online-only APIs.</span></span> <span data-ttu-id="5e9c2-112">これは、最新Excel on the webユーザーが常に 1 つのバージョンを使用できるためです。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-112">This is because Excel on the web will always have a single version available to users that is the latest version.</span></span>

<span data-ttu-id="5e9c2-113">次の表に、API の簡潔な概要を示しますが、後続の API リスト テーブルでは、現在の [API](#api-list) の詳細な一覧を `ExcelApiOnline` 示します。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-113">The following table provides a concise summary of the APIs, while the subsequent [API list](#api-list) table gives a detailed list of the current `ExcelApiOnline` APIs.</span></span>

| <span data-ttu-id="5e9c2-114">機能領域</span><span class="sxs-lookup"><span data-stu-id="5e9c2-114">Feature area</span></span> | <span data-ttu-id="5e9c2-115">説明</span><span class="sxs-lookup"><span data-stu-id="5e9c2-115">Description</span></span> | <span data-ttu-id="5e9c2-116">関連オブジェクト</span><span class="sxs-lookup"><span data-stu-id="5e9c2-116">Relevant objects</span></span> |
|:--- |:--- |:--- |
| <span data-ttu-id="5e9c2-117">名前付きシート ビュー</span><span class="sxs-lookup"><span data-stu-id="5e9c2-117">Named sheet views</span></span> | <span data-ttu-id="5e9c2-118">ユーザーごとのワークシート ビューをプログラムで制御できます。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-118">Gives programmatic control of per-user worksheet views.</span></span> | [<span data-ttu-id="5e9c2-119">NamedSheetView</span><span class="sxs-lookup"><span data-stu-id="5e9c2-119">NamedSheetView</span></span>](/javascript/api/excel/excel.namedsheetview) |

## <a name="recommended-usage"></a><span data-ttu-id="5e9c2-120">推奨される使用法</span><span class="sxs-lookup"><span data-stu-id="5e9c2-120">Recommended usage</span></span>

<span data-ttu-id="5e9c2-121">API はユーザーによってのみサポートExcel on the web、アドインは、これらの API を呼び出す前に要件セットがサポートされていない `ExcelApiOnline` か確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-121">Because `ExcelApiOnline` APIs are only supported by Excel on the web, your add-in should check if the requirement set is supported before calling these APIs.</span></span> <span data-ttu-id="5e9c2-122">これにより、別のプラットフォームでオンライン専用 API を呼び出すのを回避できます。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-122">This avoids calling an online-only API on a different platform.</span></span>

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

<span data-ttu-id="5e9c2-123">API がクロスプラットフォーム要件セットに入った後は、チェックを削除または編集する必要 `isSetSupported` があります。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-123">Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check.</span></span> <span data-ttu-id="5e9c2-124">これにより、他のプラットフォームでアドインの機能が有効になります。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-124">This will enable your add-in's feature on other platforms.</span></span> <span data-ttu-id="5e9c2-125">この変更を行う場合は、必ずこれらのプラットフォームで機能をテストしてください。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-125">Be sure to test the feature on those platforms when making this change.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5e9c2-126">マニフェストでライセンス認証 `ExcelApiOnline 1.1` 要件を指定することはできません。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-126">Your manifest cannot specify `ExcelApiOnline 1.1` as an activation requirement.</span></span> <span data-ttu-id="5e9c2-127">Set 要素で使用する有効な値 [ではありません](../manifest/set.md)。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-127">It is not a valid value to use in the [Set element](../manifest/set.md).</span></span>

## <a name="api-list"></a><span data-ttu-id="5e9c2-128">API リスト</span><span class="sxs-lookup"><span data-stu-id="5e9c2-128">API list</span></span>

<span data-ttu-id="5e9c2-129">次の表に、要件Excel含まれている JavaScript API の一覧を `ExcelApiOnline` 示します。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-129">The following table lists the Excel JavaScript APIs currently included in the `ExcelApiOnline` requirement set.</span></span> <span data-ttu-id="5e9c2-130">すべての JavaScript API (API Excel以前にリリースされた API を含む) の完全な一覧については `ExcelApiOnline` [、JavaScript](/javascript/api/excel?view=excel-js-online&preserve-view=true)API Excel参照してください。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-130">For a complete list of all Excel JavaScript APIs (including `ExcelApiOnline` APIs and previously released APIs), see [all Excel JavaScript APIs](/javascript/api/excel?view=excel-js-online&preserve-view=true).</span></span>

| <span data-ttu-id="5e9c2-131">クラス</span><span class="sxs-lookup"><span data-stu-id="5e9c2-131">Class</span></span> | <span data-ttu-id="5e9c2-132">フィールド</span><span class="sxs-lookup"><span data-stu-id="5e9c2-132">Fields</span></span> | <span data-ttu-id="5e9c2-133">説明</span><span class="sxs-lookup"><span data-stu-id="5e9c2-133">Description</span></span> |
|:---|:---|:---|
|[<span data-ttu-id="5e9c2-134">NamedSheetView</span><span class="sxs-lookup"><span data-stu-id="5e9c2-134">NamedSheetView</span></span>](/javascript/api/excel/excel.namedsheetview)|[<span data-ttu-id="5e9c2-135">activate()</span><span class="sxs-lookup"><span data-stu-id="5e9c2-135">activate()</span></span>](/javascript/api/excel/excel.namedsheetview#activate--)|<span data-ttu-id="5e9c2-136">このシート ビューをアクティブ化します。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-136">Activates this sheet view.</span></span>|
||[<span data-ttu-id="5e9c2-137">delete()</span><span class="sxs-lookup"><span data-stu-id="5e9c2-137">delete()</span></span>](/javascript/api/excel/excel.namedsheetview#delete--)|<span data-ttu-id="5e9c2-138">ワークシートからシート ビューを削除します。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-138">Removes the sheet view from the worksheet.</span></span>|
||[<span data-ttu-id="5e9c2-139">duplicate(name?: string)</span><span class="sxs-lookup"><span data-stu-id="5e9c2-139">duplicate(name?: string)</span></span>](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|<span data-ttu-id="5e9c2-140">このシート ビューのコピーを作成します。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-140">Creates a copy of this sheet view.</span></span>|
||[<span data-ttu-id="5e9c2-141">name</span><span class="sxs-lookup"><span data-stu-id="5e9c2-141">name</span></span>](/javascript/api/excel/excel.namedsheetview#name)|<span data-ttu-id="5e9c2-142">シート ビューの名前を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-142">Gets or sets the name of the sheet view.</span></span>|
|[<span data-ttu-id="5e9c2-143">NamedSheetViewCollection</span><span class="sxs-lookup"><span data-stu-id="5e9c2-143">NamedSheetViewCollection</span></span>](/javascript/api/excel/excel.namedsheetviewcollection)|[<span data-ttu-id="5e9c2-144">add(name: string)</span><span class="sxs-lookup"><span data-stu-id="5e9c2-144">add(name: string)</span></span>](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|<span data-ttu-id="5e9c2-145">指定した名前の新しいシート ビューを作成します。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-145">Creates a new sheet view with the given name.</span></span>|
||[<span data-ttu-id="5e9c2-146">enterTemporary()</span><span class="sxs-lookup"><span data-stu-id="5e9c2-146">enterTemporary()</span></span>](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|<span data-ttu-id="5e9c2-147">新しい一時シート ビューを作成してアクティブ化します。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-147">Creates and activates a new temporary sheet view.</span></span>|
||[<span data-ttu-id="5e9c2-148">exit()</span><span class="sxs-lookup"><span data-stu-id="5e9c2-148">exit()</span></span>](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|<span data-ttu-id="5e9c2-149">現在アクティブなシート ビューを終了します。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-149">Exits the currently active sheet view.</span></span>|
||[<span data-ttu-id="5e9c2-150">getActive()</span><span class="sxs-lookup"><span data-stu-id="5e9c2-150">getActive()</span></span>](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|<span data-ttu-id="5e9c2-151">ワークシートの現在アクティブなシート ビューを取得します。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-151">Gets the worksheet's currently active sheet view.</span></span>|
||[<span data-ttu-id="5e9c2-152">getCount()</span><span class="sxs-lookup"><span data-stu-id="5e9c2-152">getCount()</span></span>](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|<span data-ttu-id="5e9c2-153">このワークシートのシート ビューの数を取得します。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-153">Gets the number of sheet views in this worksheet.</span></span>|
||[<span data-ttu-id="5e9c2-154">getItem(key: string)</span><span class="sxs-lookup"><span data-stu-id="5e9c2-154">getItem(key: string)</span></span>](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|<span data-ttu-id="5e9c2-155">名前を使用してシート ビューを取得します。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-155">Gets a sheet view using its name.</span></span>|
||[<span data-ttu-id="5e9c2-156">getItemAt(index: number)</span><span class="sxs-lookup"><span data-stu-id="5e9c2-156">getItemAt(index: number)</span></span>](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|<span data-ttu-id="5e9c2-157">コレクション内のインデックスによってシート ビューを取得します。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-157">Gets a sheet view by its index in the collection.</span></span>|
||[<span data-ttu-id="5e9c2-158">items</span><span class="sxs-lookup"><span data-stu-id="5e9c2-158">items</span></span>](/javascript/api/excel/excel.namedsheetviewcollection#items)|<span data-ttu-id="5e9c2-159">このコレクション内に読み込まれた子アイテムを取得します。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-159">Gets the loaded child items in this collection.</span></span>|
|[<span data-ttu-id="5e9c2-160">Worksheet</span><span class="sxs-lookup"><span data-stu-id="5e9c2-160">Worksheet</span></span>](/javascript/api/excel/excel.worksheet)|[<span data-ttu-id="5e9c2-161">namedSheetViews</span><span class="sxs-lookup"><span data-stu-id="5e9c2-161">namedSheetViews</span></span>](/javascript/api/excel/excel.worksheet#namedsheetviews)|<span data-ttu-id="5e9c2-162">ワークシートに存在するシート ビューのコレクションを返します。</span><span class="sxs-lookup"><span data-stu-id="5e9c2-162">Returns a collection of sheet views that are present in the worksheet.</span></span>|

## <a name="see-also"></a><span data-ttu-id="5e9c2-163">関連項目</span><span class="sxs-lookup"><span data-stu-id="5e9c2-163">See also</span></span>

- [<span data-ttu-id="5e9c2-164">Excel JavaScript API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="5e9c2-164">Excel JavaScript API Reference Documentation</span></span>](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [<span data-ttu-id="5e9c2-165">Excel JavaScript プレビュー API</span><span class="sxs-lookup"><span data-stu-id="5e9c2-165">Excel JavaScript preview APIs</span></span>](excel-preview-apis.md)
- [<span data-ttu-id="5e9c2-166">Excel JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="5e9c2-166">Excel JavaScript API requirement sets</span></span>](excel-api-requirement-sets.md)
