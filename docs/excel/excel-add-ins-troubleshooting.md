---
title: Excel アドインのトラブルシューティング
description: Excel アドインの開発エラーをトラブルシューティングする方法について説明します。
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: 0efc8b4d25d9d748975146e187104972e4ad58a9
ms.sourcegitcommit: 1cdf5728102424a46998e1527508b4e7f9f74a4c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/17/2021
ms.locfileid: "50270729"
---
# <a name="troubleshooting-excel-add-ins"></a><span data-ttu-id="3bb57-103">Excel アドインのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="3bb57-103">Troubleshooting Excel Add-ins</span></span>

<span data-ttu-id="3bb57-104">この記事では、Excel に固有の問題のトラブルシューティングについて説明します。</span><span class="sxs-lookup"><span data-stu-id="3bb57-104">This article discusses troubleshooting issues that are unique to Excel.</span></span> <span data-ttu-id="3bb57-105">ページの下部にあるフィードバック ツールを使用して、記事に追加できるその他の問題を提案してください。</span><span class="sxs-lookup"><span data-stu-id="3bb57-105">Please use the feedback tool at the bottom of the page to suggest other issues that can be added to the article.</span></span>

## <a name="api-limitations-when-the-active-workbook-switches"></a><span data-ttu-id="3bb57-106">アクティブなブックが切り替えらた場合の API の制限事項</span><span class="sxs-lookup"><span data-stu-id="3bb57-106">API limitations when the active workbook switches</span></span>

<span data-ttu-id="3bb57-107">Excel 用アドインは、一度に 1 つのブックで動作することを目的とします。</span><span class="sxs-lookup"><span data-stu-id="3bb57-107">Add-ins for Excel are intended to operate on a single workbook at a time.</span></span> <span data-ttu-id="3bb57-108">アドインを実行しているブックとは別のブックがフォーカスを得た場合、エラーが発生する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="3bb57-108">Errors can arise when a workbook that is separate from the one running the add-in gains focus.</span></span> <span data-ttu-id="3bb57-109">これは、フォーカスが変更された時点で特定のメソッドが呼び出されている場合にのみ発生します。</span><span class="sxs-lookup"><span data-stu-id="3bb57-109">This only happens when particular methods are in the process of being called when the focus changes.</span></span>

<span data-ttu-id="3bb57-110">このブックの切り替えによって影響を受ける API は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="3bb57-110">The following APIs are affected by this workbook switch:</span></span>

|<span data-ttu-id="3bb57-111">Excel JavaScript API</span><span class="sxs-lookup"><span data-stu-id="3bb57-111">Excel JavaScript API</span></span> | <span data-ttu-id="3bb57-112">スローされたエラー</span><span class="sxs-lookup"><span data-stu-id="3bb57-112">Error thrown</span></span> |
|--|--|
| `Chart.activate` | <span data-ttu-id="3bb57-113">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bb57-113">GeneralException</span></span> |
| `Range.select` | <span data-ttu-id="3bb57-114">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bb57-114">GeneralException</span></span> |
| `Table.clearFilters` | <span data-ttu-id="3bb57-115">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bb57-115">GeneralException</span></span> |
| `Workbook.getActiveCell`  | <span data-ttu-id="3bb57-116">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="3bb57-116">InvalidSelection</span></span>|
| `Workbook.getSelectedRange` | <span data-ttu-id="3bb57-117">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="3bb57-117">InvalidSelection</span></span>|
| `Workbook.getSelectedRanges`  | <span data-ttu-id="3bb57-118">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="3bb57-118">InvalidSelection</span></span>|
| `Worksheet.activate` | <span data-ttu-id="3bb57-119">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bb57-119">GeneralException</span></span> |
| `Worksheet.delete`  | <span data-ttu-id="3bb57-120">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="3bb57-120">InvalidSelection</span></span>|
| `Worksheet.gridlines` | <span data-ttu-id="3bb57-121">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bb57-121">GeneralException</span></span> |
| `Worksheet.showHeadings` | <span data-ttu-id="3bb57-122">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bb57-122">GeneralException</span></span> |
| `WorksheetCollection.add` | <span data-ttu-id="3bb57-123">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bb57-123">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeAt` | <span data-ttu-id="3bb57-124">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bb57-124">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeColumns` | <span data-ttu-id="3bb57-125">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bb57-125">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeRows` | <span data-ttu-id="3bb57-126">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bb57-126">GeneralException</span></span> |
| `WorksheetFreezePanes.getLocationOrNullObject`| <span data-ttu-id="3bb57-127">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bb57-127">GeneralException</span></span> |
| `WorksheetFreezePanes.unfreeze` | <span data-ttu-id="3bb57-128">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bb57-128">GeneralException</span></span> |

> [!NOTE]
> <span data-ttu-id="3bb57-129">これは、Windows または Mac で開いている複数の Excel ブックにのみ適用されます。</span><span class="sxs-lookup"><span data-stu-id="3bb57-129">This only applies to multiple Excel workbooks open on Windows or Mac.</span></span>

## <a name="coauthoring"></a><span data-ttu-id="3bb57-130">共同編集</span><span class="sxs-lookup"><span data-stu-id="3bb57-130">Coauthoring</span></span>

<span data-ttu-id="3bb57-131">共同 [編集環境のイベントで](co-authoring-in-excel-add-ins.md) 使用するパターンについては、Excel アドインの共同編集を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3bb57-131">See [Coauthoring in Excel add-ins](co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment.</span></span> <span data-ttu-id="3bb57-132">また、次のような特定の API を使用する場合のマージ競合の可能性について説明します [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) 。</span><span class="sxs-lookup"><span data-stu-id="3bb57-132">The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).</span></span>

## <a name="known-issues"></a><span data-ttu-id="3bb57-133">既知の問題</span><span class="sxs-lookup"><span data-stu-id="3bb57-133">Known Issues</span></span>

### <a name="binding-events-return-temporary-binding-obects"></a><span data-ttu-id="3bb57-134">バインド イベントが一時的 `Binding` なオベクトを返す</span><span class="sxs-lookup"><span data-stu-id="3bb57-134">Binding events return temporary `Binding` obects</span></span>

<span data-ttu-id="3bb57-135">[BindingDataChangedEventArgs.binding と](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding)はどちらも、イベントを発生したオブジェクトの ID を含む一時オブジェクト `Binding` `Binding` を返します。</span><span class="sxs-lookup"><span data-stu-id="3bb57-135">Both [BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) and [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding) return a temporary `Binding` object that contains the ID of the `Binding` object that raised the event.</span></span> <span data-ttu-id="3bb57-136">この ID を使用して `BindingCollection.getItem(id)` 、イベントを発生 `Binding` したオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="3bb57-136">Use this ID with `BindingCollection.getItem(id)` to retrieve the `Binding` object that raised the event.</span></span>

<span data-ttu-id="3bb57-137">次のコード サンプルは、この一時バインド ID を使用して関連オブジェクトを取得する方法を示 `Binding` しています。</span><span class="sxs-lookup"><span data-stu-id="3bb57-137">The following code sample shows how to use this temporary binding ID to retrieve the related `Binding` object.</span></span> <span data-ttu-id="3bb57-138">サンプルでは、イベント リスナーがバインドに割り当てられます。</span><span class="sxs-lookup"><span data-stu-id="3bb57-138">In the sample, an event listener is assigned to a binding.</span></span> <span data-ttu-id="3bb57-139">リスナーは、イベントが `getBindingId` トリガーされるとメソッド `onDataChanged` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="3bb57-139">The listener calls the `getBindingId` method when the `onDataChanged` event is triggered.</span></span> <span data-ttu-id="3bb57-140">メソッド `getBindingId` は、一時オブジェクトの ID `Binding` を使用して、イベントを発生 `Binding` したオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="3bb57-140">The `getBindingId` method uses the ID of the temporary `Binding` object to retrieve the `Binding` object that raised the event.</span></span>

```js
Excel.run(function (context) {
    // Retrieve your binding.
    var binding = context.workbook.bindings.getItemAt(0);

    return context.sync().then(function () {
        // Register an event listener to detect changes to your binding
        // and then trigger the `getBindingId` method when the data changes. 
        binding.onDataChanged.add(getBindingId);

        return context.sync();
    });
});

function getBindingId(eventArgs) {
    return Excel.run(function (context) {
        // Get the temporary binding object and load its ID. 
        var tempBindingObject = eventArgs.binding;
        tempBindingObject.load("id");

        // Use the temporary binding object's ID to retrieve the original binding object. 
        var originalBindingObject = context.workbook.bindings.getItem(tempBindingObject.id);

        // You now have the binding object that raised the event: `originalBindingObject`. 
    });
}
```

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a><span data-ttu-id="3bb57-141">セルの `useStandardHeight` 形式 `useStandardWidth` と問題</span><span class="sxs-lookup"><span data-stu-id="3bb57-141">Cell format `useStandardHeight` and `useStandardWidth` issues</span></span>

<span data-ttu-id="3bb57-142">Web 上の Excel [では、useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) プロパティ `CellPropertiesFormat` が正しく動作しません。</span><span class="sxs-lookup"><span data-stu-id="3bb57-142">The [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) property of `CellPropertiesFormat` doesn't work properly in Excel on the web.</span></span> <span data-ttu-id="3bb57-143">Excel on the web UI の問題により、このプラットフォームでは高さが不正確に計算されるプロパティ `useStandardHeight` `true` を設定します。</span><span class="sxs-lookup"><span data-stu-id="3bb57-143">Due to an issue in the Excel on the web UI, setting the `useStandardHeight` property to `true` calculates height imprecisely on this platform.</span></span> <span data-ttu-id="3bb57-144">たとえば、Excel on the web では標準の高さ **14** **が 14.25** に変更されます。</span><span class="sxs-lookup"><span data-stu-id="3bb57-144">For example, a standard height of **14** is modified to **14.25** in Excel on the web.</span></span>

<span data-ttu-id="3bb57-145">すべてのプラットフォームで [、useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) プロパティと [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) プロパティは、設定のみを `CellPropertiesFormat` 目的とします `true` 。</span><span class="sxs-lookup"><span data-stu-id="3bb57-145">On all platforms, the [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) and [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) properties of `CellPropertiesFormat` are only intended to be set to `true`.</span></span> <span data-ttu-id="3bb57-146">これらのプロパティを設定 `false` する効果はありません。</span><span class="sxs-lookup"><span data-stu-id="3bb57-146">Setting these properties to `false` has no effect.</span></span> 

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a><span data-ttu-id="3bb57-147">`getImage`Excel for Mac でサポートされていない Range メソッド</span><span class="sxs-lookup"><span data-stu-id="3bb57-147">Range `getImage` method unsupported on Excel for Mac</span></span>

<span data-ttu-id="3bb57-148">Range [getImage メソッド](/javascript/api/excel/excel.range#getImage__) は、Excel for Mac では現在サポートされていません。</span><span class="sxs-lookup"><span data-stu-id="3bb57-148">The Range [getImage](/javascript/api/excel/excel.range#getImage__) method isn't currently supported in Excel for Mac.</span></span> <span data-ttu-id="3bb57-149">現在 [の状態については、「OfficeDev/office-js issue #235」](https://github.com/OfficeDev/office-js/issues/235) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3bb57-149">See [OfficeDev/office-js Issue #235](https://github.com/OfficeDev/office-js/issues/235) for the current status.</span></span>

### <a name="range-return-character-limit"></a><span data-ttu-id="3bb57-150">範囲の戻り値の文字制限</span><span class="sxs-lookup"><span data-stu-id="3bb57-150">Range return character limit</span></span>

<span data-ttu-id="3bb57-151">[Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_)メソッドと[Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_)メソッドのアドレス文字列の制限は 8192 文字です。</span><span class="sxs-lookup"><span data-stu-id="3bb57-151">The [Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_) and [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_) methods have an address string limit of 8192 characters.</span></span> <span data-ttu-id="3bb57-152">この制限を超えると、アドレス文字列は 8192 文字に切り捨てられて表示されます。</span><span class="sxs-lookup"><span data-stu-id="3bb57-152">When this limit is exceeded, the address string is truncated to 8192 characters.</span></span>

## <a name="see-also"></a><span data-ttu-id="3bb57-153">関連項目</span><span class="sxs-lookup"><span data-stu-id="3bb57-153">See also</span></span>

- [<span data-ttu-id="3bb57-154">アドインを使用したOfficeエラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="3bb57-154">Troubleshoot development errors with Office Add-ins</span></span>](../testing/troubleshoot-development-errors.md)
- [<span data-ttu-id="3bb57-155">Office アドインでのユーザー エラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="3bb57-155">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
