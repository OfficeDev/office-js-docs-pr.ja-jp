---
title: 一般的な問題と予期しないプラットフォームの動作に関するコーディングガイダンス
description: 開発者がよく遭遇する Office JavaScript API プラットフォームの問題の一覧です。
ms.date: 05/21/2020
localization_priority: Normal
ms.openlocfilehash: d67a069cd2b752be3fca8ce094eaacfd0db08c18
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608384"
---
# <a name="coding-guidance-for-common-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="ca12d-103">一般的な問題と予期しないプラットフォームの動作に関するコーディングガイダンス</span><span class="sxs-lookup"><span data-stu-id="ca12d-103">Coding guidance for common issues and unexpected platform behaviors</span></span>

<span data-ttu-id="ca12d-104">この記事では、予期しない動作が発生するか、必要な結果を得るために特定のコーディングパターンが必要になる可能性がある Office JavaScript API の側面について説明します。</span><span class="sxs-lookup"><span data-stu-id="ca12d-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="ca12d-105">このリストに含まれる問題が発生した場合は、記事の下部にあるフィードバックフォームを使用してお知らせください。</span><span class="sxs-lookup"><span data-stu-id="ca12d-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a><span data-ttu-id="ca12d-106">一般的な Api と Outlook Api は、約束に基づくものではありません</span><span class="sxs-lookup"><span data-stu-id="ca12d-106">Common APIs and Outlook APIs are not promise-based</span></span>

<span data-ttu-id="ca12d-107">[共通 api](/javascript/api/office) (特定の Office ホストに縛られていないもの) と[Outlook api](/javascript/api/outlook)では、コールバックベースのプログラミングモデルが使用されます。</span><span class="sxs-lookup"><span data-stu-id="ca12d-107">The [Common APIs](/javascript/api/office) (those that are not tied to a particular Office host) and [Outlook APIs](/javascript/api/outlook) use a callback-based programming model.</span></span> <span data-ttu-id="ca12d-108">基になる Office ドキュメントと対話するには、操作が完了したときに実行されるコールバックを指定する非同期の読み取りまたは書き込みの呼び出しが必要です。</span><span class="sxs-lookup"><span data-stu-id="ca12d-108">Interacting with the underlying Office document requires an asynchronous read or write call that specifies a callback to be ran when the operation completes.</span></span> <span data-ttu-id="ca12d-109">このパターンの例については、「 [getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ca12d-109">For an example of this pattern, see [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span>

<span data-ttu-id="ca12d-110">これらの共通 API および Outlook API メソッドは、[約束](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)を返しません。</span><span class="sxs-lookup"><span data-stu-id="ca12d-110">These Common API and Outlook API methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span></span> <span data-ttu-id="ca12d-111">そのため、[待機](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await)を使用して、非同期操作が完了するまで実行を一時停止することはできません。</span><span class="sxs-lookup"><span data-stu-id="ca12d-111">Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes.</span></span> <span data-ttu-id="ca12d-112">振る舞いが必要な場合は `await` 、明示的に作成した約束でメソッドの呼び出しをラップすることができます。</span><span class="sxs-lookup"><span data-stu-id="ca12d-112">If you need `await` behavior, you can wrap the method call in an explicitly created Promise.</span></span>

```js
readDocumentFileAsync(): Promise<any> {
    return new Promise((resolve, reject) => {
        const chunkSize = 65536;
        const self = this;

        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: chunkSize }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(asyncResult.error);
            } else {
                // `getAllSlices` is a Promise-wrapped implementation of File.getSliceAsync.
                self.getAllSlices(asyncResult.value).then(result => {
                    if (result.IsSuccess) {
                        resolve(result.Data);
                    } else {
                        reject(asyncResult.error);
                    }
                });
            }
        });
    });
}
```

> [!NOTE]
> <span data-ttu-id="ca12d-113">参照ドキュメントには、 [getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-)の Promise ラップによる実装が含まれています。</span><span class="sxs-lookup"><span data-stu-id="ca12d-113">The reference documentation contains the Promise-wrapped implementation of [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span></span>

## <a name="some-properties-cannot-be-set-directly"></a><span data-ttu-id="ca12d-114">一部のプロパティは直接設定できません</span><span class="sxs-lookup"><span data-stu-id="ca12d-114">Some properties cannot be set directly</span></span>

> [!NOTE]
> <span data-ttu-id="ca12d-115">このセクションは、Excel および Word のホスト固有の Api にのみ適用されます。</span><span class="sxs-lookup"><span data-stu-id="ca12d-115">This section only applies to the host-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="ca12d-116">書き込み可能であっても、一部のプロパティを設定することはできません。</span><span class="sxs-lookup"><span data-stu-id="ca12d-116">Some properties cannot be set, despite being writable.</span></span> <span data-ttu-id="ca12d-117">これらのプロパティは、1つのオブジェクトとして設定する必要がある親プロパティの一部です。</span><span class="sxs-lookup"><span data-stu-id="ca12d-117">These properties are part of a parent property that must be set as a single object.</span></span> <span data-ttu-id="ca12d-118">これは、親プロパティが特定の論理的な関係を持つサブプロパティに依存しているためです。</span><span class="sxs-lookup"><span data-stu-id="ca12d-118">This is because that parent property relies on the subproperties having specific, logical relationships.</span></span> <span data-ttu-id="ca12d-119">このような親プロパティは、オブジェクトの個々のサブプロパティを設定するのではなく、オブジェクトリテラル表記を使用して設定し、オブジェクト全体を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ca12d-119">These parent properties must be set using object literal notation to set the entire object, instead of setting that object's individual subproperties.</span></span> <span data-ttu-id="ca12d-120">この例の1つは、 [PageLayout](/javascript/api/excel/excel.pagelayout)にあります。</span><span class="sxs-lookup"><span data-stu-id="ca12d-120">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="ca12d-121">このプロパティは、次 `zoom` に示すように、1つの[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)オブジェクトで設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ca12d-121">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="ca12d-122">前の例では、値を直接割り当てることはでき***ません***。 `zoom` `sheet.pageLayout.zoom.scale = 200;`</span><span class="sxs-lookup"><span data-stu-id="ca12d-122">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="ca12d-123">が読み込まれていないため、このステートメントはエラーをスロー `zoom` します。</span><span class="sxs-lookup"><span data-stu-id="ca12d-123">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="ca12d-124">`zoom`ロードされた場合でも、スケールのセットは有効になりません。</span><span class="sxs-lookup"><span data-stu-id="ca12d-124">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="ca12d-125">すべてのコンテキスト操作が行われ `zoom` 、アドイン内のプロキシオブジェクトが更新され、ローカルに設定された値が上書きされます。</span><span class="sxs-lookup"><span data-stu-id="ca12d-125">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="ca12d-126">この動作は、[範囲形式](/javascript/api/excel/excel.range#format)などの[ナビゲーションプロパティ](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties)とは異なります。</span><span class="sxs-lookup"><span data-stu-id="ca12d-126">This behavior differs from [navigational properties](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="ca12d-127">のプロパティは `format` 、次に示すように、object ナビゲーションを使用して設定できます。</span><span class="sxs-lookup"><span data-stu-id="ca12d-127">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="ca12d-128">読み取り専用修飾子をチェックすることによって、そのサブプロパティを直接設定できないプロパティを識別できます。</span><span class="sxs-lookup"><span data-stu-id="ca12d-128">You can identify a property that cannot have its subproperties directly set by checking its read-only modifier.</span></span> <span data-ttu-id="ca12d-129">読み取り専用のプロパティは、読み取り専用でないサブプロパティを直接設定することができます。</span><span class="sxs-lookup"><span data-stu-id="ca12d-129">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="ca12d-130">書き込み可能なプロパティ `PageLayout.zoom` は、そのレベルのオブジェクトで設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ca12d-130">Writeable properties like `PageLayout.zoom` must be set with an object at that level.</span></span> <span data-ttu-id="ca12d-131">概要:</span><span class="sxs-lookup"><span data-stu-id="ca12d-131">In summary:</span></span>

- <span data-ttu-id="ca12d-132">読み取り専用プロパティ: サブプロパティは、ナビゲーションを使用して設定できます。</span><span class="sxs-lookup"><span data-stu-id="ca12d-132">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="ca12d-133">書き込み可能なプロパティ: ナビゲーションを使用してサブプロパティを設定することはできません (最初の親オブジェクトの割り当ての一部として設定する必要があります)。</span><span class="sxs-lookup"><span data-stu-id="ca12d-133">Writable property: Subproperties cannot be set through navigation (must be set as part of the initial parent object assignment).</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="ca12d-134">読み取り専用プロパティの設定</span><span class="sxs-lookup"><span data-stu-id="ca12d-134">Setting read-only properties</span></span>

<span data-ttu-id="ca12d-135">Office JS の[TypeScript 定義](referencing-the-javascript-api-for-office-library-from-its-cdn.md)は、読み取り専用のオブジェクトプロパティを指定します。</span><span class="sxs-lookup"><span data-stu-id="ca12d-135">The [TypeScript definitions](referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="ca12d-136">読み取り専用プロパティを設定しようとすると、エラーがスローされずに書き込み操作が失敗します。</span><span class="sxs-lookup"><span data-stu-id="ca12d-136">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="ca12d-137">次の例では、誤って読み取り専用プロパティ[Chart.id](/javascript/api/excel/excel.chart#id)を設定しようとしています。</span><span class="sxs-lookup"><span data-stu-id="ca12d-137">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="removing-event-handlers"></a><span data-ttu-id="ca12d-138">イベントハンドラーの削除</span><span class="sxs-lookup"><span data-stu-id="ca12d-138">Removing event handlers</span></span>

<span data-ttu-id="ca12d-139">イベントハンドラーは、追加したものと同じものを使用して削除する必要があり `RequestContext` ます。</span><span class="sxs-lookup"><span data-stu-id="ca12d-139">Event handlers must be removed using the same `RequestContext` in which they were added.</span></span> <span data-ttu-id="ca12d-140">実行中にアドインでイベントハンドラーを削除する必要がある場合は、ハンドラーを追加するために使用されるコンテキストオブジェクトを格納する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ca12d-140">If you need your add-in to remove an event handler while running, you'll need to store the context object used to add the handler.</span></span>

```js
Excel.run(async (context) => {
    [...]

    // To later remove an event handler, store the context somewhere accessible to the handler removal function.
    // You may find it helpful to also store the event handler object and associate it with the context.
    selectionChangedHandler = myWorksheet.onSelectionChanged.add(callback);
    savedContext = currentContext;
    return context.sync();
}
```

## <a name="supporting-internet-explorer"></a><span data-ttu-id="ca12d-141">Internet Explorer のサポート</span><span class="sxs-lookup"><span data-stu-id="ca12d-141">Supporting Internet Explorer</span></span>

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="excel-specific-issues"></a><span data-ttu-id="ca12d-142">Excel 固有の問題</span><span class="sxs-lookup"><span data-stu-id="ca12d-142">Excel-specific issues</span></span>

### <a name="excel-data-transfer-limits"></a><span data-ttu-id="ca12d-143">Excel データ転送の制限</span><span class="sxs-lookup"><span data-stu-id="ca12d-143">Excel data transfer limits</span></span>

<span data-ttu-id="ca12d-144">Excel アドインを作成している場合は、ブックを操作するときに以下のサイズ制限に注意してください。</span><span class="sxs-lookup"><span data-stu-id="ca12d-144">If you're building an Excel add-in, be aware of the following size limitations when interacting with the workbook:</span></span>

- <span data-ttu-id="ca12d-145">Excel on the web ではペイロードのサイズが要求と応答で 5 MB に制限されています。</span><span class="sxs-lookup"><span data-stu-id="ca12d-145">Excel on the web has a payload size limit for requests and responses of 5MB.</span></span> <span data-ttu-id="ca12d-146">その制限を超えると、`RichAPI.Error` がスローされます。</span><span class="sxs-lookup"><span data-stu-id="ca12d-146">`RichAPI.Error` will be thrown if that limit is exceeded.</span></span>
- <span data-ttu-id="ca12d-147">範囲は、取得操作に500万のセルに制限されます。</span><span class="sxs-lookup"><span data-stu-id="ca12d-147">A range is limited to five million cells for get operations.</span></span>

<span data-ttu-id="ca12d-148">ユーザー入力がこれらの制限を超えていることが予想される場合は、必ずデータを確認してから、を呼び出してください `context.sync()` 。</span><span class="sxs-lookup"><span data-stu-id="ca12d-148">If you expect user input to exceed these limits, be sure to check the data before calling `context.sync()`.</span></span> <span data-ttu-id="ca12d-149">必要に応じて、操作を小さな部分に分割します。</span><span class="sxs-lookup"><span data-stu-id="ca12d-149">Split the operation into smaller pieces as needed.</span></span> <span data-ttu-id="ca12d-150">`context.sync()`各サブ操作を呼び出して、それらの操作が再度一括されないようにしてください。</span><span class="sxs-lookup"><span data-stu-id="ca12d-150">Be sure to call `context.sync()` for each sub-operation to avoid those operations getting batched together again.</span></span>

<span data-ttu-id="ca12d-151">これらの制限は、通常、大きな範囲を超えています。</span><span class="sxs-lookup"><span data-stu-id="ca12d-151">These limitations are typically exceeded by large ranges.</span></span> <span data-ttu-id="ca12d-152">アドインでは、範囲内のセルを戦略的に更新するために[Rangeareas](/javascript/api/excel/excel.rangeareas)を使用できる場合があります。</span><span class="sxs-lookup"><span data-stu-id="ca12d-152">Your add-in might be able to use [RangeAreas](/javascript/api/excel/excel.rangeareas) to strategically update cells within a larger range.</span></span> <span data-ttu-id="ca12d-153">詳細については、「 [Excel アドインで複数の範囲を同時に操作](../excel/excel-add-ins-multiple-ranges.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ca12d-153">See [Work with multiple ranges simultaneously in Excel add-ins](../excel/excel-add-ins-multiple-ranges.md) for more information.</span></span>

### <a name="api-limitations-when-the-active-workbook-switches"></a><span data-ttu-id="ca12d-154">アクティブなブックの切り替え時の API の制限</span><span class="sxs-lookup"><span data-stu-id="ca12d-154">API limitations when the active workbook switches</span></span>

<span data-ttu-id="ca12d-155">Excel 用のアドインは、一度に1つのブックを操作することを目的としています。</span><span class="sxs-lookup"><span data-stu-id="ca12d-155">Add-ins for Excel are intended to operate on a single workbook at a time.</span></span> <span data-ttu-id="ca12d-156">アドインを実行しているブックとは別のブックがフォーカスを取得すると、エラーが発生することがあります。</span><span class="sxs-lookup"><span data-stu-id="ca12d-156">Errors can arise when a workbook that is separate from the one running the add-in gains focus.</span></span> <span data-ttu-id="ca12d-157">これは、フォーカスが変更されたときに、特定のメソッドが呼び出されたときにのみ発生します。</span><span class="sxs-lookup"><span data-stu-id="ca12d-157">This only happens when particular methods are in the process of being called when the focus changes.</span></span>

<span data-ttu-id="ca12d-158">このブックスイッチの影響を受ける Api は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="ca12d-158">The following APIs are affected by this workbook switch:</span></span>

|<span data-ttu-id="ca12d-159">Excel JavaScript API</span><span class="sxs-lookup"><span data-stu-id="ca12d-159">Excel JavaScript API</span></span> | <span data-ttu-id="ca12d-160">スローされたエラー</span><span class="sxs-lookup"><span data-stu-id="ca12d-160">Error thrown</span></span> |
|--|--|
| `Chart.activate` | <span data-ttu-id="ca12d-161">GeneralException</span><span class="sxs-lookup"><span data-stu-id="ca12d-161">GeneralException</span></span> |
| `Range.select` | <span data-ttu-id="ca12d-162">GeneralException</span><span class="sxs-lookup"><span data-stu-id="ca12d-162">GeneralException</span></span> |
| `Table.clearFilters` | <span data-ttu-id="ca12d-163">GeneralException</span><span class="sxs-lookup"><span data-stu-id="ca12d-163">GeneralException</span></span> |
| `Workbook.getActiveCell`  | <span data-ttu-id="ca12d-164">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="ca12d-164">InvalidSelection</span></span>|
| `Workbook.getSelectedRange` | <span data-ttu-id="ca12d-165">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="ca12d-165">InvalidSelection</span></span>|
| `Workbook.getSelectedRanges`  | <span data-ttu-id="ca12d-166">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="ca12d-166">InvalidSelection</span></span>|
| `Worksheet.activate` | <span data-ttu-id="ca12d-167">GeneralException</span><span class="sxs-lookup"><span data-stu-id="ca12d-167">GeneralException</span></span> |
| `Worksheet.delete`  | <span data-ttu-id="ca12d-168">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="ca12d-168">InvalidSelection</span></span>|
| `Worksheet.gridlines` | <span data-ttu-id="ca12d-169">GeneralException</span><span class="sxs-lookup"><span data-stu-id="ca12d-169">GeneralException</span></span> |
| `Worksheet.showHeadings` | <span data-ttu-id="ca12d-170">GeneralException</span><span class="sxs-lookup"><span data-stu-id="ca12d-170">GeneralException</span></span> |
| `WorksheetCollection.add` | <span data-ttu-id="ca12d-171">GeneralException</span><span class="sxs-lookup"><span data-stu-id="ca12d-171">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeAt` | <span data-ttu-id="ca12d-172">GeneralException</span><span class="sxs-lookup"><span data-stu-id="ca12d-172">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeColumns` | <span data-ttu-id="ca12d-173">GeneralException</span><span class="sxs-lookup"><span data-stu-id="ca12d-173">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeRows` | <span data-ttu-id="ca12d-174">GeneralException</span><span class="sxs-lookup"><span data-stu-id="ca12d-174">GeneralException</span></span> |
| `WorksheetFreezePanes.getLocationOrNullObject`| <span data-ttu-id="ca12d-175">GeneralException</span><span class="sxs-lookup"><span data-stu-id="ca12d-175">GeneralException</span></span> |
| `WorksheetFreezePanes.unfreeze` | <span data-ttu-id="ca12d-176">GeneralException</span><span class="sxs-lookup"><span data-stu-id="ca12d-176">GeneralException</span></span> |

> [!NOTE]
> <span data-ttu-id="ca12d-177">これは、Windows または Mac で開いている複数の Excel ブックにのみ適用されます。</span><span class="sxs-lookup"><span data-stu-id="ca12d-177">This only applies to multiple Excel workbooks open on Windows or Mac.</span></span>

## <a name="see-also"></a><span data-ttu-id="ca12d-178">関連項目</span><span class="sxs-lookup"><span data-stu-id="ca12d-178">See also</span></span>

- <span data-ttu-id="ca12d-179">[Officedev/office-js](https://github.com/OfficeDev/office-js/issues): office アドインプラットフォームおよび JavaScript api の問題を報告および表示する場所です。</span><span class="sxs-lookup"><span data-stu-id="ca12d-179">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="ca12d-180">[スタックオーバーフロー](https://stackoverflow.com/questions/tagged/office-js): Office JavaScript api に関するプログラミング上の問題を確認および表示する場所です。</span><span class="sxs-lookup"><span data-stu-id="ca12d-180">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="ca12d-181">スタックオーバーフローに投稿するときには、必ず "office-js" タグを質問に適用してください。</span><span class="sxs-lookup"><span data-stu-id="ca12d-181">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="ca12d-182">[UserVoice](https://officespdev.uservoice.com/): office アドインプラットフォームおよび Office JavaScript api の新機能を提案する場所です。</span><span class="sxs-lookup"><span data-stu-id="ca12d-182">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
