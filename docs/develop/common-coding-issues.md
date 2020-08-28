---
title: 一般的な問題と予期しないプラットフォームの動作に関するコーディングガイダンス
description: 開発者がよく遭遇する Office JavaScript API プラットフォームの問題の一覧です。
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: f6d6a31059b32550e3176ed278d7da4c2c7a6c68
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292912"
---
# <a name="coding-guidance-for-common-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="fb0f9-103">一般的な問題と予期しないプラットフォームの動作に関するコーディングガイダンス</span><span class="sxs-lookup"><span data-stu-id="fb0f9-103">Coding guidance for common issues and unexpected platform behaviors</span></span>

<span data-ttu-id="fb0f9-104">この記事では、予期しない動作が発生するか、必要な結果を得るために特定のコーディングパターンが必要になる可能性がある Office JavaScript API の側面について説明します。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="fb0f9-105">このリストに含まれる問題が発生した場合は、記事の下部にあるフィードバックフォームを使用してお知らせください。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a><span data-ttu-id="fb0f9-106">一般的な Api と Outlook Api は、約束に基づくものではありません</span><span class="sxs-lookup"><span data-stu-id="fb0f9-106">Common APIs and Outlook APIs are not promise-based</span></span>

<span data-ttu-id="fb0f9-107">[共通 api](/javascript/api/office) (特定の Office アプリケーションに縛られていないもの) と[Outlook api](/javascript/api/outlook)では、コールバックベースのプログラミングモデルが使用されます。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-107">The [Common APIs](/javascript/api/office) (those that are not tied to a particular Office application) and [Outlook APIs](/javascript/api/outlook) use a callback-based programming model.</span></span> <span data-ttu-id="fb0f9-108">基になる Office ドキュメントと対話するには、操作が完了したときに実行されるコールバックを指定する非同期の読み取りまたは書き込みの呼び出しが必要です。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-108">Interacting with the underlying Office document requires an asynchronous read or write call that specifies a callback to be run when the operation completes.</span></span> <span data-ttu-id="fb0f9-109">このパターンの例については、「 [getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-109">For an example of this pattern, see [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span>

<span data-ttu-id="fb0f9-110">これらの共通 API および Outlook API メソッドは、 [約束](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)を返しません。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-110">These Common API and Outlook API methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span></span> <span data-ttu-id="fb0f9-111">そのため、 [待機](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) を使用して、非同期操作が完了するまで実行を一時停止することはできません。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-111">Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes.</span></span> <span data-ttu-id="fb0f9-112">振る舞いが必要な場合は `await` 、明示的に作成した約束でメソッドの呼び出しをラップすることができます。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-112">If you need `await` behavior, you can wrap the method call in an explicitly created Promise.</span></span>

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
> <span data-ttu-id="fb0f9-113">参照ドキュメントには、 [getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-)の Promise ラップによる実装が含まれています。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-113">The reference documentation contains the Promise-wrapped implementation of [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span></span>

## <a name="some-properties-cannot-be-set-directly"></a><span data-ttu-id="fb0f9-114">一部のプロパティは直接設定できません</span><span class="sxs-lookup"><span data-stu-id="fb0f9-114">Some properties cannot be set directly</span></span>

> [!NOTE]
> <span data-ttu-id="fb0f9-115">このセクションは、Excel および Word 用のアプリケーション固有の Api にのみ適用されます。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-115">This section only applies to the application-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="fb0f9-116">書き込み可能であっても、一部のプロパティを設定することはできません。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-116">Some properties cannot be set, despite being writable.</span></span> <span data-ttu-id="fb0f9-117">これらのプロパティは、1つのオブジェクトとして設定する必要がある親プロパティの一部です。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-117">These properties are part of a parent property that must be set as a single object.</span></span> <span data-ttu-id="fb0f9-118">これは、親プロパティが特定の論理的な関係を持つサブプロパティに依存しているためです。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-118">This is because that parent property relies on the subproperties having specific, logical relationships.</span></span> <span data-ttu-id="fb0f9-119">このような親プロパティは、オブジェクトの個々のサブプロパティを設定するのではなく、オブジェクトリテラル表記を使用して設定し、オブジェクト全体を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-119">These parent properties must be set using object literal notation to set the entire object, instead of setting that object's individual subproperties.</span></span> <span data-ttu-id="fb0f9-120">この例の1つは、 [PageLayout](/javascript/api/excel/excel.pagelayout)にあります。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-120">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="fb0f9-121">このプロパティは、次 `zoom` に示すように、1つの [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) オブジェクトで設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-121">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="fb0f9-122">前の例では、値を直接割り当てることはでき***ません***。 `zoom` `sheet.pageLayout.zoom.scale = 200;`</span><span class="sxs-lookup"><span data-stu-id="fb0f9-122">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="fb0f9-123">が読み込まれていないため、このステートメントはエラーをスロー `zoom` します。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-123">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="fb0f9-124">`zoom`ロードされた場合でも、スケールのセットは有効になりません。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-124">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="fb0f9-125">すべてのコンテキスト操作が行われ `zoom` 、アドイン内のプロキシオブジェクトが更新され、ローカルに設定された値が上書きされます。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-125">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="fb0f9-126">この動作は、[範囲形式](/javascript/api/excel/excel.range#format)などの[ナビゲーションプロパティ](application-specific-api-model.md#scalar-and-navigation-properties)とは異なります。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-126">This behavior differs from [navigational properties](application-specific-api-model.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="fb0f9-127">のプロパティは `format` 、次に示すように、object ナビゲーションを使用して設定できます。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-127">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="fb0f9-128">読み取り専用修飾子をチェックすることによって、そのサブプロパティを直接設定できないプロパティを識別できます。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-128">You can identify a property that cannot have its subproperties directly set by checking its read-only modifier.</span></span> <span data-ttu-id="fb0f9-129">読み取り専用のプロパティは、読み取り専用でないサブプロパティを直接設定することができます。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-129">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="fb0f9-130">書き込み可能なプロパティ `PageLayout.zoom` は、そのレベルのオブジェクトで設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-130">Writeable properties like `PageLayout.zoom` must be set with an object at that level.</span></span> <span data-ttu-id="fb0f9-131">概要:</span><span class="sxs-lookup"><span data-stu-id="fb0f9-131">In summary:</span></span>

- <span data-ttu-id="fb0f9-132">読み取り専用プロパティ: サブプロパティは、ナビゲーションを使用して設定できます。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-132">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="fb0f9-133">書き込み可能なプロパティ: ナビゲーションを使用してサブプロパティを設定することはできません (最初の親オブジェクトの割り当ての一部として設定する必要があります)。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-133">Writable property: Subproperties cannot be set through navigation (must be set as part of the initial parent object assignment).</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="fb0f9-134">読み取り専用プロパティの設定</span><span class="sxs-lookup"><span data-stu-id="fb0f9-134">Setting read-only properties</span></span>

<span data-ttu-id="fb0f9-135">Office JS の [TypeScript 定義](referencing-the-javascript-api-for-office-library-from-its-cdn.md) は、読み取り専用のオブジェクトプロパティを指定します。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-135">The [TypeScript definitions](referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="fb0f9-136">読み取り専用プロパティを設定しようとすると、エラーがスローされずに書き込み操作が失敗します。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-136">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="fb0f9-137">次の例では、誤って読み取り専用プロパティ [Chart.id](/javascript/api/excel/excel.chart#id)を設定しようとしています。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-137">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="removing-event-handlers"></a><span data-ttu-id="fb0f9-138">イベントハンドラーの削除</span><span class="sxs-lookup"><span data-stu-id="fb0f9-138">Removing event handlers</span></span>

<span data-ttu-id="fb0f9-139">イベントハンドラーは、追加したものと同じものを使用して削除する必要があり `RequestContext` ます。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-139">Event handlers must be removed using the same `RequestContext` in which they were added.</span></span> <span data-ttu-id="fb0f9-140">実行中にアドインでイベントハンドラーを削除する必要がある場合は、ハンドラーを追加するために使用されるコンテキストオブジェクトを格納する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-140">If you need your add-in to remove an event handler while running, you'll need to store the context object used to add the handler.</span></span>

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

## <a name="supporting-internet-explorer"></a><span data-ttu-id="fb0f9-141">Internet Explorer のサポート</span><span class="sxs-lookup"><span data-stu-id="fb0f9-141">Supporting Internet Explorer</span></span>

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="excel-specific-issues"></a><span data-ttu-id="fb0f9-142">Excel 固有の問題</span><span class="sxs-lookup"><span data-stu-id="fb0f9-142">Excel-specific issues</span></span>

### <a name="api-limitations-when-the-active-workbook-switches"></a><span data-ttu-id="fb0f9-143">アクティブなブックの切り替え時の API の制限</span><span class="sxs-lookup"><span data-stu-id="fb0f9-143">API limitations when the active workbook switches</span></span>

<span data-ttu-id="fb0f9-144">Excel 用のアドインは、一度に1つのブックを操作することを目的としています。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-144">Add-ins for Excel are intended to operate on a single workbook at a time.</span></span> <span data-ttu-id="fb0f9-145">アドインを実行しているブックとは別のブックがフォーカスを取得すると、エラーが発生することがあります。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-145">Errors can arise when a workbook that is separate from the one running the add-in gains focus.</span></span> <span data-ttu-id="fb0f9-146">これは、フォーカスが変更されたときに、特定のメソッドが呼び出されたときにのみ発生します。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-146">This only happens when particular methods are in the process of being called when the focus changes.</span></span>

<span data-ttu-id="fb0f9-147">このブックスイッチの影響を受ける Api は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-147">The following APIs are affected by this workbook switch:</span></span>

|<span data-ttu-id="fb0f9-148">Excel JavaScript API</span><span class="sxs-lookup"><span data-stu-id="fb0f9-148">Excel JavaScript API</span></span> | <span data-ttu-id="fb0f9-149">スローされたエラー</span><span class="sxs-lookup"><span data-stu-id="fb0f9-149">Error thrown</span></span> |
|--|--|
| `Chart.activate` | <span data-ttu-id="fb0f9-150">GeneralException</span><span class="sxs-lookup"><span data-stu-id="fb0f9-150">GeneralException</span></span> |
| `Range.select` | <span data-ttu-id="fb0f9-151">GeneralException</span><span class="sxs-lookup"><span data-stu-id="fb0f9-151">GeneralException</span></span> |
| `Table.clearFilters` | <span data-ttu-id="fb0f9-152">GeneralException</span><span class="sxs-lookup"><span data-stu-id="fb0f9-152">GeneralException</span></span> |
| `Workbook.getActiveCell`  | <span data-ttu-id="fb0f9-153">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="fb0f9-153">InvalidSelection</span></span>|
| `Workbook.getSelectedRange` | <span data-ttu-id="fb0f9-154">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="fb0f9-154">InvalidSelection</span></span>|
| `Workbook.getSelectedRanges`  | <span data-ttu-id="fb0f9-155">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="fb0f9-155">InvalidSelection</span></span>|
| `Worksheet.activate` | <span data-ttu-id="fb0f9-156">GeneralException</span><span class="sxs-lookup"><span data-stu-id="fb0f9-156">GeneralException</span></span> |
| `Worksheet.delete`  | <span data-ttu-id="fb0f9-157">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="fb0f9-157">InvalidSelection</span></span>|
| `Worksheet.gridlines` | <span data-ttu-id="fb0f9-158">GeneralException</span><span class="sxs-lookup"><span data-stu-id="fb0f9-158">GeneralException</span></span> |
| `Worksheet.showHeadings` | <span data-ttu-id="fb0f9-159">GeneralException</span><span class="sxs-lookup"><span data-stu-id="fb0f9-159">GeneralException</span></span> |
| `WorksheetCollection.add` | <span data-ttu-id="fb0f9-160">GeneralException</span><span class="sxs-lookup"><span data-stu-id="fb0f9-160">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeAt` | <span data-ttu-id="fb0f9-161">GeneralException</span><span class="sxs-lookup"><span data-stu-id="fb0f9-161">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeColumns` | <span data-ttu-id="fb0f9-162">GeneralException</span><span class="sxs-lookup"><span data-stu-id="fb0f9-162">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeRows` | <span data-ttu-id="fb0f9-163">GeneralException</span><span class="sxs-lookup"><span data-stu-id="fb0f9-163">GeneralException</span></span> |
| `WorksheetFreezePanes.getLocationOrNullObject`| <span data-ttu-id="fb0f9-164">GeneralException</span><span class="sxs-lookup"><span data-stu-id="fb0f9-164">GeneralException</span></span> |
| `WorksheetFreezePanes.unfreeze` | <span data-ttu-id="fb0f9-165">GeneralException</span><span class="sxs-lookup"><span data-stu-id="fb0f9-165">GeneralException</span></span> |

> [!NOTE]
> <span data-ttu-id="fb0f9-166">これは、Windows または Mac で開いている複数の Excel ブックにのみ適用されます。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-166">This only applies to multiple Excel workbooks open on Windows or Mac.</span></span>

### <a name="coauthoring"></a><span data-ttu-id="fb0f9-167">共同編集</span><span class="sxs-lookup"><span data-stu-id="fb0f9-167">Coauthoring</span></span>

<span data-ttu-id="fb0f9-168">共同編集環境でイベントと共に使用するパターンについては、「 [Excel アドインの共同編集](../excel/co-authoring-in-excel-add-ins.md) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-168">See [Coauthoring in Excel add-ins](../excel/co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment.</span></span> <span data-ttu-id="fb0f9-169">この記事では、など、特定の Api を使用する場合のマージの競合の可能性についても説明し [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) ます。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-169">The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).</span></span>

## <a name="see-also"></a><span data-ttu-id="fb0f9-170">関連項目</span><span class="sxs-lookup"><span data-stu-id="fb0f9-170">See also</span></span>

- [<span data-ttu-id="fb0f9-171">Office アドインのリソースの制限とパフォーマンスの最適化</span><span class="sxs-lookup"><span data-stu-id="fb0f9-171">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
- <span data-ttu-id="fb0f9-172">[Officedev/office-js](https://github.com/OfficeDev/office-js/issues): office アドインプラットフォームおよび JavaScript api の問題を報告および表示する場所です。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-172">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="fb0f9-173">[スタックオーバーフロー](https://stackoverflow.com/questions/tagged/office-js): Office JavaScript api に関するプログラミング上の問題を確認および表示する場所です。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-173">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="fb0f9-174">スタックオーバーフローに投稿するときには、必ず "office-js" タグを質問に適用してください。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-174">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="fb0f9-175">[UserVoice](https://officespdev.uservoice.com/): office アドインプラットフォームおよび Office JavaScript api の新機能を提案する場所です。</span><span class="sxs-lookup"><span data-stu-id="fb0f9-175">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
