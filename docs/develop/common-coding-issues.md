---
title: 一般的なコーディングの問題と予期しないプラットフォームの動作
description: 開発者がよく遭遇する Office JavaScript API プラットフォームの問題の一覧です。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: d39c379961833cdb924628becf2c2da3f7e271b9
ms.sourcegitcommit: 59d29d01bce7543ebebf86e5a86db00cf54ca14a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/01/2019
ms.locfileid: "37924795"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="973cf-103">一般的なコーディングの問題と予期しないプラットフォームの動作</span><span class="sxs-lookup"><span data-stu-id="973cf-103">Common coding issues and unexpected platform behaviors</span></span>

<span data-ttu-id="973cf-104">この記事では、予期しない動作が発生するか、必要な結果を得るために特定のコーディングパターンが必要になる可能性がある Office JavaScript API の側面について説明します。</span><span class="sxs-lookup"><span data-stu-id="973cf-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="973cf-105">このリストに含まれる問題が発生した場合は、記事の下部にあるフィードバックフォームを使用してお知らせください。</span><span class="sxs-lookup"><span data-stu-id="973cf-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a><span data-ttu-id="973cf-106">一般的な Api と Outlook Api は、約束に基づくものではありません</span><span class="sxs-lookup"><span data-stu-id="973cf-106">Common APIs and Outlook APIs are not promise-based</span></span>

<span data-ttu-id="973cf-107">[共通 api](/javascript/api/office) (特定の Office ホストに縛られていないもの) と[Outlook api](/javascript/api/outlook)では、コールバックベースのプログラミングモデルが使用されます。</span><span class="sxs-lookup"><span data-stu-id="973cf-107">The [Common APIs](/javascript/api/office) (those that are not tied to a particular Office host) and [Outlook APIs](/javascript/api/outlook) use a callback-based programming model.</span></span> <span data-ttu-id="973cf-108">基になる Office ドキュメントと対話するには、操作が完了したときに実行されるコールバックを指定する非同期の読み取りまたは書き込みの呼び出しが必要です。</span><span class="sxs-lookup"><span data-stu-id="973cf-108">Interacting with the underlying Office document requires an asynchronous read or write call that specifies a callback to be ran when the operation completes.</span></span> <span data-ttu-id="973cf-109">このパターンの例については、「 [getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="973cf-109">For an example of this pattern, see [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span>

<span data-ttu-id="973cf-110">これらの共通 API および Outlook API メソッドは、[約束](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)を返しません。</span><span class="sxs-lookup"><span data-stu-id="973cf-110">These Common API and Outlook API methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span></span> <span data-ttu-id="973cf-111">そのため、[待機](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await)を使用して、非同期操作が完了するまで実行を一時停止することはできません。</span><span class="sxs-lookup"><span data-stu-id="973cf-111">Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes.</span></span> <span data-ttu-id="973cf-112">振る舞いが必要`await`な場合は、明示的に作成した約束でメソッドの呼び出しをラップすることができます。</span><span class="sxs-lookup"><span data-stu-id="973cf-112">If you need `await` behavior, you can wrap the method call in an explicitly created Promise.</span></span>

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
> <span data-ttu-id="973cf-113">参照ドキュメントには、 [getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-)の Promise ラップによる実装が含まれています。</span><span class="sxs-lookup"><span data-stu-id="973cf-113">The reference documentation contains the Promise-wrapped implementation of [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span></span>

## <a name="some-properties-must-be-set-with-json-structs"></a><span data-ttu-id="973cf-114">一部のプロパティは、JSON 構造体で設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="973cf-114">Some properties must be set with JSON structs</span></span>

> [!NOTE]
> <span data-ttu-id="973cf-115">このセクションは、Excel および Word のホスト固有の Api にのみ適用されます。</span><span class="sxs-lookup"><span data-stu-id="973cf-115">This section only applies to the host-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="973cf-116">一部のプロパティは、個々のサブプロパティを設定するのではなく、JSON 構造体として設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="973cf-116">Some properties must be set as JSON structs, instead of setting their individual subproperties.</span></span> <span data-ttu-id="973cf-117">この例の1つは、 [PageLayout](/javascript/api/excel/excel.pagelayout)にあります。</span><span class="sxs-lookup"><span data-stu-id="973cf-117">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="973cf-118">この`zoom`プロパティは、次に示すように、1つの[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)オブジェクトで設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="973cf-118">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="973cf-119">前の例では、値`zoom` `sheet.pageLayout.zoom.scale = 200;`を直接割り当てることはでき***ません***。</span><span class="sxs-lookup"><span data-stu-id="973cf-119">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="973cf-120">が読み込まれてい`zoom`ないため、このステートメントはエラーをスローします。</span><span class="sxs-lookup"><span data-stu-id="973cf-120">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="973cf-121">ロードさ`zoom`れた場合でも、スケールのセットは有効になりません。</span><span class="sxs-lookup"><span data-stu-id="973cf-121">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="973cf-122">すべての`zoom`コンテキスト操作が行われ、アドイン内のプロキシオブジェクトが更新され、ローカルに設定された値が上書きされます。</span><span class="sxs-lookup"><span data-stu-id="973cf-122">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="973cf-123">この動作は、[範囲形式](/javascript/api/excel/excel.range#format)などの[ナビゲーションプロパティ](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties)とは異なります。</span><span class="sxs-lookup"><span data-stu-id="973cf-123">This behavior differs from [navigational properties](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="973cf-124">の`format`プロパティは、次に示すように、object ナビゲーションを使用して設定できます。</span><span class="sxs-lookup"><span data-stu-id="973cf-124">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="973cf-125">読み取り専用修飾子をチェックすることで、そのサブプロパティを JSON 構造体で設定する必要があるプロパティを識別できます。</span><span class="sxs-lookup"><span data-stu-id="973cf-125">You can identify a property that must have its subproperties set with a JSON struct by checking its read-only modifier.</span></span> <span data-ttu-id="973cf-126">読み取り専用のプロパティは、読み取り専用でないサブプロパティを直接設定することができます。</span><span class="sxs-lookup"><span data-stu-id="973cf-126">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="973cf-127">書き込み可能な`PageLayout.zoom`プロパティは、JSON 構造体で設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="973cf-127">Writeable properties like `PageLayout.zoom` must be set with a JSON struct.</span></span> <span data-ttu-id="973cf-128">概要:</span><span class="sxs-lookup"><span data-stu-id="973cf-128">In summary:</span></span>

- <span data-ttu-id="973cf-129">読み取り専用プロパティ: サブプロパティは、ナビゲーションを使用して設定できます。</span><span class="sxs-lookup"><span data-stu-id="973cf-129">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="973cf-130">書き込み可能なプロパティ: サブプロパティは JSON 構造体で設定する必要があります (ナビゲーションで設定することはできません)。</span><span class="sxs-lookup"><span data-stu-id="973cf-130">Writable property: Subproperties must be set with a JSON struct (and cannot be set through navigation).</span></span>

## <a name="excel-range-limits"></a><span data-ttu-id="973cf-131">Excel 範囲制限</span><span class="sxs-lookup"><span data-stu-id="973cf-131">Excel Range limits</span></span>

<span data-ttu-id="973cf-132">範囲を使用する Excel アドインを作成している場合は、次のサイズ制限に注意してください。</span><span class="sxs-lookup"><span data-stu-id="973cf-132">If you're building an Excel add-in that uses ranges, be aware of the following size limitations:</span></span>

- <span data-ttu-id="973cf-133">Excel on the web ではペイロードのサイズが要求と応答で 5 MB に制限されています。</span><span class="sxs-lookup"><span data-stu-id="973cf-133">Excel on the web has a payload size limit for requests and responses of 5MB.</span></span> <span data-ttu-id="973cf-134">その制限を超えると、`RichAPI.Error` がスローされます。</span><span class="sxs-lookup"><span data-stu-id="973cf-134">`RichAPI.Error` will be thrown if that limit is exceeded.</span></span>
- <span data-ttu-id="973cf-135">範囲は、設定操作では500万のセルに制限されます。</span><span class="sxs-lookup"><span data-stu-id="973cf-135">A range is limited to five million cells for set operations.</span></span>

<span data-ttu-id="973cf-136">ユーザー入力がこれらの制限を超えていることが予想される場合は、データをチェックして、範囲を複数のオブジェクトに分割します。</span><span class="sxs-lookup"><span data-stu-id="973cf-136">If you expect user input to exceed these limits, be sure to check the data and split the ranges into multiple objects.</span></span> <span data-ttu-id="973cf-137">また、複数`context.sync()`の呼び出しを送信して、より小さな範囲の操作が同時に一括されないようにする必要もあります。</span><span class="sxs-lookup"><span data-stu-id="973cf-137">You'll also need to submit multiple `context.sync()` calls to avoid the smaller range operations getting batched together again.</span></span>

<span data-ttu-id="973cf-138">アドインでは、範囲内のセルを戦略的に更新するために[Rangeareas](/javascript/api/excel/excel.rangeareas)を使用できる場合があります。</span><span class="sxs-lookup"><span data-stu-id="973cf-138">Your add-in might be able to use [RangeAreas](/javascript/api/excel/excel.rangeareas) to strategically update cells within a larger range.</span></span> <span data-ttu-id="973cf-139">詳細については、「 [Excel アドインで複数の範囲を同時に操作](../excel/excel-add-ins-multiple-ranges.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="973cf-139">See [Work with multiple ranges simultaneously in Excel add-ins](../excel/excel-add-ins-multiple-ranges.md) for more information.</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="973cf-140">読み取り専用プロパティの設定</span><span class="sxs-lookup"><span data-stu-id="973cf-140">Setting read-only properties</span></span>

<span data-ttu-id="973cf-141">Office JS の[TypeScript 定義](/referencing-the-javascript-api-for-office-library-from-its-cdn.md)は、読み取り専用のオブジェクトプロパティを指定します。</span><span class="sxs-lookup"><span data-stu-id="973cf-141">The [TypeScript definitions](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="973cf-142">読み取り専用プロパティを設定しようとすると、エラーがスローされずに書き込み操作が失敗します。</span><span class="sxs-lookup"><span data-stu-id="973cf-142">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="973cf-143">次の例では、誤って読み取り専用プロパティ[Chart.id](/javascript/api/excel/excel.chart#id)を設定しようとしています。</span><span class="sxs-lookup"><span data-stu-id="973cf-143">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="see-also"></a><span data-ttu-id="973cf-144">関連項目</span><span class="sxs-lookup"><span data-stu-id="973cf-144">See also</span></span>

- <span data-ttu-id="973cf-145">[Officedev/office-js](https://github.com/OfficeDev/office-js/issues): office アドインプラットフォームおよび JavaScript api の問題を報告および表示する場所です。</span><span class="sxs-lookup"><span data-stu-id="973cf-145">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="973cf-146">[スタックオーバーフロー](https://stackoverflow.com/questions/tagged/office-js): Office JavaScript api に関するプログラミング上の問題を確認および表示する場所です。</span><span class="sxs-lookup"><span data-stu-id="973cf-146">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="973cf-147">スタックオーバーフローに投稿するときには、必ず "office-js" タグを質問に適用してください。</span><span class="sxs-lookup"><span data-stu-id="973cf-147">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="973cf-148">[UserVoice](https://officespdev.uservoice.com/): office アドインプラットフォームおよび Office JavaScript api の新機能を提案する場所です。</span><span class="sxs-lookup"><span data-stu-id="973cf-148">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
