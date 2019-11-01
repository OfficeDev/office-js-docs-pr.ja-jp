---
title: 一般的なコーディングの問題と予期しないプラットフォームの動作
description: 開発者がよく遭遇する Office JavaScript API プラットフォームの問題の一覧です。
ms.date: 10/29/2019
localization_priority: Normal
ms.openlocfilehash: 8cea95e3214585ba8e0b77535916f9c564dde9df
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902184"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="40d5e-103">一般的なコーディングの問題と予期しないプラットフォームの動作</span><span class="sxs-lookup"><span data-stu-id="40d5e-103">Common coding issues and unexpected platform behaviors</span></span>

<span data-ttu-id="40d5e-104">この記事では、予期しない動作が発生するか、必要な結果を得るために特定のコーディングパターンが必要になる可能性がある Office JavaScript API の側面について説明します。</span><span class="sxs-lookup"><span data-stu-id="40d5e-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="40d5e-105">このリストに含まれる問題が発生した場合は、記事の下部にあるフィードバックフォームを使用してお知らせください。</span><span class="sxs-lookup"><span data-stu-id="40d5e-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="some-properties-must-be-set-with-json-structs"></a><span data-ttu-id="40d5e-106">一部のプロパティは、JSON 構造体で設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="40d5e-106">Some properties must be set with JSON structs</span></span>

> [!NOTE]
> <span data-ttu-id="40d5e-107">このセクションは、Excel および Word のホスト固有の Api にのみ適用されます。</span><span class="sxs-lookup"><span data-stu-id="40d5e-107">This section only applies to the host-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="40d5e-108">一部のプロパティは、個々のサブプロパティを設定するのではなく、JSON 構造体として設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="40d5e-108">Some properties must be set as JSON structs, instead of setting their individual subproperties.</span></span> <span data-ttu-id="40d5e-109">この例の1つは、 [PageLayout](/javascript/api/excel/excel.pagelayout)にあります。</span><span class="sxs-lookup"><span data-stu-id="40d5e-109">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="40d5e-110">この`zoom`プロパティは、次に示すように、1つの[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)オブジェクトで設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="40d5e-110">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="40d5e-111">前の例では、値`zoom` `sheet.pageLayout.zoom.scale = 200;`を直接割り当てることはでき***ません***。</span><span class="sxs-lookup"><span data-stu-id="40d5e-111">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="40d5e-112">が読み込まれてい`zoom`ないため、このステートメントはエラーをスローします。</span><span class="sxs-lookup"><span data-stu-id="40d5e-112">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="40d5e-113">ロードさ`zoom`れた場合でも、スケールのセットは有効になりません。</span><span class="sxs-lookup"><span data-stu-id="40d5e-113">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="40d5e-114">すべての`zoom`コンテキスト操作が行われ、アドイン内のプロキシオブジェクトが更新され、ローカルに設定された値が上書きされます。</span><span class="sxs-lookup"><span data-stu-id="40d5e-114">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="40d5e-115">この動作は、[範囲形式](/javascript/api/excel/excel.range#format)などの[ナビゲーションプロパティ](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties)とは異なります。</span><span class="sxs-lookup"><span data-stu-id="40d5e-115">This behavior differs from [navigational properties](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="40d5e-116">の`format`プロパティは、次に示すように、object ナビゲーションを使用して設定できます。</span><span class="sxs-lookup"><span data-stu-id="40d5e-116">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="40d5e-117">読み取り専用修飾子をチェックすることで、そのサブプロパティを JSON 構造体で設定する必要があるプロパティを識別できます。</span><span class="sxs-lookup"><span data-stu-id="40d5e-117">You can identify a property that must have its subproperties set with a JSON struct by checking its read-only modifier.</span></span> <span data-ttu-id="40d5e-118">読み取り専用のプロパティは、読み取り専用でないサブプロパティを直接設定することができます。</span><span class="sxs-lookup"><span data-stu-id="40d5e-118">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="40d5e-119">書き込み可能な`PageLayout.zoom`プロパティは、JSON 構造体で設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="40d5e-119">Writeable properties like `PageLayout.zoom` must be set with a JSON struct.</span></span> <span data-ttu-id="40d5e-120">概要:</span><span class="sxs-lookup"><span data-stu-id="40d5e-120">In summary:</span></span>

- <span data-ttu-id="40d5e-121">読み取り専用プロパティ: サブプロパティは、ナビゲーションを使用して設定できます。</span><span class="sxs-lookup"><span data-stu-id="40d5e-121">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="40d5e-122">書き込み可能なプロパティ: サブプロパティは JSON 構造体で設定する必要があります (ナビゲーションで設定することはできません)。</span><span class="sxs-lookup"><span data-stu-id="40d5e-122">Writable property: Subproperties must be set with a JSON struct (and cannot be set through navigation).</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="40d5e-123">読み取り専用プロパティの設定</span><span class="sxs-lookup"><span data-stu-id="40d5e-123">Setting read-only properties</span></span>

<span data-ttu-id="40d5e-124">Office JS の[TypeScript 定義](/referencing-the-javascript-api-for-office-library-from-its-cdn.md)は、読み取り専用のオブジェクトプロパティを指定します。</span><span class="sxs-lookup"><span data-stu-id="40d5e-124">The [TypeScript definitions](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="40d5e-125">読み取り専用プロパティを設定しようとすると、エラーがスローされずに書き込み操作が失敗します。</span><span class="sxs-lookup"><span data-stu-id="40d5e-125">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="40d5e-126">次の例では、誤って読み取り専用プロパティ[Chart.id](/javascript/api/excel/excel.chart#id)を設定しようとしています。</span><span class="sxs-lookup"><span data-stu-id="40d5e-126">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="see-also"></a><span data-ttu-id="40d5e-127">関連項目</span><span class="sxs-lookup"><span data-stu-id="40d5e-127">See also</span></span>

- <span data-ttu-id="40d5e-128">[Officedev/office-js](https://github.com/OfficeDev/office-js/issues): office アドインプラットフォームおよび JavaScript api の問題を報告および表示する場所です。</span><span class="sxs-lookup"><span data-stu-id="40d5e-128">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="40d5e-129">[スタックオーバーフロー](https://stackoverflow.com/questions/tagged/office-js): Office JavaScript api に関するプログラミング上の問題を確認および表示する場所です。</span><span class="sxs-lookup"><span data-stu-id="40d5e-129">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="40d5e-130">スタックオーバーフローに投稿するときには、必ず "office-js" タグを質問に適用してください。</span><span class="sxs-lookup"><span data-stu-id="40d5e-130">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="40d5e-131">[UserVoice](https://officespdev.uservoice.com/): office アドインプラットフォームおよび Office JavaScript api の新機能を提案する場所です。</span><span class="sxs-lookup"><span data-stu-id="40d5e-131">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
