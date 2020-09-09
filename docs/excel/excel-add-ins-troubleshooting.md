---
title: Excel アドインのトラブルシューティング
description: Excel アドインの開発エラーをトラブルシューティングする方法について説明します。
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 1bdd96772d3a221ca3a02e3d5dfcfa16561dd5f1
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409404"
---
# <a name="troubleshooting-excel-add-ins"></a><span data-ttu-id="e228b-103">Excel アドインのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="e228b-103">Troubleshooting Excel Add-ins</span></span>

<span data-ttu-id="e228b-104">この記事では、Excel に固有の問題のトラブルシューティングについて説明します。</span><span class="sxs-lookup"><span data-stu-id="e228b-104">This article discusses troubleshooting issues that are unique to Excel.</span></span> <span data-ttu-id="e228b-105">ページの下部にあるフィードバックツールを使用して、記事に追加できるその他の問題を提案してください。</span><span class="sxs-lookup"><span data-stu-id="e228b-105">Please use the feedback tool at the bottom of the page to suggest other issues that can be added to the article.</span></span>

## <a name="api-limitations-when-the-active-workbook-switches"></a><span data-ttu-id="e228b-106">アクティブなブックの切り替え時の API の制限</span><span class="sxs-lookup"><span data-stu-id="e228b-106">API limitations when the active workbook switches</span></span>

<span data-ttu-id="e228b-107">Excel 用のアドインは、一度に1つのブックを操作することを目的としています。</span><span class="sxs-lookup"><span data-stu-id="e228b-107">Add-ins for Excel are intended to operate on a single workbook at a time.</span></span> <span data-ttu-id="e228b-108">アドインを実行しているブックとは別のブックがフォーカスを取得すると、エラーが発生することがあります。</span><span class="sxs-lookup"><span data-stu-id="e228b-108">Errors can arise when a workbook that is separate from the one running the add-in gains focus.</span></span> <span data-ttu-id="e228b-109">これは、フォーカスが変更されたときに、特定のメソッドが呼び出されたときにのみ発生します。</span><span class="sxs-lookup"><span data-stu-id="e228b-109">This only happens when particular methods are in the process of being called when the focus changes.</span></span>

<span data-ttu-id="e228b-110">このブックスイッチの影響を受ける Api は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="e228b-110">The following APIs are affected by this workbook switch:</span></span>

|<span data-ttu-id="e228b-111">Excel JavaScript API</span><span class="sxs-lookup"><span data-stu-id="e228b-111">Excel JavaScript API</span></span> | <span data-ttu-id="e228b-112">スローされたエラー</span><span class="sxs-lookup"><span data-stu-id="e228b-112">Error thrown</span></span> |
|--|--|
| `Chart.activate` | <span data-ttu-id="e228b-113">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e228b-113">GeneralException</span></span> |
| `Range.select` | <span data-ttu-id="e228b-114">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e228b-114">GeneralException</span></span> |
| `Table.clearFilters` | <span data-ttu-id="e228b-115">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e228b-115">GeneralException</span></span> |
| `Workbook.getActiveCell`  | <span data-ttu-id="e228b-116">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="e228b-116">InvalidSelection</span></span>|
| `Workbook.getSelectedRange` | <span data-ttu-id="e228b-117">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="e228b-117">InvalidSelection</span></span>|
| `Workbook.getSelectedRanges`  | <span data-ttu-id="e228b-118">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="e228b-118">InvalidSelection</span></span>|
| `Worksheet.activate` | <span data-ttu-id="e228b-119">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e228b-119">GeneralException</span></span> |
| `Worksheet.delete`  | <span data-ttu-id="e228b-120">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="e228b-120">InvalidSelection</span></span>|
| `Worksheet.gridlines` | <span data-ttu-id="e228b-121">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e228b-121">GeneralException</span></span> |
| `Worksheet.showHeadings` | <span data-ttu-id="e228b-122">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e228b-122">GeneralException</span></span> |
| `WorksheetCollection.add` | <span data-ttu-id="e228b-123">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e228b-123">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeAt` | <span data-ttu-id="e228b-124">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e228b-124">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeColumns` | <span data-ttu-id="e228b-125">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e228b-125">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeRows` | <span data-ttu-id="e228b-126">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e228b-126">GeneralException</span></span> |
| `WorksheetFreezePanes.getLocationOrNullObject`| <span data-ttu-id="e228b-127">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e228b-127">GeneralException</span></span> |
| `WorksheetFreezePanes.unfreeze` | <span data-ttu-id="e228b-128">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e228b-128">GeneralException</span></span> |

> [!NOTE]
> <span data-ttu-id="e228b-129">これは、Windows または Mac で開いている複数の Excel ブックにのみ適用されます。</span><span class="sxs-lookup"><span data-stu-id="e228b-129">This only applies to multiple Excel workbooks open on Windows or Mac.</span></span>

## <a name="coauthoring"></a><span data-ttu-id="e228b-130">共同編集</span><span class="sxs-lookup"><span data-stu-id="e228b-130">Coauthoring</span></span>

<span data-ttu-id="e228b-131">共同編集環境でイベントと共に使用するパターンについては、「 [Excel アドインの共同編集](co-authoring-in-excel-add-ins.md) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e228b-131">See [Coauthoring in Excel add-ins](co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment.</span></span> <span data-ttu-id="e228b-132">この記事では、など、特定の Api を使用する場合のマージの競合の可能性についても説明し [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) ます。</span><span class="sxs-lookup"><span data-stu-id="e228b-132">The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).</span></span>

## <a name="see-also"></a><span data-ttu-id="e228b-133">こちらもご覧ください</span><span class="sxs-lookup"><span data-stu-id="e228b-133">See also</span></span>

- [<span data-ttu-id="e228b-134">Office アドインでの開発エラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="e228b-134">Troubleshoot development errors with Office Add-ins</span></span>](../testing/troubleshoot-development-errors.md)
- [<span data-ttu-id="e228b-135">Office アドインでのユーザー エラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="e228b-135">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
