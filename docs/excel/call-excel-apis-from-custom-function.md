---
title: カスタム関数から Microsoft Excel Api を呼び出す
description: カスタム関数から呼び出すことができる Microsoft Excel Api について説明します。
ms.date: 02/06/2020
localization_priority: Normal
ms.openlocfilehash: e22ed897e95a74707bd0d8bded3f8dca724731d1
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719344"
---
# <a name="call-microsoft-excel-apis-from-a-custom-function"></a><span data-ttu-id="1c745-103">カスタム関数から Microsoft Excel Api を呼び出す</span><span class="sxs-lookup"><span data-stu-id="1c745-103">Call Microsoft Excel APIs from a custom function</span></span>

[!include[Running custom functions in a shared runtime note](../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="1c745-104">カスタム関数から Office .js Excel Api を呼び出して、範囲データを取得し、計算のためにより多くのコンテキストを取得します。</span><span class="sxs-lookup"><span data-stu-id="1c745-104">Call Office.js Excel APIs from your custom functions to get range data and obtain more context for your calculations.</span></span>

<span data-ttu-id="1c745-105">カスタム関数を使用した Office .js Api の呼び出しは、次のような場合に役立ちます。</span><span class="sxs-lookup"><span data-stu-id="1c745-105">Calling Office.js APIs through a custom function can be helpful when:</span></span>

- <span data-ttu-id="1c745-106">カスタム関数は、計算の前に Excel から情報を取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1c745-106">A custom function needs to get information from Excel before calculation.</span></span> <span data-ttu-id="1c745-107">この情報には、ドキュメントのプロパティ、範囲の書式、カスタム XML パーツ、ブック名、その他の Excel 固有の情報が含まれることがあります。</span><span class="sxs-lookup"><span data-stu-id="1c745-107">This information might include document properties, range formats, custom XML parts, a workbook name, or other Excel-specific information.</span></span>
- <span data-ttu-id="1c745-108">ユーザー設定関数は、計算後の戻り値のセルの番号書式を設定します。</span><span class="sxs-lookup"><span data-stu-id="1c745-108">A custom function will set the cell's number format for the return values after calculation.</span></span>

[!include[Excel shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="code-sample"></a><span data-ttu-id="1c745-109">コード サンプル</span><span class="sxs-lookup"><span data-stu-id="1c745-109">Code sample</span></span>

<span data-ttu-id="1c745-110">Office .js Api を呼び出すには、まずコンテキストが必要です。</span><span class="sxs-lookup"><span data-stu-id="1c745-110">To call into the Office.js APIs you first need a context.</span></span> <span data-ttu-id="1c745-111">オブジェクトを`Excel.RequestContext`使用してコンテキストを取得します。</span><span class="sxs-lookup"><span data-stu-id="1c745-111">Use the `Excel.RequestContext` object to get a context.</span></span> <span data-ttu-id="1c745-112">その後、コンテキストを使用して、ブックで必要な Api を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="1c745-112">Then use the context to call the APIs you need in the workbook.</span></span>

<span data-ttu-id="1c745-113">次のコードサンプルは、ブックから値の範囲を取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1c745-113">The following code sample shows how to get a range of values from the workbook.</span></span>

```JavaScript
/**
 * @customfunction
 * @param address range's address
 **/
async function getRangeValue (address) {
 var context = new Excel.RequestContext();
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-officejs-through-a-custom-function"></a><span data-ttu-id="1c745-114">カスタム関数を使用して Office .js を呼び出す際の制限</span><span class="sxs-lookup"><span data-stu-id="1c745-114">Limitations of calling Office.js through a custom function</span></span>

<span data-ttu-id="1c745-115">Excel の環境を変更するカスタム関数から Office .js Api を呼び出さないでください。</span><span class="sxs-lookup"><span data-stu-id="1c745-115">Don't call Office.js APIs from a custom function that change the environment of Excel.</span></span> <span data-ttu-id="1c745-116">これは、カスタム関数が以下の操作を行わないことを意味します。</span><span class="sxs-lookup"><span data-stu-id="1c745-116">This means your custom functions should not do any of the following:</span></span>

- <span data-ttu-id="1c745-117">スプレッドシートのセルを挿入、削除、または書式設定します。</span><span class="sxs-lookup"><span data-stu-id="1c745-117">Insert, delete, or format cells on the spreadsheet.</span></span>
- <span data-ttu-id="1c745-118">別のセルの値を変更します。</span><span class="sxs-lookup"><span data-stu-id="1c745-118">Change another cell's value.</span></span>
- <span data-ttu-id="1c745-119">ブックにシートを移動、名前変更、削除、または追加します。</span><span class="sxs-lookup"><span data-stu-id="1c745-119">Move, rename, delete, or add sheets to a workbook.</span></span>
- <span data-ttu-id="1c745-120">計算モードや画面表示などの環境オプションを変更します。</span><span class="sxs-lookup"><span data-stu-id="1c745-120">Change any of the environment options, such as calculation mode or screen views.</span></span>
- <span data-ttu-id="1c745-121">ブックに名前を追加します。</span><span class="sxs-lookup"><span data-stu-id="1c745-121">Add names to a workbook.</span></span>
- <span data-ttu-id="1c745-122">プロパティを設定するか、ほとんどのメソッドを実行します。</span><span class="sxs-lookup"><span data-stu-id="1c745-122">Set properties or execute most methods.</span></span>

<span data-ttu-id="1c745-123">Excel を変更すると、パフォーマンスが低下し、タイムアウトになり、無限ループが発生します。</span><span class="sxs-lookup"><span data-stu-id="1c745-123">Changing Excel can result in poor performance, time outs, and infinite loops.</span></span> <span data-ttu-id="1c745-124">Excel の再計算を実行しているときに、予期しない結果になる可能性があるため、カスタム関数の計算を実行することはできません。</span><span class="sxs-lookup"><span data-stu-id="1c745-124">Custom function calculations shouldn't run while an Excel recalculation is taking place as it will result in unpredictable results.</span></span>

<span data-ttu-id="1c745-125">代わりに、リボンボタンまたは作業ウィンドウのコンテキストから Excel に変更を加えます。</span><span class="sxs-lookup"><span data-stu-id="1c745-125">Instead, make changes to Excel from the context of a ribbon button, or task pane.</span></span>

## <a name="next-steps"></a><span data-ttu-id="1c745-126">次の手順</span><span class="sxs-lookup"><span data-stu-id="1c745-126">Next steps</span></span>

- [<span data-ttu-id="1c745-127">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="1c745-127">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a><span data-ttu-id="1c745-128">関連項目</span><span class="sxs-lookup"><span data-stu-id="1c745-128">See also</span></span>

- [<span data-ttu-id="1c745-129">Excel カスタム関数と作業ウィンドウチュートリアルの間でデータとイベントを共有する</span><span class="sxs-lookup"><span data-stu-id="1c745-129">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)