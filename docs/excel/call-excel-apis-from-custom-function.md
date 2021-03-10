---
title: カスタム関数から Excel JavaScript API を呼び出す
description: カスタム関数から呼び出す Excel JavaScript API について説明します。
ms.date: 03/05/2021
localization_priority: Normal
ms.openlocfilehash: 4be1b1ee8ea4ae8b2f5d1d27195be18f7aa841da
ms.sourcegitcommit: d153f6d4c3e01d63ed24aa1349be16fa8ad51218
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/10/2021
ms.locfileid: "50613907"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a><span data-ttu-id="dac3f-103">カスタム関数から Excel JavaScript API を呼び出す</span><span class="sxs-lookup"><span data-stu-id="dac3f-103">Call Excel JavaScript APIs from a custom function</span></span>

<span data-ttu-id="dac3f-104">カスタム関数から Excel JavaScript API を呼び出して、範囲データを取得し、計算のコンテキストを追加します。</span><span class="sxs-lookup"><span data-stu-id="dac3f-104">Call Excel JavaScript APIs from your custom functions to get range data and obtain more context for your calculations.</span></span> <span data-ttu-id="dac3f-105">カスタム関数を使用して Excel JavaScript API を呼び出すのは、次の場合に役立ちます。</span><span class="sxs-lookup"><span data-stu-id="dac3f-105">Calling Excel JavaScript APIs through a custom function can be helpful when:</span></span>

- <span data-ttu-id="dac3f-106">カスタム関数は、計算前に Excel から情報を取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="dac3f-106">A custom function needs to get information from Excel before calculation.</span></span> <span data-ttu-id="dac3f-107">この情報には、ドキュメントのプロパティ、範囲の形式、カスタム XML パーツ、ブック名、その他の Excel 固有の情報が含まれる場合があります。</span><span class="sxs-lookup"><span data-stu-id="dac3f-107">This information might include document properties, range formats, custom XML parts, a workbook name, or other Excel-specific information.</span></span>
- <span data-ttu-id="dac3f-108">カスタム関数は、計算後の戻り値のセルの数値形式を設定します。</span><span class="sxs-lookup"><span data-stu-id="dac3f-108">A custom function will set the cell's number format for the return values after calculation.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="dac3f-109">カスタム関数から Excel JavaScript API を呼び出す場合は、共有 JavaScript ランタイムを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="dac3f-109">To call Excel JavaScript APIs from your custom function, you'll need to use a shared JavaScript runtime.</span></span> <span data-ttu-id="dac3f-110">詳細については、「[Office アドインを構成して共有 JavaScript ランタイムを使用する ](../develop/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="dac3f-110">See [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md) to learn more.</span></span>

## <a name="code-sample"></a><span data-ttu-id="dac3f-111">コード サンプル</span><span class="sxs-lookup"><span data-stu-id="dac3f-111">Code sample</span></span>

<span data-ttu-id="dac3f-112">カスタム関数から Excel JavaScript API を呼び出す場合は、まずコンテキストが必要です。</span><span class="sxs-lookup"><span data-stu-id="dac3f-112">To call Excel JavaScript APIs from a custom function, you first need a context.</span></span> <span data-ttu-id="dac3f-113">コンテキストを [取得するには、Excel.RequestContext](/javascript/api/excel/excel.requestcontext) オブジェクトを使用します。</span><span class="sxs-lookup"><span data-stu-id="dac3f-113">Use the [Excel.RequestContext](/javascript/api/excel/excel.requestcontext) object to get a context.</span></span> <span data-ttu-id="dac3f-114">次に、ブックで必要な API を呼び出すコンテキストを使用します。</span><span class="sxs-lookup"><span data-stu-id="dac3f-114">Then use the context to call the APIs you need in the workbook.</span></span>

<span data-ttu-id="dac3f-115">次のコード サンプルは、ブック内のセルから値を取得する `Excel.RequestContext` 方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="dac3f-115">The following code sample shows how to use `Excel.RequestContext` to get a value from a cell in the workbook.</span></span> <span data-ttu-id="dac3f-116">このサンプルでは、パラメーター `address` は Excel JavaScript API [Worksheet.getRange](/javascript/api/excel/excel.worksheet#getRange_address_) メソッドに渡され、文字列として入力する必要があります。</span><span class="sxs-lookup"><span data-stu-id="dac3f-116">In this sample, the `address` parameter is passed into the Excel JavaScript API [Worksheet.getRange](/javascript/api/excel/excel.worksheet#getRange_address_) method and must be entered as a string.</span></span> <span data-ttu-id="dac3f-117">たとえば、Excel UI に入力されたカスタム関数は、値を取得するセルのアドレスであるパターンに従 `=CONTOSO.GETRANGEVALUE("A1")` `"A1"` う必要があります。</span><span class="sxs-lookup"><span data-stu-id="dac3f-117">For example, the custom function entered into the Excel UI must follow the pattern `=CONTOSO.GETRANGEVALUE("A1")`, where `"A1"` is the address of the cell from which to retrieve the value.</span></span>

```JavaScript
/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
async function getRangeValue(address) {
 // Retrieve the context object. 
 var context = new Excel.RequestContext();
 
 // Use the context object to access the cell at the input address. 
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a><span data-ttu-id="dac3f-118">カスタム関数を使用して Excel JavaScript API を呼び出す場合の制限事項</span><span class="sxs-lookup"><span data-stu-id="dac3f-118">Limitations of calling Excel JavaScript APIs through a custom function</span></span>

<span data-ttu-id="dac3f-119">Excel の環境を変更するカスタム関数から Excel JavaScript API を呼び出さない。</span><span class="sxs-lookup"><span data-stu-id="dac3f-119">Don't call Excel JavaScript APIs from a custom function that change the environment of Excel.</span></span> <span data-ttu-id="dac3f-120">つまり、カスタム関数は次の操作を行う必要があります。</span><span class="sxs-lookup"><span data-stu-id="dac3f-120">This means your custom functions should not do any of the following:</span></span>

- <span data-ttu-id="dac3f-121">スプレッドシートにセルを挿入、削除、または書式設定します。</span><span class="sxs-lookup"><span data-stu-id="dac3f-121">Insert, delete, or format cells on the spreadsheet.</span></span>
- <span data-ttu-id="dac3f-122">別のセルの値を変更します。</span><span class="sxs-lookup"><span data-stu-id="dac3f-122">Change another cell's value.</span></span>
- <span data-ttu-id="dac3f-123">ブックにシートを移動、名前変更、削除、または追加します。</span><span class="sxs-lookup"><span data-stu-id="dac3f-123">Move, rename, delete, or add sheets to a workbook.</span></span>
- <span data-ttu-id="dac3f-124">計算モードや画面ビューなど、環境オプションを変更します。</span><span class="sxs-lookup"><span data-stu-id="dac3f-124">Change any of the environment options, such as calculation mode or screen views.</span></span>
- <span data-ttu-id="dac3f-125">ブックに名前を追加します。</span><span class="sxs-lookup"><span data-stu-id="dac3f-125">Add names to a workbook.</span></span>
- <span data-ttu-id="dac3f-126">プロパティを設定するか、ほとんどのメソッドを実行します。</span><span class="sxs-lookup"><span data-stu-id="dac3f-126">Set properties or execute most methods.</span></span>

<span data-ttu-id="dac3f-127">Excel を変更すると、パフォーマンスが低下し、タイム アウトが発生し、無限ループが発生する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="dac3f-127">Changing Excel can result in poor performance, time outs, and infinite loops.</span></span> <span data-ttu-id="dac3f-128">Excel の再計算が行なっている間は、予期しない結果になるので、カスタム関数の計算は実行できません。</span><span class="sxs-lookup"><span data-stu-id="dac3f-128">Custom function calculations shouldn't run while an Excel recalculation is taking place as it will result in unpredictable results.</span></span>

<span data-ttu-id="dac3f-129">代わりに、リボン ボタンまたは作業ウィンドウのコンテキストから Excel に変更を加えます。</span><span class="sxs-lookup"><span data-stu-id="dac3f-129">Instead, make changes to Excel from the context of a ribbon button, or task pane.</span></span>

## <a name="next-steps"></a><span data-ttu-id="dac3f-130">次の手順</span><span class="sxs-lookup"><span data-stu-id="dac3f-130">Next steps</span></span>

- [<span data-ttu-id="dac3f-131">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="dac3f-131">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a><span data-ttu-id="dac3f-132">関連項目</span><span class="sxs-lookup"><span data-stu-id="dac3f-132">See also</span></span>

- [<span data-ttu-id="dac3f-133">Excel カスタム関数と作業ウィンドウチュートリアルの間でデータとイベントを共有する</span><span class="sxs-lookup"><span data-stu-id="dac3f-133">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [<span data-ttu-id="dac3f-134">Office アドインを構成して共有 JavaScript ランタイムを使用する</span><span class="sxs-lookup"><span data-stu-id="dac3f-134">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
