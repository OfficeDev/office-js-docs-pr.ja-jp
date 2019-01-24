---
title: Excel JavaScript API を使用して範囲を操作する (基本)
description: ''
ms.date: 12/28/2018
localization_priority: Priority
ms.openlocfilehash: 505c22d2a3230aeafaf4d0c62a371a2ab93b3a9a
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386786"
---
# <a name="work-with-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="5d6f8-102">Excel JavaScript API を使用して範囲を操作する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-102">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="5d6f8-103">この記事では、Excel JavaScript API を使用して、範囲に関する一般的なタスクを実行する方法を示すサンプル コードを提供します。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-103">This article provides code samples that show how to perform common tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="5d6f8-104">**Range** オブジェクトがサポートするプロパティとメソッドの完全な一覧については、「[Range オブジェクト (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span></span>

> [!NOTE]
> <span data-ttu-id="5d6f8-105">範囲を指定してより詳細なタスクを実行する方法のサンプル コードについては、「[Excel JavaScript API を使用して範囲を操作する (詳細)](excel-add-ins-ranges-advanced.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-105">For code samples that show how to perform more advanced tasks with ranges, see [Work with ranges using the Excel JavaScript API (advanced)](excel-add-ins-ranges-advanced.md).</span></span>

## <a name="get-a-range"></a><span data-ttu-id="5d6f8-106">範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-106">Get a range</span></span>

<span data-ttu-id="5d6f8-107">次の例では、ワークシート内の範囲への参照を取得する、さまざまな方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-107">The following examples show different ways to get a reference to a range within a worksheet.</span></span>

### <a name="get-range-by-address"></a><span data-ttu-id="5d6f8-108">アドレスによって範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-108">Get range by address</span></span>

<span data-ttu-id="5d6f8-109">次のコード サンプルでは、**Sample** という名前のワークシートからアドレス **B2:B5** の範囲を取得し、**address** プロパティを読み込んで、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-109">The following code sample gets the range with address **B2:B5** from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:C5");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range B2:C5 is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-range-by-name"></a><span data-ttu-id="5d6f8-110">名前によって範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-110">Get range by name</span></span>

<span data-ttu-id="5d6f8-111">次のコード サンプルでは、**Sample** という名前のワークシートから **MyRange** という名前の範囲を取得し、**address** プロパティを読み込んで、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-111">The following code sample gets the range named **MyRange** from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("MyRange");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range "MyRange" is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-used-range"></a><span data-ttu-id="5d6f8-112">使用範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-112">Get used range</span></span>

<span data-ttu-id="5d6f8-113">次のコード サンプルでは、**Sample** という名前のワークシートから使用範囲を取得し、**address** プロパティを読み込んで、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-113">The following code sample gets the used range from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span> <span data-ttu-id="5d6f8-114">使用範囲とは、値または書式設定が割り当てられているワークシート内のセルを含む、最小の範囲です。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-114">The used range is the smallest range that encompasses any cells in the worksheet that have a value or formatting assigned to them.</span></span> <span data-ttu-id="5d6f8-115">ワークシート全体が空白の場合、**getUsedRange()** メソッドは、ワークシートの左上のセルのみで構成される範囲を返します。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-115">If the entire worksheet is blank, the **getUsedRange()** method returns a range that consists of only the top-left cell in the worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getUsedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the used range in the worksheet is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-entire-range"></a><span data-ttu-id="5d6f8-116">範囲全体を取得する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-116">Get entire range</span></span>

<span data-ttu-id="5d6f8-117">次のコード サンプルでは、**Sample** という名前のワークシートからワークシートの範囲全体を取得し、**address** プロパティを読み込んで、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-117">The following code sample gets the entire worksheet range from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the entire worksheet range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="insert-a-range-of-cells"></a><span data-ttu-id="5d6f8-118">セルの範囲を挿入する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-118">Insert a range of cells</span></span>

<span data-ttu-id="5d6f8-119">次のコードサンプルは、場所 **B4:E4** にセルの範囲を挿入し、他のセルを下にシフトして、新しいセルのためのスペースを提供します。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-119">The following code sample inserts a range of cells in location **B4:E4** and shifts other cells down to provide space for the new cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);
    
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="5d6f8-120">**範囲を挿入する前のデータ**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-120">**Data before range is inserted**</span></span>

![範囲を挿入する前の Excel のデータ](../images/excel-ranges-start.png)

<span data-ttu-id="5d6f8-122">**範囲を挿入した後のデータ**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-122">**Data after range is inserted**</span></span>

![範囲を挿入した後の Excel のデータ](../images/excel-ranges-after-insert.png)

## <a name="clear-a-range-of-cells"></a><span data-ttu-id="5d6f8-124">セルの範囲をクリアする</span><span class="sxs-lookup"><span data-stu-id="5d6f8-124">Clear a range of cells</span></span>

<span data-ttu-id="5d6f8-125">次のコード サンプルは、範囲 **E2：E5** のセルの内容と書式をすべてクリアします。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-125">The following code sample clears all contents and formatting of cells in the range **E2:E5**.</span></span>  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="5d6f8-126">**範囲をクリアする前のデータ**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-126">**Data before range is cleared**</span></span>

![範囲をクリアする前の Excel のデータ](../images/excel-ranges-start.png)

<span data-ttu-id="5d6f8-128">**範囲をクリアした後のデータ**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-128">**Data after range is cleared**</span></span>

![範囲をクリアした後の Excel のデータ](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a><span data-ttu-id="5d6f8-130">セルの範囲を削除する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-130">Delete a range of cells</span></span>

<span data-ttu-id="5d6f8-131">次のコード サンプルは、範囲 **B4:E4** のセルを削除し、他のセルを上にシフトして、削除されたセルのために空いたスペースに入力します。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-131">The following code sample deletes the cells in the range **B4:E4** and shift other cells up to fill the space that was vacated by the deleted cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="5d6f8-132">**範囲を削除する前のデータ**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-132">**Data before range is deleted**</span></span>

![範囲を削除する前の Excel のデータ](../images/excel-ranges-start.png)

<span data-ttu-id="5d6f8-134">**範囲を削除した後のデータ**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-134">**Data after range is deleted**</span></span>

![範囲を削除した後の Excel のデータ](../images/excel-ranges-after-delete.png)

## <a name="set-the-selected-range"></a><span data-ttu-id="5d6f8-136">選択範囲を設定する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-136">Set the selected range</span></span>

<span data-ttu-id="5d6f8-137">次のコード サンプルは、作業中のワークシートの範囲 **B2:E6** を選択します。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-137">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="5d6f8-138">**選択範囲 B2:E6**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-138">**Selected range B2:E6**</span></span>

![Excel の選択範囲](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="5d6f8-140">選択範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-140">Get the selected range</span></span>

<span data-ttu-id="5d6f8-141">次のコード サンプルでは、選択範囲を取得し、**address** プロパティを読み込んで、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-141">The following code sample gets the selected range, loads its **address** property, and writes a message to the console.</span></span> 

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the selected range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="set-values-or-formulas"></a><span data-ttu-id="5d6f8-142">値または数式を設定する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-142">Set values or formulas</span></span>

<span data-ttu-id="5d6f8-143">次の例は、1 つのセルまたはセルの範囲の値と数式を設定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-143">The following examples show how to set values and formulas for a single cell or a range of cells.</span></span>

### <a name="set-value-for-a-single-cell"></a><span data-ttu-id="5d6f8-144">1 つのセルの値を設定する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-144">Set value for a single cell</span></span>

<span data-ttu-id="5d6f8-145">次のコード サンプルでは、セル **C3** の値を "5" に設定し、データに最も適した列の幅を設定します。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-145">The following code sample sets the value of cell **C3** to "5" and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="5d6f8-146">**セルの値が更新される前のデータ**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-146">**Data before cell value is updated**</span></span>

![セルの値が更新される前の Excel のデータ](../images/excel-ranges-set-start.png)

<span data-ttu-id="5d6f8-148">**セルの値が更新された後のデータ**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-148">**Data after cell value is updated**</span></span>

![セルの値が更新された後の Excel のデータ](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a><span data-ttu-id="5d6f8-150">複数のセルの範囲の値を設定する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-150">Set values for a range of cells</span></span>

<span data-ttu-id="5d6f8-151">次のコード サンプルでは、範囲 **B5：D5** のセルの値を設定し、データに最も適した列の幅を設定します。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-151">The following code sample sets values for the cells in the range **B5:D5** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["Potato Chips", 10, 1.80],
    ];
    
    var range = sheet.getRange("B5:D5");
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="5d6f8-152">**複数のセルの値が更新される前のデータ**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-152">**Data before cell values are updated**</span></span>

![複数のセルの値が更新される前の Excel のデータ](../images/excel-ranges-set-start.png)

<span data-ttu-id="5d6f8-154">**複数のセルの値が更新された後のデータ**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-154">**Data after cell values are updated**</span></span>

![複数のセルの値が更新された後の Excel のデータ](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a><span data-ttu-id="5d6f8-156">1 つのセルの数式を設定する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-156">Set formula for a single cell</span></span>

<span data-ttu-id="5d6f8-157">次のコード サンプルでは、セル **E3** の数式を設定し、データに最も適した列の幅を設定します。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-157">The following code sample sets a formula for cell **E3** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="5d6f8-158">**セルの数式が設定される前のデータ**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-158">**Data before cell formula is set**</span></span>

![セルの数式が設定される前の Excel のデータ](../images/excel-ranges-start-set-formula.png)

<span data-ttu-id="5d6f8-160">**セルの数式が設定された後のデータ**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-160">**Data after cell formula is set**</span></span>

![セルの数式が設定された後の Excel のデータ](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a><span data-ttu-id="5d6f8-162">セルの範囲の数式を設定する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-162">Set formulas for a range of cells</span></span>

<span data-ttu-id="5d6f8-163">次のコード サンプルでは、範囲 **E2:E6** のセルの数式を設定し、データに最も適した列の幅を設定します。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-163">The following code sample sets formulas for cells in the range **E2:E6** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"]
    ];
    
    var range = sheet.getRange("E3:E6");
    range.formulas = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="5d6f8-164">**複数のセルの数式が設定される前のデータ**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-164">**Data before cell formulas are set**</span></span>

![複数のセルの数式が設定される前の Excel のデータ](../images/excel-ranges-start-set-formula.png)

<span data-ttu-id="5d6f8-166">**複数のセルの数式が設定された後のデータ**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-166">**Data after cell formulas are set**</span></span>

![複数のセルの数式が設定された後の Excel のデータ](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a><span data-ttu-id="5d6f8-168">値、テキスト、または数式を取得する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-168">Get values, text, or formulas</span></span>

<span data-ttu-id="5d6f8-169">次の例は、セルの範囲から値、テキスト、および数式を取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-169">These examples show how to get values, text, and formulas from a range of cells.</span></span>

### <a name="get-values-from-a-range-of-cells"></a><span data-ttu-id="5d6f8-170">セルの範囲から値を取得する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-170">Get values from a range of cells</span></span>

<span data-ttu-id="5d6f8-171">次のコード サンプルでは、範囲 **B2:E6** を取得し、**values** プロパティを読み込んで、コンソールに値を書き込みます。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-171">The following code sample gets the range **B2:E6**, loads its **values** property, and writes the values to the console.</span></span> <span data-ttu-id="5d6f8-172">範囲の **values** プロパティは、セルに含まれる未処理の値を指定します。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-172">The **values** property of a range specifies the raw values that the cells contain.</span></span> <span data-ttu-id="5d6f8-173">範囲内の一部のセルに数式が含まれている場合でも、範囲の **values** プロパティは、それらのセルの未処理の値 (数式ではなく) を指定します。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-173">Even if some cells in a range contain formulas, the **values** property of the range specifies the raw values for those cells, not any of the formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("values");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.values, null, 4));
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="5d6f8-174">**範囲内のデータ (列 E の値は数式の結果)**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-174">**Data in range (values in column E are a result of formulas)**</span></span>

![複数のセルの数式が設定された後の Excel のデータ](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="5d6f8-176">**range.values (上記のコード サンプルによりコンソールに記録される)**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-176">**range.values (as logged to the console by the code sample above)**</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        15
    ],
    [
        "Coffee",
        1,
        34.5,
        34.5
    ],
    [
        "Chocolate",
        5,
        9.56,
        47.8
    ],
    [
        "",
        "",
        "",
        97.3
    ]
]
```

### <a name="get-text-from-a-range-of-cells"></a><span data-ttu-id="5d6f8-177">セルの範囲からテキストを取得する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-177">Get text from a range of cells</span></span>

<span data-ttu-id="5d6f8-178">次のコード サンプルでは、範囲 **B2:E6** を取得し、**text** プロパティを読み込んでコンソールに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-178">The following code sample gets the range **B2:E6**, loads its **text** property, and writes it to the console.</span></span>  <span data-ttu-id="5d6f8-179">範囲の **text** プロパティは、範囲内のセルの表示値を指定します。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-179">The **text** property of a range specifies the display values for cells in the range.</span></span> <span data-ttu-id="5d6f8-180">範囲内の一部のセルに数式が含まれている場合でも、範囲の **text** プロパティは、それらのセルの表示値 (数式ではなく) を指定します。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-180">Even if some cells in a range contain formulas, the **text** property of the range specifies the display values for those cells, not any of the formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("text");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.text, null, 4));
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="5d6f8-181">**範囲内のデータ (列 E の値は数式の結果)**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-181">**Data in range (values in column E are a result of formulas)**</span></span>

![複数のセルの数式が設定された後の Excel のデータ](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="5d6f8-183">**range.text (上記のコード サンプルによりコンソールに記録される)**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-183">**range.text (as logged to the console by the code sample above)**</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        "2",
        "7.5",
        "15"
    ],
    [
        "Coffee",
        "1",
        "34.5",
        "34.5"
    ],
    [
        "Chocolate",
        "5",
        "9.56",
        "47.8"
    ],
    [
        "",
        "",
        "",
        "97.3"
    ]
]
```

### <a name="get-formulas-from-a-range-of-cells"></a><span data-ttu-id="5d6f8-184">セルの範囲から数式を取得する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-184">Get formulas from a range of cells</span></span>

<span data-ttu-id="5d6f8-185">次のコード サンプルでは、範囲 **B2:E6** を取得し、**formulas** プロパティを読み込んでコンソールに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-185">The following code sample gets the range **B2:E6**, loads its **formulas** property, and writes it to the console.</span></span>  <span data-ttu-id="5d6f8-186">範囲の **formulas** プロパティは、数式と数式を含まない範囲のセルの未処理の値が含まれる、範囲内のセルの数式を指定します。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-186">The **formulas** property of a range specifies the formulas for cells in the range that contain formulas and the raw values for cells in the range that do not contain formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("formulas");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.formulas, null, 4));
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="5d6f8-187">**範囲内のデータ (列 E の値は数式の結果)**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-187">**Data in range (values in column E are a result of formulas)**</span></span>

![複数のセルの数式が設定された後の Excel のデータ](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="5d6f8-189">**range.formulas (上記のコード サンプルによりコンソールに記録される)**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-189">**range.formulas (as logged to the console by the code sample above)**</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        "=C3 * D3"
    ],
    [
        "Coffee",
        1,
        34.5,
        "=C4 * D4"
    ],
    [
        "Chocolate",
        5,
        9.56,
        "=C5 * D5"
    ],
    [
        "",
        "",
        "",
        "=SUM(E3:E5)"
    ]
]
```

## <a name="set-range-format"></a><span data-ttu-id="5d6f8-190">範囲の書式を設定する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-190">Set range format</span></span>

<span data-ttu-id="5d6f8-191">次の例は、範囲内のセルのフォントの色、塗りつぶしの色、および数値の書式を設定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-191">The following examples show how to set font color, fill color, and number format for cells in a range.</span></span>

### <a name="set-font-color-and-fill-color"></a><span data-ttu-id="5d6f8-192">フォントの色と塗りつぶしの色を設定する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-192">Set font color and fill color</span></span>

<span data-ttu-id="5d6f8-193">次のコード サンプルは、範囲 **B2：E2** のセルのフォントの色と塗りつぶしの色を設定します。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-193">The following code sample sets the font color and fill color for cells in range **B2:E2**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";;
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="5d6f8-194">**フォントの色と塗りつぶしの色を設定する前の範囲内のデータ**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-194">**Data in range before font color and fill color are set**</span></span>

![書式設定する前の Excel のデータ](../images/excel-ranges-format-before.png)

<span data-ttu-id="5d6f8-196">**フォントの色と塗りつぶしの色を設定した後の範囲内のデータ**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-196">**Data in range after font color and fill color are set**</span></span>

![書式設定した後の Excel のデータ](../images/excel-ranges-format-font-and-fill.png)

### <a name="set-number-format"></a><span data-ttu-id="5d6f8-198">数値の書式を設定する</span><span class="sxs-lookup"><span data-stu-id="5d6f8-198">Set number format</span></span>

<span data-ttu-id="5d6f8-199">次のコード サンプルは、範囲 **D3：E5** のセルの数値を書式を設定します。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-199">The following code sample sets the number format for the cells in range **D3:E5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ];

    var range = sheet.getRange("D3:E5");
    range.numberFormat = formats;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="5d6f8-200">**数値の書式を設定する前の範囲内のデータ**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-200">**Data in range before number format is set**</span></span>

![書式設定する前の Excel のデータ](../images/excel-ranges-format-font-and-fill.png)

<span data-ttu-id="5d6f8-202">**数値の書式を設定した後の範囲内のデータ**</span><span class="sxs-lookup"><span data-stu-id="5d6f8-202">**Data in range after number format is set**</span></span>

![書式設定した後の Excel のデータ](../images/excel-ranges-format-numbers.png)

### <a name="conditional-formatting-of-ranges"></a><span data-ttu-id="5d6f8-204">範囲の条件付き書式</span><span class="sxs-lookup"><span data-stu-id="5d6f8-204">Conditional formatting of ranges</span></span>

<span data-ttu-id="5d6f8-205">範囲には、条件に基づいて個々のセルに適用する書式設定を含めることができます。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-205">Ranges can have formats applied to individual cells based on conditions.</span></span> <span data-ttu-id="5d6f8-206">この詳細については、「[Excel の範囲に条件付き書式を適用する](excel-add-ins-conditional-formatting.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-206">For more information about this, see [Apply conditional formatting to Excel ranges](excel-add-ins-conditional-formatting.md).</span></span>

## <a name="find-a-cell-using-string-matching-preview"></a><span data-ttu-id="5d6f8-207">文字列のマッチングを使用してセルを検索する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="5d6f8-207">Find a cell using string matching (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="5d6f8-208">現在、Range オブジェクトの `find` 関数は、パブリック プレビュー (ベータ版) でのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-208">The Range object's `find` function is currently available only in public preview (beta).</span></span> <span data-ttu-id="5d6f8-209">この機能を使用するには、Office.js CDN のベータ版のライブラリを使用する必要があります: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-209">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="5d6f8-210">TypeScript を使用している場合、または IntelliSense に TypeScript 型定義ファイルを使用するコード エディターを使用している場合は、https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts を使用してください。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-210">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="5d6f8-211">`Range` オブジェクトには、範囲内で指定された文字列を検索するための `find` メソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-211">The `Range` object has a `find` method to search for a specified string within the range.</span></span> <span data-ttu-id="5d6f8-212">このメソッドは、一致するテキストがある最初のセルの範囲を返します。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-212">It returns the range of the first cell with matching text.</span></span> <span data-ttu-id="5d6f8-213">次のコード サンプルは、文字列 **Food** と等しい値を持つ最初のセルを検索して、そのアドレスをコンソールに記録します。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-213">The following code sample finds the first cell with a value equal to the string **Food** and logs its address to the console.</span></span> <span data-ttu-id="5d6f8-214">指定した文字列が範囲に存在しない場合、`ItemNotFound` エラーが `find` によってスローされます。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-214">Note that `find` throws an `ItemNotFound` error if the specified string doesn't exist in the range.</span></span> <span data-ttu-id="5d6f8-215">指定した文字列が範囲に存在しない可能性がある場合は、自分のコードで適切にシナリオを処理できるように、[findOrNullObject](excel-add-ins-advanced-concepts.md#42ornullobject-methods) メソッドを使用するようにしてください。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-215">If you expect that the specified string may not exist in the range, use the [findOrNullObject](excel-add-ins-advanced-concepts.md#42ornullobject-methods) method instead, so your code gracefully handles that scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var table = sheet.tables.getItem("ExpensesTable");
    var searchRange = table.getRange();
    var foundRange = searchRange.find("Food", {
        completeMatch: true, // find will match the whole cell value
        matchCase: false, // find will not match case
        searchDirection: Excel.SearchDirection.forward // find will start searching at the beginning of the range
    });

    foundRange.load("address");
    return context.sync()
        .then(function() {
            console.log(foundRange.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="5d6f8-216">単一のセルを表す範囲に対して `find` メソッドが呼び出されると、ワークシート全体が検索されます。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-216">When the `find` method is called on a range representing a single cell, the entire worksheet is searched.</span></span> <span data-ttu-id="5d6f8-217">検索はその単一のセルから始まり、`SearchCriteria.searchDirection` によって指定された方向へ行われ、場合によってはワークシートの最終部分で折り返されます。</span><span class="sxs-lookup"><span data-stu-id="5d6f8-217">The search begins at that cell and goes in the direction specified by `SearchCriteria.searchDirection`, wrapping around the ends of the worksheet if needed.</span></span>

## <a name="see-also"></a><span data-ttu-id="5d6f8-218">関連項目</span><span class="sxs-lookup"><span data-stu-id="5d6f8-218">See also</span></span>

- [<span data-ttu-id="5d6f8-219">Excel JavaScript API を使用して範囲を操作する (高度)</span><span class="sxs-lookup"><span data-stu-id="5d6f8-219">Work with ranges using the Excel JavaScript API (advanced)</span></span>](excel-add-ins-ranges-advanced.md)
- [<span data-ttu-id="5d6f8-220">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="5d6f8-220">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
