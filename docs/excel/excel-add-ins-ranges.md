---
title: Excel JavaScript API を使用して範囲を操作する (基本)
description: Excel JavaScript API を使用して、範囲に関する一般的なタスクを実行する方法を示すコードサンプルです。
ms.date: 07/28/2020
localization_priority: Normal
ms.openlocfilehash: 2cb5640fa9c647c0e6d1f8083986455f2b5a59a3
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408482"
---
# <a name="work-with-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="1f5e8-103">Excel JavaScript API を使用して範囲を操作する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-103">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="1f5e8-104">この記事では、Excel JavaScript API を使用して、範囲に関する一般的なタスクを実行する方法を示すサンプル コードを提供します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-104">This article provides code samples that show how to perform common tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="1f5e8-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧につい `Range` ては、「 [Range オブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.range)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-105">For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

> [!NOTE]
> <span data-ttu-id="1f5e8-106">範囲を指定してより詳細なタスクを実行する方法のサンプル コードについては、「[Excel JavaScript API を使用して範囲を操作する (詳細)](excel-add-ins-ranges-advanced.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-106">For code samples that show how to perform more advanced tasks with ranges, see [Work with ranges using the Excel JavaScript API (advanced)](excel-add-ins-ranges-advanced.md).</span></span>

## <a name="get-a-range"></a><span data-ttu-id="1f5e8-107">範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-107">Get a range</span></span>

<span data-ttu-id="1f5e8-108">次の例では、ワークシート内の範囲への参照を取得する、さまざまな方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-108">The following examples show different ways to get a reference to a range within a worksheet.</span></span>

### <a name="get-range-by-address"></a><span data-ttu-id="1f5e8-109">アドレスによって範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-109">Get range by address</span></span>

<span data-ttu-id="1f5e8-110">次のコードサンプルでは、 **sample**という名前のワークシートからアドレス**B2: C5**の範囲を取得し、そのプロパティを読み込んで、 `address` コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-110">The following code sample gets the range with address **B2:C5** from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

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

### <a name="get-range-by-name"></a><span data-ttu-id="1f5e8-111">名前によって範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-111">Get range by name</span></span>

<span data-ttu-id="1f5e8-112">次のコードサンプルでは、Sample という名前のワークシートから指定された範囲を取得し、 `MyRange` そのプロパティを読み込んで、 **Sample** `address` コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-112">The following code sample gets the range named `MyRange` from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

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

### <a name="get-used-range"></a><span data-ttu-id="1f5e8-113">使用範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-113">Get used range</span></span>

<span data-ttu-id="1f5e8-114">次のコードサンプルでは、 **sample**という名前のワークシートから使用された範囲を取得し、その `address` プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-114">The following code sample gets the used range from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span> <span data-ttu-id="1f5e8-115">使用範囲とは、値または書式設定が割り当てられているワークシート内のセルを含む、最小の範囲です。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-115">The used range is the smallest range that encompasses any cells in the worksheet that have a value or formatting assigned to them.</span></span> <span data-ttu-id="1f5e8-116">ワークシート全体が空白の場合、このメソッドは、ワークシートの左上の `getUsedRange()` セルのみで構成される範囲を返します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-116">If the entire worksheet is blank, the `getUsedRange()` method returns a range that consists of only the top-left cell in the worksheet.</span></span>

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

### <a name="get-entire-range"></a><span data-ttu-id="1f5e8-117">範囲全体を取得する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-117">Get entire range</span></span>

<span data-ttu-id="1f5e8-118">次のコードサンプルでは、 **sample**という名前のワークシートからワークシートの範囲全体を取得し、その `address` プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-118">The following code sample gets the entire worksheet range from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

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

## <a name="insert-a-range-of-cells"></a><span data-ttu-id="1f5e8-119">セルの範囲を挿入する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-119">Insert a range of cells</span></span>

<span data-ttu-id="1f5e8-120">次のコードサンプルは、場所 **B4:E4** にセルの範囲を挿入し、他のセルを下にシフトして、新しいセルのためのスペースを提供します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-120">The following code sample inserts a range of cells in location **B4:E4** and shifts other cells down to provide space for the new cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-inserted"></a><span data-ttu-id="1f5e8-121">範囲を挿入する前のデータ</span><span class="sxs-lookup"><span data-stu-id="1f5e8-121">Data before range is inserted</span></span>

![範囲を挿入する前の Excel のデータ](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a><span data-ttu-id="1f5e8-123">範囲を挿入した後のデータ</span><span class="sxs-lookup"><span data-stu-id="1f5e8-123">Data after range is inserted</span></span>

![範囲を挿入した後の Excel のデータ](../images/excel-ranges-after-insert.png)

## <a name="clear-a-range-of-cells"></a><span data-ttu-id="1f5e8-125">セルの範囲をクリアする</span><span class="sxs-lookup"><span data-stu-id="1f5e8-125">Clear a range of cells</span></span>

<span data-ttu-id="1f5e8-126">次のコード サンプルは、範囲 **E2：E5** のセルの内容と書式をすべてクリアします。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-126">The following code sample clears all contents and formatting of cells in the range **E2:E5**.</span></span>  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-cleared"></a><span data-ttu-id="1f5e8-127">範囲をクリアする前のデータ</span><span class="sxs-lookup"><span data-stu-id="1f5e8-127">Data before range is cleared</span></span>

![範囲をクリアする前の Excel のデータ](../images/excel-ranges-start.png)

### <a name="data-after-range-is-cleared"></a><span data-ttu-id="1f5e8-129">範囲をクリアした後のデータ</span><span class="sxs-lookup"><span data-stu-id="1f5e8-129">Data after range is cleared</span></span>

![範囲をクリアした後の Excel のデータ](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a><span data-ttu-id="1f5e8-131">セルの範囲を削除する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-131">Delete a range of cells</span></span>

<span data-ttu-id="1f5e8-132">次のコード サンプルは、範囲 **B4:E4** のセルを削除し、他のセルを上にシフトして、削除されたセルのために空いたスペースに入力します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-132">The following code sample deletes the cells in the range **B4:E4** and shift other cells up to fill the space that was vacated by the deleted cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-deleted"></a><span data-ttu-id="1f5e8-133">範囲を削除する前のデータ</span><span class="sxs-lookup"><span data-stu-id="1f5e8-133">Data before range is deleted</span></span>

![範囲を削除する前の Excel のデータ](../images/excel-ranges-start.png)

### <a name="data-after-range-is-deleted"></a><span data-ttu-id="1f5e8-135">範囲を削除した後のデータ</span><span class="sxs-lookup"><span data-stu-id="1f5e8-135">Data after range is deleted</span></span>

![範囲を削除した後の Excel のデータ](../images/excel-ranges-after-delete.png)

## <a name="set-the-selected-range"></a><span data-ttu-id="1f5e8-137">選択範囲を設定する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-137">Set the selected range</span></span>

<span data-ttu-id="1f5e8-138">次のコード サンプルは、作業中のワークシートの範囲 **B2:E6** を選択します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-138">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="1f5e8-139">選択範囲 B2:E6</span><span class="sxs-lookup"><span data-stu-id="1f5e8-139">Selected range B2:E6</span></span>

![Excel の選択範囲](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="1f5e8-141">選択範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-141">Get the selected range</span></span>

<span data-ttu-id="1f5e8-142">次のコードサンプルでは、選択されている範囲を取得し、その `address` プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-142">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

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

## <a name="set-values-or-formulas"></a><span data-ttu-id="1f5e8-143">値または数式を設定する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-143">Set values or formulas</span></span>

<span data-ttu-id="1f5e8-144">次の例は、1 つのセルまたはセルの範囲の値と数式を設定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-144">The following examples show how to set values and formulas for a single cell or a range of cells.</span></span>

### <a name="set-value-for-a-single-cell"></a><span data-ttu-id="1f5e8-145">1 つのセルの値を設定する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-145">Set value for a single cell</span></span>

<span data-ttu-id="1f5e8-146">次のコード サンプルでは、セル **C3** の値を "5" に設定し、データに最も適した列の幅を設定します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-146">The following code sample sets the value of cell **C3** to "5" and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-value-is-updated"></a><span data-ttu-id="1f5e8-147">セルの値が更新される前のデータ</span><span class="sxs-lookup"><span data-stu-id="1f5e8-147">Data before cell value is updated</span></span>

![セルの値が更新される前の Excel のデータ](../images/excel-ranges-set-start.png)

#### <a name="data-after-cell-value-is-updated"></a><span data-ttu-id="1f5e8-149">セルの値が更新された後のデータ</span><span class="sxs-lookup"><span data-stu-id="1f5e8-149">Data after cell value is updated</span></span>

![セルの値が更新された後の Excel のデータ](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a><span data-ttu-id="1f5e8-151">複数のセルの範囲の値を設定する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-151">Set values for a range of cells</span></span>

<span data-ttu-id="1f5e8-152">次のコード サンプルでは、範囲 **B5：D5** のセルの値を設定し、データに最も適した列の幅を設定します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-152">The following code sample sets values for the cells in the range **B5:D5** and then sets the width of the columns to best fit the data.</span></span>

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

#### <a name="data-before-cell-values-are-updated"></a><span data-ttu-id="1f5e8-153">複数のセルの値が更新される前のデータ</span><span class="sxs-lookup"><span data-stu-id="1f5e8-153">Data before cell values are updated</span></span>

![複数のセルの値が更新される前の Excel のデータ](../images/excel-ranges-set-start.png)

#### <a name="data-after-cell-values-are-updated"></a><span data-ttu-id="1f5e8-155">複数のセルの値が更新された後のデータ</span><span class="sxs-lookup"><span data-stu-id="1f5e8-155">Data after cell values are updated</span></span>

![複数のセルの値が更新された後の Excel のデータ](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a><span data-ttu-id="1f5e8-157">1 つのセルの数式を設定する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-157">Set formula for a single cell</span></span>

<span data-ttu-id="1f5e8-158">次のコード サンプルでは、セル **E3** の数式を設定し、データに最も適した列の幅を設定します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-158">The following code sample sets a formula for cell **E3** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-formula-is-set"></a><span data-ttu-id="1f5e8-159">セルの数式が設定される前のデータ</span><span class="sxs-lookup"><span data-stu-id="1f5e8-159">Data before cell formula is set</span></span>

![セルの数式が設定される前の Excel のデータ](../images/excel-ranges-start-set-formula.png)

#### <a name="data-after-cell-formula-is-set"></a><span data-ttu-id="1f5e8-161">セルの数式が設定された後のデータ</span><span class="sxs-lookup"><span data-stu-id="1f5e8-161">Data after cell formula is set</span></span>

![セルの数式が設定された後の Excel のデータ](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a><span data-ttu-id="1f5e8-163">セルの範囲の数式を設定する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-163">Set formulas for a range of cells</span></span>

<span data-ttu-id="1f5e8-164">次のコード サンプルでは、範囲 **E2:E6** のセルの数式を設定し、データに最も適した列の幅を設定します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-164">The following code sample sets formulas for cells in the range **E2:E6** and then sets the width of the columns to best fit the data.</span></span>

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

#### <a name="data-before-cell-formulas-are-set"></a><span data-ttu-id="1f5e8-165">複数のセルの数式が設定される前のデータ</span><span class="sxs-lookup"><span data-stu-id="1f5e8-165">Data before cell formulas are set</span></span>

![複数のセルの数式が設定される前の Excel のデータ](../images/excel-ranges-start-set-formula.png)

#### <a name="data-after-cell-formulas-are-set"></a><span data-ttu-id="1f5e8-167">複数のセルの数式が設定された後のデータ</span><span class="sxs-lookup"><span data-stu-id="1f5e8-167">Data after cell formulas are set</span></span>

![複数のセルの数式が設定された後の Excel のデータ](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a><span data-ttu-id="1f5e8-169">値、テキスト、または数式を取得する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-169">Get values, text, or formulas</span></span>

<span data-ttu-id="1f5e8-170">次の例は、セルの範囲から値、テキスト、および数式を取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-170">These examples show how to get values, text, and formulas from a range of cells.</span></span>

### <a name="get-values-from-a-range-of-cells"></a><span data-ttu-id="1f5e8-171">セルの範囲から値を取得する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-171">Get values from a range of cells</span></span>

<span data-ttu-id="1f5e8-172">次のコードサンプルでは、範囲 **B2: E6**を取得し、その `values` プロパティを読み込んで、その値をコンソールに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-172">The following code sample gets the range **B2:E6**, loads its `values` property, and writes the values to the console.</span></span> <span data-ttu-id="1f5e8-173">`values`範囲のプロパティは、セルに含まれる生の値を指定します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-173">The `values` property of a range specifies the raw values that the cells contain.</span></span> <span data-ttu-id="1f5e8-174">範囲内の一部のセルに数式が含まれている場合でも、 `values` 範囲のプロパティは、それらのセルの生の値 (数式ではなく) を指定します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-174">Even if some cells in a range contain formulas, the `values` property of the range specifies the raw values for those cells, not any of the formulas.</span></span>

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

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a><span data-ttu-id="1f5e8-175">範囲内のデータ (列 E の値は数式の結果)</span><span class="sxs-lookup"><span data-stu-id="1f5e8-175">Data in range (values in column E are a result of formulas)</span></span>

![複数のセルの数式が設定された後の Excel のデータ](../images/excel-ranges-set-formulas.png)

#### <a name="rangevalues-as-logged-to-the-console-by-the-code-sample-above"></a><span data-ttu-id="1f5e8-177">range.values (上記のコード サンプルによりコンソールに記録される)</span><span class="sxs-lookup"><span data-stu-id="1f5e8-177">range.values (as logged to the console by the code sample above)</span></span>

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

### <a name="get-text-from-a-range-of-cells"></a><span data-ttu-id="1f5e8-178">セルの範囲からテキストを取得する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-178">Get text from a range of cells</span></span>

<span data-ttu-id="1f5e8-179">次のコードサンプルでは、範囲 **B2: E6**を取得し、その `text` プロパティを読み込んでコンソールに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-179">The following code sample gets the range **B2:E6**, loads its `text` property, and writes it to the console.</span></span> <span data-ttu-id="1f5e8-180">`text`範囲のプロパティは、範囲内のセルの表示値を指定します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-180">The `text` property of a range specifies the display values for cells in the range.</span></span> <span data-ttu-id="1f5e8-181">範囲内の一部のセルに数式が含まれている場合でも、 `text` 範囲のプロパティは、それらのセルの表示値を指定します。数式は使用できません。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-181">Even if some cells in a range contain formulas, the `text` property of the range specifies the display values for those cells, not any of the formulas.</span></span>

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

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a><span data-ttu-id="1f5e8-182">範囲内のデータ (列 E の値は数式の結果)</span><span class="sxs-lookup"><span data-stu-id="1f5e8-182">Data in range (values in column E are a result of formulas)</span></span>

![複数のセルの数式が設定された後の Excel のデータ](../images/excel-ranges-set-formulas.png)

#### <a name="rangetext-as-logged-to-the-console-by-the-code-sample-above"></a><span data-ttu-id="1f5e8-184">range.text (上記のコード サンプルによりコンソールに記録される)</span><span class="sxs-lookup"><span data-stu-id="1f5e8-184">range.text (as logged to the console by the code sample above)</span></span>

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

### <a name="get-formulas-from-a-range-of-cells"></a><span data-ttu-id="1f5e8-185">セルの範囲から数式を取得する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-185">Get formulas from a range of cells</span></span>

<span data-ttu-id="1f5e8-186">次のコードサンプルでは、範囲 **B2: E6**を取得し、その `formulas` プロパティを読み込んでコンソールに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-186">The following code sample gets the range **B2:E6**, loads its `formulas` property, and writes it to the console.</span></span> <span data-ttu-id="1f5e8-187">`formulas`範囲のプロパティは、数式を含む範囲内のセルの数式と、数式を含まない範囲のセルの生の値を指定します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-187">The `formulas` property of a range specifies the formulas for cells in the range that contain formulas and the raw values for cells in the range that do not contain formulas.</span></span>

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

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a><span data-ttu-id="1f5e8-188">範囲内のデータ (列 E の値は数式の結果)</span><span class="sxs-lookup"><span data-stu-id="1f5e8-188">Data in range (values in column E are a result of formulas)</span></span>

![複数のセルの数式が設定された後の Excel のデータ](../images/excel-ranges-set-formulas.png)

#### <a name="rangeformulas-as-logged-to-the-console-by-the-code-sample-above"></a><span data-ttu-id="1f5e8-190">range.formulas (上記のコード サンプルによりコンソールに記録される)</span><span class="sxs-lookup"><span data-stu-id="1f5e8-190">range.formulas (as logged to the console by the code sample above)</span></span>

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

## <a name="set-range-format"></a><span data-ttu-id="1f5e8-191">範囲の書式を設定する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-191">Set range format</span></span>

<span data-ttu-id="1f5e8-192">次の例は、範囲内のセルのフォントの色、塗りつぶしの色、および数値の書式を設定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-192">The following examples show how to set font color, fill color, and number format for cells in a range.</span></span>

### <a name="set-font-color-and-fill-color"></a><span data-ttu-id="1f5e8-193">フォントの色と塗りつぶしの色を設定する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-193">Set font color and fill color</span></span>

<span data-ttu-id="1f5e8-194">次のコード サンプルは、範囲 **B2：E2** のセルのフォントの色と塗りつぶしの色を設定します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-194">The following code sample sets the font color and fill color for cells in range **B2:E2**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-before-font-color-and-fill-color-are-set"></a><span data-ttu-id="1f5e8-195">フォントの色と塗りつぶしの色を設定する前の範囲内のデータ</span><span class="sxs-lookup"><span data-stu-id="1f5e8-195">Data in range before font color and fill color are set</span></span>

![書式設定する前の Excel のデータ](../images/excel-ranges-format-before.png)

#### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a><span data-ttu-id="1f5e8-197">フォントの色と塗りつぶしの色を設定した後の範囲内のデータ</span><span class="sxs-lookup"><span data-stu-id="1f5e8-197">Data in range after font color and fill color are set</span></span>

![書式設定した後の Excel のデータ](../images/excel-ranges-format-font-and-fill.png)

### <a name="set-number-format"></a><span data-ttu-id="1f5e8-199">数値の書式を設定する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-199">Set number format</span></span>

<span data-ttu-id="1f5e8-200">次のコード サンプルは、範囲 **D3：E5** のセルの数値を書式を設定します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-200">The following code sample sets the number format for the cells in range **D3:E5**.</span></span>

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

#### <a name="data-in-range-before-number-format-is-set"></a><span data-ttu-id="1f5e8-201">数値の書式を設定する前の範囲内のデータ</span><span class="sxs-lookup"><span data-stu-id="1f5e8-201">Data in range before number format is set</span></span>

![数値形式が設定される前の Excel のデータ](../images/excel-ranges-format-font-and-fill.png)

#### <a name="data-in-range-after-number-format-is-set"></a><span data-ttu-id="1f5e8-203">数値の書式を設定した後の範囲内のデータ</span><span class="sxs-lookup"><span data-stu-id="1f5e8-203">Data in range after number format is set</span></span>

![数値形式が設定された後の Excel のデータ](../images/excel-ranges-format-numbers.png)

## <a name="read-or-write-to-an-unbounded-range"></a><span data-ttu-id="1f5e8-205">無制限の範囲への読み取りまたは書き込み</span><span class="sxs-lookup"><span data-stu-id="1f5e8-205">Read or write to an unbounded range</span></span>

### <a name="read-an-unbounded-range"></a><span data-ttu-id="1f5e8-206">無制限の範囲の読み取り</span><span class="sxs-lookup"><span data-stu-id="1f5e8-206">Read an unbounded range</span></span>

<span data-ttu-id="1f5e8-207">非制限範囲アドレスは、列全体または行全体を指定する範囲アドレスです。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-207">An unbounded range address is a range address that specifies either entire columns or entire rows.</span></span> <span data-ttu-id="1f5e8-208">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-208">For example:</span></span>

- <span data-ttu-id="1f5e8-209">範囲のアドレスは列全体で構成されます。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-209">Range addresses comprised of entire columns:</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
- <span data-ttu-id="1f5e8-210">行全体から成る範囲アドレス:</span><span class="sxs-lookup"><span data-stu-id="1f5e8-210">Range addresses comprised of entire rows:</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>

<span data-ttu-id="1f5e8-p107">API が無制限の範囲を取得する要求を行う場合 (`getRange('C:C')` など)、返される応答では、`null`、`values`、`text`、または `numberFormat` などのセル レベルのプロパティに `formula` 値が含まれます。 `address` または `cellCount` など、範囲のその他のプロパティには、無制限の範囲に有効な値が含まれます。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-p107">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`. Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>

### <a name="write-to-an-unbounded-range"></a><span data-ttu-id="1f5e8-213">無制限の範囲への書き込み</span><span class="sxs-lookup"><span data-stu-id="1f5e8-213">Write to an unbounded range</span></span>

<span data-ttu-id="1f5e8-214">`values` `numberFormat` `formula` 入力要求が大きすぎるため、、、などのセルレベルのプロパティを無制限の範囲に設定することはできません。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-214">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on an unbounded range because the input request is too large.</span></span> <span data-ttu-id="1f5e8-215">たとえば、次のコード スニペットは、無制限の範囲に対して `values` を指定しようとしているため無効です。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-215">For example, the following code snippet is not valid because it attempts to specify `values` for an unbounded range.</span></span> <span data-ttu-id="1f5e8-216">無制限の範囲のセルレベルのプロパティを設定しようとすると、API はエラーを返します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-216">The API returns an error if you attempt to set cell-level properties for an unbounded range.</span></span>

```js
var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="read-or-write-to-a-large-range"></a><span data-ttu-id="1f5e8-217">広い範囲に対する読み取りまたは書き込み</span><span class="sxs-lookup"><span data-stu-id="1f5e8-217">Read or write to a large range</span></span>

<span data-ttu-id="1f5e8-p109">範囲に多数のセル、値、数値書式、数式などが含まれる場合、その範囲では API 操作を実行できない場合があります。 API は常に範囲に要求された操作 (特定のデータを取得または書き込む) を実行しようとしますが、広い範囲に対する読み取りや書き込みの操作は、過剰なリソース使用によるエラーになる場合があります。 このようなエラーを避けるため、広い範囲に対して読み取りや書き取り操作を 1 回で実行するのではなく、その範囲の小さいサブセットに対して個別に読み取りまたは書き込み操作を実行することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-p109">If a range contains a large number of cells, values, number formats, and/or formulas, it may not be possible to run API operations on that range. The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization. To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>

<span data-ttu-id="1f5e8-221">システム制限の詳細については、「 [リソースの制限」と「Office アドインのパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)」の「Excel アドイン」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-221">For details on the system limitations, see the "Excel add-ins" section of [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).</span></span>

### <a name="conditional-formatting-of-ranges"></a><span data-ttu-id="1f5e8-222">範囲の条件付き書式</span><span class="sxs-lookup"><span data-stu-id="1f5e8-222">Conditional formatting of ranges</span></span>

<span data-ttu-id="1f5e8-223">範囲には、条件に基づいて個々のセルに適用する書式設定を含めることができます。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-223">Ranges can have formats applied to individual cells based on conditions.</span></span> <span data-ttu-id="1f5e8-224">この詳細については、「[Excel の範囲に条件付き書式を適用する](excel-add-ins-conditional-formatting.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-224">For more information about this, see [Apply conditional formatting to Excel ranges](excel-add-ins-conditional-formatting.md).</span></span>

## <a name="find-a-cell-using-string-matching"></a><span data-ttu-id="1f5e8-225">文字列のマッチングを使用してセルを検索する</span><span class="sxs-lookup"><span data-stu-id="1f5e8-225">Find a cell using string matching</span></span>

<span data-ttu-id="1f5e8-226">`Range` オブジェクトには、範囲内で指定された文字列を検索するための `find` メソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-226">The `Range` object has a `find` method to search for a specified string within the range.</span></span> <span data-ttu-id="1f5e8-227">このメソッドは、一致するテキストがある最初のセルの範囲を返します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-227">It returns the range of the first cell with matching text.</span></span> <span data-ttu-id="1f5e8-228">次のコード サンプルは、文字列 **Food** と等しい値を持つ最初のセルを検索して、そのアドレスをコンソールに記録します。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-228">The following code sample finds the first cell with a value equal to the string **Food** and logs its address to the console.</span></span> <span data-ttu-id="1f5e8-229">指定した文字列が範囲に存在しない場合、`ItemNotFound` エラーが `find` によってスローされます。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-229">Note that `find` throws an `ItemNotFound` error if the specified string doesn't exist in the range.</span></span> <span data-ttu-id="1f5e8-230">指定した文字列が範囲に存在しない可能性がある場合は、自分のコードで適切にシナリオを処理できるように、[findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) メソッドを使用するようにしてください。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-230">If you expect that the specified string may not exist in the range, use the [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) method instead, so your code gracefully handles that scenario.</span></span>

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

<span data-ttu-id="1f5e8-231">単一のセルを表す範囲に対して `find` メソッドが呼び出されると、ワークシート全体が検索されます。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-231">When the `find` method is called on a range representing a single cell, the entire worksheet is searched.</span></span> <span data-ttu-id="1f5e8-232">検索はその単一のセルから始まり、`SearchCriteria.searchDirection` によって指定された方向へ行われ、場合によってはワークシートの最終部分で折り返されます。</span><span class="sxs-lookup"><span data-stu-id="1f5e8-232">The search begins at that cell and goes in the direction specified by `SearchCriteria.searchDirection`, wrapping around the ends of the worksheet if needed.</span></span>

## <a name="see-also"></a><span data-ttu-id="1f5e8-233">関連項目</span><span class="sxs-lookup"><span data-stu-id="1f5e8-233">See also</span></span>

- [<span data-ttu-id="1f5e8-234">Excel JavaScript API を使用して範囲を操作する (高度)</span><span class="sxs-lookup"><span data-stu-id="1f5e8-234">Work with ranges using the Excel JavaScript API (advanced)</span></span>](excel-add-ins-ranges-advanced.md)
- [<span data-ttu-id="1f5e8-235">Office アドインでの Excel JavaScript オブジェクトモデル</span><span class="sxs-lookup"><span data-stu-id="1f5e8-235">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
