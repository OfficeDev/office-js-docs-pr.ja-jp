---
title: Excel JavaScript API を使用してワークシートを操作する
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 9ceb2187cdd7f503fb39171e420adabcc2f13041
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459134"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a><span data-ttu-id="3a5da-102">Excel JavaScript API を使用してワークシートを操作する</span><span class="sxs-lookup"><span data-stu-id="3a5da-102">Work with Worksheets using the Excel JavaScript API</span></span>

<span data-ttu-id="3a5da-103">この記事では、Excel JavaScript API を使用して、ワークシートでタスクを実行する方法のコード サンプルを示しています。</span><span class="sxs-lookup"><span data-stu-id="3a5da-103">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API.</span></span> <span data-ttu-id="3a5da-104">**Worksheet** オブジェクトおよび **WorksheetCollection** オブジェクトがサポートするプロパティとメソッドの完全なリストについては、「[Worksheet オブジェクト (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=office-js)」および「[WorksheetCollection オブジェクト (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection?view=office-js)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3a5da-104">For the complete list of properties and methods that the **Worksheet** and **WorksheetCollection** objects support, see [Worksheet Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=office-js) and [WorksheetCollection Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection?view=office-js).</span></span>

> [!NOTE]
> <span data-ttu-id="3a5da-105">この記事の情報は標準のワークシートにのみ適用されます。"グラフ" シートや "マクロ" シートには適用されません。</span><span class="sxs-lookup"><span data-stu-id="3a5da-105">The information in this article applies only to regular worksheets; it does not apply to "chart" sheets or "macro" sheets.</span></span>

## <a name="get-worksheets"></a><span data-ttu-id="3a5da-106">ワークシートを取得する</span><span class="sxs-lookup"><span data-stu-id="3a5da-106">Get worksheets</span></span>

<span data-ttu-id="3a5da-107">次のコード サンプルでは、ワークシートのコレクションを取得し、各ワークシートの **name** プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="3a5da-107">The following code sample gets the collection of worksheets, loads the **name** property of each worksheet, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    return context.sync()
        .then(function () {
            if (sheets.items.length > 1) {
                console.log(`There are ${sheets.items.length} worksheets in the workbook:`);
            } else {
                console.log(`There is one worksheet in the workbook:`);
            }
            for (var i in sheets.items) {
                console.log(sheets.items[i].name);
            }
        });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="3a5da-108">ワークシートの **id** プロパティは、指定されたブックのワークシートを一意に識別します。その値は、ワークシートの名前変更や移動をしても同じままです。</span><span class="sxs-lookup"><span data-stu-id="3a5da-108">The **id** property of a worksheet uniquely identifies the worksheet in a given workbook and its value will remain the same even when the worksheet is renamed or moved.</span></span> <span data-ttu-id="3a5da-109">Excel for Mac のブックからワークシートを削除すると、削除されたワークシートの **id** はそれ以降に作成される新規ワークシートに再割り当てされる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="3a5da-109">When a worksheet is deleted from a workbook in Excel for Mac, the **id** of the deleted worksheet may be reassigned to a new worksheet that is subsequently created.</span></span>

## <a name="get-the-active-worksheet"></a><span data-ttu-id="3a5da-110">作業中のワークシートを取得する</span><span class="sxs-lookup"><span data-stu-id="3a5da-110">Get the active worksheet</span></span>

<span data-ttu-id="3a5da-111">次のコード サンプルでは、作業中のワークシートを取得し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="3a5da-111">The following code sample gets the active worksheet, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    
    return context.sync()
        .then(function () {
            console.log(`The active worksheet is "${sheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="set-the-active-worksheet"></a><span data-ttu-id="3a5da-112">作業中のワークシートを設定する</span><span class="sxs-lookup"><span data-stu-id="3a5da-112">Set the active worksheet</span></span>

<span data-ttu-id="3a5da-113">次のコード サンプルでは、作業中のワークシートを **Sample** という名前のワークシートに設定し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="3a5da-113">The following code sample sets the active worksheet to the worksheet named **Sample**, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="3a5da-114">その名前を持つワークシートが存在しない場合、**activate()** メソッドにより **ItemNotFound** エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="3a5da-114">If there is no worksheet with that name, the **activate()** method throws an **ItemNotFound** error.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    sheet.activate();
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The active worksheet is "${sheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="reference-worksheets-by-relative-position"></a><span data-ttu-id="3a5da-115">相対位置でワークシートを参照する</span><span class="sxs-lookup"><span data-stu-id="3a5da-115">Reference worksheets by relative position</span></span>

<span data-ttu-id="3a5da-116">以下の例は、相対位置でワークシートを参照する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="3a5da-116">These examples show how to reference a worksheet by its relative position.</span></span>

### <a name="get-the-first-worksheet"></a><span data-ttu-id="3a5da-117">最初のワークシートを取得する</span><span class="sxs-lookup"><span data-stu-id="3a5da-117">Get the first worksheet</span></span>

<span data-ttu-id="3a5da-118">次のコード サンプルでは、ブックの最初のワークシートを取得し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="3a5da-118">The following code sample gets the first worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var firstSheet = context.workbook.worksheets.getFirst();
    firstSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the first worksheet is "${firstSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-the-last-worksheet"></a><span data-ttu-id="3a5da-119">最後のワークシートを取得する</span><span class="sxs-lookup"><span data-stu-id="3a5da-119">Get the last worksheet</span></span>

<span data-ttu-id="3a5da-120">次のコード サンプルでは、ブックの最後のワークシートを取得し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="3a5da-120">The following code sample gets the last worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var lastSheet = context.workbook.worksheets.getLast();
    lastSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the last worksheet is "${lastSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-the-next-worksheet"></a><span data-ttu-id="3a5da-121">次のワークシートを取得する</span><span class="sxs-lookup"><span data-stu-id="3a5da-121">Get the next worksheet</span></span>

<span data-ttu-id="3a5da-122">次のコード サンプルでは、ブックで作業中のワークシートの後のワークシートを取得し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="3a5da-122">The following code sample gets the worksheet that follows the active worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="3a5da-123">作業中のワークシートの後にワークシートがない場合、**getNext()** メソッドにより **ItemNotFound** エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="3a5da-123">If there is no worksheet after the active worksheet, the **getNext()** method throws an **ItemNotFound** error.</span></span>

```js
 Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    var nextSheet = currentSheet.getNext();
    nextSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the sheet that follows the active worksheet is "${nextSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-the-previous-worksheet"></a><span data-ttu-id="3a5da-124">前のワークシートを取得する</span><span class="sxs-lookup"><span data-stu-id="3a5da-124">Get the previous worksheet</span></span>

<span data-ttu-id="3a5da-125">次のコード サンプルでは、ブックで作業中のワークシートの前のワークシートを取得し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="3a5da-125">The following code sample gets the worksheet that precedes the active worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="3a5da-126">作業中のワークシートの前にワークシートが存在しない場合、**getPrevious()** メソッドにより **ItemNotFound** エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="3a5da-126">If there is no worksheet before the active worksheet, the **getPrevious()** method throws an **ItemNotFound** error.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    var previousSheet = currentSheet.getPrevious();
    previousSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the sheet that precedes the active worksheet is "${previousSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="add-a-worksheet"></a><span data-ttu-id="3a5da-127">ワークシートを追加する</span><span class="sxs-lookup"><span data-stu-id="3a5da-127">Add a worksheet</span></span>

<span data-ttu-id="3a5da-p106">次のコード サンプルでは、**Sample** という名前の新しいワークシートをブックに追加し、**name** プロパティと **position** プロパティを読み込み、コンソールにメッセージを書き込みます。新しいワークシートは既存の全ワークシートの後に追加されます。</span><span class="sxs-lookup"><span data-stu-id="3a5da-p106">The following code sample adds a new worksheet named **Sample** to the workbook, loads its **name** and **position** properties, and writes a message to the console. The new worksheet is added after all existing worksheets.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;

    var sheet = sheets.add("Sample");
    sheet.load("name, position");
    
    return context.sync()
        .then(function () {
            console.log(`Added worksheet named "${sheet.name}" in position ${sheet.position}`);
        });
}).catch(errorHandlerFunction);
```

## <a name="delete-a-worksheet"></a><span data-ttu-id="3a5da-130">ワークシートの削除</span><span class="sxs-lookup"><span data-stu-id="3a5da-130">Delete a worksheet</span></span>

<span data-ttu-id="3a5da-131">次のコード サンプルでは、ブックの最後のワークシートを (ただし、ブック内の唯一のシートでない場合に) 削除し、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="3a5da-131">The following code sample deletes the final worksheet in the workbook (as long as it's not the only sheet in the workbook) and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    return context.sync()
        .then(function () {
            if (sheets.items.length === 1) {
                console.log("Unable to delete the only worksheet in the workbook");
            } else {
                var lastSheet = sheets.items[sheets.items.length - 1];

                console.log(`Deleting worksheet named "${lastSheet.name}"`);
                lastSheet.delete();

                return context.sync();
            };
        });
}).catch(errorHandlerFunction);
```

## <a name="rename-a-worksheet"></a><span data-ttu-id="3a5da-132">ワークシートの名前を変更する</span><span class="sxs-lookup"><span data-stu-id="3a5da-132">Rename a worksheet</span></span>

<span data-ttu-id="3a5da-133">次のコード サンプルでは、作業中のワークシートの名前を **New Name** に変更します。</span><span class="sxs-lookup"><span data-stu-id="3a5da-133">The following code sample changes the name of the active worksheet to **New Name**.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-a-worksheet"></a><span data-ttu-id="3a5da-134">ワークシートを移動する</span><span class="sxs-lookup"><span data-stu-id="3a5da-134">Move a worksheet</span></span>

<span data-ttu-id="3a5da-135">次のコード サンプルでは、ブックの最後の位置からブックの最初の位置にワークシートを移動します。</span><span class="sxs-lookup"><span data-stu-id="3a5da-135">The following code sample moves a worksheet from the last position in the workbook to the first position in the workbook.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items");

    return context.sync()
        .then(function () {
            var lastSheet = sheets.items[sheets.items.length - 1];
            lastSheet.position = 0;

            return context.sync();
        });
}).catch(errorHandlerFunction);
```

## <a name="set-worksheet-visibility"></a><span data-ttu-id="3a5da-136">ワークシートの可視性を設定する</span><span class="sxs-lookup"><span data-stu-id="3a5da-136">Set worksheet visibility</span></span>

<span data-ttu-id="3a5da-137">これらの例では、ワークシートの可視性を設定する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="3a5da-137">These examples show how to set the visibility of a worksheet.</span></span>

### <a name="hide-a-worksheet"></a><span data-ttu-id="3a5da-138">ワークシートを非表示にする</span><span class="sxs-lookup"><span data-stu-id="3a5da-138">Hide a worksheet</span></span>

<span data-ttu-id="3a5da-139">次のコード サンプルでは、**Sample** という名前のワークシートの可視性を非表示に設定し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="3a5da-139">The following code sample sets the visibility of worksheet named **Sample** to hidden, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    sheet.visibility = Excel.SheetVisibility.hidden;
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`Worksheet with name "${sheet.name}" is hidden`);
        });
}).catch(errorHandlerFunction);
```

### <a name="unhide-a-worksheet"></a><span data-ttu-id="3a5da-140">ワークシートを再表示する</span><span class="sxs-lookup"><span data-stu-id="3a5da-140">Unhide a worksheet</span></span>

<span data-ttu-id="3a5da-141">次のコード サンプルでは、**Sample** という名前のワークシートの可視性を表示に設定し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="3a5da-141">The following code sample sets the visibility of worksheet named **Sample** to visible, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    sheet.visibility = Excel.SheetVisibility.visible;
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`Worksheet with name "${sheet.name}" is visible`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-a-cell-within-a-worksheet"></a><span data-ttu-id="3a5da-142">ワークシート内でセルを取得する</span><span class="sxs-lookup"><span data-stu-id="3a5da-142">Get a cell within a worksheet</span></span>

<span data-ttu-id="3a5da-143">次のコード サンプルでは、**Sample** という名前のワークシートの 2 行目、5 列目にあるセルを取得し、**address** プロパティと **values** プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="3a5da-143">The following code sample gets the cell that is located in row 2, column 5 of the worksheet named **Sample**, loads its **address** and **values** properties, and writes a message to the console.</span></span> <span data-ttu-id="3a5da-144">**getCell(行番号、列番号)** メソッドに渡される値は、取得するセルの 0 から始まる行番号および列番号です。</span><span class="sxs-lookup"><span data-stu-id="3a5da-144">The values that are passed into the **getCell(row: number, column:number)** method are the zero-indexed row number and column number for the cell that is being retrieved.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var cell = sheet.getCell(1, 4);
    cell.load("address, values");
    
    return context.sync()
        .then(function() {
            console.log(`The value of the cell in row 2, column 5 is "${cell.values[0][0]}" and the address of that cell is "${cell.address}"`);
        })
}).catch(errorHandlerFunction);
```

## <a name="get-a-range-within-a-worksheet"></a><span data-ttu-id="3a5da-145">ワークシート内の範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="3a5da-145">Get a range within a worksheet</span></span>

<span data-ttu-id="3a5da-146">ワークシート内の範囲を取得する方法を示す例については、「[Excel の JavaScript API を使用して範囲を操作する](excel-add-ins-ranges.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3a5da-146">For examples that show how to get a range within a worksheet, see [Work with Ranges using the Excel JavaScript API](excel-add-ins-ranges.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="3a5da-147">関連項目</span><span class="sxs-lookup"><span data-stu-id="3a5da-147">See also</span></span>

- [<span data-ttu-id="3a5da-148">Excel の JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="3a5da-148">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)

