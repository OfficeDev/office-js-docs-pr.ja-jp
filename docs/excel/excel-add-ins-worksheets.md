---
title: Excel JavaScript API を使用してワークシートを操作する
description: ''
ms.date: 02/20/2019
localization_priority: Priority
ms.openlocfilehash: 825ae88afd98afbcd268716c93ddcb13d24a9a1e
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871557"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a><span data-ttu-id="ffa75-102">Excel JavaScript API を使用してワークシートを操作する</span><span class="sxs-lookup"><span data-stu-id="ffa75-102">Work with worksheets using the Excel JavaScript API</span></span>

<span data-ttu-id="ffa75-p101">この記事では、Excel JavaScript API を使用して、ワークシートでタスクを実行する方法のコード サンプルを示しています。 **Worksheet** オブジェクトおよび **WorksheetCollection** オブジェクトがサポートするプロパティとメソッドの完全なリストについては、「[Worksheet オブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.worksheet)」および「[WorksheetCollection オブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.worksheetcollection)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ffa75-p101">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API. For the complete list of properties and methods that the **Worksheet** and **WorksheetCollection** objects support, see [Worksheet Object (JavaScript API for Excel)](/javascript/api/excel/excel.worksheet) and [WorksheetCollection Object (JavaScript API for Excel)](/javascript/api/excel/excel.worksheetcollection).</span></span>

> [!NOTE]
> <span data-ttu-id="ffa75-105">この記事の情報は標準のワークシートにのみ適用されます。"グラフ" シートや "マクロ" シートには適用されません。</span><span class="sxs-lookup"><span data-stu-id="ffa75-105">The information in this article applies only to regular worksheets; it does not apply to "chart" sheets or "macro" sheets.</span></span>

## <a name="get-worksheets"></a><span data-ttu-id="ffa75-106">ワークシートを取得する</span><span class="sxs-lookup"><span data-stu-id="ffa75-106">Get worksheets</span></span>

<span data-ttu-id="ffa75-107">次のコード サンプルでは、ワークシートのコレクションを取得し、各ワークシートの **name** プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="ffa75-107">The following code sample gets the collection of worksheets, loads the **name** property of each worksheet, and writes a message to the console.</span></span>

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
> <span data-ttu-id="ffa75-p102">ワークシートの **id** プロパティは、指定されたブックのワークシートを一意に識別します。その値は、ワークシートの名前変更や移動をしても同じままです。 Excel for Mac のブックからワークシートを削除すると、削除されたワークシートの **id** はそれ以降に作成される新規ワークシートに再割り当てされる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="ffa75-p102">The **id** property of a worksheet uniquely identifies the worksheet in a given workbook and its value will remain the same even when the worksheet is renamed or moved. When a worksheet is deleted from a workbook in Excel for Mac, the **id** of the deleted worksheet may be reassigned to a new worksheet that is subsequently created.</span></span>

## <a name="get-the-active-worksheet"></a><span data-ttu-id="ffa75-110">作業中のワークシートを取得する</span><span class="sxs-lookup"><span data-stu-id="ffa75-110">Get the active worksheet</span></span>

<span data-ttu-id="ffa75-111">次のコード サンプルでは、作業中のワークシートを取得し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="ffa75-111">The following code sample gets the active worksheet, loads its **name** property, and writes a message to the console.</span></span>

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

## <a name="set-the-active-worksheet"></a><span data-ttu-id="ffa75-112">作業中のワークシートを設定する</span><span class="sxs-lookup"><span data-stu-id="ffa75-112">Set the active worksheet</span></span>

<span data-ttu-id="ffa75-p103">次のコード サンプルでは、作業中のワークシートを **Sample** という名前のワークシートに設定し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。 その名前を持つワークシートが存在しない場合、**activate()** メソッドにより **ItemNotFound** エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="ffa75-p103">The following code sample sets the active worksheet to the worksheet named **Sample**, loads its **name** property, and writes a message to the console. If there is no worksheet with that name, the **activate()** method throws an **ItemNotFound** error.</span></span>

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

## <a name="reference-worksheets-by-relative-position"></a><span data-ttu-id="ffa75-115">相対位置でワークシートを参照する</span><span class="sxs-lookup"><span data-stu-id="ffa75-115">Reference worksheets by relative position</span></span>

<span data-ttu-id="ffa75-116">以下の例は、相対位置でワークシートを参照する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="ffa75-116">These examples show how to reference a worksheet by its relative position.</span></span>

### <a name="get-the-first-worksheet"></a><span data-ttu-id="ffa75-117">最初のワークシートを取得する</span><span class="sxs-lookup"><span data-stu-id="ffa75-117">Get the first worksheet</span></span>

<span data-ttu-id="ffa75-118">次のコード サンプルでは、ブックの最初のワークシートを取得し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="ffa75-118">The following code sample gets the first worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

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

### <a name="get-the-last-worksheet"></a><span data-ttu-id="ffa75-119">最後のワークシートを取得する</span><span class="sxs-lookup"><span data-stu-id="ffa75-119">Get the last worksheet</span></span>

<span data-ttu-id="ffa75-120">次のコード サンプルでは、ブックの最後のワークシートを取得し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="ffa75-120">The following code sample gets the last worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

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

### <a name="get-the-next-worksheet"></a><span data-ttu-id="ffa75-121">次のワークシートを取得する</span><span class="sxs-lookup"><span data-stu-id="ffa75-121">Get the next worksheet</span></span>

<span data-ttu-id="ffa75-p104">次のコード サンプルでは、ブックで作業中のワークシートの後のワークシートを取得し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。 作業中のワークシートの後にワークシートがない場合、**getNext()** メソッドにより **ItemNotFound** エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="ffa75-p104">The following code sample gets the worksheet that follows the active worksheet in the workbook, loads its **name** property, and writes a message to the console. If there is no worksheet after the active worksheet, the **getNext()** method throws an **ItemNotFound** error.</span></span>

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

### <a name="get-the-previous-worksheet"></a><span data-ttu-id="ffa75-124">前のワークシートを取得する</span><span class="sxs-lookup"><span data-stu-id="ffa75-124">Get the previous worksheet</span></span>

<span data-ttu-id="ffa75-p105">次のコード サンプルでは、ブックで作業中のワークシートの前のワークシートを取得し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。 作業中のワークシートの前にワークシートが存在しない場合、**getPrevious()** メソッドにより **ItemNotFound** エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="ffa75-p105">The following code sample gets the worksheet that precedes the active worksheet in the workbook, loads its **name** property, and writes a message to the console. If there is no worksheet before the active worksheet, the **getPrevious()** method throws an **ItemNotFound** error.</span></span>

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

## <a name="add-a-worksheet"></a><span data-ttu-id="ffa75-127">ワークシートを追加する</span><span class="sxs-lookup"><span data-stu-id="ffa75-127">Add a worksheet</span></span>

<span data-ttu-id="ffa75-p106">次のコード サンプルでは、**Sample** という名前の新しいワークシートをブックに追加し、**name** プロパティと **position** プロパティを読み込み、コンソールにメッセージを書き込みます。新しいワークシートは既存の全ワークシートの後に追加されます。</span><span class="sxs-lookup"><span data-stu-id="ffa75-p106">The following code sample adds a new worksheet named **Sample** to the workbook, loads its **name** and **position** properties, and writes a message to the console. The new worksheet is added after all existing worksheets.</span></span>

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

## <a name="delete-a-worksheet"></a><span data-ttu-id="ffa75-130">ワークシートの削除</span><span class="sxs-lookup"><span data-stu-id="ffa75-130">Delete a worksheet</span></span>

<span data-ttu-id="ffa75-131">次のコード サンプルでは、ブックの最後のワークシートを (ただし、ブック内の唯一のシートでない場合に) 削除し、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="ffa75-131">The following code sample deletes the final worksheet in the workbook (as long as it's not the only sheet in the workbook) and writes a message to the console.</span></span>

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

> [!NOTE]
> <span data-ttu-id="ffa75-132">可視性が [Very Hidden](/javascript/api/excel/excel.sheetvisibility) のワークシートは、`delete` メソッドで削除することはできません。</span><span class="sxs-lookup"><span data-stu-id="ffa75-132">A worksheet with a visibility of "[Very Hidden](/javascript/api/excel/excel.sheetvisibility)" cannot be deleted with the `delete` method.</span></span> <span data-ttu-id="ffa75-133">このワークシートを削除する場合には、最初に可視性を変更する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ffa75-133">If you wish to delete the worksheet anyway, you must first change the visibility.</span></span>

## <a name="rename-a-worksheet"></a><span data-ttu-id="ffa75-134">ワークシートの名前を変更する</span><span class="sxs-lookup"><span data-stu-id="ffa75-134">Rename a worksheet</span></span>

<span data-ttu-id="ffa75-135">次のコード サンプルでは、作業中のワークシートの名前を **New Name** に変更します。</span><span class="sxs-lookup"><span data-stu-id="ffa75-135">The following code sample changes the name of the active worksheet to **New Name**.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-a-worksheet"></a><span data-ttu-id="ffa75-136">ワークシートを移動する</span><span class="sxs-lookup"><span data-stu-id="ffa75-136">Move a worksheet</span></span>

<span data-ttu-id="ffa75-137">次のコード サンプルでは、ブックの最後の位置からブックの最初の位置にワークシートを移動します。</span><span class="sxs-lookup"><span data-stu-id="ffa75-137">The following code sample moves a worksheet from the last position in the workbook to the first position in the workbook.</span></span>

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

## <a name="set-worksheet-visibility"></a><span data-ttu-id="ffa75-138">ワークシートの可視性を設定する</span><span class="sxs-lookup"><span data-stu-id="ffa75-138">Set worksheet visibility</span></span>

<span data-ttu-id="ffa75-139">これらの例では、ワークシートの可視性を設定する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="ffa75-139">These examples show how to set the visibility of a worksheet.</span></span>

### <a name="hide-a-worksheet"></a><span data-ttu-id="ffa75-140">ワークシートを非表示にする</span><span class="sxs-lookup"><span data-stu-id="ffa75-140">Hide a worksheet</span></span>

<span data-ttu-id="ffa75-141">次のコード サンプルでは、**Sample** という名前のワークシートの可視性を非表示に設定し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="ffa75-141">The following code sample sets the visibility of worksheet named **Sample** to hidden, loads its **name** property, and writes a message to the console.</span></span>

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

### <a name="unhide-a-worksheet"></a><span data-ttu-id="ffa75-142">ワークシートを再表示する</span><span class="sxs-lookup"><span data-stu-id="ffa75-142">Unhide a worksheet</span></span>

<span data-ttu-id="ffa75-143">次のコード サンプルでは、**Sample** という名前のワークシートの可視性を表示に設定し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="ffa75-143">The following code sample sets the visibility of worksheet named **Sample** to visible, loads its **name** property, and writes a message to the console.</span></span>

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

## <a name="get-a-single-cell-within-a-worksheet"></a><span data-ttu-id="ffa75-144">ワークシート内で単一のセルを取得する</span><span class="sxs-lookup"><span data-stu-id="ffa75-144">Get a single cell within a worksheet</span></span>

<span data-ttu-id="ffa75-145">次のコード サンプルでは、**Sample** という名前のワークシートの 2 行目、5 列目にあるセルを取得し、**address** プロパティと **values** プロパティを読み込み、コンソールにメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="ffa75-145">The following code sample gets the cell that is located in row 2, column 5 of the worksheet named **Sample**, loads its **address** and **values** properties, and writes a message to the console.</span></span> <span data-ttu-id="ffa75-146">`getCell(row: number, column:number)` メソッドに渡される値は、取得するセルの 0 から始まる行番号および列番号です。</span><span class="sxs-lookup"><span data-stu-id="ffa75-146">The values that are passed into the `getCell(row: number, column:number)` method are the zero-indexed row number and column number for the cell that is being retrieved.</span></span>

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

## <a name="find-all-cells-with-matching-text-preview"></a><span data-ttu-id="ffa75-147">一致するテキストがあるすべてのセルを検索する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="ffa75-147">Find all cells with matching text (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="ffa75-148">現在、Worksheet オブジェクトの `findAll` 関数は、パブリック プレビューでのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="ffa75-148">The Worksheet object's `findAll` function is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="ffa75-149">`Worksheet` オブジェクトには、ワークシート内の指定された文字列を検索するための `find` メソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="ffa75-149">The `Worksheet` object has a `find` method to search for a specified string within the worksheet.</span></span> <span data-ttu-id="ffa75-150">このメソッドは `RangeAreas` オブジェクトを返します。これは、一度に編集できる `Range` オブジェクトのコレクションとなります。</span><span class="sxs-lookup"><span data-stu-id="ffa75-150">It returns a `RangeAreas` object, which is a collection of `Range` objects that can be edited all at once.</span></span> <span data-ttu-id="ffa75-151">以下のコード サンプルは、文字列 **Complete** と等しいすべてのセルを検索し、そのセルの色を緑色にします。</span><span class="sxs-lookup"><span data-stu-id="ffa75-151">The following code sample finds all cells with values equal to the string **Complete** and colors them green.</span></span> <span data-ttu-id="ffa75-152">指定した文字列がワークシートに存在しない場合、`ItemNotFound` エラーが `findAll` によってスローされます。</span><span class="sxs-lookup"><span data-stu-id="ffa75-152">Note that `findAll` will throw an `ItemNotFound` error if the specified string doesn't exist in the worksheet.</span></span> <span data-ttu-id="ffa75-153">指定した文字列がワークシートに存在しない可能性がある場合は、自分のコードで適切にシナリオを処理できるように、[findAllOrNullObject](excel-add-ins-advanced-concepts.md#ornullobject-methods) メソッドを使用するようにしてください。</span><span class="sxs-lookup"><span data-stu-id="ffa75-153">If you expect that the specified string may not exist in the worksheet, use the [findAllOrNullObject](excel-add-ins-advanced-concepts.md#ornullobject-methods) method instead, so your code gracefully handles that scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var foundRanges = sheet.findAll("Complete", {
        completeMatch: true, // findAll will match the whole cell value
        matchCase: false // findAll will not match case
    });

    return context.sync()
        .then(function() {
            foundRanges.format.fill.color = "green"
    });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="ffa75-154">このセクションでは、`Worksheet` オブジェクトの関数を使用してセルと範囲を検索する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="ffa75-154">This section describes how to find cells and ranges using the `Worksheet` object's functions.</span></span> <span data-ttu-id="ffa75-155">範囲の取得の詳細については、オブジェクト専用の記事で確認することができます。</span><span class="sxs-lookup"><span data-stu-id="ffa75-155">More range retrieval information can be found in object-specific articles.</span></span>
> - <span data-ttu-id="ffa75-156">`Range` オブジェクトを使用して、ワークシート内の範囲を取得する方法を示す例については、「[Excel JavaScript API を使用して範囲を操作する](excel-add-ins-ranges.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ffa75-156">For examples that show how to get a range within a worksheet using the `Range` object, see [Work with ranges using the Excel JavaScript API](excel-add-ins-ranges.md).</span></span>
> - <span data-ttu-id="ffa75-157">`Table` オブジェクトから範囲を取得する方法を示す例については、「[Excel JavaScript API を使用して表を操作する](excel-add-ins-tables.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ffa75-157">For examples that show how to get ranges from a `Table` object, see [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md).</span></span>
> - <span data-ttu-id="ffa75-158">セルの特性に基づいて複数の副範囲を幅広く検索する方法の例については、「[Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ffa75-158">For examples that show how to search a large range for multiple sub-ranges based on cell characteristics, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="data-protection"></a><span data-ttu-id="ffa75-159">データの保護</span><span class="sxs-lookup"><span data-stu-id="ffa75-159">Data protection</span></span>

<span data-ttu-id="ffa75-160">ご使用のアドインでは、ワークシート内のデータを編集するユーザー機能を制御できます。</span><span class="sxs-lookup"><span data-stu-id="ffa75-160">Your add-in can control a user's ability to edit data in a worksheet.</span></span> <span data-ttu-id="ffa75-161">ワークシートの `protection` プロパティは [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) オブジェクトであり、`protect()` メソッドを備えています。</span><span class="sxs-lookup"><span data-stu-id="ffa75-161">The worksheet's `protection` property is a [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) object with a `protect()` method.</span></span> <span data-ttu-id="ffa75-162">次の例では、アクティブなワークシートの完全な保護を切り替える基本的なシナリオを示します。</span><span class="sxs-lookup"><span data-stu-id="ffa75-162">The following example shows a basic scenario toggling the complete protection of the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var activeSheet = context.workbook.worksheets.getActiveWorksheet();
    activeSheet.load("protection/protected");

    return context.sync().then(function() {
        if (!activeSheet.protection.protected) {
            activeSheet.protection.protect();
        }
    })
}).catch(errorHandlerFunction);
```

<span data-ttu-id="ffa75-163">`protect` メソッドには、2 つの省略可能なパラメーターがあります。</span><span class="sxs-lookup"><span data-stu-id="ffa75-163">The `protect` method has two optional parameters:</span></span>

- <span data-ttu-id="ffa75-164">`options`: 特定の編集制限を定義する [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions) オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="ffa75-164">`options`: A [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions) object defining specific editing restrictions.</span></span>
- <span data-ttu-id="ffa75-165">`password`: ユーザーが保護をバイパスしてワークシートを編集するために必要なパスワードを表す文字列。</span><span class="sxs-lookup"><span data-stu-id="ffa75-165">`password`: A string representing the password needed for a user to bypass protection and edit the worksheet.</span></span>

<span data-ttu-id="ffa75-166">ワークシートの保護と、Excel の UI を使用してそれを変更する方法の詳細については、記事「[ワークシートを保護する](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ffa75-166">The article [Protect a worksheet](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) has more information about worksheet protection and how to change it through the Excel UI.</span></span>

## <a name="see-also"></a><span data-ttu-id="ffa75-167">関連項目</span><span class="sxs-lookup"><span data-stu-id="ffa75-167">See also</span></span>

- [<span data-ttu-id="ffa75-168">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="ffa75-168">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
