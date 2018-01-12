# <a name="work-with-worksheets-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してワークシートを操作する

この記事では、Excel JavaScript API を使用して、ワークシートでタスクを実行する方法のコード サンプルを示しています。 **Worksheet** オブジェクトおよび **WorksheetCollection** オブジェクトがサポートするプロパティとメソッドの完全なリストについては、「[Worksheet オブジェクト (JavaScript API for Excel)](../../reference/excel/worksheet.md)」および「[WorksheetCollection オブジェクト (JavaScript API for Excel)](../../reference/excel/worksheetcollection.md)」を参照してください。

**注**:この記事の情報は標準のワークシートにのみ適用されます。"グラフ" シートや "マクロ" シートには適用されません。

## <a name="get-worksheets"></a>ワークシートを取得する

次のコード サンプルでは、ワークシートのコレクションを取得し、各ワークシートの **name** プロパティを読み込み、コンソールにメッセージを書き込みます。

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

**注**:ワークシートの **id** プロパティは、指定されたブックのワークシートを一意に識別します。その値は、ワークシートの名前変更や移動をしても同じままです。 Excel for Mac のブックからワークシートを削除すると、削除されたワークシートの **id** はそれ以降に作成される新規ワークシートに再割り当てされる可能性があります。

## <a name="get-the-active-worksheet"></a>作業中のワークシートを取得する

次のコード サンプルでは、作業中のワークシートを取得し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。

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

## <a name="set-the-active-worksheet"></a>作業中のワークシートを設定する

次のコード サンプルでは、作業中のワークシートを **Sample** という名前のワークシートに設定し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。 その名前を持つワークシートが存在しない場合、**activate()** メソッドにより **ItemNotFound** エラーがスローされます。

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

## <a name="reference-worksheets-by-relative-position"></a>相対位置でワークシートを参照する

以下の例は、相対位置でワークシートを参照する方法を示しています。

### <a name="get-the-first-worksheet"></a>最初のワークシートを取得する

次のコード サンプルでは、ブックの最初のワークシートを取得し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。

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

### <a name="get-the-last-worksheet"></a>最後のワークシートを取得する

次のコード サンプルでは、ブックの最後のワークシートを取得し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。

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

### <a name="get-the-next-worksheet"></a>次のワークシートを取得する

次のコード サンプルでは、ブックで作業中のワークシートの後のワークシートを取得し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。 作業中のワークシートの後にワークシートがない場合、**getNext()** メソッドにより **ItemNotFound** エラーがスローされます。

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

### <a name="get-the-previous-worksheet"></a>前のワークシートを取得する

次のコード サンプルでは、ブックで作業中のワークシートの前のワークシートを取得し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。 作業中のワークシートの前にワークシートが存在しない場合、**getPrevious()** メソッドにより **ItemNotFound** エラーがスローされます。

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

## <a name="add-a-worksheet"></a>ワークシートを追加する

次のコード サンプルでは、**Sample** という名前の新しいワークシートをブックに追加し、**name** プロパティと **position** プロパティを読み込み、コンソールにメッセージを書き込みます。 既存のすべてのワークシートの後に、新しいワークシートが追加されます。

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

## <a name="delete-a-worksheet"></a>ワークシートを削除する

次のコード サンプルでは、ブックの最後のワークシートを (ただし、ブック内の唯一のシートでない場合に) 削除し、コンソールにメッセージを書き込みます。

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

## <a name="rename-a-worksheet"></a>ワークシートの名前を変更する

次のコード サンプルでは、作業中のワークシートの名前を **New Name** に変更します。

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-a-worksheet"></a>ワークシートを移動する

次のコード サンプルでは、ブックの最後の位置からブックの最初の位置にワークシートを移動します。

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

## <a name="set-worksheet-visibility"></a>ワークシートの可視性を設定する

これらの例では、ワークシートの可視性を設定する方法を示します。

### <a name="hide-a-worksheet"></a>ワークシートを非表示にする

次のコード サンプルでは、**Sample** という名前のワークシートの可視性を非表示に設定し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。

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

### <a name="unhide-a-worksheet"></a>ワークシートを再表示する

次のコード サンプルでは、**Sample** という名前のワークシートの可視性を表示に設定し、**name** プロパティを読み込み、コンソールにメッセージを書き込みます。

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

## <a name="get-a-cell-within-a-worksheet"></a>ワークシート内でセルを取得する

次のコード サンプルでは、**Sample** という名前のワークシートの 2 行目、5 列目にあるセルを取得し、**address** プロパティと **values** プロパティを読み込み、コンソールにメッセージを書き込みます。 **getCell(行番号、列番号)** メソッドに渡される値は、取得するセルの 0 から始まる行番号および列番号です。

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

## <a name="get-a-range-within-a-worksheet"></a>ワークシート内の範囲を取得する

ワークシート内の範囲を取得する方法を示す例については、「[Excel の JavaScript API を使用して範囲を操作する](excel-add-ins-ranges.md)」を参照してください。

## <a name="additional-resources"></a>その他のリソース

- [Excel JavaScript API の中心概念](excel-add-ins-core-concepts.md)
- [Worksheet オブジェクト (JavaScript API for Excel)](../../reference/excel/worksheet.md)
- [WorksheetCollection オブジェクト (JavaScript API for Excel)](../../reference/excel/worksheetcollection.md)
