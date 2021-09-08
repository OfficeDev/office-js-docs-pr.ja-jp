---
title: Excel JavaScript API を使用してワークシートを操作する
description: JavaScript API を使用してワークシートで一般的なタスクを実行する方法を示Excelコード サンプル。
ms.date: 07/02/2021
localization_priority: Normal
ms.openlocfilehash: 2c0c046d060e9ed32d872307f27784ff8337b100
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936943"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してワークシートを操作する

この記事では、Excel JavaScript API を使用して、ワークシートでタスクを実行する方法のコード サンプルを示しています。 `Worksheet` オブジェクトおよび `WorksheetCollection` オブジェクトがサポートするプロパティとメソッドの完全なリストについては、「[Worksheet オブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.worksheet)」および「[WorksheetCollection オブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.worksheetcollection)」を参照してください。

> [!NOTE]
> この記事の情報は標準のワークシートにのみ適用されます。"グラフ" シートや "マクロ" シートには適用されません。

## <a name="get-worksheets"></a>ワークシートを取得する

次のコード サンプルでは、ワークシートのコレクションを取得し、各ワークシートの `name` プロパティを読み込み、コンソールにメッセージを書き込みます。

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
            sheets.items.forEach(function (sheet) {
              console.log(sheet.name);
            });
        });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> ワークシートの `id` プロパティは、指定されたブックのワークシートを一意に識別します。その値は、ワークシートの名前変更や移動をしても同じままです。Mac 版の Excel のブックからワークシートを削除すると、削除されたワークシートの `id` はそれ以降に作成される新規ワークシートに再割り当てされる可能性があります。

## <a name="get-the-active-worksheet"></a>作業中のワークシートを取得する

次のコード サンプルでは、作業中のワークシートを取得し、`name` プロパティを読み込み、コンソールにメッセージを書き込みます。

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

次のコード サンプルでは、作業中のワークシートを **Sample** という名前のワークシートに設定し、`name` プロパティを読み込み、コンソールにメッセージを書き込みます。 その名前を持つワークシートが存在しない場合、`activate()` メソッドにより `ItemNotFound` エラーがスローされます。

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

次のコード サンプルでは、ブックの最初のワークシートを取得し、`name` プロパティを読み込み、コンソールにメッセージを書き込みます。

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

次のコード サンプルでは、ブックの最後のワークシートを取得し、`name` プロパティを読み込み、コンソールにメッセージを書き込みます。

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

次のコード サンプルでは、ブックで作業中のワークシートの後のワークシートを取得し、`name` プロパティを読み込み、コンソールにメッセージを書き込みます。 作業中のワークシートの後にワークシートがない場合、`getNext()` メソッドにより `ItemNotFound` エラーがスローされます。

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

次のコード サンプルでは、ブックで作業中のワークシートの前のワークシートを取得し、`name` プロパティを読み込み、コンソールにメッセージを書き込みます。 作業中のワークシートの前にワークシートが存在しない場合、`getPrevious()` メソッドにより `ItemNotFound` エラーがスローされます。

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

次のコード サンプルでは、**Sample** という名前の新しいワークシートをブックに追加し、`name` プロパティと `position` プロパティを読み込み、コンソールにメッセージを書き込みます。新しいワークシートは既存の全ワークシートの後に追加されます。

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

### <a name="copy-an-existing-worksheet"></a>既存のワークシートをコピーする

`Worksheet.copy` は、既存のワークシートのコピーである新しいワークシートを追加します。 新しいワークシートの名前には、Excel UI を介してワークシートをコピーするのと一貫した方法で、末尾に番号が追加されます (たとえば、**MySheet (2)**)。 `Worksheet.copy` は 2 つのパラメーターを取ることができますが、どちらもオプションです。

- `positionType` - ブック内の新しいワークシートを追加する場所を指定する [WorksheetPositionType](/javascript/api/excel/excel.worksheetpositiontype) 列挙。
- `relativeTo` - `positionType` が `Before` または `After` である場合、新しいシートを追加するワークシートを指定する必要があります (このパラメーターは、「何の前か後に?」という質問に答えます)。

次のコード サンプルは、現在のワークシートをコピーし、現在のワークシートの直後に新しいシートを挿入します。

```js
Excel.run(function (context) {
    var myWorkbook = context.workbook;
    var sampleSheet = myWorkbook.worksheets.getActiveWorksheet();
    var copiedSheet = sampleSheet.copy(Excel.WorksheetPositionType.after, sampleSheet);
    return context.sync();
});
```

## <a name="delete-a-worksheet"></a>ワークシートの削除

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

> [!NOTE]
> 可視性が [Very Hidden](/javascript/api/excel/excel.sheetvisibility) のワークシートは、`delete` メソッドで削除することはできません。 このワークシートを削除する場合には、最初に可視性を変更する必要があります。

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

次のコード サンプルでは、**Sample** という名前のワークシートの可視性を非表示に設定し、`name` プロパティを読み込み、コンソールにメッセージを書き込みます。

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

次のコード サンプルでは、**Sample** という名前のワークシートの可視性を表示に設定し、`name` プロパティを読み込み、コンソールにメッセージを書き込みます。

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

## <a name="get-a-single-cell-within-a-worksheet"></a>ワークシート内で単一のセルを取得する

次のコード サンプルでは、**Sample** という名前のワークシートの 2 行目、5 列目にあるセルを取得し、`address` プロパティと `values` プロパティを読み込み、コンソールにメッセージを書き込みます。 `getCell(row: number, column:number)` メソッドに渡される値は、取得するセルの 0 から始まる行番号および列番号です。

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

## <a name="detect-data-changes"></a>データ変更を検出します

表のデータをユーザーが変更した場合に、アドインを使用して対応する必要がある場合があります。 これらの変更を検出するために、`onChanged`ワークシートのイベントに対する[イベントハンドラを登録できます](excel-add-ins-events.md#register-an-event-handler)。 `onChanged`イベントのイベント ハンドラーは、そのイベントが発生した際に [TableChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs) オブジェクトを受け取ります。

この`WorksheetChangedEventArgs`オブジェクトは、変更とソースに関する情報を提供します。 `onChanged` が発生するのは書式設定またはデータの値が変更された時であるため、値が実際に変更されたかどうかを確認するのにアドインを使用すると便利です。 `details`プロパティは、この情報を [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail) としてカプセル化します。 次のコード サンプルでは、変更前と変更後の値および変更されたセルの種類を表示する方法を表示します。

```js
// This function would be used as an event handler for the Worksheet.onChanged event.
function onWorksheetChanged(eventArgs) {
    Excel.run(function (context) {
        var details = eventArgs.details;
        var address = eventArgs.address;

        // Print the before and after types and values to the console.
        console.log(`Change at ${address}: was ${details.valueBefore}(${details.valueTypeBefore}),`
            + ` now is ${details.valueAfter}(${details.valueTypeAfter})`);
        return context.sync();
    });
}
```

## <a name="detect-formula-changes"></a>数式の変更を検出する

アドインは、ワークシート内の数式の変更を追跡できます。 これは、ワークシートが外部データベースに接続されている場合に便利です。 ワークシート内の数式が変更されると、このシナリオのイベントによって外部データベースの対応する更新プログラムがトリガーされます。

数式の変更を検出するには、 [ワークシート](excel-add-ins-events.md#register-an-event-handler) の [onFormulaChanged](/javascript/api/excel/excel.worksheet#onFormulaChanged) イベントのイベント ハンドラーを登録します。 イベントのイベント ハンドラーは、イベントが発生すると `onFormulaChanged` [WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs) オブジェクトを受け取る。

> [!IMPORTANT]
> このイベントは、数式自体が変更された場合を検出します。数式の計算に起因する `onFormulaChanged` データ値は検出しません。

次のコード サンプルは、イベント ハンドラーを登録し、オブジェクトを使用して変更された数式の `onFormulaChanged` `WorksheetFormulaChangedEventArgs` [formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formulaDetails) 配列を取得し [、FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail) プロパティを使用して変更された数式の詳細を出力する方法を示しています。

> [!NOTE]
> このコード サンプルは、1 つの数式が変更された場合にのみ機能します。

```js
Excel.run(function (context) {
    // Retrieve the worksheet named "Sample".
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Register the formula changed event handler for this worksheet.
    sheet.onFormulaChanged.add(formulaChangeHandler);

    return context.sync();
});

function formulaChangeHandler(event) {
    Excel.run(function (context) {
        // Retrieve details about the formula change event.
        // Note: This method assumes only a single formula is changed at a time. 
        var cellAddress = event.formulaDetails[0].cellAddress;
        var previousFormula = event.formulaDetails[0].previousFormula;
        var source = event.source;
    
        // Print out the change event details.
        console.log(
          `The formula in cell ${cellAddress} changed. 
          The previous formula was: ${previousFormula}. 
          The source of the change was: ${source}.`
        );         
    });
}
```

## <a name="handle-sorting-events"></a>並べ替えイベントを処理する

`onColumnSorted` および `onRowSorted` イベントは、ワークシート データがいつ並べ替えられるかを示します。 これらのイベントは、個々の `Worksheet` オブジェクトおよびブックの `WorkbookCollection` に接続されています。 これらは、並べ替えがプログラムで実行されるか、Excel ユーザー インターフェイスを介して手動で実行されるかに関係なく起動します。

> [!NOTE]
> `onColumnSorted` は、左から右への並べ替え操作の結果として列が並べ替えされたときに起動します。 `onRowSorted` は、上から下への並べ替え操作の結果として行が並べ替えされたときに起動します。 列のヘッダーのドロップダウン メニューを使用してテーブルを並べ替えると、`onRowSorted` イベントが発生します。 イベントは、並べ替え条件として考慮されているものではなく、移動しているものに対応します。

`onColumnSorted` および `onRowSorted` イベントは、それぞれ [WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs) または [WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs) でコールバックを提供します。 これらは、イベントの詳細を提供します。 特に、両方の `EventArgs` には、並べ替え操作の結果として移動された行または列を表す `address` プロパティがあります。 セルの値が並べ替え条件の一部ではない場合でも、並べ替えられたコンテンツを持つセルが含まれます。

以下の画像は、並べ替えイベントの `address` プロパティによって返される範囲を示しています。 まず、並べ替えの前のサンプル データを次に示します。

![並べ替え前Excelテーブル データ。](../images/excel-sort-event-before.png)

**"Q1" ("B"** の値) に対して上から下への並べ替えを実行すると、次の強調表示された行がによって返されます `WorksheetRowSortedEventArgs.address` 。

![上から下への並べ替えの後の Excel のテーブル データ。 移動した行が強調表示されます。](../images/excel-sort-event-after-row.png)

元のデータの **"Quinces&quot;**(&quot;**4**") の値に対して左から右への並べ替えを実行すると、次の強調表示された列が返されます `WorksheetColumnsSortedEventArgs.address` 。

![左から右への並べ替えの後の Excel のテーブル データ。 移動した列が強調表示されます。](../images/excel-sort-event-after-column.png)

次のコード サンプルは、`Worksheet.onRowSorted` イベントのイベント ハンドラーを登録する方法を示しています。 ハンドラーのコールバックは、範囲の塗りつぶしの色をクリアし、移動した行のセルを塗りつぶします。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // This will fire whenever a row has been moved as the result of a sort action.
    sheet.onRowSorted.add(function (event) {
        return Excel.run(function (context) {
            console.log("Row sorted: " + event.address);
            var sheet = context.workbook.worksheets.getActiveWorksheet();

            // Clear formatting for section, then highlight the sorted area.
            sheet.getRange("A1:E5").format.fill.clear();
            if (event.address !== "") {
                sheet.getRanges(event.address).format.fill.color = "yellow";
            }

            return context.sync();
        });
    });

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="find-all-cells-with-matching-text"></a>一致するテキストがあるすべてのセルを検索する

`Worksheet` オブジェクトには、ワークシート内の指定された文字列を検索するための `find` メソッドがあります。 このメソッドは `RangeAreas` オブジェクトを返します。これは、一度に編集できる `Range` オブジェクトのコレクションとなります。 以下のコード サンプルは、文字列 **Complete** と等しいすべてのセルを検索し、そのセルの色を緑色にします。 指定した文字列がワークシートに存在しない場合、`ItemNotFound` エラーが `findAll` によってスローされます。 指定した文字列がワークシートに存在しない可能性がある場合は、自分のコードで適切にシナリオを処理できるように、[findAllOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) メソッドを使用するようにしてください。

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
> このセクションでは、`Worksheet` オブジェクトの関数を使用してセルと範囲を検索する方法について説明します。 範囲の取得の詳細については、オブジェクト専用の記事で確認することができます。
> - オブジェクトを使用してワークシート内の範囲を取得する方法を示す例については、「JavaScript API を使用して範囲を取得するExcel `Range` [参照してください](excel-add-ins-ranges-get.md)。
> - `Table` オブジェクトから範囲を取得する方法を示す例については、「[Excel JavaScript API を使用して表を操作する](excel-add-ins-tables.md)」を参照してください。
> - セルの特性に基づいて複数の副範囲を幅広く検索する方法の例については、「[Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)」を参照してください。

## <a name="filter-data"></a>データをフィルター処理する

[AutoFilter](/javascript/api/excel/excel.autofilter) はワークシート内の範囲にわたってデータ フィルターを適用します。 これは、次の `Worksheet.autoFilter.apply` パラメーターを持つ、で作成されます。

- `range`: フィルターが適用される範囲を、`Range` オブジェクトまたは文字列の範囲として指定します。
- `columnIndex`: フィルター条件が評価される 0 から始まる列インデックス。
- `criteria`: 列のセルに基づいてどの行をフィルター処理するかを決定する [FilterCriteria](/javascript/api/excel/excel.filtercriteria) オブジェクト。

最初のコード サンプルは、ワークシートの使用範囲にフィルターを追加する方法を示しています。 このフィルタは、列 **3** の値に基づいて、上位 25% にないエントリを非表示にします。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var farmData = sheet.getUsedRange();

    // This filter will only show the rows with the top 25% of values in column 3.
    sheet.autoFilter.apply(farmData, 3, { criterion1: "25", filterOn: Excel.FilterOn.topPercent });
    return context.sync();
}).catch(errorHandlerFunction);
```

次のコード サンプルでは、`reapply` メソッドを使用してオート フィルターを更新する方法を示します。 これは、範囲内のデータが変更されたときに実行する必要があります。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.autoFilter.reapply();
    return context.sync();
}).catch(errorHandlerFunction);
```

最後のオート フィルター コード サンプルでは、`remove` メソッドを使用してワークシートからオート フィルターを削除する方法を示します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.autoFilter.remove();
    return context.sync();
}).catch(errorHandlerFunction);
```

`AutoFilter` を個々のテーブルに適用することもできます。 詳しくは、[Excel JavaScript API を使用して表を操作する](excel-add-ins-tables.md#autofilter)を参照してください。

## <a name="data-protection"></a>データの保護

ご使用のアドインでは、ワークシート内のデータを編集するユーザー機能を制御できます。 ワークシートの `protection` プロパティは [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) オブジェクトであり、`protect()` メソッドを備えています。 次の例では、アクティブなワークシートの完全な保護を切り替える基本的なシナリオを示します。

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

`protect` メソッドには、2 つの省略可能なパラメーターがあります。

- `options`: 特定の編集制限を定義する [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions) オブジェクト。
- `password`: ユーザーが保護をバイパスしてワークシートを編集するために必要なパスワードを表す文字列。

ワークシートの保護と、Excel の UI を使用してそれを変更する方法の詳細については、記事「[ワークシートを保護する](https://support.microsoft.com/office/3179efdb-1285-4d49-a9c3-f4ca36276de6)」を参照してください。

## <a name="page-layout-and-print-settings"></a>ページ レイアウトと印刷の設定

アドインは、ワークシート レベルでページ レイアウトの設定にアクセスできます。 シートの印刷方法は、これらの設定により制御されます。 `Worksheet` オブジェクトには、レイアウト関連のプロパティが 3 つ含まれます: `horizontalPageBreaks`、`verticalPageBreaks`、`pageLayout`。

`Worksheet.horizontalPageBreaks` と `Worksheet.verticalPageBreaks` は [PageBreakCollections](/javascript/api/excel/excel.pagebreakcollection) です。 これらは [PageBreak](/javascript/api/excel/excel.pagebreak) のコレクションで、手動改ページを挿入する範囲を指定します。 次のコード サンプルは、水平の改ページを行 **21** の上に追加します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.horizontalPageBreaks.add("A21:E21"); // The page break is added above this range.
    return context.sync();
}).catch(errorHandlerFunction);
```

`Worksheet.pageLayout` は、[PageLayout](/javascript/api/excel/excel.pagelayout) オブジェクトです。 このオブジェクトには、プリンター固有の実装に依存しないレイアウト設定とプリント設定が含まれています。 これらの設定には、余白、印刷の向き、ページ番号、タイトル行、および印刷範囲が含まれます。

次のコード サンプルは、ページを縦方向と横方向ともに中央に配置、すべてのページの上部に印刷するタイトル行を設定し、ワークシートのサブセクションで印刷範囲を設定します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // Center the page in both directions.
    sheet.pageLayout.centerHorizontally = true;
    sheet.pageLayout.centerVertically = true;

    // Set the first row as the title row for every page.
    sheet.pageLayout.setPrintTitleRows("$1:$1");

    // Limit the area to be printed to the range "A1:D100".
    sheet.pageLayout.setPrintArea("A1:D100");

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
