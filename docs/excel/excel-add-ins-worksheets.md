---
title: Excel JavaScript API を使用してワークシートを操作する
description: JavaScript API を使用してワークシートで一般的なタスクを実行する方法を示Excelサンプル。
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 565a186220fb9b9a33d97ad73954fe405658cf97
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743394"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してワークシートを操作する

この記事では、Excel JavaScript API を使用して、ワークシートでタスクを実行する方法のコード サンプルを示しています。 `Worksheet` オブジェクトおよび `WorksheetCollection` オブジェクトがサポートするプロパティとメソッドの完全なリストについては、「[Worksheet オブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.worksheet)」および「[WorksheetCollection オブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.worksheetcollection)」を参照してください。

> [!NOTE]
> この記事の情報は標準のワークシートにのみ適用されます。"グラフ" シートや "マクロ" シートには適用されません。

## <a name="get-worksheets"></a>ワークシートを取得する

次のコード サンプルでは、ワークシートのコレクションを取得し、各ワークシートの `name` プロパティを読み込み、コンソールにメッセージを書き込みます。

```js
await Excel.run(async (context) => {
    let sheets = context.workbook.worksheets;
    sheets.load("items/name");

    await context.sync();
    
    if (sheets.items.length > 1) {
        console.log(`There are ${sheets.items.length} worksheets in the workbook:`);
    } else {
        console.log(`There is one worksheet in the workbook:`);
    }

    sheets.items.forEach(function (sheet) {
        console.log(sheet.name);
    });
});
```

> [!NOTE]
> ワークシートの `id` プロパティは、指定されたブックのワークシートを一意に識別します。その値は、ワークシートの名前変更や移動をしても同じままです。Mac 版の Excel のブックからワークシートを削除すると、削除されたワークシートの `id` はそれ以降に作成される新規ワークシートに再割り当てされる可能性があります。

## <a name="get-the-active-worksheet"></a>作業中のワークシートを取得する

次のコード サンプルでは、作業中のワークシートを取得し、`name` プロパティを読み込み、コンソールにメッセージを書き込みます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");

    await context.sync();
    console.log(`The active worksheet is "${sheet.name}"`);
});
```

## <a name="set-the-active-worksheet"></a>作業中のワークシートを設定する

次のコード サンプルでは、作業中のワークシートを **Sample** という名前のワークシートに設定し、`name` プロパティを読み込み、コンソールにメッセージを書き込みます。 その名前を持つワークシートが存在しない場合、`activate()` メソッドにより `ItemNotFound` エラーがスローされます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    sheet.activate();
    sheet.load("name");

    await context.sync();
    console.log(`The active worksheet is "${sheet.name}"`);
});
```

## <a name="reference-worksheets-by-relative-position"></a>相対位置でワークシートを参照する

以下の例は、相対位置でワークシートを参照する方法を示しています。

### <a name="get-the-first-worksheet"></a>最初のワークシートを取得する

次のコード サンプルでは、ブックの最初のワークシートを取得し、`name` プロパティを読み込み、コンソールにメッセージを書き込みます。

```js
await Excel.run(async (context) => {
    let firstSheet = context.workbook.worksheets.getFirst();
    firstSheet.load("name");

    await context.sync();
    console.log(`The name of the first worksheet is "${firstSheet.name}"`);
});
```

### <a name="get-the-last-worksheet"></a>最後のワークシートを取得する

次のコード サンプルでは、ブックの最後のワークシートを取得し、`name` プロパティを読み込み、コンソールにメッセージを書き込みます。

```js
await Excel.run(async (context) => {
    let lastSheet = context.workbook.worksheets.getLast();
    lastSheet.load("name");

    await context.sync();
    console.log(`The name of the last worksheet is "${lastSheet.name}"`);
});
```

### <a name="get-the-next-worksheet"></a>次のワークシートを取得する

次のコード サンプルでは、ブックで作業中のワークシートの後のワークシートを取得し、`name` プロパティを読み込み、コンソールにメッセージを書き込みます。 作業中のワークシートの後にワークシートがない場合、`getNext()` メソッドにより `ItemNotFound` エラーがスローされます。

```js
await Excel.run(async (context) => {
    let currentSheet = context.workbook.worksheets.getActiveWorksheet();
    let nextSheet = currentSheet.getNext();
    nextSheet.load("name");

    await context.sync();
    console.log(`The name of the sheet that follows the active worksheet is "${nextSheet.name}"`);
});
```

### <a name="get-the-previous-worksheet"></a>前のワークシートを取得する

次のコード サンプルでは、ブックで作業中のワークシートの前のワークシートを取得し、`name` プロパティを読み込み、コンソールにメッセージを書き込みます。 作業中のワークシートの前にワークシートが存在しない場合、`getPrevious()` メソッドにより `ItemNotFound` エラーがスローされます。

```js
await Excel.run(async (context) => {
    let currentSheet = context.workbook.worksheets.getActiveWorksheet();
    let previousSheet = currentSheet.getPrevious();
    previousSheet.load("name");

    await context.sync();
    console.log(`The name of the sheet that precedes the active worksheet is "${previousSheet.name}"`);
});
```

## <a name="add-a-worksheet"></a>ワークシートを追加する

次のコード サンプルでは、**Sample** という名前の新しいワークシートをブックに追加し、`name` プロパティと `position` プロパティを読み込み、コンソールにメッセージを書き込みます。新しいワークシートは既存の全ワークシートの後に追加されます。

```js
await Excel.run(async (context) => {
    let sheets = context.workbook.worksheets;

    let sheet = sheets.add("Sample");
    sheet.load("name, position");

    await context.sync();
    console.log(`Added worksheet named "${sheet.name}" in position ${sheet.position}`);
});
```

### <a name="copy-an-existing-worksheet"></a>既存のワークシートをコピーする

`Worksheet.copy` は、既存のワークシートのコピーである新しいワークシートを追加します。 新しいワークシートの名前には、Excel UI を介してワークシートをコピーするのと一貫した方法で、末尾に番号が追加されます (たとえば、**MySheet (2)**)。 `Worksheet.copy` は 2 つのパラメーターを取ることができますが、どちらもオプションです。

- `positionType` - ブック内の新しいワークシートを追加する場所を指定する [WorksheetPositionType](/javascript/api/excel/excel.worksheetpositiontype) 列挙。
- `relativeTo` - `positionType` が `Before` または `After` である場合、新しいシートを追加するワークシートを指定する必要があります (このパラメーターは、「何の前か後に?」という質問に答えます)。

次のコード サンプルは、現在のワークシートをコピーし、現在のワークシートの直後に新しいシートを挿入します。

```js
await Excel.run(async (context) => {
    let myWorkbook = context.workbook;
    let sampleSheet = myWorkbook.worksheets.getActiveWorksheet();
    let copiedSheet = sampleSheet.copy(Excel.WorksheetPositionType.after, sampleSheet);
    await context.sync();
});
```

## <a name="delete-a-worksheet"></a>ワークシートの削除

次のコード サンプルでは、ブックの最後のワークシートを (ただし、ブック内の唯一のシートでない場合に) 削除し、コンソールにメッセージを書き込みます。

```js
await Excel.run(async (context) => {
    let sheets = context.workbook.worksheets;
    sheets.load("items/name");

    await context.sync();
    if (sheets.items.length === 1) {
        console.log("Unable to delete the only worksheet in the workbook");
    } else {
        let lastSheet = sheets.items[sheets.items.length - 1];

        console.log(`Deleting worksheet named "${lastSheet.name}"`);
        lastSheet.delete();

        await context.sync();
    }
});
```

> [!NOTE]
> 可視性が [Very Hidden](/javascript/api/excel/excel.sheetvisibility) のワークシートは、`delete` メソッドで削除することはできません。 このワークシートを削除する場合には、最初に可視性を変更する必要があります。

## <a name="rename-a-worksheet"></a>ワークシートの名前を変更する

次のコード サンプルでは、作業中のワークシートの名前を **New Name** に変更します。

```js
await Excel.run(async (context) => {
    let currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    await context.sync();
});
```

## <a name="move-a-worksheet"></a>ワークシートを移動する

次のコード サンプルでは、ブックの最後の位置からブックの最初の位置にワークシートを移動します。

```js
await Excel.run(async (context) => {
    let sheets = context.workbook.worksheets;
    sheets.load("items");
    await context.sync();

    let lastSheet = sheets.items[sheets.items.length - 1];
    lastSheet.position = 0;
    await context.sync();
});
```

## <a name="set-worksheet-visibility"></a>ワークシートの可視性を設定する

これらの例では、ワークシートの可視性を設定する方法を示します。

### <a name="hide-a-worksheet"></a>ワークシートを非表示にする

次のコード サンプルでは、**Sample** という名前のワークシートの可視性を非表示に設定し、`name` プロパティを読み込み、コンソールにメッセージを書き込みます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    sheet.visibility = Excel.SheetVisibility.hidden;
    sheet.load("name");

    await context.sync();
    console.log(`Worksheet with name "${sheet.name}" is hidden`);
});
```

### <a name="unhide-a-worksheet"></a>ワークシートを再表示する

次のコード サンプルでは、**Sample** という名前のワークシートの可視性を表示に設定し、`name` プロパティを読み込み、コンソールにメッセージを書き込みます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    sheet.visibility = Excel.SheetVisibility.visible;
    sheet.load("name");

    await context.sync();
    console.log(`Worksheet with name "${sheet.name}" is visible`);
});
```

## <a name="get-a-single-cell-within-a-worksheet"></a>ワークシート内で単一のセルを取得する

次のコード サンプルでは、**Sample** という名前のワークシートの 2 行目、5 列目にあるセルを取得し、`address` プロパティと `values` プロパティを読み込み、コンソールにメッセージを書き込みます。 `getCell(row: number, column:number)` メソッドに渡される値は、取得するセルの 0 から始まる行番号および列番号です。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let cell = sheet.getCell(1, 4);
    cell.load("address, values");

    await context.sync();
    console.log(`The value of the cell in row 2, column 5 is "${cell.values[0][0]}" and the address of that cell is "${cell.address}"`);
});
```

## <a name="detect-data-changes"></a>データ変更を検出します

表のデータをユーザーが変更した場合に、アドインを使用して対応する必要がある場合があります。 これらの変更を検出するために、`onChanged`ワークシートのイベントに対する[イベントハンドラを登録できます](excel-add-ins-events.md#register-an-event-handler)。 `onChanged`イベントのイベント ハンドラーは、そのイベントが発生した際に [TableChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs) オブジェクトを受け取ります。

この`WorksheetChangedEventArgs`オブジェクトは、変更とソースに関する情報を提供します。 `onChanged` が発生するのは書式設定またはデータの値が変更された時であるため、値が実際に変更されたかどうかを確認するのにアドインを使用すると便利です。 `details`プロパティは、この情報を [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail) としてカプセル化します。 次のコード サンプルでは、変更前と変更後の値および変更されたセルの種類を表示する方法を表示します。

```js
// This function would be used as an event handler for the Worksheet.onChanged event.
function onWorksheetChanged(eventArgs) {
    Excel.run(function (context) {
        let details = eventArgs.details;
        let address = eventArgs.address;

        // Print the before and after types and values to the console.
        console.log(`Change at ${address}: was ${details.valueBefore}(${details.valueTypeBefore}),`
            + ` now is ${details.valueAfter}(${details.valueTypeAfter})`);
        return context.sync();
    });
}
```

## <a name="detect-formula-changes"></a>数式の変更を検出する

アドインは、ワークシート内の数式の変更を追跡できます。 これは、ワークシートが外部データベースに接続されている場合に便利です。 ワークシート内の数式が変更されると、このシナリオのイベントによって外部データベースの対応する更新プログラムがトリガーされます。

数式の変更を検出するには、 [ワークシート](excel-add-ins-events.md#register-an-event-handler) の [onFormulaChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member) イベントのイベント ハンドラーを登録します。 イベントのイベント ハンドラーは、 `onFormulaChanged` イベントが発生すると [WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs) オブジェクトを受け取る。

> [!IMPORTANT]
> この `onFormulaChanged` イベントは、数式自体が変更された場合を検出します。数式の計算に起因するデータ値は検出しません。

次のコード サンプルは、イベント ハンドラーを登録し、オブジェクトを使用して変更された数式の [formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-formuladetails-member) 配列を取得し、[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail) プロパティを使用して変更された数式の詳細を出力する方法を示しています。`onFormulaChanged` `WorksheetFormulaChangedEventArgs`

> [!NOTE]
> このコード サンプルは、1 つの数式が変更された場合にのみ機能します。

```js
async function run() {
    await Excel.run(async (context) => {
        // Retrieve the worksheet named "Sample".
        let sheet = context.workbook.worksheets.getItem("Sample");
    
        // Register the formula changed event handler for this worksheet.
        sheet.onFormulaChanged.add(formulaChangeHandler);
    
        await context.sync();
    });
}

async function formulaChangeHandler(event) {
    await Excel.run(async (context) => {
        // Retrieve details about the formula change event.
        // Note: This method assumes only a single formula is changed at a time. 
        let cellAddress = event.formulaDetails[0].cellAddress;
        let previousFormula = event.formulaDetails[0].previousFormula;
        let source = event.source;
    
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

![並べ替える前Excelテーブル のデータ。](../images/excel-sort-event-before.png)

"**Q1**" ("**B**" の値) で上から下への並べ替えを実行すると、次の強調表示された行がによって返されます `WorksheetRowSortedEventArgs.address`。

![上から下への並べ替えの後の Excel のテーブル データ。 移動した行が強調表示されます。](../images/excel-sort-event-after-row.png)

元のデータの "**Quinces**" (**"4**" の値) に対して左から右への並べ替えを実行すると、次の強調表示された列が返されます `WorksheetColumnsSortedEventArgs.address`。

![左から右への並べ替えの後の Excel のテーブル データ。 移動した列が強調表示されます。](../images/excel-sort-event-after-column.png)

次のコード サンプルは、`Worksheet.onRowSorted` イベントのイベント ハンドラーを登録する方法を示しています。 ハンドラーのコールバックは、範囲の塗りつぶしの色をクリアし、移動した行のセルを塗りつぶします。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    // This will fire whenever a row has been moved as the result of a sort action.
    sheet.onRowSorted.add(async (event) => {
        await Excel.run(async (context) => {
            console.log("Row sorted: " + event.address);
            let sheet = context.workbook.worksheets.getActiveWorksheet();

            // Clear formatting for section, then highlight the sorted area.
            sheet.getRange("A1:E5").format.fill.clear();
            if (event.address !== "") {
                sheet.getRanges(event.address).format.fill.color = "yellow";
            }

            await context.sync();
        });
    });

    await context.sync();
});
```

## <a name="find-all-cells-with-matching-text"></a>一致するテキストがあるすべてのセルを検索する

`Worksheet` オブジェクトには、ワークシート内の指定された文字列を検索するための `find` メソッドがあります。 このメソッドは `RangeAreas` オブジェクトを返します。これは、一度に編集できる `Range` オブジェクトのコレクションとなります。 以下のコード サンプルは、文字列 **Complete** と等しいすべてのセルを検索し、そのセルの色を緑色にします。 指定した文字列がワークシートに存在しない場合、`ItemNotFound` エラーが `findAll` によってスローされます。 指定した文字列がワークシートに存在しない可能性がある場合は、自分のコードで適切にシナリオを処理できるように、[findAllOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) メソッドを使用するようにしてください。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let foundRanges = sheet.findAll("Complete", {
        completeMatch: true, // findAll will match the whole cell value
        matchCase: false // findAll will not match case
    });

    await context.sync();
    foundRanges.format.fill.color = "green"
});
```

> [!NOTE]
> このセクションでは、`Worksheet` オブジェクトの関数を使用してセルと範囲を検索する方法について説明します。 範囲の取得の詳細については、オブジェクト専用の記事で確認することができます。
>
> - オブジェクトを使用してワークシート内`Range`の範囲を取得する方法を示す例については、「[JavaScript API](excel-add-ins-ranges-get.md) を使用して範囲を取得するExcel参照してください。
> - `Table` オブジェクトから範囲を取得する方法を示す例については、「[Excel JavaScript API を使用して表を操作する](excel-add-ins-tables.md)」を参照してください。
> - セルの特性に基づいて複数の副範囲を幅広く検索する方法の例については、「[Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)」を参照してください。

## <a name="filter-data"></a>データをフィルター処理する

[AutoFilter](/javascript/api/excel/excel.autofilter) はワークシート内の範囲にわたってデータ フィルターを適用します。 これは、次のパラメーター `Worksheet.autoFilter.apply`を持つ、で作成されます。

- `range`: フィルターが適用される範囲を、`Range` オブジェクトまたは文字列の範囲として指定します。
- `columnIndex`: フィルター条件が評価される 0 から始まる列インデックス。
- `criteria`: 列のセルに基づいてどの行をフィルター処理するかを決定する [FilterCriteria](/javascript/api/excel/excel.filtercriteria) オブジェクト。

最初のコード サンプルは、ワークシートの使用範囲にフィルターを追加する方法を示しています。 このフィルタは、列 **3** の値に基づいて、上位 25% にないエントリを非表示にします。

```js
// This method adds a custom AutoFilter to the active worksheet
// and applies the filter to a column of the used range.
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let farmData = sheet.getUsedRange();

    // This filter will only show the rows with the top 25% of values in column 3.
    sheet.autoFilter.apply(farmData, 3, { criterion1: "25", filterOn: Excel.FilterOn.topPercent });
    await context.sync();
});
```

次のコード サンプルでは、`reapply` メソッドを使用してオート フィルターを更新する方法を示します。 これは、範囲内のデータが変更されたときに実行する必要があります。

```js
// This method refreshes the AutoFilter to ensure that changes are captured.
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.autoFilter.reapply();
    await context.sync();
});
```

次のコード サンプルは、 `clearColumnCriteria` メソッドを使用して、他の列でアクティブなフィルターを残しながら、1 つの列から自動フィルターをクリアする方法を示しています。

```js
// This method clears the AutoFilter setting from one column.
await Excel.run(async (context) => {
    // Retrieve the active worksheet.
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    // Clear the filter from only column 3.
    sheet.autoFilter.clearColumnCriteria(3);
    await context.sync();
});
```

最後のオート フィルター コード サンプルでは、`remove` メソッドを使用してワークシートからオート フィルターを削除する方法を示します。

```js
// This method removes all AutoFilters from the active worksheet.
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.autoFilter.remove();
    await context.sync();
});
```

`AutoFilter` を個々のテーブルに適用することもできます。 詳しくは、[Excel JavaScript API を使用して表を操作する](excel-add-ins-tables.md#autofilter)を参照してください。

## <a name="data-protection"></a>データの保護

ご使用のアドインでは、ワークシート内のデータを編集するユーザー機能を制御できます。 ワークシートの `protection` プロパティは [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) オブジェクトであり、`protect()` メソッドを備えています。 次の例では、アクティブなワークシートの完全な保護を切り替える基本的なシナリオを示します。

```js
await Excel.run(async (context) => {
    let activeSheet = context.workbook.worksheets.getActiveWorksheet();
    activeSheet.load("protection/protected");
    await context.sync();

    if (!activeSheet.protection.protected) {
        activeSheet.protection.protect();
    }
});
```

`protect` メソッドには、2 つの省略可能なパラメーターがあります。

- `options`: 特定の編集制限を定義する [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions) オブジェクト。
- `password`: ユーザーが保護をバイパスしてワークシートを編集するために必要なパスワードを表す文字列。

ワークシートの保護と、Excel の UI を使用してそれを変更する方法の詳細については、記事「[ワークシートを保護する](https://support.microsoft.com/office/3179efdb-1285-4d49-a9c3-f4ca36276de6)」を参照してください。

### <a name="detect-changes-to-the-worksheet-protection-state"></a>ワークシートの保護状態の変更を検出する

ワークシートの保護状態は、アドインまたは UI を使用してExcelできます。 保護状態の変更を検出するには [、ワークシートのイベント](excel-add-ins-events.md#register-an-event-handler) の [`onProtectionChanged`](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onprotectionchanged-member) イベント ハンドラーを登録します。 イベントのイベント ハンドラーは、 `onProtectionChanged` イベントが発生すると [`WorksheetProtectionChangedEventArgs`](/javascript/api/excel/excel.worksheetprotectionchangedeventargs) オブジェクトを受け取る。

次のコード サンプルは、イベント ハンドラーを登録`onProtectionChanged` `WorksheetProtectionChangedEventArgs` `isProtected`し、オブジェクトを使用してイベントの 、 、`worksheetId``source`およびプロパティを取得する方法を示しています。

```js
// This method registers an event handler for the onProtectionChanged event of a worksheet.
async function run() {
    await Excel.run(async (context) => {
        // Retrieve the worksheet named "Sample".
        let sheet = context.workbook.worksheets.getItem("Sample");
    
        // Register the onProtectionChanged event handler.
        sheet.onProtectionChanged.add(checkProtection);
        await context.sync();
    });
}

// This method is an event handler that returns the protection state of a worksheet 
// and information about the changed worksheet.
async function checkProtection(event) {
    await Excel.run(async (context) => {
        // Retrieve the protection, worksheet ID, and source properties of the event.
        let protectionStatus = event.isProtected;
        let worksheetId = event.worksheetId;
        let source = event.source;

        // Print the event properties to the console.
        console.log("Protection status changed. Protection status is now: " + protectionStatus);
        console.log("    ID of changed worksheet: " + worksheetId);
        console.log("    Source of change event: " + source);    
    });
}
```

## <a name="page-layout-and-print-settings"></a>ページ レイアウトと印刷の設定

アドインは、ワークシート レベルでページ レイアウトの設定にアクセスできます。 シートの印刷方法は、これらの設定により制御されます。 `Worksheet` オブジェクトには、レイアウト関連のプロパティが 3 つ含まれます: `horizontalPageBreaks`、`verticalPageBreaks`、`pageLayout`。

`Worksheet.horizontalPageBreaks` と `Worksheet.verticalPageBreaks` は [PageBreakCollections](/javascript/api/excel/excel.pagebreakcollection) です。 これらは [PageBreak](/javascript/api/excel/excel.pagebreak) のコレクションで、手動改ページを挿入する範囲を指定します。 次のコード サンプルは、水平の改ページを行 **21** の上に追加します。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.horizontalPageBreaks.add("A21:E21"); // The page break is added above this range.
    await context.sync();
});
```

`Worksheet.pageLayout` は、[PageLayout](/javascript/api/excel/excel.pagelayout) オブジェクトです。 このオブジェクトには、プリンター固有の実装に依存しないレイアウト設定とプリント設定が含まれています。 これらの設定には、余白、印刷の向き、ページ番号、タイトル行、および印刷範囲が含まれます。

次のコード サンプルは、ページを縦方向と横方向ともに中央に配置、すべてのページの上部に印刷するタイトル行を設定し、ワークシートのサブセクションで印刷範囲を設定します。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    // Center the page in both directions.
    sheet.pageLayout.centerHorizontally = true;
    sheet.pageLayout.centerVertically = true;

    // Set the first row as the title row for every page.
    sheet.pageLayout.setPrintTitleRows("$1:$1");

    // Limit the area to be printed to the range "A1:D100".
    sheet.pageLayout.setPrintArea("A1:D100");

    await context.sync();
});
```

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
