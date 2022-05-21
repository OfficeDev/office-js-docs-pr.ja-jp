---
title: Excel JavaScript API を使用して表を操作する
description: Excel JavaScript API を使用してテーブルで一般的なタスクを実行する方法を示すコード サンプル。
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: f4cbed134c8ca9f53e89fa97bd4c7ccaa35e45c7
ms.sourcegitcommit: 4ca3334f3cefa34e6b391eb92a429a308229fe89
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2022
ms.locfileid: "65628111"
---
# <a name="work-with-tables-using-the-excel-javascript-api"></a>Excel JavaScript API を使用して表を操作する

この記事では、Excel JavaScript API を使用して、表に関する一般的なタスクを実行する方法を示すサンプル コードを提供します。 オブジェクトがサポートする`Table`プロパティとメソッドの完全な一覧については、「[Table Object (Excel 用 JavaScript API)](/javascript/api/excel/excel.table)」と[「TableCollection オブジェクト (Excel用の JavaScript API)」](/javascript/api/excel/excel.tablecollection)を参照`TableCollection`してください。

## <a name="create-a-table"></a>表を作成する

次のコード サンプルでは、**Sample** というワークシートに表を作成します。 表にはヘッダーがあり、4 つの列と 7 つのデータ行が含まれています。 コードが実行されているExcel アプリケーションで [要件セット](/javascript/api/requirement-sets/excel/excel-api-requirement-sets) **ExcelApi 1.2** がサポートされている場合、列の幅と行の高さは、テーブル内の現在のデータに最も適するように設定されます。

> [!NOTE]
> テーブルの名前を指定するには、次の例に示すように、最初にテーブルを作成してから、その `name` プロパティを設定する必要があります。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add rows to the end of the table*/, [
        ["1/1/2017", "The Phone Company", "Communications", "$120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "$33"],
        ["1/11/2017", "Bellows College", "Education", "$350"],
        ["1/15/2017", "Trey Research", "Other", "$135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"]
    ]);

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    await context.sync();
});
```

### <a name="new-table"></a>新しい表

![Excelの新しいテーブル。](../images/excel-tables-create.png)

## <a name="add-rows-to-a-table"></a>表に行を追加する

次のコード サンプルでは、**Sample** ワークシート内の **ExpensesTable** という表に 7 つの新しい行を追加します。 `index`メソッドの[`add`](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1))パラメーターは `null`、テーブル内の既存の行の後に行を追加することを指定します。 `alwaysInsert`このパラメーターは 、テーブルの下ではなく、テーブルに新しい行を挿入することを示す値に設定`true`されます。 次に、列の幅と行の高さが、テーブル内の現在のデータに最も適するように設定されます。

> [!NOTE]
> [TableRow](/javascript/api/excel/excel.tablerow) オブジェクトのプロパティは`index`、テーブルの行コレクション内の行のインデックス番号を示します。 オブジェクトには `TableRow` 、行を `id` 識別するために一意のキーとして使用できるプロパティが含まれていません。

```js
// This code sample shows how to add rows to a table that already exists 
// on a worksheet named Sample.
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.rows.add(
        null, // index, Adds rows to the end of the table.
        [
            ["1/16/2017", "THE PHONE COMPANY", "Communications", "$120"],
            ["1/20/2017", "NORTHWIND ELECTRIC CARS", "Transportation", "$142"],
            ["1/20/2017", "BEST FOR YOU ORGANICS COMPANY", "Groceries", "$27"],
            ["1/21/2017", "COHO VINEYARD", "Restaurant", "$33"],
            ["1/25/2017", "BELLOWS COLLEGE", "Education", "$350"],
            ["1/28/2017", "TREY RESEARCH", "Other", "$135"],
            ["1/31/2017", "BEST FOR YOU ORGANICS COMPANY", "Groceries", "$97"]
        ], 
        true, // alwaysInsert, Specifies that the new rows be inserted into the table.
    );

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    await context.sync();
});
```

### <a name="table-with-new-rows"></a>新しい行を含む表

![Excelの新しい行を含むテーブル。](../images/excel-tables-add-rows.png)

## <a name="add-a-column-to-a-table"></a>表に列を追加する

以下の例では、表に列を追加する方法を示します。 最初の例では、新しい列に静的な値を作成し、2 番目の例では新しい列に数式を作成します。

> [!NOTE]
> **TableColumn** オブジェクトの [index](/javascript/api/excel/excel.tablecolumn) プロパティは、表の列コレクション内の列のインデックス番号を示しています。 **TableColumn** オブジェクトの **id** プロパティには、列を識別する一意のキーが含まれています。

### <a name="add-a-column-that-contains-static-values"></a>静的な値を含む列を追加する

次のコード サンプルでは、**Sample** ワークシート内の **ExpensesTable** という表に新しい列を追加します。 新しい列は、表内の既存の列すべての後に追加され、ヘッダー (「曜日」) を含み、列内のセルにデータが作成されます。 次に、列の幅と行の高さが、テーブル内の現在のデータに最も適するように設定されます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.columns.add(null /*add columns to the end of the table*/, [
        ["Day of the Week"],
        ["Saturday"],
        ["Friday"],
        ["Monday"],
        ["Thursday"],
        ["Sunday"],
        ["Saturday"],
        ["Monday"]
    ]);

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    await context.sync();
});
```

#### <a name="table-with-new-column"></a>新しい列を含む表

![Excelの新しい列を含むテーブル。](../images/excel-tables-add-column.png)

### <a name="add-a-column-that-contains-formulas"></a>数式を含む列を追加する

次のコード サンプルでは、**Sample** ワークシート内の **ExpensesTable** という表に新しい列を追加します。 新しい列は表の末尾に追加され、ヘッダー (「曜日」) を含み、数式を使用して列内のそれぞれのデータ セルを作成します。 次に、列の幅と行の高さが、テーブル内の現在のデータに最も適するように設定されます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.columns.add(null /*add columns to the end of the table*/, [
        ["Type of the Day"],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")']
    ]);

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    await context.sync();
});
```

#### <a name="table-with-new-calculated-column"></a>新しい集計列を含む表

![Excelの新しい計算列を含むテーブル。](../images/excel-tables-add-calculated-column.png)

## <a name="resize-a-table"></a>テーブルのサイズを変更する

アドインは、テーブルにデータを追加したり、セル値を変更したりすることなく、テーブルのサイズを変更できます。 テーブルのサイズを変更するには、 [Table.resize](/javascript/api/excel/excel.table#excel-excel-table-resize-member(1)) メソッドを使用します。 次のコード サンプルは、テーブルのサイズを変更する方法を示しています。 このコード サンプルでは、この記事の前半の「[テーブルの作成](#create-a-table)」セクションの **ExpensesTable** を使用し、テーブルの新しい範囲を **A1:D20** に設定します。

```js
await Excel.run(async (context) => {
    // Retrieve the worksheet and a table on that worksheet.
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    // Resize the table.
    expensesTable.resize("A1:D20");

    await context.sync();
});
```

> [!IMPORTANT]
> テーブルの新しい範囲は元の範囲と重複している必要があり、ヘッダー (またはテーブルの先頭) は同じ行に含まれている必要があります。

### <a name="table-after-resize"></a>サイズ変更後のテーブル

![Excel内に複数の空の行を含むテーブル。](../images/excel-tables-resize.png)

## <a name="update-column-name"></a>列名を更新する

次のコード サンプルでは、表の最初の列の名前を **Purchase date** に更新します。 次に、列の幅と行の高さが、テーブル内の現在のデータに最も適するように設定されます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let expensesTable = sheet.tables.getItem("ExpensesTable");
    expensesTable.columns.load("items");

    await context.sync();
        
    expensesTable.columns.items[0].name = "Purchase date";

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    await context.sync();
});
```

### <a name="table-with-new-column-name"></a>新しい列名を含む表

![Excelの新しい列名を持つテーブル。](../images/excel-tables-update-column-name.png)

## <a name="get-data-from-a-table"></a>表からデータを取得する

次のコード サンプルでは、**Sample** ワークシートから **ExpensesTable** という表のデータを読み取り、そのデータを同じワークシートの表の下に出力します。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    // Get data from the header row.
    let headerRange = expensesTable.getHeaderRowRange().load("values");

    // Get data from the table.
    let bodyRange = expensesTable.getDataBodyRange().load("values");

    // Get data from a single column.
    let columnRange = expensesTable.columns.getItem("Merchant").getDataBodyRange().load("values");

    // Get data from a single row.
    let rowRange = expensesTable.rows.getItemAt(1).load("values");

    // Sync to populate proxy objects with data from Excel.
    await context.sync();

    let headerValues = headerRange.values;
    let bodyValues = bodyRange.values;
    let merchantColumnValues = columnRange.values;
    let secondRowValues = rowRange.values;

    // Write data from table back to the sheet
    sheet.getRange("A11:A11").values = [["Results"]];
    sheet.getRange("A13:D13").values = headerValues;
    sheet.getRange("A14:D20").values = bodyValues;
    sheet.getRange("B23:B29").values = merchantColumnValues;
    sheet.getRange("A32:D32").values = secondRowValues;

    // Sync to update the sheet in Excel.
    await context.sync();
});
```

### <a name="table-and-data-output"></a>表とデータの出力

![Excelのテーブル データ。](../images/excel-tables-get-data.png)

## <a name="detect-data-changes"></a>データの変更の検出

表のデータをユーザーが変更した場合に、アドインを使用して対応する必要がある場合があります。 そのような変更を検出するには、表の `onChanged` イベントについて[イベント ハンドラーを登録](excel-add-ins-events.md#register-an-event-handler)します。 `onChanged`イベントのイベント ハンドラーは、そのイベントが発生した際に [TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs) オブジェクトを受け取ります。

`TableChangedEventArgs` オブジェクトは、変更内容とソースに関する情報を提供します。 `onChanged` が発生するのは書式設定またはデータの値が変更された時であるため、値が実際に変更されたかどうかを確認するのにアドインを使用すると便利です。 `details`プロパティは、この情報を [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail) としてカプセル化します。 次のコード サンプルでは、変更前と変更後の値および変更されたセルの種類を表示する方法を表示します。

```js
// This function would be used as an event handler for the Table.onChanged event.
async function onTableChanged(eventArgs) {
    await Excel.run(async (context) => {
        let details = eventArgs.details;
        let address = eventArgs.address;

        // Print the before and after types and values to the console.
        console.log(`Change at ${address}: was ${details.valueBefore}(${details.valueTypeBefore}),`
            + ` now is ${details.valueAfter}(${details.valueTypeAfter})`);
        await context.sync();
    });
}
```

## <a name="sort-data-in-a-table"></a>表のデータを並べ替える

次のコード サンプルでは、表の 4 番目の列の値に従って降順で表データを並べ替えます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    // Queue a command to sort data by the fourth column of the table (descending).
    let sortRange = expensesTable.getDataBodyRange();
    sortRange.sort.apply([
        {
            key: 3,
            ascending: false,
        },
    ]);

    // Sync to run the queued command in Excel.
    await context.sync();
});
```

### <a name="table-data-sorted-by-amount-descending"></a>Amount (降順) で並べ替えた表データ

![Excel内の並べ替えられたテーブル データ。](../images/excel-tables-sort.png)

ワークシートでデータを並べ替えると、イベント通知が発生します。 並べ替え関連のイベントと、アドインがイベント ハンドラーを登録してそのようなイベントに応答する方法の詳細については、「[並べ替えイベントを処理する](excel-add-ins-worksheets.md#handle-sorting-events)」を参照してください。

## <a name="apply-filters-to-a-table"></a>表にフィルターを適用する

次のコード サンプルでは、表内の **Amount** 列と **Category** 列にフィルターを適用しています。 フィルター処理の結果、**Category** が指定した値であり、**Amount** が表示されている行の平均値未満の行のみが表示されます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    // Queue a command to apply a filter on the Category column.
    let categoryFilter = expensesTable.columns.getItem("Category").filter;
    categoryFilter.apply({
      filterOn: Excel.FilterOn.values,
      values: ["Restaurant", "Groceries"]
    });

    // Queue a command to apply a filter on the Amount column.
    let amountFilter = expensesTable.columns.getItem("Amount").filter;
    amountFilter.apply({
      filterOn: Excel.FilterOn.dynamic,
      dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage
    });

    // Sync to run the queued commands in Excel.
    await context.sync();
});
```

### <a name="table-data-with-filters-applied-for-category-and-amount"></a>Category と Amount にフィルターを適用した表データ

![Excelでフィルター処理されたテーブル データ。](../images/excel-tables-filters-apply.png)

## <a name="clear-table-filters"></a>表フィルターのクリア

次のコード サンプルでは、表に現在適用されているフィルターをクリアします。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.clearFilters();

    await context.sync();
});
```

### <a name="table-data-with-no-filters-applied"></a>フィルターが適用されていない表データ

![Excelでフィルター処理されていないテーブル データ。](../images/excel-tables-filters-clear.png)

## <a name="get-the-visible-range-from-a-filtered-table"></a>フィルター処理された表から、表示されている範囲を取得します。

次のコード サンプルでは、指定した表内で現在表示されているセルのデータのみを含む範囲を取得し、その範囲の値をコンソールに書き込みます。 次に示すように、このメソッドを `getVisibleView()` 使用すると、列フィルターが適用されるたびにテーブルの目に見える内容を取得できます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    let visibleRange = expensesTable.getDataBodyRange().getVisibleView();
    visibleRange.load("values");

    await context.sync();
    console.log(visibleRange.values);
});
```

## <a name="autofilter"></a>オートフィルター

アドインは、テーブルの[AutoFilter](/javascript/api/excel/excel.autofilter)オブジェクトを使用してデータをフィルタ処理できます。 `AutoFilter` オブジェクトは、テーブルまたは範囲のフィルタ構造全体です。 この記事で前述したすべてのフィルター操作は、自動フィルターと互換性があります。 単一のアクセスポイントにより、複数のフィルタへのアクセスと管理が簡単になります。

次のコードサンプルは、[前のコードサンプルと同じデータフィルタリング](#apply-filters-to-a-table)を示していますが、完全に自動フィルタを介して行われます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.autoFilter.apply(expensesTable.getRange(), 2, {
        filterOn: Excel.FilterOn.values,
        values: ["Restaurant", "Groceries"]
    });
    expensesTable.autoFilter.apply(expensesTable.getRange(), 3, {
        filterOn: Excel.FilterOn.dynamic,
        dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage
    });

    await context.sync();
});
```

`AutoFilter` は、ワークシートレベルでの範囲にも適用できます。 詳しくは、[Excel JavaScript APIを使用したワークシートの処理](excel-add-ins-worksheets.md#filter-data)を参照してください。

## <a name="format-a-table"></a>表を書式設定する

次のコード サンプルでは、表に書式を適用します。 表のヘッダー行、表の本体、表の 2 行目、表の 1 列目にそれぞれ別の塗りつぶし色を指定します。 書式の指定に使用できるプロパティの詳細については、「[RangeFormat オブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.rangeformat)」を参照してください。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.getHeaderRowRange().format.fill.color = "#C70039";
    expensesTable.getDataBodyRange().format.fill.color = "#DAF7A6";
    expensesTable.rows.getItemAt(1).getRange().format.fill.color = "#FFC300";
    expensesTable.columns.getItemAt(0).getDataBodyRange().format.fill.color = "#FFA07A";

    await context.sync();
});
```

### <a name="table-after-formatting-is-applied"></a>書式設定を適用後の表

![Excelで書式設定が適用された後のテーブル。](../images/excel-tables-formatting-after.png)

## <a name="convert-a-range-to-a-table"></a>範囲を表に変換する

次のコード サンプルでは、データ範囲を作成し、その範囲を表に変換します。 次に、列の幅と行の高さが、テーブル内の現在のデータに最も適するように設定されます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    // Define values for the range.
    let values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"],
    ["Frames", 5000, 7000, 6544, 4377],
    ["Saddles", 400, 323, 276, 651],
    ["Brake levers", 12000, 8766, 8456, 9812],
    ["Chains", 1550, 1088, 692, 853],
    ["Mirrors", 225, 600, 923, 544],
    ["Spokes", 6005, 7634, 4589, 8765]];

    // Create the range.
    let range = sheet.getRange("A1:E7");
    range.values = values;

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    sheet.activate();

    // Convert the range to a table.
    let expensesTable = sheet.tables.add('A1:E7', true);
    expensesTable.name = "ExpensesTable";

    await context.sync();
});
```

### <a name="data-in-the-range-before-the-range-is-converted-to-a-table"></a>範囲データ (範囲を表に変換する前)

![Excelの範囲内のデータ。](../images/excel-ranges.png)

### <a name="data-in-the-table-after-the-range-is-converted-to-a-table"></a>表内のデータ (範囲を表に変換した後)

![Excelのテーブル内のデータ。](../images/excel-tables-from-range.png)

## <a name="import-json-data-into-a-table"></a>JSON データを表にインポートする

次のコード サンプルでは、**Sample** ワークシートに表を作成し、2 行のデータを定義する JSON オブジェクトを使用して表にデータを入力します。 次に、列の幅と行の高さが、テーブル内の現在のデータに最も適するように設定されます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    let transactions = [
      {
        "DATE": "1/1/2017",
        "MERCHANT": "The Phone Company",
        "CATEGORY": "Communications",
        "AMOUNT": "$120"
      },
      {
        "DATE": "1/1/2017",
        "MERCHANT": "Southridge Video",
        "CATEGORY": "Entertainment",
        "AMOUNT": "$40"
      }
    ];

    let newData = transactions.map(item =>
        [item.DATE, item.MERCHANT, item.CATEGORY, item.AMOUNT]);

    expensesTable.rows.add(null, newData);

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    sheet.activate();

    await context.sync();
});
```

### <a name="new-table"></a>新しい表

![Excelでインポートされた JSON データからの新しいテーブル。](../images/excel-tables-create-from-json.png)

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
