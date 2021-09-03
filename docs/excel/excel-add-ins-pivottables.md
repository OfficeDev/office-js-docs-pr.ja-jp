---
title: JavaScript API を使用してピボットテーブルをExcelする
description: JavaScript API Excel使用してピボットテーブルを作成し、それらのコンポーネントを操作します。
ms.date: 07/02/2021
localization_priority: Normal
ms.openlocfilehash: d9ccaf72be4fa23b73f1f91d38d240ea02569eca
ms.sourcegitcommit: 69f6492de8a4c91e734250c76681c44b3f349440
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/03/2021
ms.locfileid: "58868639"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>JavaScript API を使用してピボットテーブルをExcelする

PivotTables は、より大きなデータ セットを合理化します。 グループ化されたデータを簡単に操作できます。 JavaScript API Excelにより、アドインはピボットテーブルを作成し、それらのコンポーネントを操作できます。 この記事では、ピボットテーブルが JavaScript API の Officeされる方法について説明し、主要なシナリオのコード サンプルを提供します。

ピボットテーブルの機能に慣れていない場合は、エンド ユーザーとして探索を検討してください。
これらの [ツールの優れた入門については、「ピボットテーブル](https://support.microsoft.com/office/ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EBBD=PivotTables) を作成してワークシート データを分析する」を参照してください。

> [!IMPORTANT]
> OLAP で作成されたピボットテーブルは現在サポートされていません。 Power Pivot もサポートされていません。

## <a name="object-model"></a>オブジェクト モデル

ピボット[テーブルは](/javascript/api/excel/excel.pivottable)、JavaScript API のピボットテーブルOfficeです。

- `Workbook.pivotTables`は、それぞれブックとワークシートにピボットテーブルを含む `Worksheet.pivotTables` [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection)です。 [](/javascript/api/excel/excel.pivottable)
- ピボット [テーブルには、](/javascript/api/excel/excel.pivottable) 複数の [PivotHierarchies を持つ PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) [が含まれる](/javascript/api/excel/excel.pivothierarchy)。
- これらの [PivotHierarchies を](/javascript/api/excel/excel.pivothierarchy) 特定の階層コレクションに追加して、ピボットテーブルがデータをピボットする方法を定義できます (次のセクション [で説明します](#hierarchies))。
- [PivotHierarchy には](/javascript/api/excel/excel.pivothierarchy)、ピボットフィールドが 1 つ正確に含まれる[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)が[含まれる](/javascript/api/excel/excel.pivotfield)。 OLAP ピボットテーブルを含むデザインが展開された場合、変更される可能性があります。
- ピボット[フィールドには、](/javascript/api/excel/excel.pivotfield)フィールドの[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)が階層カテゴリに割り当てられている限り、1 つ以上のピボットフィルターを適用できます。 [](/javascript/api/excel/excel.pivotfilters)
- PivotField [には、](/javascript/api/excel/excel.pivotfield) 複数の PivotItem を持つ [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) [が含まれる](/javascript/api/excel/excel.pivotitem)。
- ピボット[テーブルには、](/javascript/api/excel/excel.pivottable)[ピボットフィールドと](/javascript/api/excel/excel.pivotlayout)ピボットアイテムがワークシートに表示される[](/javascript/api/excel/excel.pivotfield)場所を定義する[ピボット](/javascript/api/excel/excel.pivotitem)レイアウトが含まれる。 レイアウトでは、ピボットテーブルの一部の表示設定も制御します。

これらのリレーションシップがデータの例に適用される方法について説明します。 次のデータは、さまざまなファームからの果物の販売について説明します。 この記事全体の例を示します。

![異なるファームから異なる種類の果物の販売のコレクション。](../images/excel-pivots-raw-data.png)

このフルーツ ファームの販売データは、ピボットテーブルの作成に使用されます。 Types などの各列 **は**、 です `PivotHierarchy` 。 [ **型]** 階層には、[種類] **フィールドが含** まれます。 [**種類]** フィールドには、Apple、Kiwi、Lemon、Lime、**および Orange** の項目が **含まれます**。  

### <a name="hierarchies"></a>Hierarchies

ピボットテーブルは、行、列、データ、およびフィルター[](/javascript/api/excel/excel.rowcolumnpivothierarchy)の 4 つの[階層カテゴリに](/javascript/api/excel/excel.rowcolumnpivothierarchy)[基づいて編成](/javascript/api/excel/excel.datapivothierarchy)[されます](/javascript/api/excel/excel.filterpivothierarchy)。

前に示したファーム データには **、Farms**、 **Type**、 **Classification**、 **Crates Sold** at Farm 、 および Crates Sold Wholesale の 5 **つの階層があります**。 各階層は、4 つのカテゴリの 1 つにのみ存在できます。 列 **階層** に Type を追加した場合は、行、データ、またはフィルター階層に追加することはできません。 Type **が** 後で行階層に追加されると、列階層から削除されます。 この動作は、階層の割り当てが UI または JavaScript API のExcel同Excelです。

行階層と列階層は、データのグループ化方法を定義します。 たとえば **、Farms** の行階層は、同じファームのすべてのデータ セットをグループ化します。 行階層と列階層の選択によって、ピボットテーブルの向きが定義されます。

データ階層は、行階層と列階層に基づいて集計される値です。 ファームの行階層とクレート販売済みホールセールのデータ階層を持つピボットテーブルには、ファームごとに異なるすべての果物の合計 (既定) が表示されます。

フィルター階層には、フィルター処理された型内の値に基づいてピボットからデータが含まれるか除外されます。 [分類] の **フィルター階層で** [オーガニック] **が選択** されている場合は、オーガニック フルーツのデータだけが表示されます。

ピボットテーブルと共に、もう一度ファーム データを次に示します。 ピボットテーブルは、行階層として Farm と **Type** を使用し、データ階層として [ファームで販売されたクレート] と [販売済みクレートの販売済みホールセール] をデータ階層 (合計の既定の集計関数を使用)、および分類をフィルター階層として使用します ([オーガニック] が選択されている場合)。 

![行、データ、およびフィルター階層を持つピボットテーブルの横にある果物の販売データの選択。](../images/excel-pivot-table-and-data.png)

このピボットテーブルは、JavaScript API または UI を使用してExcelできます。 どちらのオプションでも、アドインを介してさらに操作できます。

## <a name="create-a-pivottable"></a>ピボットテーブルの作成

ピボットテーブルには、名前、ソース、および宛先が必要です。 ソースには、範囲アドレスまたはテーブル名 (、、または型として渡される) `Range` `string` `Table` を指定できます。 宛先は範囲アドレス (a または ) のいずれかとして `Range` 指定 `string` されます。
次のサンプルは、さまざまなピボットテーブル作成手法を示しています。

### <a name="create-a-pivottable-with-range-addresses"></a>範囲アドレスを使用してピボットテーブルを作成する

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    return context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a>Range オブジェクトを使用してピボットテーブルを作成する

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data comes from the worksheet "DataWorksheet" across the range A1:E21.
    var rangeToAnalyze = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
    var rangeToPlacePivot = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
    context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add(
      "Farm Sales", rangeToAnalyze, rangeToPlacePivot);

    return context.sync();
});
```

### <a name="create-a-pivottable-at-the-workbook-level"></a>ブック レベルでピボットテーブルを作成する

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    return context.sync();
});
```

## <a name="use-an-existing-pivottable"></a>既存のピボットテーブルを使用する

手動で作成されたピボットテーブルには、ブックの PivotTable コレクションまたは個々のワークシートからアクセスすることもできます。 次のコードは、ブックから My Pivot という **名前のピボットテーブル** を取得します。

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>ピボットテーブルに行と列を追加する

行と列は、これらのフィールドの値を中心にデータをピボットします。

[ファーム] **列を** 追加すると、各ファームの周りのすべての売上がピボットされます。 Type 行 **と Classification** **行を追加** すると、販売されたフルーツと、それがオーガニックかどうかに基づいてデータがさらに分類されます。

![[ファーム] 列と [種類] 行と [分類] 行を含むピボットテーブル。](../images/excel-pivots-table-rows-and-columns.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    return context.sync();
});
```

行または列のみを含むピボットテーブルを使用することもできます。

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a>ピボットテーブルにデータ階層を追加する

データ階層は、ピボットテーブルに行と列に基づいて結合する情報を入力します。 ファームで販売されたクレートと **ク** レートの販売済みホールセールのデータ階層を追加すると、各行と列に対してそれらの数値の合計が表示されます。

この例では **、Farm と** **Type の両方** が行であり、クレート売上をデータとして使用します。

![彼らが来たファームに基づいて異なる果物の総売上を示すピボットテーブル。](../images/excel-pivots-data-hierarchy.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // "Farm" and "Type" are the hierarchies on which the aggregation is based.
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

    // "Crates Sold at Farm" and "Crates Sold Wholesale" are the hierarchies
    // that will have their data aggregated (summed in this case).
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    return context.sync();
});
```

## <a name="pivottable-layouts-and-getting-pivoted-data"></a>ピボットテーブルレイアウトとピボットデータの取得

[PivotLayout は](/javascript/api/excel/excel.pivotlayout)、階層とそのデータの配置を定義します。 レイアウトにアクセスして、データが格納される範囲を決定します。

次の図は、ピボットテーブルの範囲に対応するレイアウト関数呼び出しを示しています。

![レイアウトの取得範囲関数によって返されるピボットテーブルのセクションを示す図。](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a>ピボットテーブルからデータを取得する

レイアウトは、ワークシートでのピボットテーブルの表示方法を定義します。 つまり、オブジェクト `PivotLayout` はピボットテーブル要素に使用される範囲を制御します。 ピボットテーブルによって収集および集計されたデータを取得するには、レイアウトによって提供される範囲を使用します。 特に、ピボット `PivotLayout.getDataBodyRange` テーブルによって生成されたデータにアクセスするために使用します。

次のコードは、レイアウトを実行してピボットテーブル データの最後の行を取得する方法を示しています (前の例では、[ファームで販売されたクレートの合計] 列と [販売済みクレートの合計] 列の両方の総計)。  これらの値は、セル **E30** (ピボットテーブルの外側) に表示される最終的な合計に合わせて合計されます。

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // Get the totals for each data hierarchy from the layout.
    var range = pivotTable.layout.getDataBodyRange();
    var grandTotalRange = range.getLastRow();
    grandTotalRange.load("address");
    return context.sync().then(function () {
        // Sum the totals from the PivotTable data hierarchies and place them in a new range, outside of the PivotTable.
        var masterTotalRange = context.workbook.worksheets.getActiveWorksheet().getRange("E30");
        masterTotalRange.formulas = [["=SUM(" + grandTotalRange.address + ")"]];
    });
});
```

### <a name="layout-types"></a>レイアウトの種類

ピボットテーブルには、コンパクト、アウトライン、表形式の 3 つのレイアウト スタイルがあります。 前の例では、コンパクトなスタイルを見ていました。

次の例では、アウトラインスタイルと表形式スタイルをそれぞれ使用します。 コード サンプルは、異なるレイアウト間を切り替える方法を示しています。

#### <a name="outline-layout"></a>アウトライン レイアウト

![アウトライン レイアウトを使用するピボットテーブル。](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a>表形式のレイアウト

![表形式のレイアウトを使用するピボットテーブル。](../images/excel-pivots-tabular-layout.png)

#### <a name="pivotlayout-type-switch-code-sample"></a>PivotLayout の種類のスイッチ コードのサンプル

```js
Excel.run(function (context) {
    // Change the PivotLayout.type to a new type.
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.layout.load("layoutType");
    return context.sync().then(function () {
        // Cycle between the three layout types.
        if (pivotTable.layout.layoutType === "Compact") {
            pivotTable.layout.layoutType = "Outline";
        } else if (pivotTable.layout.layoutType === "Outline") {
            pivotTable.layout.layoutType = "Tabular";
        } else {
            pivotTable.layout.layoutType = "Compact";
        }
    
        return context.sync();
    });
});
```

### <a name="other-pivotlayout-functions"></a>その他の PivotLayout 関数

既定では、ピボットテーブルは必要に応じて行と列のサイズを調整します。 これは、ピボットテーブルが更新された場合に実行されます。 `PivotLayout.autoFormat` その動作を指定します。 アドインによって行われた行または列のサイズの変更は、次の場合も保持 `autoFormat` されます `false` 。 さらに、ピボットテーブルの既定の設定では、ピボットテーブル内のカスタム書式 (塗りつぶしやフォントの変更など) が保持されます。 更新 `PivotLayout.preserveFormatting` 時 `false` に既定の形式を適用する場合に設定します。

また `PivotLayout` 、ヘッダーと行の合計設定、空のデータ セルの表示方法、および代替テキスト オプション [も制御](https://support.microsoft.com/topic/44989b2a-903c-4d9a-b742-6a75b451c669) します。 [PivotLayout 参照は](/javascript/api/excel/excel.pivotlayout)、これらの機能の完全な一覧を提供します。

次のコード サンプルでは、空のデータ セルに文字列を表示し、本文範囲を一貫性のある水平方向の配置に書式設定し、ピボットテーブルが更新された後も書式設定の変更が維持されます `"--"` 。

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("Farm Sales");
    var pivotLayout = pivotTable.layout;

    // Set a default value for an empty cell in the PivotTable. This doesn't include cells left blank by the layout.
    pivotLayout.emptyCellText = "--";

    // Set the text alignment to match the rest of the PivotTable.
    pivotLayout.getDataBodyRange().format.horizontalAlignment = Excel.HorizontalAlignment.right;

    // Ensure empty cells are filled with a default value.
    pivotLayout.fillEmptyCells = true;

    // Ensure that the format settings persist, even after the PivotTable is refreshed and recalculated.
    pivotLayout.preserveFormatting = true;
    return context.sync();
});
```

## <a name="delete-a-pivottable"></a>ピボットテーブルの削除

ピボットテーブルは、その名前を使用して削除されます。

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="filter-a-pivottable"></a>ピボットテーブルをフィルター処理する

ピボットテーブル データをフィルター処理する主な方法は、PivotFilters です。 スライサーは、柔軟性の低い代替フィルター方法を提供します。

[PivotFilters は](/javascript/api/excel/excel.pivotfilters) 、ピボットテーブルの 4 つの階層カテゴリ [(フィルター](#hierarchies) 、列、行、値) に基づいてデータをフィルター処理します。 PivotFilter には 4 つの種類があります。予定表の日付ベースのフィルター処理、文字列解析、数値比較、およびカスタム入力に基づくフィルター処理が可能です。

[スライサー](/javascript/api/excel/excel.slicer)は、ピボットテーブルと通常のテーブルの両方Excelできます。 ピボットテーブルに適用すると、スライサーは [PivotManualFilter](#pivotmanualfilter) のように機能し、カスタム入力に基づいてフィルター処理を許可します。 PivotFilters とは異なり、スライサーには UI[コンポーネントExcelがあります](https://support.microsoft.com/office/249f966b-a9d5-4b0f-b31a-12651785d29d)。 クラスを `Slicer` 使用して、この UI コンポーネントを作成し、フィルター処理を管理し、その外観を制御します。

### <a name="filter-with-pivotfilters"></a>PivotFilters を使用したフィルター

[PivotFilters を使用](/javascript/api/excel/excel.pivotfilters) すると、4 つの階層カテゴリ [(フィルター](#hierarchies) 、列、行、値) に基づいてピボットテーブル データをフィルター処理できます。 PivotTable オブジェクト モデルでは、PivotField に適用され、それぞれが 1 つ以上 `PivotFilters` の値を割[](/javascript/api/excel/excel.pivotfield) `PivotField` り当てることができます `PivotFilters` 。 PivotField にピボットフィルターを適用するには、フィールドの対応する [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) を階層カテゴリに割り当てる必要があります。

#### <a name="types-of-pivotfilters"></a>PivotFilters の種類

| フィルターの種類 | フィルターの目的 | Excel JavaScript API リファレンス |
|:--- |:--- |:--- |
| DateFilter | 予定表の日付ベースのフィルター処理。 | [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) |
| LabelFilter | テキスト比較フィルター。 | [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) |
| ManualFilter | カスタム入力フィルター。 | [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) |
| ValueFilter | 数値比較フィルター。 | [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter) |

#### <a name="create-a-pivotfilter"></a>ピボットフィルターの作成

ピボットテーブル データを (a など) でフィルター処理するには、 `Pivot*Filter` `PivotDateFilter` ピボットフィールドにフィルターを [適用します](/javascript/api/excel/excel.pivotfield)。 次の 4 つのコード サンプルは、4 種類の PivotFilter のそれぞれを使用する方法を示しています。

##### <a name="pivotdatefilter"></a>PivotDateFilter

最初のコード サンプルでは [、PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) を Date **Updated** PivotField に適用し **、2020-08-01** より前のデータを非表示にしています。

> [!IMPORTANT]
> そのフィールドの PivotHierarchy が階層カテゴリに割り当てられていない限り、ピボットフィールドに A を `Pivot*Filter` 適用することはできません。 次のコード サンプルでは、ピボットテーブルをフィルター処理に使用する前に、ピボットテーブルのカテゴリに追加 `dateHierarchy` `rowHierarchies` する必要があります。

```js
Excel.run(function (context) {
    // Get the PivotTable and the date hierarchy.
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var dateHierarchy = pivotTable.rowHierarchies.getItemOrNullObject("Date Updated");
    
    return context.sync().then(function () {
        // PivotFilters can only be applied to PivotHierarchies that are being used for pivoting.
        // If it's not already there, add "Date Updated" to the hierarchies.
        if (dateHierarchy.isNullObject) {
          dateHierarchy = pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Date Updated"));
        }

        // Apply a date filter to filter out anything logged before August.
        var filterField = dateHierarchy.fields.getItem("Date Updated");
        var dateFilter = {
          condition: Excel.DateFilterCondition.afterOrEqualTo,
          comparator: {
            date: "2020-08-01",
            specificity: Excel.FilterDatetimeSpecificity.month
          }
        };
        filterField.applyFilter({ dateFilter: dateFilter });
        
        return context.sync();
    });
});
```

> [!NOTE]
> 次の 3 つのコード スニペットは、完全な呼び出しではなく、フィルター固有の抜粋のみを表示 `Excel.run` します。

##### <a name="pivotlabelfilter"></a>PivotLabelFilter

2 番目のコード スニペットは、プロパティを使用して文字 L で始まるラベルを除外して [、PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) を Type **PivotField** に適用する方法 `LabelFilterCondition.beginsWith` を **示しています**。

```js
    // Get the "Type" field.
    var filterField = pivotTable.hierarchies.getItem("Type").fields.getItem("Type");

    // Filter out any types that start with "L" ("Lemons" and "Limes" in this case).
    var filter: Excel.PivotLabelFilter = {
      condition: Excel.LabelFilterCondition.beginsWith,
      substring: "L",
      exclusive: true
    };

    // Apply the label filter to the field.
    filterField.applyFilter({ labelFilter: filter });
```

##### <a name="pivotmanualfilter"></a>PivotManualFilter

3 番目のコード スニペットは [、PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)を含む手動フィルターを [分類] フィールドに適用し、分類オーガニック を含むデータをフィルター処理 **します**。

```js
    // Apply a manual filter to include only a specific PivotItem (the string "Organic").
    var filterField = classHierarchy.fields.getItem("Classification");
    var manualFilter = { selectedItems: ["Organic"] };
    filterField.applyFilter({ manualFilter: manualFilter });
```

##### <a name="pivotvaluefilter"></a>PivotValueFilter

数値を比較するには、最終的なコード スニペットに示すように [、PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)と値フィルターを使用します。 ファーム ピボットフィールドのデータと、販売されたクレートの合計が値 `PivotValueFilter` **500** を超えるファームのみを含む、クレート販売済みホールセール ピボットフィールドのデータとを比較します。

```js
    // Get the "Farm" field.
    var filterField = pivotTable.hierarchies.getItem("Farm").fields.getItem("Farm");
    
    // Filter to only include rows with more than 500 wholesale crates sold.
    var filter: Excel.PivotValueFilter = {
      condition: Excel.ValueFilterCondition.greaterThan,
      comparator: 500,
      value: "Sum of Crates Sold Wholesale"
    };
    
    // Apply the value filter to the field.
    filterField.applyFilter({ valueFilter: filter });
```

#### <a name="remove-pivotfilters"></a>ピボットフィルターの削除

すべての PivotFilter を削除するには、次のコード サンプルに示すように、各 PivotField にメソッド `clearAllFilters` を適用します。

```js
Excel.run(function (context) {
    // Get the PivotTable.
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.hierarchies.load("name");
    
    return context.sync().then(function () {
        // Clear the filters on each PivotField.
        pivotTable.hierarchies.items.forEach(function (hierarchy) {
          hierarchy.fields.getItem(hierarchy.name).clearAllFilters();
        });
        return context.sync();
    });
});
```

### <a name="filter-with-slicers"></a>スライサーを使用したフィルター

[スライサー](/javascript/api/excel/excel.slicer)を使用すると、ピボットテーブルまたはテーブルExcelデータをフィルター処理できます。 スライサーは、指定した列または PivotField の値を使用して、対応する行をフィルター処理します。 これらの値は、 [に SlicerItem](/javascript/api/excel/excel.sliceritem) オブジェクトとして格納されます `Slicer` 。 アドインは、ユーザーと同様に、これらのフィルターを調整できます (ユーザーは、Excel[UI を使用します](https://support.microsoft.com/office/249f966b-a9d5-4b0f-b31a-12651785d29d))。 次のスクリーンショットに示すように、スライサーは図面レイヤーのワークシートの上に配置されます。

![ピボットテーブル上のデータをスライサー フィルター処理します。](../images/excel-slicer.png)

> [!NOTE]
> このセクションで説明する手法は、ピボットテーブルに接続されたスライサーの使い方に焦点を当ててします。 同じ手法は、テーブルに接続されたスライサーの使用にも適用されます。

#### <a name="create-a-slicer"></a>スライサーを作成する

メソッドまたはメソッドを使用して、ブックまたはワークシートにスライサー `Workbook.slicers.add` を作成 `Worksheet.slicers.add` できます。 指定したオブジェクトまたはオブジェクトの [SlicerCollection](/javascript/api/excel/excel.slicercollection) にスライサーを `Workbook` 追加 `Worksheet` します。 メソッド `SlicerCollection.add` には、次の 3 つのパラメーターがあります。

- `slicerSource`: 新しいスライサーが基づくデータ ソース。 名前または ID を表す 、 、または文字列を指定 `PivotTable` `Table` `PivotTable` できます `Table` 。
- `sourceField`: フィルター処理するデータ ソースのフィールド。 名前または ID を表す 、 、または文字列を指定 `PivotField` `TableColumn` `PivotField` できます `TableColumn` 。
- `slicerDestination`: 新しいスライサーが作成されるワークシート。 オブジェクト、または `Worksheet` . の名前または ID を指定できます `Worksheet` 。 を使用してアクセスする場合 `SlicerCollection` 、このパラメーターは不要です `Worksheet.slicers` 。 この場合、コレクションのワークシートが移動先として使用されます。

次のコード サンプルでは、ピボット ワークシートに新しいスライサー **を追加** します。 スライサーのソースは **、Farm Sales ピボット** テーブルであり、Type データを使用して **フィルター処理** します。 スライサーは、将来の参照 **のために、Fruit Slicer という** 名前も付けます。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Pivot");
    var slicer = sheet.slicers.add(
        "Farm Sales" /* The slicer data source. For PivotTables, this can be the PivotTable object reference or name. */,
        "Type" /* The field in the data to filter by. For PivotTables, this can be a PivotField object reference or ID. */
    );
    slicer.name = "Fruit Slicer";
    return context.sync();
});
```

#### <a name="filter-items-with-a-slicer"></a>スライサーを使用してアイテムをフィルター処理する

スライサーはピボットテーブルにフィルターを適用し、. `sourceField` この `Slicer.selectItems` メソッドは、スライサーに残るアイテムを設定します。 これらのアイテムは、アイテムのキーを表す `string[]` 、 としてメソッドに渡されます。 これらのアイテムを含む行は、ピボットテーブルの集計に残ります。 以降の呼び `selectItems` 出しでは、リストをそれらの呼び出しで指定されたキーに設定します。

> [!NOTE]
> データ ソースに含めされていないアイテムが渡された場合は `Slicer.selectItems` 、 `InvalidArgument` エラーがスローされます。 コンテンツは `Slicer.slicerItems` [、SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)であるプロパティを通じて確認できます。

次のコード サンプルは、スライサーで選択されている 3つの項目を示 **しています。**

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

スライサーからすべてのフィルターを削除するには、次のサンプルに示 `Slicer.clearFilters` すように、メソッドを使用します。

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

#### <a name="style-and-format-a-slicer"></a>スライサーのスタイルと書式設定

アドインは、プロパティを使用してスライサーの表示設定を調整 `Slicer` できます。 次のコード サンプルでは、スタイルを **SlicerStyleLight6** に設定し、スライサーの上部にあるテキストを **[フルーツ** の種類] に設定し、スライサーを描画レイヤーの **位置 (395、 15)** に配置し、スライサーのサイズを **135x150** ピクセルに設定します。

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.caption = "Fruit Types";
    slicer.left = 395;
    slicer.top = 15;
    slicer.height = 135;
    slicer.width = 150;
    slicer.style = "SlicerStyleLight6";
    return context.sync();
});
```

#### <a name="delete-a-slicer"></a>スライサーを削除する

スライサーを削除するには、メソッドを呼び出 `Slicer.delete` します。 次のコード サンプルでは、現在のワークシートから最初のスライサーを削除します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a>変更集約関数

データ階層の値は集計されます。 数値のデータセットの場合、これは既定では合計です。 この `summarizeBy` プロパティは、AggregationFunction 型に基づいて [この動作を定義](/javascript/api/excel/excel.aggregationfunction) します。

現在サポートされている集計関数の種類は `Sum` `Count` 、、、( `Average` `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP` `Automatic` 既定値) です。

次のコード サンプルでは、集計をデータの平均に変更します。

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.dataHierarchies.load("no-properties-needed");
    return context.sync().then(function() {

        // Change the aggregation from the default sum to an average of all the values in the hierarchy.
        pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.average;
        pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.average;
        return context.sync();
    });
});
```

## <a name="change-calculations-with-a-showasrule"></a>ShowAsRule を使用して計算を変更する

ピボットテーブルは、既定では、行階層と列階層のデータを個別に集計します。 [ShowAsRule は](/javascript/api/excel/excel.showasrule)、ピボットテーブル内の他のアイテムに基づいてデータ階層を出力値に変更します。

オブジェクト `ShowAsRule` には、次の 3 つのプロパティがあります。

- `calculation`: データ階層に適用する相対計算の種類 (既定値は `none` ) です。
- `baseField`: 計算が適用される前の基本データを含む階層内の[PivotField。](/javascript/api/excel/excel.pivotfield) ピボットテーブルExcel、階層とフィールドの 1 対 1 のマッピングを持つので、同じ名前を使用して階層とフィールドの両方にアクセスします。
- `baseItem`: 計算の種類に基づいて基本フィールドの値と比較される個々の[PivotItem。](/javascript/api/excel/excel.pivotitem) すべての計算でこのフィールドが必要な場合があります。

次の使用例は、ファームデータ階層で販売されたクレートの合計の計算を、列の合計に対する割合に設定します。
この場合も、粒度をフルーツの種類レベルまで拡張する必要があります。そのため **、Type** 行階層とその基になるフィールドを使用します。
この例では、 **最初** の行階層として Farm も含まれています。そのため、ファームの合計エントリには、各ファームが生成する割合も表示されます。

![個々のファームと各ファーム内の個々の果物の種類の両方の総計に対する果物の売上の割合を示すピボットテーブル。](../images/excel-pivots-showas-percentage.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    return context.sync().then(function () {

        // Show the crates of each fruit type sold at the farm as a percentage of the column's total.
        var farmShowAs = farmDataHierarchy.showAs;
        farmShowAs.calculation = Excel.ShowAsCalculation.percentOfColumnTotal;
        farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");
        farmDataHierarchy.showAs = farmShowAs;
        farmDataHierarchy.name = "Percentage of Total Farm Sales";
    });
});
```

前の使用例は、個々の行階層のフィールドを基準として、列に計算を設定します。 計算が個々のアイテムに関連する場合は、プロパティを使用 `baseItem` します。

次の例は、計算を示 `differenceFrom` しています。 ファームクレート販売データ階層エントリの違いを **、A Farms** のエントリと相対的に表示します。
is Farm です。したがって、他のファーム間の違い、および同様のフルーツの種類ごとに内訳が表示されます `baseField` (この例では **、Type** も行階層です)。

!["A Farms" と他のファームの果物販売の違いを示すピボットテーブル。 これは、ファームの果物の総売上と種類の果物の販売の両方の違いを示しています。 "A Farms" が特定の種類の果物を販売していない場合は、"#N/A" が表示されます。](../images/excel-pivots-showas-differencefrom.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    return context.sync().then(function () {
        // Show the difference between crate sales of the "A Farms" and the other farms.
        // This difference is both aggregated and shown for individual fruit types (where applicable).
        var farmShowAs = farmDataHierarchy.showAs;
        farmShowAs.calculation = Excel.ShowAsCalculation.differenceFrom;
        farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");
        farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");
        farmDataHierarchy.showAs = farmShowAs;
        farmDataHierarchy.name = "Difference from A Farms";
    });
});
```

## <a name="change-hierarchy-names"></a>階層名の変更

階層フィールドは編集可能です。 次のコードは、2 つのデータ階層の表示名を変更する方法を示しています。

```js
Excel.run(function (context) {
    var dataHierarchies = context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.getItem("Farm Sales").dataHierarchies;
    dataHierarchies.load("no-properties-needed");
    return context.sync().then(function () {
        // changing the displayed names of these entries
        dataHierarchies.items[0].name = "Farm Sales";
        dataHierarchies.items[1].name = "Wholesale";
    });
});
```

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [ExcelJavaScript API リファレンス](/javascript/api/excel)
