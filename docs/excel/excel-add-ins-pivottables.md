---
title: Excel JavaScript API を使用してピボットテーブルを使用する
description: Excel JavaScript API を使用してピボットテーブルを作成し、それらのコンポーネントを操作します。
ms.date: 01/26/2021
localization_priority: Normal
ms.openlocfilehash: 9832322d40bbeb247685ff2498bdce42975c0377
ms.sourcegitcommit: 3123b9819c5225ee45a5312f64be79e46cbd0e3c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/29/2021
ms.locfileid: "50043912"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してピボットテーブルを使用する

ピボットテーブルは、より大きなデータ セットを合理化します。 グループ化されたデータを迅速に操作できます。 Excel JavaScript API を使用すると、アドインでピボットテーブルを作成し、それらのコンポーネントを操作できます。 この記事では、Office JavaScript API によってピボットテーブルがどのように表現されるのかについて説明し、主要なシナリオのコード サンプルを提供します。

ピボットテーブルの機能に慣れていない場合は、エンド ユーザーとして探索を検討してください。
これらの [ツールの優れた入門情報については、「](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) ピボットテーブルを作成してワークシート データを分析する」を参照してください。

> [!IMPORTANT]
> OLAP で作成されたピボットテーブルは現在サポートされていません。 Power Pivot もサポートされていません。

## <a name="object-model"></a>オブジェクト モデル

ピボット [テーブルは](/javascript/api/excel/excel.pivottable) 、JavaScript API のピボットテーブルOfficeオブジェクトです。

- `Workbook.pivotTables`は、ブックとワークシートのピボットテーブルをそれぞれ含む `Worksheet.pivotTables` [PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)です。 [](/javascript/api/excel/excel.pivottable)
- ピボット[テーブルには、](/javascript/api/excel/excel.pivottable)複数の PivotHierarchies を持つ[PivotHierarchyCollection が含まれる](/javascript/api/excel/excel.pivothierarchy)。 [](/javascript/api/excel/excel.pivothierarchycollection)
- これらの [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy) を特定の階層コレクションに追加して、ピボットテーブルのピボットデータの方法を定義できます (以下のセクション [で説明します](#hierarchies))。
- [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)には、PivotField が 1 つ正確に含まれる[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) [が含まれる](/javascript/api/excel/excel.pivotfield)。 OLAP ピボットテーブルを含むデザインが拡張された場合、これは変更される可能性があります。
- [](/javascript/api/excel/excel.pivotfilters)ピボット[フィールドの](/javascript/api/excel/excel.pivotfield) [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)が階層カテゴリに割り当てられている限り、ピボットフィールドには 1 つ以上のピボットフィルターを適用できます。 
- ピボット [フィールドには](/javascript/api/excel/excel.pivotfield) 、複数の [PivotItem を持つ PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) [が含まれる](/javascript/api/excel/excel.pivotitem)。
- ピボット[テーブルには](/javascript/api/excel/excel.pivottable)、ピボットフィールドとピボットアイテムがワークシート内で表示[](/javascript/api/excel/excel.pivotfield)される場所を定義する[PivotLayout](/javascript/api/excel/excel.pivotlayout)が含まれる。 [](/javascript/api/excel/excel.pivotitem)

これらの関係がいくつかのサンプル データに適用される方法を見てみしましょう。 次のデータは、さまざまなファームからの青果売上を示しています。 この記事全体の例を示します。

![さまざまな種類のファームからのさまざまな種類の青果売上のコレクション。](../images/excel-pivots-raw-data.png)

このファームの売上データは、ピボットテーブルの作成に使用されます。 型などの **各列は**、 `PivotHierarchy` . Types **階層には** 、Types フィールド **が含** まれます。 Types **フィールドには**、Apple、Kiwi、Orange、**および Orange** の各項目 **が含****まれます**。 

### <a name="hierarchies"></a>Hierarchies

ピボットテーブルは、行、列、データ、およびフィルター[](/javascript/api/excel/excel.rowcolumnpivothierarchy)の[](/javascript/api/excel/excel.rowcolumnpivothierarchy)4 つの階層[カテゴリに基](/javascript/api/excel/excel.datapivothierarchy)づいて編成[されます](/javascript/api/excel/excel.filterpivothierarchy)。

前に示したファーム データには 5 つの階層があります。**ファーム**、種類、**分類**、ファームで販売されたクレート、および販売された区分 **。**  各階層は、4 つのカテゴリの 1 つにのみ存在できます。 Type **が** 列階層に追加された場合は、行、データ、またはフィルター階層にも追加できません。 その **後** 、Type が行階層に追加されると、列階層から削除されます。 この動作は、階層の割り当てが Excel UI または Excel JavaScript API を使用して行われる場合でも同じです。

行と列の階層は、データをグループ化する方法を定義します。 たとえば、ファームの行階層 **は** 、同じファームのすべてのデータ セットをグループ化します。 行と列の階層を選択すると、ピボットテーブルの向きが定義されます。

データ階層は、行と列の階層に基づいて集計される値です。 ファームの行階層と **Crates Sold Sold の** データ階層を持つピボットテーブルには、ファームごとに異なるすべての青果の合計 (既定) が表示されます。

フィルター階層は、フィルター処理された型内の値に基づいてピボットからデータを含めるか除外します。 分類のフィルター階層で、種類が **[****組織**] で選択されている場合は、青い青果のデータだけが表示されます。

ここでは、ピボットテーブルと共にファーム データを再び示します。 ピボットテーブルでは、行階層として **Farm** と **Type** を使用し、データ階層として (合計の既定の集計関数を使用して) ファームで販売された **Crates** **と Crates Sold Hierarchie** を使用し、分類をフィルター階層として ([組織的] を選択した場合) 使用します。 

![行、データ、およびフィルター階層を持つピボットテーブルの横にある、青果売上データの選択。](../images/excel-pivot-table-and-data.png)

このピボットテーブルは、JavaScript API または Excel UI を使用して生成できます。 どちらのオプションでも、アドインを介してさらに操作できます。

## <a name="create-a-pivottable"></a>ピボットテーブルを作成する

ピボットテーブルには、名前、ソース、および変換先が必要です。 ソースには、範囲のアドレスまたはテーブル名を指定できます (型 `Range` として `string` 渡 `Table` されます)。 宛先は、範囲アドレス (a または ) `Range` です `string` 。
次のサンプルは、さまざまなピボットテーブル作成の手法を示しています。

### <a name="create-a-pivottable-with-range-addresses"></a>範囲アドレスを含むピボットテーブルを作成する

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

手動で作成されたピボットテーブルには、ブックのピボットテーブル コレクションまたは個々のワークシートからアクセスすることもできます。 次のコードは、ブックから **My Pivot という名前のピボット** テーブルを取得します。

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>ピボットテーブルに行と列を追加する

行と列は、これらのフィールドの値を中心にデータをピボットします。

ファーム列 **を追加** すると、各ファームのすべての売上がピボットされます。 [種類 **] 行と** **[分類** ] 行を追加すると、販売された青の種類と、その商品の種類や種類に基づいてデータがさらに分類されます。

![[ファーム] 列と [種類] 行と [分類] 行があるピボットテーブル。](../images/excel-pivots-table-rows-and-columns.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    return context.sync();
});
```

行または列のみを含むピボットテーブルを作成することもできます。

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

データ階層は、行と列に基づいて結合する情報でピボットテーブルを埋め込む。 ファームで販売されたクレートと販売された製品版のクレートのデータ階層を追加すると、各行と列のこれらの数値の合計が提供されます。

この例では **、Farm** と Type の **両方** が行であり、クレート売上がデータです。

![出所ファームに基づくさまざまな青果の総売上高を示すピボットテーブル。](../images/excel-pivots-data-hierarchy.png)

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

[PivotLayout は、](/javascript/api/excel/excel.pivotlayout)階層とそのデータの配置を定義します。 レイアウトにアクセスして、データが格納される範囲を決定します。

次の図は、ピボットテーブルの範囲に対応するレイアウト関数の呼び出しを示しています。

![レイアウトの範囲取得関数によって返されるピボットテーブルのセクションを示す図。](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a>ピボットテーブルからデータを取得する

レイアウトは、ワークシートでのピボットテーブルの表示方法を定義します。 つまり、オブジェクト `PivotLayout` はピボットテーブル要素に使用される範囲を制御します。 ピボットテーブルによって収集および集計されるデータを取得するには、レイアウトによって提供される範囲を使用します。 特に、ピボット `PivotLayout.getDataBodyRange` テーブルが生成するデータにアクセスするために使用します。

次のコードは、レイアウト (前の例では[ファームで販売されたクレートの合計] 列と [販売されたクレートの合計] 列の両方の総計) を使用して、ピボットテーブル データの最後の行を取得する方法を示しています。  これらの値は、セル **E30** (ピボットテーブルの外側) に表示される最終的な合計に合わせて合計されます。

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

ピボットテーブルには、コンパクト、アウトライン、表形式の 3 つのレイアウト スタイルがあります。 前の例では、コンパクトなスタイルを確認しました。

次の例では、アウトラインスタイルと表形式スタイルをそれぞれ使用します。 コード サンプルは、さまざまなレイアウト間を切り替える方法を示しています。

#### <a name="outline-layout"></a>アウトライン レイアウト

![アウトライン レイアウトを使用するピボットテーブル。](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a>表形式レイアウト

![表形式レイアウトを使用するピボットテーブル。](../images/excel-pivots-tabular-layout.png)

## <a name="delete-a-pivottable"></a>ピボットテーブルを削除する

ピボットテーブルは、名前を使用して削除されます。

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="filter-a-pivottable"></a>ピボットテーブルをフィルター処理する

ピボットテーブル データをフィルター処理する主な方法は、PivotFilters です。 スライサーは、代替の柔軟性の低いフィルタリング方法を提供します。 

[ピボットフィルターは、](/javascript/api/excel/excel.pivotfilters) ピボットテーブルの 4 つの階層カテゴリ [(フィルター](#hierarchies) 、列、行、値) に基づいてデータをフィルター処理します。 PivotFilter には、カレンダーの日付ベースのフィルター処理、文字列解析、数値比較、カスタム入力に基づくフィルター処理の 4 種類があります。 

[スライサー](/javascript/api/excel/excel.slicer) は、ピボットテーブルと通常の Excel テーブルの両方に適用できます。 ピボットテーブルに適用すると、スライサーは [PivotManualFilter](#pivotmanualfilter) のように機能し、カスタム入力に基づいてフィルター処理を実行できます。 PivotFilters とは異なり、スライサーには [Excel UI コンポーネントがあります](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)。 クラスを `Slicer` 使って、この UI コンポーネントを作成し、フィルター処理を管理し、その外観を制御します。 

### <a name="filter-with-pivotfilters"></a>PivotFilters でフィルター処理する

[ピボットフィルターを使用](/javascript/api/excel/excel.pivotfilters) すると、4 つの階層カテゴリ [(フィルター](#hierarchies) 、列、行、値) に基づいてピボットテーブル データをフィルター処理できます。 ピボットテーブル オブジェクト モデルでは、ピボットフィールドに適用され、それぞれが 1 つ以上を割り `PivotFilters` [](/javascript/api/excel/excel.pivotfield) `PivotField` 当てることができます `PivotFilters` 。 ピボットフィールドにピボットフィルターを適用するには、フィールドに対応する [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) を階層カテゴリに割り当てる必要があります。 

#### <a name="types-of-pivotfilters"></a>PivotFilters の種類

| フィルターの種類 | フィルターの目的 | Excel JavaScript API リファレンス |
|:--- |:--- |:--- |
| DateFilter | カレンダーの日付ベースのフィルター。 | [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) |
| LabelFilter | テキスト比較フィルター。 | [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) |
| ManualFilter | カスタム入力フィルター。 | [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) |
| ValueFilter | 数値比較フィルター。 | [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter) |

#### <a name="create-a-pivotfilter"></a>PivotFilter を作成する

ピボットテーブル データをフィルター処理する (a など) には、ピボット `Pivot*Filter` `PivotDateFilter` フィールドにフィルター [を適用します](/javascript/api/excel/excel.pivotfield)。 次の 4 つのコード サンプルは、4 種類の PivotFilter のそれぞれを使用する方法を示しています。 

##### <a name="pivotdatefilter"></a>PivotDateFilter

最初のコード サンプルでは [、PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) を Date **Updated** PivotField に適用し **、2020-08-01** より前のデータを非表示にしています。 

> [!IMPORTANT] 
> A は、そのフィールドの PivotHierarchy が階層カテゴリに割り当てられていない限り、 `Pivot*Filter` ピボットフィールドに適用できません。 次のコード サンプルでは、フィルター処理に使用する前にピボットテーブルのカテゴリに追加 `dateHierarchy` `rowHierarchies` する必要があります。

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

2 番目のコード スニペットは、プロパティを使用して文字 L で始まるラベルを除外して [、PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) を **タイプ** ピボットフィールドに適用する方法 `LabelFilterCondition.beginsWith` を **示しています**。 

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

3 番目のコード スニペットでは [、PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) を含む手動フィルターを **Classification** フィールドに適用し、分類の [グループ化] を含むデータをフィルター **処理します**。 

```js
    // Apply a manual filter to include only a specific PivotItem (the string "Organic").
    var filterField = classHierarchy.fields.getItem("Classification");
    var manualFilter = { selectedItems: ["Organic"] };
    filterField.applyFilter({ manualFilter: manualFilter });
```

##### <a name="pivotvaluefilter"></a>PivotValueFilter

数値を比較するには、最終的なコード スニペットに示すように、値フィルターと [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)を使用します。 The `PivotValueFilter` compares the data in the **Farm** PivotField to the data in the **Crates Sold PivotField,** including only farms whose sum of crates sold exceeds the value **500**. 

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

#### <a name="remove-pivotfilters"></a>PivotFilters を削除する

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

### <a name="filter-with-slicers"></a>スライサーでフィルター処理する

[スライサー](/javascript/api/excel/excel.slicer) を使用すると、Excel ピボットテーブルまたはテーブルからデータをフィルター処理できます。 スライサーは、指定された列またはピボットフィールドの値を使用して、対応する行をフィルター処理します。 これらの値は [、SlicerItem オブジェクトとして](/javascript/api/excel/excel.sliceritem) 格納されます `Slicer` 。 アドインは、(Excel UI を使用して) ユーザーと同様に[、これらのフィルターを調整できます](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)。 スライサーは、次のスクリーンショットに示すように、描画レイヤーのワークシートの上に配置されます。

![ピボットテーブルのデータをフィルター処理するスライサー。](../images/excel-slicer.png)

> [!NOTE]
> このセクションで説明する手法では、ピボットテーブルに接続されたスライサーの使い方に重点を置いて説明します。 テーブルに接続されたスライサーの使用にも同じ手法が適用されます。

#### <a name="create-a-slicer"></a>スライサーを作成する

このメソッドまたはメソッドを使用して、ブックまたはワークシートにスライサー `Workbook.slicers.add` を作成 `Worksheet.slicers.add` できます。 これにより、指定したオブジェクトの [SlicerCollection](/javascript/api/excel/excel.slicercollection) にスライサーが `Workbook` 追加 `Worksheet` されます。 この `SlicerCollection.add` メソッドには、次の 3 つのパラメーターがあります。

- `slicerSource`: 新しいスライサーが基づくデータ ソース。 名前または ID を表す 、または文字列 `PivotTable` `Table` を指定 `PivotTable` できます `Table` 。
- `sourceField`: フィルター処理に使用するデータ ソースのフィールドです。 名前または ID を表す 、または文字列 `PivotField` `TableColumn` を指定 `PivotField` できます `TableColumn` 。
- `slicerDestination`: 新しいスライサーが作成されるワークシートです。 オブジェクト、 `Worksheet` またはオブジェクトの名前または ID を指定できます `Worksheet` 。 このパラメーターは、アクセスするときに `SlicerCollection` 不要です `Worksheet.slicers` 。 この場合、コレクションのワークシートが移動先として使用されます。

次のコード サンプルでは、ピボット ワークシートに新しいスライサー **を追加** します。 スライサーのソースは **Farm Sales ピボット** テーブルで、種類データを使用して **フィルター処理** します。 スライサーは、今後の参照用 **に"Slicer"** という名前も付けます。

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

#### <a name="filter-items-with-a-slicer"></a>スライサーでアイテムをフィルター処理する

スライサーは、 `sourceField` . この `Slicer.selectItems` メソッドは、スライサーに残っているアイテムを設定します。 これらの項目は、項目のキーを表すメソッドとして `string[]` 渡されます。 これらのアイテムを含む行は、ピボットテーブルの集計に残ります。 後続の呼 `selectItems` び出しで、それらの呼び出しで指定されたキーにリストを設定します。

> [!NOTE]
> データ `Slicer.selectItems` ソースに含めされていないアイテムが渡された場合は、 `InvalidArgument` エラーがスローされます。 内容は、プロパティ `Slicer.slicerItems` [(SlicerItemCollection)](/javascript/api/excel/excel.sliceritemcollection)を通じて確認できます。

次のコード サンプルは、スライサーに対して選択されている 3 つの項目を示 **しています。** 

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

スライサーからすべてのフィルターを削除するには、次のサンプルに示 `Slicer.clearFilters` すようにメソッドを使用します。

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

#### <a name="style-and-format-a-slicer"></a>スライサーのスタイルと書式設定

アドインでは、プロパティを使用してスライサーの表示設定を調整 `Slicer` できます。 次のコード サンプルでは、スタイルを **SlicerStyleLight6** に設定し、スライサーの上部にあるテキストを **[種類**] に設定し、スライサーを描画レイヤー上の位置 **(395、15)** に配置し、スライサーのサイズを **135x150** ピクセルに設定します。

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

## <a name="change-aggregation-function"></a>変更集計関数

データ階層の値は集計されます。 数値のデータセットの場合、これは既定で合計です。 この `summarizeBy` プロパティは [、AggregationFunction](/javascript/api/excel/excel.aggregationfunction) 型に基づいてこの動作を定義します。

現在サポートされている集計関数の種類は、,, `Sum` `Count` `Average` `Max` and `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP` `Automatic` (既定値) です。

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

ピボットテーブルは、既定では、行階層と列階層のデータを個別に集計します。 [ShowAsRule は](/javascript/api/excel/excel.showasrule)、ピボットテーブル内の他のアイテムに基づいて、データ階層を出力値に変更します。

オブジェクト `ShowAsRule` には次の 3 つのプロパティがあります。

- `calculation`: データ階層に適用する相対的な計算の種類を指定します (既定値は次の値です `none` )。
- `baseField`: 計算 [を適用](/javascript/api/excel/excel.pivotfield) する前の基本データを含む階層内のピボットフィールド。 Excel のピボットテーブルには、階層とフィールドの 1 対 1 のマッピングが含まれます。この名前を使用して、階層とフィールドの両方にアクセスします。
- `baseItem`: 計算の種類に基づいて、各 [PivotItem](/javascript/api/excel/excel.pivotitem) が基本フィールドの値と比較されます。 すべての計算でこのフィールドが必要な場合があります。

次の使用例は、ファームデータ階層で販売された商品の合計の計算を、列の合計に対する割合に設定します。
この場合も、粒度を青の種類のレベルまで拡張する必要があります。したがって **、Type** 行階層とその基になるフィールドを使用します。
この例では **、1** 行目の階層として Farm も含まれています。したがって、ファームの合計エントリ数には、各ファームが生成する割合も表示されます。

![個々のファームおよび各ファーム内の個々の種類の青果の総計に対する、青果売上の割合を示すピボットテーブル。](../images/excel-pivots-showas-percentage.png)

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

前の例では、個々の行階層のフィールドを基準に計算を列に設定します。 計算が個々のアイテムに関連する場合は、プロパティを使用 `baseItem` します。

次の例は、計算を示 `differenceFrom` しています。 ファームの売上データ階層エントリの違いを、A ファームのエントリに対して **表示します**。
The `baseField` is **Farm,** so we see the differences between the other farms, as well as breakdowns for each type of like farm **(Type** is also a row hierarchy in this example).

!["A Farms" と他のファームとの間の青果売上の違いを示すピボットテーブル。 これは、ファームの総青果売上の違いと、青果の種類の売り上げの両方を示しています。 "A Farms" が特定の種類のファームを販売しなかった場合は、"#N/A" が表示されます。](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="change-hierarchy-names"></a>階層名を変更する

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
- [Excel JavaScript API リファレンス](/javascript/api/excel)
