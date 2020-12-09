---
title: Excel JavaScript API を使用してピボットテーブルを操作する
description: Excel JavaScript API を使用して、ピボットテーブルを作成し、それらのコンポーネントを操作します。
ms.date: 12/07/2020
localization_priority: Normal
ms.openlocfilehash: 0a1fefa6a855ab9ee1ccd71fd0dc60f282d2944b
ms.sourcegitcommit: fecad2afa7938d7178456c11ba52b558224813b4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/09/2020
ms.locfileid: "49603800"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してピボットテーブルを操作する

ピボットテーブルは、より大きなデータセットを合理化します。 グループ化されたデータのクイック操作を可能にします。 Excel JavaScript API を使用すると、アドインでピボットテーブルを作成し、それらのコンポーネントを操作できます。 この記事では、Office JavaScript API によってピボットテーブルがどのように表現されるかについて説明し、主要なシナリオのコードサンプルを示します。

ピボットテーブルの機能についてよく知らない場合は、エンドユーザーとしての調査を検討してください。
これらのツールの詳細については、「 [ワークシートデータを分析するためのピボットテーブルを作成する](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) 」を参照してください。

> [!IMPORTANT]
> OLAP を使用して作成されたピボットテーブルは現在サポートされていません。 Power Pivot もサポートされていません。

## <a name="object-model"></a>オブジェクト モデル

[PivotTable](/javascript/api/excel/excel.pivottable)は、OFFICE JavaScript API のピボットテーブルの中心的なオブジェクトです。

- `Workbook.pivotTables`および `Worksheet.pivotTables` は、ブックとワークシートの[ピボットテーブル](/javascript/api/excel/excel.pivottable)をそれぞれ含む[PivotTableCollections](/javascript/api/excel/excel.pivottablecollection)です。
- [ピボットテーブル](/javascript/api/excel/excel.pivottable)に、複数の[PivotHierarchies](/javascript/api/excel/excel.pivothierarchy)を持つ[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)が含まれています。
- これらの [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy) は、 [次のセクション](#hierarchies)で説明するように、PivotTable がデータをピボットする方法を定義するために、特定の階層コレクションに追加できます。
- [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)には、1つだけの[PivotField](/javascript/api/excel/excel.pivotfield)を持つ[pivotfieldcollection](/javascript/api/excel/excel.pivotfieldcollection)が含まれています。 デザインを拡張して OLAP ピボットテーブルが含まれる場合は、これが変更されることがあります。
- [PivotField](/javascript/api/excel/excel.pivotfield)には、フィールドの[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)が階層カテゴリに割り当てられている限り、1つまたは複数の[PivotFilters](/javascript/api/excel/excel.pivotfilters)を適用できます。 
- [PivotField](/javascript/api/excel/excel.pivotfield)には、複数の[PivotItems](/javascript/api/excel/excel.pivotitem)を持つ[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)が含まれています。
- [ピボットテーブル](/javascript/api/excel/excel.pivottable)には、ピボット[フィールド](/javascript/api/excel/excel.pivotfield)と[PivotItems](/javascript/api/excel/excel.pivotitem)をワークシートのどこに表示するかを定義する[PivotLayout](/javascript/api/excel/excel.pivotlayout)が含まれています。

これらの関係がいくつかの例のデータにどのように適用されるかを見てみましょう。 次のデータは、さまざまなファームからの果物販売を示しています。 この記事全体の例を示します。

![さまざまなファームからのさまざまな種類の果物販売のコレクション。](../images/excel-pivots-raw-data.png)

この果物 farm sales データは、ピボットテーブルを作成するために使用されます。 **Types** などの各列は、 `PivotHierarchy` です。 **種類** の階層には、[**種類**] フィールドが含まれています。 [ **種類** ] フィールドには、 **Apple**、 **Kiwi**、 **レモン**、 **黄**、 **オレンジ色** の項目が含まれています。

### <a name="hierarchies"></a>Hierarchies

ピボットテーブルは、 [行](/javascript/api/excel/excel.rowcolumnpivothierarchy)、 [列](/javascript/api/excel/excel.rowcolumnpivothierarchy)、 [データ](/javascript/api/excel/excel.datapivothierarchy)、および [フィルター](/javascript/api/excel/excel.filterpivothierarchy)の4つの階層カテゴリに基づいて編成されます。

前に示したファームデータには、ファーム、**種類**、**分類**、 **Crates で販売** されたファーム、 **Crates 販売** された卸売の5つの階層が **あります。** 各階層は、4つのカテゴリのいずれかにのみ存在できます。 **型** が列階層に追加されている場合は、行、データ、またはフィルター階層に配置することもできません。 その後、 **型** が行階層に追加されると、列階層から削除されます。 この動作は、階層の割り当てが Excel UI または Excel JavaScript Api のどちらで行われた場合でも同じです。

行と列の階層は、データをグループ化する方法を定義します。 たとえば、 **ファーム** の行階層は、同じファームのすべてのデータセットをグループ化します。 行と列の階層を選択すると、ピボットテーブルの向きが定義されます。

データ階層は、行と列の階層に基づいて集計される値です。 **ファームの行** 階層があり、 **Crates** のデータ階層があるピボットテーブルには、各ファームのすべての異なる fruits の合計 (既定では) が表示されます。

フィルター階層では、フィルター処理された種類の値に基づいて、ピボットのデータが含まれるか、除外されます。 **有機** 的に選択された種類の **分類** のフィルター階層は、有機フルーツのデータのみを表示します。

次に、ファームデータをピボットテーブルと共に示します。 ピボットテーブルは、**ファーム** と **タイプ** を行階層として使用し、**ファームで販売** された Crates と **Crates** がデータ階層として (既定の集計関数を使用して)、データ階層として (**有機** が選択された状態で)**分類** しています。

![行、データ、およびフィルター階層を使用したピボットテーブルの横の、果物 sales データの選択。](../images/excel-pivot-table-and-data.png)

このピボットテーブルは、JavaScript API または Excel UI を使用して生成できます。 両方のオプションを使用すると、アドインをさらに操作できます。

## <a name="create-a-pivottable"></a>ピボットテーブルを作成する

ピボットテーブルには、名前、ソース、および出力先が必要です。 ソースは、範囲内のアドレスまたはテーブル名 (、、、または型として渡さ `Range` `string` `Table` れます) を指定できます。 宛先は、またはのいずれかとして指定された範囲のアドレスです `Range` `string` 。
次のサンプルは、さまざまなピボットテーブル作成手法を示しています。

### <a name="create-a-pivottable-with-range-addresses"></a>範囲のアドレスを使用してピボットテーブルを作成する

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

### <a name="create-a-pivottable-at-the-workbook-level"></a>ブックレベルでピボットテーブルを作成する

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

手動で作成したピボットテーブルは、ブックまたは個々のワークシートの PivotTable コレクションからアクセスすることもできます。 次のコードは、ブックから **My Pivot** という名前のピボットテーブルを取得します。

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>ピボットテーブルに行と列を追加する

行と列は、それらのフィールド値を中心にデータをピボットします。

[ **ファーム** ] 列を追加すると、各ファームのすべての売上が回転します。 **種類** と **分類** 行を追加すると、果物が販売されたものと、それが有機であったかどうかに基づいてデータがさらに分解されます。

![ファーム列と種類と分類行を含む PivotTable。](../images/excel-pivots-table-rows-and-columns.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    return context.sync();
});
```

行または列だけのピボットテーブルを作成することもできます。

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a>データ階層をピボットテーブルに追加する

データ階層は、行と列に基づいて結合する情報で、ピボットテーブルに格納されます。 **ファームで販売** された Crates のデータ階層を追加し、 **Crates に販売** されたものは、行と列ごとにこれらの数値を合計します。

この例では、 **ファーム** と **種類** の両方が行で、箱 sales がデータとして含まれています。

![元のファームに基づいたさまざまな果物の総売上高を示すピボットテーブル。](../images/excel-pivots-data-hierarchy.png)

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

## <a name="pivottable-layouts-and-getting-pivoted-data"></a>ピボットテーブルのレイアウトとピボットデータの取得

[PivotLayout](/javascript/api/excel/excel.pivotlayout)は、階層とそのデータの配置を定義します。 レイアウトにアクセスして、データが格納される範囲を決定します。

次の図は、ピボットテーブルの範囲に対応するどの layout 関数呼び出しを示しています。

![レイアウトの範囲取得機能によって返されるピボットテーブルのセクションを示す図。](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a>ピボットテーブルからデータを取得する

レイアウトは、ピボットテーブルをワークシートに表示する方法を定義します。 これは、 `PivotLayout` オブジェクトがピボットテーブル要素で使用される範囲を制御することを意味します。 レイアウトによって提供される範囲を使用して、ピボットテーブルによって収集および集計されるデータを取得します。 特に、を使用し `PivotLayout.getDataBodyRange` て、ピボットテーブルによって生成されるものにアクセスします。

次のコードでは、レイアウト (**ファームで販売される Crates の合計** と、前の例で **Crates に販売** された卸売列の **合計の両方**) によって、ピボットテーブルデータの最後の行を取得する方法を示します。 これらの値は、最終的な合計として集計され、セル **E30** (ピボットテーブルの外側) に表示されます。

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

ピボットテーブルには、コンパクト、アウトライン、表形式という3つのレイアウトスタイルがあります。 前の例ではコンパクトなスタイルを見てきました。

次の例では、アウトラインスタイルと表形式スタイルをそれぞれ使用します。 このコードサンプルは、さまざまなレイアウト間で循環する方法を示しています。

#### <a name="outline-layout"></a>アウトラインレイアウト

![アウトラインレイアウトを使用したピボットテーブル。](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a>表形式レイアウト

![表形式レイアウトを使用したピボットテーブル。](../images/excel-pivots-tabular-layout.png)

## <a name="delete-a-pivottable"></a>ピボットテーブルを削除する

ピボットテーブルは、名前を使用して削除されます。

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="filter-a-pivottable"></a>ピボットテーブルのフィルター処理

ピボットテーブルのデータをフィルター処理するための主な方法は、PivotFilters を使用する方法です。 スライサーは、柔軟な代替のフィルター方法を提供します。 

[PivotFilters](/javascript/api/excel/excel.pivotfilters) は、ピボットテーブルの4つの [階層カテゴリ](#hierarchies) (フィルター、列、行、および値) に基づいてデータをフィルター処理します。 PivotFilters には4つの種類があり、カレンダーの日付に基づくフィルター処理、文字列解析、数字比較、およびカスタム入力に基づくフィルター処理を行うことができます。 

[スライサー](/javascript/api/excel/excel.slicer) は、ピボットテーブルと通常の Excel テーブルの両方に適用できます。 ピボットテーブル (PivotTable) に適用すると、 [PivotManualFilter](#pivotmanualfilter) のように機能し、カスタム入力に基づいてフィルターを適用することができます。 PivotFilters とは異なり、スライサーには [EXCEL UI コンポーネント](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)があります。 クラスを使用して、 `Slicer` この UI コンポーネントを作成し、フィルター処理を管理して、視覚的な外観を制御します。 

### <a name="filter-with-pivotfilters"></a>PivotFilters を使用してフィルターを適用する

[PivotFilters](/javascript/api/excel/excel.pivotfilters) では、4つの [階層カテゴリ](#hierarchies) (フィルター、列、行、および値) に基づいてピボットテーブルデータをフィルターできます。 PivotTable オブジェクトモデルでは、 `PivotFilters` [PivotField](/javascript/api/excel/excel.pivotfield)に適用され、それぞれに `PivotField` 1 つ以上の割り当てることができ `PivotFilters` ます。 PivotFilters を PivotField に適用するには、フィールドに対応する [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) を階層カテゴリに割り当てる必要があります。 

#### <a name="types-of-pivotfilters"></a>PivotFilters の種類

| フィルターの種類 | フィルターの目的 | Excel JavaScript API リファレンス |
|:--- |:--- |:--- |
| DateFilter | カレンダーの日付ベースのフィルター処理。 | [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) |
| LabelFilter | テキスト比較フィルター処理。 | [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) |
| ManualFilter | カスタム入力フィルター。 | [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) |
| ValueFilter | 数値比較フィルター処理。 | [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter) |

#### <a name="create-a-pivotfilter"></a>PivotFilter を作成する

ピボットテーブルの * フィルター (PivotDateFilter など) を使用してピボットテーブルデータをフィルター処理するには、ピボット [フィールド](/javascript/api/excel/excel.pivotfield)にフィルターを適用します。 次の4つのコードサンプルは、4種類の PivotFilters を使用する方法を示しています。 

##### <a name="pivotdatefilter"></a>PivotDateFilter

最初のコード例では、**更新さ** れた PivotField 日付に [pivotdatefilter](/javascript/api/excel/excel.pivotdatefilter)を適用し、 **2020-08-01** より前のデータを非表示にします。 

> [!IMPORTANT] 
> ピボット * フィルターは、そのフィールドの PivotHierarchy が階層カテゴリに割り当てられていない限り、PivotField に適用できません。 次のコードサンプルでは、を `dateHierarchy` ピボットテーブルのカテゴリに追加してから、 `rowHierarchies` フィルター処理に使用できるようにする必要があります。

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
> 次の3つのコードスニペットでは、完全な呼び出しではなく、フィルター固有の抜粋のみが表示され `Excel.run` ます。

##### <a name="pivotlabelfilter"></a>PivotLabelFilter

2番目のコードスニペットでは、プロパティを使用して **Type** 、文字 L で始まるラベルを除外することにより、 [Pivotlabelfilter](/javascript/api/excel/excel.pivotlabelfilter)を型の PivotField に適用する方法を示します `LabelFilterCondition.beginsWith` **L**。 

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

3番目のコードスニペットは、 [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) を使用した手動フィルターを **分類** フィールドに適用し、その分類の **有機** を含まないデータをフィルターで除外します。 

```js
    // Apply a manual filter to include only a specific PivotItem (the string "Organic").
    var filterField = classHierarchy.fields.getItem("Classification");
    var manualFilter = { selectedItems: ["Organic"] };
    filterField.applyFilter({ manualFilter: manualFilter });
```

##### <a name="pivotvaluefilter"></a>PivotValueFilter

数値を比較するには、最後のコードスニペットに示されているように、 [Pivotvaluefilter](/javascript/api/excel/excel.pivotvaluefilter)で値フィルターを使用します。 は、 `PivotValueFilter` **ファーム** のピボットテーブル内のデータと **Crates 販売** された卸売のピボットフィールドのデータを比較します。これには、Crates の合計が **500** を超えるファームのみが含まれます。 

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

#### <a name="remove-pivotfilters"></a>PivotFilters の削除

すべての PivotFilters を削除するには、 `clearAllFilters` 次のコードサンプルに示すように、各 PivotField にメソッドを適用します。 

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

### <a name="filter-with-slicers"></a>スライサーを使用してフィルターを適用する

[スライサー](/javascript/api/excel/excel.slicer) を使用すると、Excel のピボットテーブルまたはテーブルからデータをフィルターできます。 スライサーは、指定された列またはピボットテーブルの値を使用して、対応する行にフィルターを適用します。 これらの値は、 [SlicerItem](/javascript/api/excel/excel.sliceritem) オブジェクトとしてに格納され `Slicer` ます。 アドインでは、ユーザーと同様に ([EXCEL UI を介し](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)て) これらのフィルターを調整できます。 スライサーは、次のスクリーンショットに示されているように、描画層のワークシートの一番上にあります。

![ピボットテーブルのスライサーフィルターデータ。](../images/excel-slicer.png)

> [!NOTE]
> このセクションで説明する手法は、ピボットテーブルに接続されたスライサーの使用方法に重点を置いています。 テーブルに接続されたスライサーを使用する場合にも同じ方法が適用されます。

#### <a name="create-a-slicer"></a>スライサーを作成する

メソッドまたはメソッドを使用して、ブックまたはワークシートにスライサーを作成でき `Workbook.slicers.add` `Worksheet.slicers.add` ます。 これにより、指定したオブジェクトまたはオブジェクトの [SlicerCollection](/javascript/api/excel/excel.slicercollection) にスライサーが追加され `Workbook` `Worksheet` ます。 `SlicerCollection.add`メソッドには、次の3つのパラメーターがあります。

- `slicerSource`: 新しいスライサーの基になるデータソース。 、、または `PivotTable` `Table` の名前または ID を表す文字列を指定でき `PivotTable` `Table` ます。
- `sourceField`: フィルター処理の対象となるデータソース内のフィールド。 、、または `PivotField` `TableColumn` の名前または ID を表す文字列を指定でき `PivotField` `TableColumn` ます。
- `slicerDestination`: 新しいスライサーを作成するワークシートを指定します。 オブジェクト、または `Worksheet` の名前または ID を指定でき `Worksheet` ます。 を経由してアクセスする場合、このパラメーターは必要あり `SlicerCollection` `Worksheet.slicers` ません。 この例では、コレクションのワークシートがコピー先として使用されます。

次のコードサンプルでは、新しいスライサーを **ピボット** ワークシートに追加します。 スライサーのソースは、 **ファームの売上** ピボットテーブルで、 **型** データを使用してフィルター処理されます。 スライサーは、後で参照するために **フルーツスライサー** という名前も付けられています。

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

#### <a name="filter-items-with-a-slicer"></a>スライサーを使用してアイテムをフィルターにかける

スライサーは、からのアイテムを使用して、ピボットテーブルをフィルターし `sourceField` ます。 この `Slicer.selectItems` メソッドは、スライサーに残っているアイテムを設定します。 これらのアイテムは、アイテムのキーを表すとしてメソッドに渡され `string[]` ます。 これらのアイテムを含む行は、ピボットテーブルの集計に残ります。 以降の呼び出し `selectItems` では、これらの呼び出しで指定されたキーにリストを設定します。

> [!NOTE]
> `Slicer.selectItems`データソースに含まれていないアイテムが渡されると、 `InvalidArgument` エラーがスローされます。 このプロパティを使用して、 `Slicer.slicerItems` [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)の内容を確認できます。

次のコードサンプルでは、スライサーに対して選択されている3つのアイテム ( **レモン**、 **黄**、 **オレンジ色**) を示します。

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

スライサーからすべてのフィルターを削除するには、 `Slicer.clearFilters` 次の例に示すようにメソッドを使用します。

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

#### <a name="style-and-format-a-slicer"></a>スライサーのスタイルと書式設定

アドインでは、プロパティを使用してスライサーの表示設定を調整でき `Slicer` ます。 次のコードサンプルでは、スタイルを **SlicerStyleLight6** に設定し、スライサーの上部のテキストを **果物の種類** に設定し、スライサーを描画レイヤーの位置 **(395, 15)** に配置し、スライサーのサイズを **135x150** ピクセルに設定します。

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

スライサーを削除するには、 `Slicer.delete` メソッドを呼び出します。 次のコードサンプルでは、現在のワークシートから最初のスライサーを削除します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a>集計関数を変更する

データ階層の値が集計されます。 数値のデータセットの場合は、既定でこれが合計になります。 `summarizeBy`このプロパティは、この動作を[集約 ationfunction](/javascript/api/excel/excel.aggregationfunction)型に基づいて定義します。

現在サポートされている集計関数の種類は、、、、、、、、、、、、 `Sum` `Count` `Average` `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP` および `Automatic` (既定値) です。

次のコードサンプルでは、集計をデータの平均値に変更します。

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

既定では、ピボットテーブルでは、行と列の階層のデータが個別に集計されます。 [Showasrule](/javascript/api/excel/excel.showasrule)は、データ階層を、ピボットテーブル内の他のアイテムに基づいて出力値に変更します。

`ShowAsRule`オブジェクトには、次の3つのプロパティがあります。

- `calculation`: データ階層に適用する相対的な計算の種類 (既定値は `none` )。
- `baseField`: 計算を適用する前に、基本データを含む階層の [PivotField](/javascript/api/excel/excel.pivotfield) 。 Excel ピボットテーブルには、階層とフィールドの間に一対一のマッピングがあるため、階層とフィールドの両方にアクセスするには同じ名前を使用します。
- `baseItem`: 計算の種類に基づいて、基準フィールドの値と比較した個々の [ピボット](/javascript/api/excel/excel.pivotitem) テーブル。 すべての計算にこのフィールドが必要なわけではありません。

次の例では、ファームデータ階層で販売された **Crates の合計** の計算を列の合計のパーセンテージに設定します。
さらに、粒度をフルーツの種類レベルにまで拡張する必要があるので、 **type** 行階層とその基になるフィールドを使用します。
この例は、最初の行階層としても **ファーム** を持っているので、farm total エントリには各ファームがそれぞれを生成する割合が表示されます。

![各ファーム内の個々のファームと個々の果物の種類の総計を基準とした果物 sales の割合を示すピボットテーブル。](../images/excel-pivots-showas-percentage.png)

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

前の例では、列に対して、個々の行階層のフィールドを基準にして計算を行います。 計算が個々のアイテムに関連している場合は、プロパティを使用し `baseItem` ます。

次の例は、計算を示して `differenceFrom` います。 ファームの箱売上データ階層エントリの違いが **、ファームの** ものと比較して表示されます。
`baseField`は **ファーム** ですので、他のファームの違いと、果物などの種類ごとの内訳 (この例では、**type** が行階層になっています) を確認しています。

!["畑" とその他の果物の売上の違いを示すピボットテーブル。 これは、畑の総売上合計と果物の種類の売上の違いを示しています。 "畑" が特定の種類の果物を販売していない場合は、"#N/A" が表示されます。](../images/excel-pivots-showas-differencefrom.png)

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

階層フィールドは編集できます。 次のコードは、2つのデータ階層の表示名を変更する方法を示しています。

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

- [Office アドインでの Excel JavaScript オブジェクトモデル](excel-add-ins-core-concepts.md)
- [Excel JavaScript API リファレンス](/javascript/api/excel)
