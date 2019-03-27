---
title: Excel JavaScript API を使用してピボットテーブルを操作する
description: Excel JavaScript API を使用して、ピボットテーブルを作成し、それらのコンポーネントを操作します。
ms.date: 03/21/2019
localization_priority: Normal
ms.openlocfilehash: b53d734e676417a6438f1008bac720a38a244d1f
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870325"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してピボットテーブルを操作する

ピボットテーブルは、より大きなデータセットを合理化します。 グループ化されたデータのクイック操作を可能にします。 Excel JavaScript API を使用すると、アドインでピボットテーブルを作成し、それらのコンポーネントを操作できます。

ピボットテーブルの機能についてよく知らない場合は、エンドユーザーとしての調査を検討してください。 これらのツールの詳細については、「[ワークシートデータを分析するためのピボットテーブルを作成する](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables)」を参照してください。 

この記事では、一般的なシナリオのコードサンプルを示します。 ピボットテーブル API について理解するには、「 [**pivottable**](/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](/javascript/api/excel/excel.pivottable)」を参照してください。

> [!IMPORTANT]
> OLAP を使用して作成されたピボットテーブルは現在サポートされていません。 Power Pivot もサポートされていません。

## <a name="hierarchies"></a>Hierarchies

ピボットテーブルは、行、列、データ、およびフィルターの4つの階層カテゴリに基づいて編成されます。 この記事では、さまざまなファームからの果物 sales について説明する次のデータが使用されます。

![さまざまなファームからのさまざまな種類の果物販売のコレクション。](../images/excel-pivots-raw-data.png)

このデータには、**畑**、 **Type**、**分類**、 **Crates で販売**されたファーム、 **Crates 販売**された卸売の5つの階層があります。 各階層は、4つのカテゴリのいずれかにのみ存在できます。 **Type**が列階層に追加されてから、行階層に追加されても、後者には残ります。

行と列の階層は、データをグループ化する方法を定義します。 たとえば、**ファーム**の行階層は、同じファームのすべてのデータセットをグループ化します。 行と列の階層を選択すると、ピボットテーブルの向きが定義されます。

データ階層は、行と列の階層に基づいて集計される値です。 ファームの行階層があり、 **** **Crates**のデータ階層があるピボットテーブルには、各ファームのすべての異なる fruits の合計 (既定では) が表示されます。

フィルター階層では、フィルター処理された種類の値に基づいて、ピボットのデータが含まれるか、除外されます。 **有機**的に選択された種類の**分類**のフィルター階層は、有機フルーツのデータのみを表示します。

次に、ファームデータをピボットテーブルと共に示します。 ピボットテーブルでは、**ファーム**と**タイプ**を行階層として使用し、**ファームで販売**された Crates と Crates がデータ階層として**卸売販売**され、フィルターとして**分類**されています。階層 (**有機**が選択されている)。 

![行、データ、およびフィルター階層を使用したピボットテーブルの横の、果物 sales データの選択。](../images/excel-pivot-table-and-data.png)

このピボットテーブルは、JavaScript API または Excel UI を使用して生成できます。 両方のオプションを使用すると、アドインをさらに操作できます。

## <a name="create-a-pivottable"></a>ピボットテーブルを作成する

ピボットテーブルには、名前、ソース、および出力先が必要です。 ソースは、範囲内のアドレスまたはテーブル名 (、、 `Range`、 `string`または`Table`型として渡されます) を指定できます。 宛先は、または`Range` `string`のいずれかとして指定された範囲のアドレスです。 次のサンプルは、さまざまなピボットテーブル作成手法を示しています。

### <a name="create-a-pivottable-with-range-addresses"></a>範囲のアドレスを使用してピボットテーブルを作成する

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a>Range オブジェクトを使用してピボットテーブルを作成する

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data comes from the worksheet "DataWorksheet" across the range A1:E21
    const rangeToAnalyze = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
    const rangeToPlacePivot = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
    context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add(
        "Farm Sales", rangeToAnalyze, rangeToPlacePivot);

    await context.sync();
});
```

### <a name="create-a-pivottable-at-the-workbook-level"></a>ブックレベルでピボットテーブルを作成する

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a>既存のピボットテーブルを使用する

手動で作成したピボットテーブルは、ブックまたは個々のワークシートの PivotTable コレクションからアクセスすることもできます。 

次のコードは、ブック内の最初のピボットテーブルを取得します。 その後、表の名前を後で簡単に参照できるようにします。

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>ピボットテーブルに行と列を追加する

行と列は、それらのフィールド値を中心にデータをピボットします。

[**ファーム**] 列を追加すると、各ファームのすべての売上が回転します。 **種類**と**分類**行を追加すると、果物が販売されたものと、それが有機であったかどうかに基づいてデータがさらに分解されます。

![ファーム列と種類と分類行を含む PivotTable。](../images/excel-pivots-table-rows-and-columns.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    await context.sync();
});
```

行または列だけのピボットテーブルを作成することもできます。

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a>データ階層をピボットテーブルに追加する

データ階層は、行と列に基づいて結合する情報で、ピボットテーブルに格納されます。 **ファームで販売**された Crates のデータ階層を追加し、 **Crates に販売**されたものは、行と列ごとにこれらの数値を合計します。 

この例では、**ファーム**と**種類**の両方が行で、箱 sales がデータとして含まれています。 

![元のファームに基づいたさまざまな果物の総売上高を示すピボットテーブル。](../images/excel-pivots-data-hierarchy.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // "Farm" and "Type" are the hierarchies on which the aggregation is based
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

    // "Crates Sold at Farm" and "Crates Sold Wholesale" are the hierarchies
    // that will have their data aggregated (summed in this case)
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    await context.sync();
});
```

## <a name="change-aggregation-function"></a>集計関数を変更する

データ階層の値が集計されます。 数値のデータセットの場合は、既定でこれが合計になります。 この`summarizeBy`プロパティは、この動作を[集約 ationfunction](/javascript/api/excel/excel.aggregationfunction)型に基づいて定義します。

現在サポートされている集計`Sum`関数`Count`の`Average`種類`Max`は`Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP`、、、、、、、、 `Automatic` 、、、、および (既定値) です。

次のコードサンプルでは、集計をデータの平均値に変更します。

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.dataHierarchies.load("no-properties-needed");
    await context.sync();

    // changing the aggregation from the default sum to an average of all the values in the hierarchy
    pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.average;
    pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.average;
    await context.sync();
});
```

## <a name="change-calculations-with-a-showasrule"></a>showasrule を使用して計算を変更する

既定では、ピボットテーブルでは、行と列の階層のデータが個別に集計されます。 [showasrule](/javascript/api/excel/excel.showasrule)は、データ階層を、ピボットテーブル内の他のアイテムに基づいて出力値に変更します。

オブジェクト`ShowAsRule`には、次の3つのプロパティがあります。

-   `calculation`: データ階層に適用する相対的な計算の種類 (既定値は`none`)。
-   `baseField`: 計算を適用する前に、基本データを含む階層内のフィールド。 通常、[ピボットフィールド](/javascript/api/excel/excel.pivotfield)の名前は親階層と同じです。
-   `baseItem`: 計算の種類に基づいて、基準フィールドの値と比較した個々の[ピボット](/javascript/api/excel/excel.pivotitem)テーブル。 すべての計算にこのフィールドが必要なわけではありません。

次の例では、ファームデータ階層で販売された**Crates の合計**の計算を列の合計のパーセンテージに設定します。 さらに、粒度をフルーツの種類レベルにまで拡張する必要があるので、 **type**行階層とその基になるフィールドを使用します。 この例は、最初の行階層としても**ファーム**を持っているので、farm total エントリには各ファームがそれぞれを生成する割合が表示されます。

![各ファーム内の個々のファームと個々の果物の種類の総計を基準とした果物 sales の割合を示すピボットテーブル。](../images/excel-pivots-showas-percentage.png)

``` TypeScript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    const farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    await context.sync();

    // show the crates of each fruit type sold at the farm as a percentage of the column's total
    let farmShowAs = farmDataHierarchy.showAs;
    farmShowAs.calculation = Excel.ShowAsCalculation.percentOfColumnTotal;
    farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");
    farmDataHierarchy.showAs = farmShowAs; 
    farmDataHierarchy.name = "Percentage of Total Farm Sales";

    await context.sync();
});
```

前の例では、列の各行階層を基準にして計算を設定しています。 計算が個々のアイテムに関連している場合`baseItem`は、プロパティを使用します。

次の例は、 `differenceFrom`計算を示しています。 この例では、"ファーム" と比較した場合の、ファームの箱売上データ階層エントリの違いを示します。
`baseField`は**ファーム**ですので、他のファームの違いと、果物などの種類ごとの内訳 (この例では、**type**が行階層になっています) を確認しています。

!["畑" とその他の果物の売上の違いを示すピボットテーブル。 これは、畑の総売上合計と果物の種類の売上の違いを示しています。 "畑" が特定の種類の果物を販売していない場合は、"#N/a" が表示されます。](../images/excel-pivots-showas-differencefrom.png)

``` TypeScript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    const farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    await context.sync();

    // show the difference between crate sales of the "A Farms" and the other farms
    // this difference is both aggregated and shown for individual fruit types (where applicable)
    let farmShowAs = farmDataHierarchy.showAs;
    farmShowAs.calculation = Excel.ShowAsCalculation.differenceFrom;
    farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");
    farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");
    farmDataHierarchy.showAs = farmShowAs;
    farmDataHierarchy.name = "Difference from A Farms";
    await context.sync();
});
```

## <a name="pivottable-layouts"></a>ピボットテーブルのレイアウト

[PivotLayout](/javascript/api/excel/excel.pivotlayout)は、階層とそのデータの配置を定義します。 レイアウトにアクセスして、データが格納される範囲を決定します。

次の図は、ピボットテーブルの範囲に対応するどの layout 関数呼び出しを示しています。

![レイアウトの範囲取得機能によって返されるピボットテーブルのセクションを示す図。](../images/excel-pivots-layout-breakdown.png)

次のコードは、レイアウトを使用してピボットテーブルデータの最後の行を取得する方法を示しています。 これらの値は総計に対して合計されます。

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // get the totals for each data hierarchy from the layout
    const range = pivotTable.layout.getDataBodyRange();
    const grandTotalRange = range.getLastRow();
    grandTotalRange.load("address");
    await context.sync();

    // sum the totals from the PivotTable data hierarchies and place them in a new range
    const masterTotalRange = context.workbook.worksheets.getActiveWorksheet().getRange("B27:C27");
    masterTotalRange.formulas = [["All Crates", "=SUM(" + grandTotalRange.address + ")"]];
    await context.sync();
});
```

ピボットテーブルには、コンパクト、アウトライン、表形式という3つのレイアウトスタイルがあります。 前の例ではコンパクトなスタイルを見てきました。 

次の例では、アウトラインスタイルと表形式スタイルをそれぞれ使用します。 このコードサンプルは、さまざまなレイアウト間で循環する方法を示しています。

### <a name="outline-layout"></a>アウトラインレイアウト

![アウトラインレイアウトを使用したピボットテーブル。](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a>表形式レイアウト

![表形式レイアウトを使用したピボットテーブル。](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a>階層名を変更する

階層フィールドは編集できます。 次のコードは、2つのデータ階層の表示名を変更する方法を示しています。

```typescript
await Excel.run(async (context) => {
    const dataHierarchies = context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.getItem("Farm Sales").dataHierarchies;
    dataHierarchies.load("no-properties-needed");
    await context.sync();

    // changing the displayed names of these entries
    dataHierarchies.items[0].name = "Farm Sales";
    dataHierarchies.items[1].name = "Wholesale";
    await context.sync();
});
```

## <a name="delete-a-pivottable"></a>ピボットテーブルを削除する

ピボットテーブルは、名前を使用して削除されます。

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a>関連項目

- [Excel JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)
- [Excel JavaScript API リファレンス](/javascript/api/excel)
