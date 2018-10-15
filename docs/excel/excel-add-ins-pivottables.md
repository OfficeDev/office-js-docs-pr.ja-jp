---
title: Excel の JavaScript API を使用してピボット テーブルで作業します
description: Excel JavaScript API を使用してピボットテーブルを作成し、そのコンポーネントと対話します。
ms.date: 09/21/2018
ms.openlocfilehash: 00dd982d4ba4de0db34277cd546b572d4394e258
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459281"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="7e66c-103">Excel の JavaScript API を使用してピボット テーブルで作業します</span><span class="sxs-lookup"><span data-stu-id="7e66c-103">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="7e66c-p101">ピボット テーブルより大きなデータ セットを効率化します。このことにより、グループ化されたデータのクイック操作が可能になります。Excel の JavaScript API では、アドインにピボット テーブルを作成させ、それらのコンポーネントと対話することができます。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p101">PivotTables streamline larger data sets. They allow the quick manipulation of grouped data. The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> 

<span data-ttu-id="7e66c-p102">ピボット テーブルの機能に慣れていない場合は、エンド ユーザーとしてこれらの調査を検討してください。これらのツールの適切な入門書には、 [ワークシートのデータを分析するピボット テーブルの作成](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p102">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user. See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span> 

<span data-ttu-id="7e66c-p103">この資料では、一般的なシナリオのコード サンプルを提供します。ピボットテーブル API の理解をさらに深めるには、 [**PivotTable**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) と [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p103">This article provides code samples for common scenarios. To further your understanding of the PivotTable API, see [**PivotTable**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7e66c-111">OLAP で作成されたピボット テーブルは、現在サポートされていません。</span><span class="sxs-lookup"><span data-stu-id="7e66c-111">PivotTables created with OLAP are not currently supported.</span></span>

## <a name="hierarchies"></a><span data-ttu-id="7e66c-112">階層</span><span class="sxs-lookup"><span data-stu-id="7e66c-112">Hierarchies</span></span>

<span data-ttu-id="7e66c-p104">ピボット テーブルは、行、列、データ、フィルターの 4 つの階層カテゴリに基づいて構成されています。この記事全体を通して、さまざまな農場の果物の売り上げを記述した次のデータを使用します。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p104">PivotTables are organized based on four hierarchy categories: row, column, data, and filter. The following data describing fruit sales from various farms will be used throughout this article.</span></span>

![さまざまな農場における、多種の果物の売り上げのコレクション。](../images/excel-pivots-raw-data.png)

<span data-ttu-id="7e66c-p105">このデータには **農家**、 **種類**、 **分類**、**農場で販売された箱数**、および **卸売りで販売された箱数** の 5 つの階層があります。各階層は、 4 つのカテゴリのいずれかにのみ存在できます。 \*\* 種類\*\* が列の階層に追加され、さらに行の階層に追加された場合、行の階層にのみ残ります。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p105">This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**. Each hierarchy can only exist in one of the four categories. If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.</span></span>

<span data-ttu-id="7e66c-p106">行と列の階層は、データをグループ化する方法を定義します。たとえば、 **農場** の行の階層は、同じ農場のすべてのデータ セットをまとめてグループ化します。行と列の階層から選択すると、ピボット テーブルの向きが定義されます。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p106">Row and column hierarchies define how data will be grouped. For example, a row hierarchy of **Farms** will group together all the data sets from the same farm. The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="7e66c-p107">データ階層は、行と列の階層に基づいて集計する値です。**農場** の行の階層と **卸売りで販売された木箱** のデータ階層からなるピボット テーブルは、各農場のさまざまな種類の果物の総計 (既定) を示します。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p107">Data hierarchies are the values to be aggregated based on the row and column hierarchies. A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="7e66c-p108">フィルター階層は、フィルターされた種類の中の値に基づいてピボットにデータを取り込むか、取り除きます。**分類** のフィルター階層で **有機栽培** を選択すると、有機栽培の果物のデータのみが表示されます。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p108">Filter hierarchies include or exclude data from the pivot based on values within that filtered type. A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="7e66c-p109">こちらにもまた、ピボット テーブルを添えた農場のデータがあります。ピボット テーブルは、**農場** と **種類**を行階層、 **農場で販売された箱数** と**卸売りで販売された箱数** をデータ階層 (既定の合計の集計関数)、**分類**  をフィルター階層 (**有機栽培**を選択) として使用しています。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p109">Here is the farm data again, alongside a PivotTable. The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span> 

![行、データ、フィルターの階層で構成したピボット テーブルの次に果物の売り上げデータの選択範囲があります。](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="7e66c-p110">このピボット テーブルは、JavaScript API または Excel の UI で作られた可能性があります。両方のオプションで、アドインを通じ、さらに操作することができます。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p110">This PivotTable could be generated through the JavaScript API or through the Excel UI. Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="7e66c-131">ピボット テーブルの作成</span><span class="sxs-lookup"><span data-stu-id="7e66c-131">Create a PivotTable and PivotChart</span></span>

<span data-ttu-id="7e66c-p111">ピボット テーブルには、名前、ソース、および宛先を必要とします。ソースは、範囲アドレス、またはテーブル名を指定できます ( `Range`、 `string`、`Table` 型として渡されます)。同期先は、範囲アドレスです (`Range`  または `string`  のいずれかとして付与されます)。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p111">PivotTables need a name, source, and destination. The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type). The destination is a range address (given as either a `Range` or `string`). The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="7e66c-136">範囲アドレスを使用してピボット テーブルを作成</span><span class="sxs-lookup"><span data-stu-id="7e66c-136">Create a PivotTable with range addresses</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="7e66c-137">Range オブジェクトを使用してピボット テーブルを作成</span><span class="sxs-lookup"><span data-stu-id="7e66c-137">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="7e66c-138">ワークブック レベルでピボット テーブルを作成</span><span class="sxs-lookup"><span data-stu-id="7e66c-138">Create a PivotTable at the workbook level</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="7e66c-139">既存のピボット テーブルの使用</span><span class="sxs-lookup"><span data-stu-id="7e66c-139">Use an existing PivotTable</span></span>

<span data-ttu-id="7e66c-140">手動で作成したピボット テーブルも、ブックのピボット テーブルのコレクションまたはここのワークシートを使用してアクセス可能です。</span><span class="sxs-lookup"><span data-stu-id="7e66c-140">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> 

<span data-ttu-id="7e66c-p112">次のコードは、ブックに最初のピボットテーブルを追加します。以降に参照しやすくするため、テーブルに名前を付与します。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p112">The following code gets the first PivotTable in the workbook. It then gives the table a name for easy future reference.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="7e66c-143">ピボット テーブルに行と列を追加</span><span class="sxs-lookup"><span data-stu-id="7e66c-143">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="7e66c-144">行と列は、これらのフィールドの値の周りでデータをピボットします。</span><span class="sxs-lookup"><span data-stu-id="7e66c-144">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="7e66c-p113">**農場** 列を追加すると、各農場のすべての売り上げをピボットします。\*\* 種類\*\* と \*\* 分類\*\* 行を追加すると、どの果物が販売されたか、そしてそれが有機栽培かどうかに基づいて、データがさらに分解されます。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p113">Adding the **Farm** column pivots all the sales around each farm. Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

![農場の列、種類と、分類の行を含むピボット テーブル。](../images/excel-pivots-table-rows-and-columns.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    await context.sync();
});
```

<span data-ttu-id="7e66c-148">行または列のみを含むピボット テーブルも可能です。</span><span class="sxs-lookup"><span data-stu-id="7e66c-148">You can also have a PivotTable with only rows or columns.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="7e66c-149">ピボット テーブルへのデータ階層の追加</span><span class="sxs-lookup"><span data-stu-id="7e66c-149">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="7e66c-p114">データ階層は、行と列に基づいて組み合わせる情報でピボット テーブルを入力します。 **農場で販売された箱数** と **卸売りで販売された箱数** のデータ階層を追加すると、各行と列にそれらの数値の合計が表示されます。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p114">Data hierarchies fill the PivotTable with information to combine based on the rows and columns. Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span> 

<span data-ttu-id="7e66c-152">この例では、 **農場** と **種類** はともに行となり、箱の販売数をデータとして表示します。</span><span class="sxs-lookup"><span data-stu-id="7e66c-152">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span> 

![出荷された農場別に果物の総売り上げを示すピボット テーブル。](../images/excel-pivots-data-hierarchy.png)

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

## <a name="change-aggregation-function"></a><span data-ttu-id="7e66c-154">集計関数を変更する</span><span class="sxs-lookup"><span data-stu-id="7e66c-154">Change aggregation function</span></span>

<span data-ttu-id="7e66c-p115">データの階層の値は集計されています。数値のデータセットの場合、規定で和となります。`summarizeBy` プロパティは `AggregrationFunction`  タイプに基づいてこの動作を定義します。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p115">Data hierarchies have their values aggregated. For datasets of numbers, this is a sum by default. The `summarizeBy` property defines this behavior based on an `AggregrationFunction` type.</span></span> 

<span data-ttu-id="7e66c-158">現在サポートされている集計関数のタイプは、 `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`,  `Automatic` (既定値) です。</span><span class="sxs-lookup"><span data-stu-id="7e66c-158">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="7e66c-159">次のコード サンプルでは、データの平均値を使用する集計を変更します。</span><span class="sxs-lookup"><span data-stu-id="7e66c-159">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="7e66c-160">ShowAsRule を使用しての計算の変更</span><span class="sxs-lookup"><span data-stu-id="7e66c-160">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="7e66c-p116">規定でピボット テーブルは、行と列の階層のデータを個別に集計します。 `ShowAsRule` は、ピボット テーブル内の他の項目に基づいた値を出力するために、データの階層を変更します。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p116">PivotTables, by default, aggregate the data of their row and column hierarchies independently. A `ShowAsRule` changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="7e66c-163">`ShowAsRule` オブジェクトには次の 3 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="7e66c-163">The `ShowAsRule` object has three properties:</span></span>
-   <span data-ttu-id="7e66c-164">`calculation`: データの階層に適用する相対的な計算の種類 (既定値は `none`)。</span><span class="sxs-lookup"><span data-stu-id="7e66c-164">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
-   <span data-ttu-id="7e66c-p117">`baseField`: 計算が適用される前の基本データを含む階層内のフィールド。通常、`PivotField` は親階層と同じ名前です。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p117">`baseField`: The field within the hierarchy containing the base data before the calculation is applied. The `PivotField` usually has the same name as its parent hierarchy.</span></span>
-   <span data-ttu-id="7e66c-p118">`baseItem`:計算の種類に基づいた基本フィールドの値と比較した個々の項目。すべての計算にこのフィールドが必要なわけではありません。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p118">`baseItem`: The individual item compared against the values of the base fields based on the calculation type. Not all calculations require this field.</span></span>

<span data-ttu-id="7e66c-p119">次の例では、 **農場で販売された木箱の合計** のデータ階層の計算を、列合計のパーセント値に設定します。粒度を果物の種類レベルに拡張するため、 \*\* 種類\*\* の行の階層と基になるフィールドを使用するようにします。この例でも、最初の行の階層として **農場**  も示しているため、農場の合計エントリは、各農場の生産責任の割合も表示します。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p119">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total. We still want the granularity to extend to the fruit type level, so we’ll use the **Type** row hierarchy and its underlying field. The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

![個別の農場別、そして各農場内の果物別両方総計に関する、果物の売り上げ高のパーセント値を示すピボット テーブル。](../images/excel-pivots-showas-percentage.png)

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

<span data-ttu-id="7e66c-p120">前の例では、個別の行階層に関して、列に計算を設定します。計算が個々の項目に関連する場合は、 `baseItem` プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p120">The previous example set the calculation to the column, relative to an individual row hierarchy. When the calculation relates to an individual item, use the `baseItem` property.</span></span> 

<span data-ttu-id="7e66c-p121">次の例では、 `differenceFrom` 計算を示します。「A農場」に関する、農場で販売された木箱のデータ階層の入力値の差を表示します。`baseField`  は **農場**なので、各果物の種類のブレークダウン図形と同様に、他の農場間の差がわかります (この例では**種類** も行の階層) 。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p121">The following example shows the `differenceFrom` calculation. It displays the difference of the farm crate sales data hierarchy entries relative to those of “A Farms”. The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![「A 農場」と他の農場の果物売上高の差を表示するピボット テーブル。これは、農場の果物総売上高と種類別の果物販売高の、両方の差を示しています。「A 農場」で特定の種類の果物が販売されなかった場合、"#N/A"と表示されます。](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="pivottable-layouts"></a><span data-ttu-id="7e66c-181">ピボット テーブルのレイアウト</span><span class="sxs-lookup"><span data-stu-id="7e66c-181">PivotTable layouts</span></span>

<span data-ttu-id="7e66c-p123">ピボット テーブルのレイアウトは、階層とそのデータの配置を定義します。データが保存されている範囲を決定するためレイアウトにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p123">A PivotTable layout defines the placement of hierarchies and their data. You access the layout to determine the ranges where data is stored.</span></span> 

<span data-ttu-id="7e66c-184">次のダイアグラムは、どのレイアウト関数の呼び出しがピボット テーブルのどの範囲に対応しているか示しています。</span><span class="sxs-lookup"><span data-stu-id="7e66c-184">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![ピボット テーブルのどの部分がレイアウトの取得範囲の関数によって返されるかを示す図。](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="7e66c-p124">次のコードでは、レイアウトを使用するピボット テーブルのデータの最後の行を取得する方法を示します。これらの値は、総計用にまとめて集計されます。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p124">The following code demonstrates how to get the last row of the PivotTable data by going through the layout. Those values are then summed together for a grand total.</span></span>

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

<span data-ttu-id="7e66c-p125">ピボット テーブルには、コンパクト、アウトラインおよび表形式の3 つのレイアウトがあります。前の例はコンパクトスタイルです。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p125">PivotTables have three layout styles: Compact, Outline, and Tabular. We’ve seen the compact style in the previous examples.</span></span> 

<span data-ttu-id="7e66c-p126">次の例は、アウトライン、表形式のスタイルをそれぞれ使用します。コード サンプルでは、さまざまなレイアウトを交互に表示する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p126">The following examples use the outline and tabular styles, respectively. The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="7e66c-192">アウトライン レイアウト表示</span><span class="sxs-lookup"><span data-stu-id="7e66c-192">Outline layout</span></span>

![アウトライン表示のレイアウトを使用するピボットテーブル。](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="7e66c-194">表形式のレイアウト</span><span class="sxs-lookup"><span data-stu-id="7e66c-194">Tabular layout</span></span>

![表形式のレイアウトを使用するピボットテーブル。](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a><span data-ttu-id="7e66c-196">階層名の変更</span><span class="sxs-lookup"><span data-stu-id="7e66c-196">Change hierarchy names</span></span>

<span data-ttu-id="7e66c-p127">階層のフィールドは編集できます。次のコードでは、2 つのデータ階層の表示名をどのように変更するかを説明します。</span><span class="sxs-lookup"><span data-stu-id="7e66c-p127">Hierarchy fields are editable. The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="7e66c-199">ピボット テーブルの削除</span><span class="sxs-lookup"><span data-stu-id="7e66c-199">Delete a PivotTable</span></span>

<span data-ttu-id="7e66c-200">ピボットテーブルをその名前を用いて削除します。</span><span class="sxs-lookup"><span data-stu-id="7e66c-200">PivotTables are deleted by using their name.</span></span>

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="7e66c-201">関連項目</span><span class="sxs-lookup"><span data-stu-id="7e66c-201">See also</span></span>

- [<span data-ttu-id="7e66c-202">Excel の JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="7e66c-202">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="7e66c-203">Excel の JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="7e66c-203">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/api/excel)
