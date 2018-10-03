---
title: Excel の JavaScript API を使用してピボット テーブルで作業します
description: Excel JavaScript API を使用してピボットテーブルを作成し、そのコンポーネントと対話します。
ms.date: 09/21/2018
ms.openlocfilehash: 5245665bad2933df205bcda29e226a965de1c356
ms.sourcegitcommit: 64da9ed76d22b14df745b1f0ef97a8f5194400e4
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/03/2018
ms.locfileid: "25361025"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="b48fb-103">Excel の JavaScript API を使用してピボット テーブルで作業します</span><span class="sxs-lookup"><span data-stu-id="b48fb-103">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="b48fb-104">ピボット テーブルより大きなデータ セットを合理化します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="b48fb-105">グループ化されたデータのクイック操作が可能です。</span><span class="sxs-lookup"><span data-stu-id="b48fb-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="b48fb-106">Excel の JavaScript API では、アドインにピボット テーブルを作成させ、それらのコンポーネントと対話することができます。</span><span class="sxs-lookup"><span data-stu-id="b48fb-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> 

<span data-ttu-id="b48fb-107">ピボット テーブルの機能に慣れていない場合は、エンド ユーザーとしてこれらの操作を検討してください。</span><span class="sxs-lookup"><span data-stu-id="b48fb-107">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end-user.</span></span> <span data-ttu-id="b48fb-108">これらのツールの良い入門書については、[ピボットテーブルを作成してワークシートのデータを分析する ](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b48fb-108">See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span> 

<span data-ttu-id="b48fb-109">この記事では、一般的なシナリオのコード サンプルを提供します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-109">This article provides code samples for common scenarios.</span></span> <span data-ttu-id="b48fb-110">ピボットテーブルAPI の理解をさらに深めるには、 [**PivotTable**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) と [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b48fb-110">To further your understanding of the PivotTable API, see [**PivotTable**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b48fb-111">OLAP で作成されたピボット テーブルは、現在サポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b48fb-111">PivotTables created with OLAP are not currently supported.</span></span>

## <a name="hierarchies"></a><span data-ttu-id="b48fb-112">階層</span><span class="sxs-lookup"><span data-stu-id="b48fb-112">Hierarchies</span></span>

<span data-ttu-id="b48fb-113">ピボット テーブルは、行、列、データ、フィルターの 4 つの階層カテゴリに基づいて構成されています。</span><span class="sxs-lookup"><span data-stu-id="b48fb-113">PivotTables are organized based on four hierarchy categories: row, column, data, and filter.</span></span> <span data-ttu-id="b48fb-114">この記事全体を通して、さまざまな農場の果物の売り上げを記述した次のデータを使用します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-114">The following data describing fruit sales from various farms will be used throughout this article.</span></span>

![さまざまな農場のさまざまな種類の果物の売り上げのコレクション。](../images/excel-pivots-raw-data.png)

<span data-ttu-id="b48fb-116">このデータには **農家**、 **種類**、 **分類**、**農場で販売された箱数**、および **卸売りで販売された箱数** の 5 つの階層があります。</span><span class="sxs-lookup"><span data-stu-id="b48fb-116">This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="b48fb-117">各階層は、4 つの分類項目のうちの 1 つにのみ存在することができます。</span><span class="sxs-lookup"><span data-stu-id="b48fb-117">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="b48fb-118">**種類** が 列の階層に追加され、さらに行の階層に追加された場合、行の階層にのみ残ります。</span><span class="sxs-lookup"><span data-stu-id="b48fb-118">If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.</span></span>

<span data-ttu-id="b48fb-119">行と列の階層は、データをグループ化する方法を定義します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-119">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="b48fb-120">たとえば、 **農場** の行の階層は、同じ農場のすべてのデータ セットをまとめてグループ化します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-120">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="b48fb-121">行と列の階層から選択すると、ピボット テーブルの向きが定義されます。</span><span class="sxs-lookup"><span data-stu-id="b48fb-121">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="b48fb-122">データ階層は、行と列の階層に基づいて集計される値です。</span><span class="sxs-lookup"><span data-stu-id="b48fb-122">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="b48fb-123">**農場** の行の階層と **卸売りで販売された木箱** のデータ階層からなるピボット テーブルは、各農場のさまざまな種類の果物の総計 (既定) を示します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-123">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="b48fb-124">フィルター階層は、フィルターされた種類の中の値に基づいてピボットにデータを取り込むか、取り除きます。</span><span class="sxs-lookup"><span data-stu-id="b48fb-124">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="b48fb-125">**分類** のフィルター階層で **有機栽培** を選択すると、有機栽培の果物のデータのみが表示されます。</span><span class="sxs-lookup"><span data-stu-id="b48fb-125">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="b48fb-126">これで再び農場のデータができ、ピボット テーブルに表示されます。</span><span class="sxs-lookup"><span data-stu-id="b48fb-126">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="b48fb-127">ピボット テーブルは、**農場** と **種類**を行階層、 **農場で販売された箱数** と**卸売りで販売された箱数** をデータ階層 (既定の合計の集計関数)、**分類**  をフィルター階層 (**有機栽培**を選択) として使用しています。</span><span class="sxs-lookup"><span data-stu-id="b48fb-127">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span> 

![行、データ、フィルターの階層で構成したピボット テーブルの次に果物の売り上げデータの選択範囲があります。](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="b48fb-129">このピボットテーブルは、 JavaScript API  または Excel  の UI を用いて生成できました。</span><span class="sxs-lookup"><span data-stu-id="b48fb-129">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="b48fb-130">両方のオプションで、アドインでさらに操作することができます。</span><span class="sxs-lookup"><span data-stu-id="b48fb-130">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="b48fb-131">ピボット テーブルの作成</span><span class="sxs-lookup"><span data-stu-id="b48fb-131">Create a PivotTable and PivotChart</span></span>

<span data-ttu-id="b48fb-132">ピボット テーブルには、名前、ソース、同期先が必要です。</span><span class="sxs-lookup"><span data-stu-id="b48fb-132">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="b48fb-133">ソースは、範囲アドレス、またはテーブル名を指定できます ( `Range`、 `string`、`Table` 型として渡されます)。</span><span class="sxs-lookup"><span data-stu-id="b48fb-133">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="b48fb-134">同期先は、範囲アドレスです (`Range` または `string` のいずれかとして付与されます)。</span><span class="sxs-lookup"><span data-stu-id="b48fb-134">The destination is a range address (given as either a `Range` or `string`).</span></span> <span data-ttu-id="b48fb-135">次のサンプルでは、さまざまなピボット テーブルの作成方法を示します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-135">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="b48fb-136">範囲アドレスを使用してピボット テーブルを作成</span><span class="sxs-lookup"><span data-stu-id="b48fb-136">Create a PivotTable with range addresses</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="b48fb-137">Range オブジェクトを使用してピボット テーブルを作成</span><span class="sxs-lookup"><span data-stu-id="b48fb-137">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="b48fb-138">ワークブック レベルでピボット テーブルを作成</span><span class="sxs-lookup"><span data-stu-id="b48fb-138">Create a PivotTable at the workbook level</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="b48fb-139">既存のピボット テーブルの使用</span><span class="sxs-lookup"><span data-stu-id="b48fb-139">Use an existing PivotTable</span></span>

<span data-ttu-id="b48fb-140">手動で作成したピボット テーブルも、ブックのピボット テーブルのコレクションまたはここのワークシートを使用してアクセス可能です。</span><span class="sxs-lookup"><span data-stu-id="b48fb-140">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> 

<span data-ttu-id="b48fb-141">次のコードは、ブックに最初のピボットテーブルを追加します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-141">The following code gets the first PivotTable in the workbook.</span></span> <span data-ttu-id="b48fb-142">以降に参照しやすくするため、テーブルに名前を付与します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-142">It then gives the table a name for easy future reference.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="b48fb-143">ピボット テーブルに行と列を追加</span><span class="sxs-lookup"><span data-stu-id="b48fb-143">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="b48fb-144">行と列は、これらのフィールドの値の周りでデータをピボットします。</span><span class="sxs-lookup"><span data-stu-id="b48fb-144">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="b48fb-145">**農場** 列を追加すると、各農場のすべての売り上げをピボットします。</span><span class="sxs-lookup"><span data-stu-id="b48fb-145">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="b48fb-146">**種類** と **分類** 行を追加すると、どの果物が販売されたか、そしてそれが有機栽培かどうかに基づいて、データがさらに分解されます。</span><span class="sxs-lookup"><span data-stu-id="b48fb-146">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

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

<span data-ttu-id="b48fb-148">行または列のみを含むピボット テーブルも可能です。</span><span class="sxs-lookup"><span data-stu-id="b48fb-148">You can also have a PivotTable with only rows or columns.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="b48fb-149">ピボット テーブルへのデータ階層の追加</span><span class="sxs-lookup"><span data-stu-id="b48fb-149">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="b48fb-150">データ階層は、行と列に基づいて組み合わせる情報でピボット テーブルを入力します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-150">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="b48fb-151">**農場で販売された箱数** と **卸売りで販売された箱数** のデータ階層を追加すると、各行と列にそれらの数値の合計が表示されます。</span><span class="sxs-lookup"><span data-stu-id="b48fb-151">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span> 

<span data-ttu-id="b48fb-152">この例では、 **農場** と **種類** はともに行となり、箱の販売数をデータとして表示します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-152">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span> 

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

## <a name="change-aggregation-function"></a><span data-ttu-id="b48fb-154">集計関数を変更する</span><span class="sxs-lookup"><span data-stu-id="b48fb-154">Change aggregation function</span></span>

<span data-ttu-id="b48fb-155">データの階層では、値を集計します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-155">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="b48fb-156">数値のデータセットの場合、既定ではこれは合計となります。</span><span class="sxs-lookup"><span data-stu-id="b48fb-156">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="b48fb-157">タイプ `summarizeBy` に基づいてプロパティはこの動作を定義します 。`AggregrationFunction`</span><span class="sxs-lookup"><span data-stu-id="b48fb-157">The `summarizeBy` property defines this behavior based on an `AggregrationFunction` type.</span></span> 

<span data-ttu-id="b48fb-158">現在サポートされている集計関数のタイプは、 `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`,  `Automatic` (既定値) です。</span><span class="sxs-lookup"><span data-stu-id="b48fb-158">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="b48fb-159">次のコード サンプルでは、データの平均値を使用する集計を変更します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-159">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="b48fb-160">ShowAsRule を使用して計算を変更します</span><span class="sxs-lookup"><span data-stu-id="b48fb-160">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="b48fb-161">ピボット テーブルでは、既定では、行と列の階層のデータを個別に集約します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-161">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="b48fb-162">A `ShowAsRule` ピボット テーブル内の他の項目に基づいた値を出力するために、データの階層を変更します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-162">A `ShowAsRule` changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="b48fb-163"> `ShowAsRule` オブジェクトには次の 3 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="b48fb-163">The `ShowAsRule` object has three properties:</span></span>
-   <span data-ttu-id="b48fb-164">`calculation`: データの階層に適用する相対的な計算の種類 (既定値は `none`)。</span><span class="sxs-lookup"><span data-stu-id="b48fb-164">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
-   <span data-ttu-id="b48fb-165">`baseField`: 計算が適用される前の基本データを含む階層内のフィールド。</span><span class="sxs-lookup"><span data-stu-id="b48fb-165">`baseField`: The field within the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="b48fb-166">通常、 `PivotField`は 親の階層と同じ名前を持ちます。</span><span class="sxs-lookup"><span data-stu-id="b48fb-166">The `PivotField` usually has the same name as its parent hierarchy.</span></span>
-   <span data-ttu-id="b48fb-167">`baseItem`: 計算の種類に基づいた基本フィールドの値と比較した個々の項目。</span><span class="sxs-lookup"><span data-stu-id="b48fb-167">`baseItem`: The individual item compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="b48fb-168">すべての計算がこのフィールドを必要とするわけではありません。</span><span class="sxs-lookup"><span data-stu-id="b48fb-168">Not all calculations require this field.</span></span>

<span data-ttu-id="b48fb-169">列合計のパーセント値で指定する **ファームで販売される木箱の合計** データ階層の計算を設定する例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-169">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span> <span data-ttu-id="b48fb-170">粒度を果物の種類レベルに拡張するため、 **種類** の行の階層と基になるフィールドを使用するようにします。</span><span class="sxs-lookup"><span data-stu-id="b48fb-170">We still want the granularity to extend to the fruit type level, so we’ll use the **Type** row hierarchy and its underlying field.</span></span> <span data-ttu-id="b48fb-171">例では、最初の行の階層として **ファーム** も示しているため、ファームの合計エントリは、各ファームが生産の責任を負うパーセント値も表示します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-171">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

![各ファーム内の個々のファームと個々の果物の種類の両方の総計と比べて果物の売り上げ高のパーセント値を示すピボット テーブル。](../images/excel-pivots-showas-percentage.png)

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

<span data-ttu-id="b48fb-173">以前の例では、個々の行の階層と比べて、列に計算を設定します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-173">The previous example set the calculation to the column, relative to an individual row hierarchy.</span></span> <span data-ttu-id="b48fb-174">計算が個々の項目に関連する場合は、 `baseItem` プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-174">When the calculation relates to an individual item, use the `baseItem` property.</span></span> 

<span data-ttu-id="b48fb-175">次の例は、 `differenceFrom` 計算を示します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-175">The following example shows the request.</span></span> <span data-ttu-id="b48fb-176">「A ファーム」のファーム木箱販売データ階層エントリの差を表示します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-176">It displays the difference of the farm crate sales data hierarchy entries relative to those of “A Farms”.</span></span> <span data-ttu-id="b48fb-177"> `baseField\`は *\*ファーム*\*なので、各果物の種類のブレークダウン図形と同様に、他のファーム間の差がわかります (この例では*\*種類** も行の階層) 。</span><span class="sxs-lookup"><span data-stu-id="b48fb-177">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![「A ファーム」と他のユーザーの果物販売の差を示すピボット テーブル。](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="pivottable-layouts"></a><span data-ttu-id="b48fb-181">ピボット テーブルのレイアウト</span><span class="sxs-lookup"><span data-stu-id="b48fb-181">PivotTable layouts</span></span>

<span data-ttu-id="b48fb-182">ピボットテーブルのレイアウトは、階層とそのデータの配置を定義します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-182">A PivotTable layout defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="b48fb-183">データが保存されている範囲を決定するレイアウトにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="b48fb-183">You access the layout to determine the ranges where data is stored.</span></span> 

<span data-ttu-id="b48fb-184">レイアウト関数を呼び出す次の図は、ピボット テーブルの範囲に対応します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-184">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![ピボット テーブルのどの部分がレイアウトの取得範囲の関数によって返されるかを示す図。](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="b48fb-186">次のコードでは、レイアウトを使用するピボット テーブルのデータの最後の行を取得する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-186">The following code demonstrates how to get the last row of the PivotTable data by going through the layout.</span></span> <span data-ttu-id="b48fb-187">これらの値は、総計用にまとめて集計されます。</span><span class="sxs-lookup"><span data-stu-id="b48fb-187">Those values are then summed together for a grand total.</span></span>

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

<span data-ttu-id="b48fb-188">ピボット テーブルには、3 つのレイアウト スタイル: コンパクト、アウトライン、および表形式があります。</span><span class="sxs-lookup"><span data-stu-id="b48fb-188">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="b48fb-189">前の例でコンパクトなスタイルを使用しました。</span><span class="sxs-lookup"><span data-stu-id="b48fb-189">We’ve seen the compact style in the previous examples.</span></span> 

<span data-ttu-id="b48fb-190">次の例では、アウトライン、表形式のスタイルをそれぞれ使用します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-190">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="b48fb-191">コード サンプルでは、さまざまなレイアウトが交互に表示する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-191">The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="b48fb-192">アウトライン レイアウト表示</span><span class="sxs-lookup"><span data-stu-id="b48fb-192">Outline layout</span></span>

![アウトライン表示のレイアウトを使用するピボットテーブル。](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="b48fb-194">表形式のレイアウト</span><span class="sxs-lookup"><span data-stu-id="b48fb-194">Tabular layout</span></span>

![表形式のレイアウトを使用するピボットテーブル。](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a><span data-ttu-id="b48fb-196">階層名の変更</span><span class="sxs-lookup"><span data-stu-id="b48fb-196">Change hierarchy names</span></span>

<span data-ttu-id="b48fb-197">階層のフィールドは、編集できます。</span><span class="sxs-lookup"><span data-stu-id="b48fb-197">Hierarchy fields are editable.</span></span> <span data-ttu-id="b48fb-198">次のコードでは、二つのデータ階層の表示された名前をどのように変更するかを説明します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-198">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="b48fb-199">ピボット テーブルを削除します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-199">Delete a PivotTable</span></span>

<span data-ttu-id="b48fb-200">ピボットテーブルをその名前を用いて削除します。</span><span class="sxs-lookup"><span data-stu-id="b48fb-200">PivotTables are deleted by using their name.</span></span>

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="b48fb-201">関連項目</span><span class="sxs-lookup"><span data-stu-id="b48fb-201">See also</span></span>

- [<span data-ttu-id="b48fb-202">Excel JavaScript API の中心概念</span><span class="sxs-lookup"><span data-stu-id="b48fb-202">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="b48fb-203">Excel の JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="b48fb-203">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/api/excel)
