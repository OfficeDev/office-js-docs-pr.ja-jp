---
title: Excel JavaScript API を使用してピボットテーブルを操作する
description: Excel JavaScript API を使用して、ピボットテーブルを作成し、それらのコンポーネントを操作します。
ms.date: 03/21/2019
localization_priority: Normal
ms.openlocfilehash: b53d734e676417a6438f1008bac720a38a244d1f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449376"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="d0a35-103">Excel JavaScript API を使用してピボットテーブルを操作する</span><span class="sxs-lookup"><span data-stu-id="d0a35-103">Work with PivotTables using the Excel JavaScript API</span></span>

<span data-ttu-id="d0a35-104">ピボットテーブルは、より大きなデータセットを合理化します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="d0a35-105">グループ化されたデータのクイック操作を可能にします。</span><span class="sxs-lookup"><span data-stu-id="d0a35-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="d0a35-106">Excel JavaScript API を使用すると、アドインでピボットテーブルを作成し、それらのコンポーネントを操作できます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span>

<span data-ttu-id="d0a35-107">ピボットテーブルの機能についてよく知らない場合は、エンドユーザーとしての調査を検討してください。</span><span class="sxs-lookup"><span data-stu-id="d0a35-107">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user.</span></span> <span data-ttu-id="d0a35-108">これらのツールの詳細については、「[ワークシートデータを分析するためのピボットテーブルを作成する](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d0a35-108">See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span> 

<span data-ttu-id="d0a35-109">この記事では、一般的なシナリオのコードサンプルを示します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-109">This article provides code samples for common scenarios.</span></span> <span data-ttu-id="d0a35-110">ピボットテーブル API について理解するには、「 [**pivottable**](/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](/javascript/api/excel/excel.pivottable)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d0a35-110">To further your understanding of the PivotTable API, see [**PivotTable**](/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](/javascript/api/excel/excel.pivottable).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d0a35-111">OLAP を使用して作成されたピボットテーブルは現在サポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d0a35-111">PivotTables created with OLAP are not currently supported.</span></span> <span data-ttu-id="d0a35-112">Power Pivot もサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d0a35-112">There is also no support for Power Pivot.</span></span>

## <a name="hierarchies"></a><span data-ttu-id="d0a35-113">Hierarchies</span><span class="sxs-lookup"><span data-stu-id="d0a35-113">Hierarchies</span></span>

<span data-ttu-id="d0a35-114">ピボットテーブルは、行、列、データ、およびフィルターの4つの階層カテゴリに基づいて編成されます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-114">PivotTables are organized based on four hierarchy categories: row, column, data, and filter.</span></span> <span data-ttu-id="d0a35-115">この記事では、さまざまなファームからの果物 sales について説明する次のデータが使用されます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-115">The following data describing fruit sales from various farms will be used throughout this article.</span></span>

![さまざまなファームからのさまざまな種類の果物販売のコレクション。](../images/excel-pivots-raw-data.png)

<span data-ttu-id="d0a35-117">このデータには、**畑**、 **Type**、**分類**、 **Crates で販売**されたファーム、 **Crates 販売**された卸売の5つの階層があります。</span><span class="sxs-lookup"><span data-stu-id="d0a35-117">This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="d0a35-118">各階層は、4つのカテゴリのいずれかにのみ存在できます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-118">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="d0a35-119">**Type**が列階層に追加されてから、行階層に追加されても、後者には残ります。</span><span class="sxs-lookup"><span data-stu-id="d0a35-119">If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.</span></span>

<span data-ttu-id="d0a35-120">行と列の階層は、データをグループ化する方法を定義します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-120">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="d0a35-121">たとえば、**ファーム**の行階層は、同じファームのすべてのデータセットをグループ化します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-121">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="d0a35-122">行と列の階層を選択すると、ピボットテーブルの向きが定義されます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-122">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="d0a35-123">データ階層は、行と列の階層に基づいて集計される値です。</span><span class="sxs-lookup"><span data-stu-id="d0a35-123">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="d0a35-124">ファームの行階層があり、 \*\*\*\* **Crates**のデータ階層があるピボットテーブルには、各ファームのすべての異なる fruits の合計 (既定では) が表示されます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-124">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="d0a35-125">フィルター階層では、フィルター処理された種類の値に基づいて、ピボットのデータが含まれるか、除外されます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-125">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="d0a35-126">**有機**的に選択された種類の**分類**のフィルター階層は、有機フルーツのデータのみを表示します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-126">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="d0a35-127">次に、ファームデータをピボットテーブルと共に示します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-127">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="d0a35-128">ピボットテーブルでは、**ファーム**と**タイプ**を行階層として使用し、**ファームで販売**された Crates と Crates がデータ階層として**卸売販売**され、フィルターとして**分類**されています。階層 (**有機**が選択されている)。</span><span class="sxs-lookup"><span data-stu-id="d0a35-128">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span> 

![行、データ、およびフィルター階層を使用したピボットテーブルの横の、果物 sales データの選択。](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="d0a35-130">このピボットテーブルは、JavaScript API または Excel UI を使用して生成できます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-130">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="d0a35-131">両方のオプションを使用すると、アドインをさらに操作できます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-131">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="d0a35-132">ピボットテーブルを作成する</span><span class="sxs-lookup"><span data-stu-id="d0a35-132">Create a PivotTable</span></span>

<span data-ttu-id="d0a35-133">ピボットテーブルには、名前、ソース、および出力先が必要です。</span><span class="sxs-lookup"><span data-stu-id="d0a35-133">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="d0a35-134">ソースは、範囲内のアドレスまたはテーブル名 (、、 `Range`、 `string`または`Table`型として渡されます) を指定できます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-134">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="d0a35-135">宛先は、または`Range` `string`のいずれかとして指定された範囲のアドレスです。</span><span class="sxs-lookup"><span data-stu-id="d0a35-135">The destination is a range address (given as either a `Range` or `string`).</span></span> <span data-ttu-id="d0a35-136">次のサンプルは、さまざまなピボットテーブル作成手法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d0a35-136">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="d0a35-137">範囲のアドレスを使用してピボットテーブルを作成する</span><span class="sxs-lookup"><span data-stu-id="d0a35-137">Create a PivotTable with range addresses</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="d0a35-138">Range オブジェクトを使用してピボットテーブルを作成する</span><span class="sxs-lookup"><span data-stu-id="d0a35-138">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="d0a35-139">ブックレベルでピボットテーブルを作成する</span><span class="sxs-lookup"><span data-stu-id="d0a35-139">Create a PivotTable at the workbook level</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="d0a35-140">既存のピボットテーブルを使用する</span><span class="sxs-lookup"><span data-stu-id="d0a35-140">Use an existing PivotTable</span></span>

<span data-ttu-id="d0a35-141">手動で作成したピボットテーブルは、ブックまたは個々のワークシートの PivotTable コレクションからアクセスすることもできます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-141">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> 

<span data-ttu-id="d0a35-142">次のコードは、ブック内の最初のピボットテーブルを取得します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-142">The following code gets the first PivotTable in the workbook.</span></span> <span data-ttu-id="d0a35-143">その後、表の名前を後で簡単に参照できるようにします。</span><span class="sxs-lookup"><span data-stu-id="d0a35-143">It then gives the table a name for easy future reference.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="d0a35-144">ピボットテーブルに行と列を追加する</span><span class="sxs-lookup"><span data-stu-id="d0a35-144">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="d0a35-145">行と列は、それらのフィールド値を中心にデータをピボットします。</span><span class="sxs-lookup"><span data-stu-id="d0a35-145">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="d0a35-146">[**ファーム**] 列を追加すると、各ファームのすべての売上が回転します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-146">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="d0a35-147">**種類**と**分類**行を追加すると、果物が販売されたものと、それが有機であったかどうかに基づいてデータがさらに分解されます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-147">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

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

<span data-ttu-id="d0a35-149">行または列だけのピボットテーブルを作成することもできます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-149">You can also have a PivotTable with only rows or columns.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="d0a35-150">データ階層をピボットテーブルに追加する</span><span class="sxs-lookup"><span data-stu-id="d0a35-150">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="d0a35-151">データ階層は、行と列に基づいて結合する情報で、ピボットテーブルに格納されます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-151">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="d0a35-152">**ファームで販売**された Crates のデータ階層を追加し、 **Crates に販売**されたものは、行と列ごとにこれらの数値を合計します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-152">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span> 

<span data-ttu-id="d0a35-153">この例では、**ファーム**と**種類**の両方が行で、箱 sales がデータとして含まれています。</span><span class="sxs-lookup"><span data-stu-id="d0a35-153">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span> 

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

## <a name="change-aggregation-function"></a><span data-ttu-id="d0a35-155">集計関数を変更する</span><span class="sxs-lookup"><span data-stu-id="d0a35-155">Change aggregation function</span></span>

<span data-ttu-id="d0a35-156">データ階層の値が集計されます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-156">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="d0a35-157">数値のデータセットの場合は、既定でこれが合計になります。</span><span class="sxs-lookup"><span data-stu-id="d0a35-157">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="d0a35-158">この`summarizeBy`プロパティは、この動作を[集約 ationfunction](/javascript/api/excel/excel.aggregationfunction)型に基づいて定義します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-158">The `summarizeBy` property defines this behavior based on an [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) type.</span></span>

<span data-ttu-id="d0a35-159">現在サポートされている集計`Sum`関数`Count`の`Average`種類`Max`は`Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP`、、、、、、、、 `Automatic` 、、、、および (既定値) です。</span><span class="sxs-lookup"><span data-stu-id="d0a35-159">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="d0a35-160">次のコードサンプルでは、集計をデータの平均値に変更します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-160">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="d0a35-161">showasrule を使用して計算を変更する</span><span class="sxs-lookup"><span data-stu-id="d0a35-161">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="d0a35-162">既定では、ピボットテーブルでは、行と列の階層のデータが個別に集計されます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-162">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="d0a35-163">[showasrule](/javascript/api/excel/excel.showasrule)は、データ階層を、ピボットテーブル内の他のアイテムに基づいて出力値に変更します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-163">A [ShowAsRule](/javascript/api/excel/excel.showasrule) changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="d0a35-164">オブジェクト`ShowAsRule`には、次の3つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="d0a35-164">The `ShowAsRule` object has three properties:</span></span>

-   <span data-ttu-id="d0a35-165">`calculation`: データ階層に適用する相対的な計算の種類 (既定値は`none`)。</span><span class="sxs-lookup"><span data-stu-id="d0a35-165">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
-   <span data-ttu-id="d0a35-166">`baseField`: 計算を適用する前に、基本データを含む階層内のフィールド。</span><span class="sxs-lookup"><span data-stu-id="d0a35-166">`baseField`: The field within the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="d0a35-167">通常、[ピボットフィールド](/javascript/api/excel/excel.pivotfield)の名前は親階層と同じです。</span><span class="sxs-lookup"><span data-stu-id="d0a35-167">The [PivotField](/javascript/api/excel/excel.pivotfield) usually has the same name as its parent hierarchy.</span></span>
-   <span data-ttu-id="d0a35-168">`baseItem`: 計算の種類に基づいて、基準フィールドの値と比較した個々の[ピボット](/javascript/api/excel/excel.pivotitem)テーブル。</span><span class="sxs-lookup"><span data-stu-id="d0a35-168">`baseItem`: The individual [PivotItem](/javascript/api/excel/excel.pivotitem) compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="d0a35-169">すべての計算にこのフィールドが必要なわけではありません。</span><span class="sxs-lookup"><span data-stu-id="d0a35-169">Not all calculations require this field.</span></span>

<span data-ttu-id="d0a35-170">次の例では、ファームデータ階層で販売された**Crates の合計**の計算を列の合計のパーセンテージに設定します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-170">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span> <span data-ttu-id="d0a35-171">さらに、粒度をフルーツの種類レベルにまで拡張する必要があるので、 **type**行階層とその基になるフィールドを使用します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-171">We still want the granularity to extend to the fruit type level, so we’ll use the **Type** row hierarchy and its underlying field.</span></span> <span data-ttu-id="d0a35-172">この例は、最初の行階層としても**ファーム**を持っているので、farm total エントリには各ファームがそれぞれを生成する割合が表示されます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-172">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

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

<span data-ttu-id="d0a35-174">前の例では、列の各行階層を基準にして計算を設定しています。</span><span class="sxs-lookup"><span data-stu-id="d0a35-174">The previous example set the calculation to the column, relative to an individual row hierarchy.</span></span> <span data-ttu-id="d0a35-175">計算が個々のアイテムに関連している場合`baseItem`は、プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-175">When the calculation relates to an individual item, use the `baseItem` property.</span></span>

<span data-ttu-id="d0a35-176">次の例は、 `differenceFrom`計算を示しています。</span><span class="sxs-lookup"><span data-stu-id="d0a35-176">The following example shows the `differenceFrom` calculation.</span></span> <span data-ttu-id="d0a35-177">この例では、"ファーム" と比較した場合の、ファームの箱売上データ階層エントリの違いを示します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-177">It displays the difference of the farm crate sales data hierarchy entries relative to those of “A Farms”.</span></span>
<span data-ttu-id="d0a35-178">`baseField`は**ファーム**ですので、他のファームの違いと、果物などの種類ごとの内訳 (この例では、**type**が行階層になっています) を確認しています。</span><span class="sxs-lookup"><span data-stu-id="d0a35-178">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

!["畑" とその他の果物の売上の違いを示すピボットテーブル。](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="pivottable-layouts"></a><span data-ttu-id="d0a35-182">ピボットテーブルのレイアウト</span><span class="sxs-lookup"><span data-stu-id="d0a35-182">PivotTable layouts</span></span>

<span data-ttu-id="d0a35-183">[PivotLayout](/javascript/api/excel/excel.pivotlayout)は、階層とそのデータの配置を定義します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-183">A [PivotLayout](/javascript/api/excel/excel.pivotlayout) defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="d0a35-184">レイアウトにアクセスして、データが格納される範囲を決定します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-184">You access the layout to determine the ranges where data is stored.</span></span>

<span data-ttu-id="d0a35-185">次の図は、ピボットテーブルの範囲に対応するどの layout 関数呼び出しを示しています。</span><span class="sxs-lookup"><span data-stu-id="d0a35-185">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![レイアウトの範囲取得機能によって返されるピボットテーブルのセクションを示す図。](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="d0a35-187">次のコードは、レイアウトを使用してピボットテーブルデータの最後の行を取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d0a35-187">The following code demonstrates how to get the last row of the PivotTable data by going through the layout.</span></span> <span data-ttu-id="d0a35-188">これらの値は総計に対して合計されます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-188">Those values are then summed together for a grand total.</span></span>

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

<span data-ttu-id="d0a35-189">ピボットテーブルには、コンパクト、アウトライン、表形式という3つのレイアウトスタイルがあります。</span><span class="sxs-lookup"><span data-stu-id="d0a35-189">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="d0a35-190">前の例ではコンパクトなスタイルを見てきました。</span><span class="sxs-lookup"><span data-stu-id="d0a35-190">We’ve seen the compact style in the previous examples.</span></span> 

<span data-ttu-id="d0a35-191">次の例では、アウトラインスタイルと表形式スタイルをそれぞれ使用します。</span><span class="sxs-lookup"><span data-stu-id="d0a35-191">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="d0a35-192">このコードサンプルは、さまざまなレイアウト間で循環する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d0a35-192">The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="d0a35-193">アウトラインレイアウト</span><span class="sxs-lookup"><span data-stu-id="d0a35-193">Outline layout</span></span>

![アウトラインレイアウトを使用したピボットテーブル。](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="d0a35-195">表形式レイアウト</span><span class="sxs-lookup"><span data-stu-id="d0a35-195">Tabular layout</span></span>

![表形式レイアウトを使用したピボットテーブル。](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a><span data-ttu-id="d0a35-197">階層名を変更する</span><span class="sxs-lookup"><span data-stu-id="d0a35-197">Change hierarchy names</span></span>

<span data-ttu-id="d0a35-198">階層フィールドは編集できます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-198">Hierarchy fields are editable.</span></span> <span data-ttu-id="d0a35-199">次のコードは、2つのデータ階層の表示名を変更する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d0a35-199">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="d0a35-200">ピボットテーブルを削除する</span><span class="sxs-lookup"><span data-stu-id="d0a35-200">Delete a PivotTable</span></span>

<span data-ttu-id="d0a35-201">ピボットテーブルは、名前を使用して削除されます。</span><span class="sxs-lookup"><span data-stu-id="d0a35-201">PivotTables are deleted by using their name.</span></span>

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="d0a35-202">関連項目</span><span class="sxs-lookup"><span data-stu-id="d0a35-202">See also</span></span>

- [<span data-ttu-id="d0a35-203">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="d0a35-203">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="d0a35-204">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="d0a35-204">Excel JavaScript API Reference</span></span>](/javascript/api/excel)
