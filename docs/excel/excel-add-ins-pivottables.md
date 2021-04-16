---
title: Excel JavaScript API を使用してピボットテーブルを使用する
description: Excel JavaScript API を使用してピボットテーブルを作成し、それらのコンポーネントを操作します。
ms.date: 04/09/2021
localization_priority: Normal
ms.openlocfilehash: a76d2401784c7ca52c2c54342ccce21b53097a58
ms.sourcegitcommit: 094caf086c2696e78fbdfdc6030cb0c89d32b585
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/16/2021
ms.locfileid: "51862345"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="64d7d-103">Excel JavaScript API を使用してピボットテーブルを使用する</span><span class="sxs-lookup"><span data-stu-id="64d7d-103">Work with PivotTables using the Excel JavaScript API</span></span>

<span data-ttu-id="64d7d-104">PivotTables は、より大きなデータ セットを合理化します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="64d7d-105">グループ化されたデータを簡単に操作できます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="64d7d-106">Excel JavaScript API を使用すると、アドインでピボットテーブルを作成し、それらのコンポーネントを操作できます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> <span data-ttu-id="64d7d-107">この記事では、ピボットテーブルが JavaScript API の Officeされる方法について説明し、主要なシナリオのコード サンプルを提供します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-107">This article describes how PivotTables are represented by the Office JavaScript API and provides code samples for key scenarios.</span></span>

<span data-ttu-id="64d7d-108">ピボットテーブルの機能に慣れていない場合は、エンド ユーザーとして探索を検討してください。</span><span class="sxs-lookup"><span data-stu-id="64d7d-108">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user.</span></span>
<span data-ttu-id="64d7d-109">これらの [ツールの優れた入門については、「ピボットテーブル](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) を作成してワークシート データを分析する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="64d7d-109">See [Create a PivotTable to analyze worksheet data](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="64d7d-110">OLAP で作成されたピボットテーブルは現在サポートされていません。</span><span class="sxs-lookup"><span data-stu-id="64d7d-110">PivotTables created with OLAP are not currently supported.</span></span> <span data-ttu-id="64d7d-111">Power Pivot もサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="64d7d-111">There is also no support for Power Pivot.</span></span>

## <a name="object-model"></a><span data-ttu-id="64d7d-112">オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="64d7d-112">Object model</span></span>

<span data-ttu-id="64d7d-113">ピボット [テーブルは](/javascript/api/excel/excel.pivottable) 、JavaScript API のピボットテーブルOfficeです。</span><span class="sxs-lookup"><span data-stu-id="64d7d-113">The [PivotTable](/javascript/api/excel/excel.pivottable) is the central object for PivotTables in the Office JavaScript API.</span></span>

- <span data-ttu-id="64d7d-114">`Workbook.pivotTables`は、それぞれブックとワークシートにピボットテーブルを含む `Worksheet.pivotTables` [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection)です。 [](/javascript/api/excel/excel.pivottable)</span><span class="sxs-lookup"><span data-stu-id="64d7d-114">`Workbook.pivotTables` and `Worksheet.pivotTables` are [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) that contain the [PivotTables](/javascript/api/excel/excel.pivottable) in the workbook and worksheet, respectively.</span></span>
- <span data-ttu-id="64d7d-115">ピボット [テーブルには、](/javascript/api/excel/excel.pivottable) 複数の [PivotHierarchies を持つ PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) [が含まれる](/javascript/api/excel/excel.pivothierarchy)。</span><span class="sxs-lookup"><span data-stu-id="64d7d-115">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) that has multiple [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).</span></span>
- <span data-ttu-id="64d7d-116">これらの [PivotHierarchies を](/javascript/api/excel/excel.pivothierarchy) 特定の階層コレクションに追加して、ピボットテーブルがデータをピボットする方法を定義できます (次のセクション [で説明します](#hierarchies))。</span><span class="sxs-lookup"><span data-stu-id="64d7d-116">These [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy) can be added to specific hierarchy collections to define how the PivotTable pivots data (as explained in the [following section](#hierarchies)).</span></span>
- <span data-ttu-id="64d7d-117">[PivotHierarchy には](/javascript/api/excel/excel.pivothierarchy)、ピボットフィールドが 1 つ正確に含まれる[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)が[含まれる](/javascript/api/excel/excel.pivotfield)。</span><span class="sxs-lookup"><span data-stu-id="64d7d-117">A [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contains a [PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) that has exactly one [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="64d7d-118">OLAP ピボットテーブルを含むデザインが展開された場合、変更される可能性があります。</span><span class="sxs-lookup"><span data-stu-id="64d7d-118">If the design expands to include OLAP PivotTables, this may change.</span></span>
- <span data-ttu-id="64d7d-119">ピボット[フィールドには、](/javascript/api/excel/excel.pivotfield)フィールドの[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)が階層カテゴリに割り当てられている限り、1 つ以上のピボットフィルターを適用できます。 [](/javascript/api/excel/excel.pivotfilters)</span><span class="sxs-lookup"><span data-stu-id="64d7d-119">A [PivotField](/javascript/api/excel/excel.pivotfield) can have one or more [PivotFilters](/javascript/api/excel/excel.pivotfilters) applied, as long as the field's [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) is assigned to a hierarchy category.</span></span>
- <span data-ttu-id="64d7d-120">PivotField [には、](/javascript/api/excel/excel.pivotfield) 複数の PivotItem を持つ [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) [が含まれる](/javascript/api/excel/excel.pivotitem)。</span><span class="sxs-lookup"><span data-stu-id="64d7d-120">A [PivotField](/javascript/api/excel/excel.pivotfield) contains a [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) that has multiple [PivotItems](/javascript/api/excel/excel.pivotitem).</span></span>
- <span data-ttu-id="64d7d-121">ピボット[テーブルには、](/javascript/api/excel/excel.pivottable)[ピボットフィールドと](/javascript/api/excel/excel.pivotlayout)ピボットアイテムがワークシートに表示される[](/javascript/api/excel/excel.pivotfield)場所を定義する[ピボット](/javascript/api/excel/excel.pivotitem)レイアウトが含まれる。</span><span class="sxs-lookup"><span data-stu-id="64d7d-121">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotLayout](/javascript/api/excel/excel.pivotlayout) that defines where the [PivotFields](/javascript/api/excel/excel.pivotfield) and [PivotItems](/javascript/api/excel/excel.pivotitem) are displayed in the worksheet.</span></span> <span data-ttu-id="64d7d-122">レイアウトでは、ピボットテーブルの一部の表示設定も制御します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-122">The layout also controls some display settings for the PivotTable.</span></span>

<span data-ttu-id="64d7d-123">これらのリレーションシップがデータの例に適用される方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-123">Let's look at how these relationships apply to some example data.</span></span> <span data-ttu-id="64d7d-124">次のデータは、さまざまなファームからの果物の販売について説明します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-124">The following data describes fruit sales from various farms.</span></span> <span data-ttu-id="64d7d-125">この記事全体の例を示します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-125">It will be the example throughout this article.</span></span>

![異なるファームから異なる種類の果物の販売のコレクション。](../images/excel-pivots-raw-data.png)

<span data-ttu-id="64d7d-127">このフルーツ ファームの販売データは、ピボットテーブルの作成に使用されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-127">This fruit farm sales data will be used to make a PivotTable.</span></span> <span data-ttu-id="64d7d-128">Types などの各列 **は**、 です `PivotHierarchy` 。</span><span class="sxs-lookup"><span data-stu-id="64d7d-128">Each column, such as **Types**, is a `PivotHierarchy`.</span></span> <span data-ttu-id="64d7d-129">[ **型]** 階層には、[種類] **フィールドが含** まれます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-129">The **Types** hierarchy contains the **Types** field.</span></span> <span data-ttu-id="64d7d-130">[**種類]** フィールドには、Apple、Kiwi、Lemon、Lime、**および Orange** の項目が **含まれます**。  </span><span class="sxs-lookup"><span data-stu-id="64d7d-130">The **Types** field contains the items **Apple**, **Kiwi**, **Lemon**, **Lime**, and **Orange**.</span></span>

### <a name="hierarchies"></a><span data-ttu-id="64d7d-131">Hierarchies</span><span class="sxs-lookup"><span data-stu-id="64d7d-131">Hierarchies</span></span>

<span data-ttu-id="64d7d-132">ピボットテーブルは、行、列、データ、およびフィルター[](/javascript/api/excel/excel.rowcolumnpivothierarchy)の 4 つの[階層カテゴリに](/javascript/api/excel/excel.rowcolumnpivothierarchy)[基づいて編成](/javascript/api/excel/excel.datapivothierarchy)[されます](/javascript/api/excel/excel.filterpivothierarchy)。</span><span class="sxs-lookup"><span data-stu-id="64d7d-132">PivotTables are organized based on four hierarchy categories: [row](/javascript/api/excel/excel.rowcolumnpivothierarchy), [column](/javascript/api/excel/excel.rowcolumnpivothierarchy), [data](/javascript/api/excel/excel.datapivothierarchy), and [filter](/javascript/api/excel/excel.filterpivothierarchy).</span></span>

<span data-ttu-id="64d7d-133">前に示したファーム データには **、Farms**、 **Type**、 **Classification**、 **Crates Sold** at Farm 、 および Crates Sold Wholesale の 5 **つの階層があります**。</span><span class="sxs-lookup"><span data-stu-id="64d7d-133">The farm data shown earlier has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="64d7d-134">各階層は、4 つのカテゴリの 1 つにのみ存在できます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-134">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="64d7d-135">列 **階層** に Type を追加した場合は、行、データ、またはフィルター階層に追加することはできません。</span><span class="sxs-lookup"><span data-stu-id="64d7d-135">If **Type** is added to column hierarchies, it cannot also be in the row, data, or filter hierarchies.</span></span> <span data-ttu-id="64d7d-136">Type **が** 後で行階層に追加されると、列階層から削除されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-136">If **Type** is subsequently added to row hierarchies, it is removed from the column hierarchies.</span></span> <span data-ttu-id="64d7d-137">この動作は、Excel UI または Excel JavaScript API を使用して階層の割り当てを行う場合と同じです。</span><span class="sxs-lookup"><span data-stu-id="64d7d-137">This behavior is the same whether hierarchy assignment is done through the Excel UI or the Excel JavaScript APIs.</span></span>

<span data-ttu-id="64d7d-138">行階層と列階層は、データのグループ化方法を定義します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-138">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="64d7d-139">たとえば **、Farms** の行階層は、同じファームのすべてのデータ セットをグループ化します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-139">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="64d7d-140">行階層と列階層の選択によって、ピボットテーブルの向きが定義されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-140">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="64d7d-141">データ階層は、行階層と列階層に基づいて集計される値です。</span><span class="sxs-lookup"><span data-stu-id="64d7d-141">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="64d7d-142">ファームの行階層とクレート販売済みホールセールのデータ階層を持つピボットテーブルには、ファームごとに異なるすべての果物の合計 (既定) が表示されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-142">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="64d7d-143">フィルター階層には、フィルター処理された型内の値に基づいてピボットからデータが含まれるか除外されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-143">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="64d7d-144">[分類] の **フィルター階層で** [オーガニック] **が選択** されている場合は、オーガニック フルーツのデータだけが表示されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-144">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="64d7d-145">ピボットテーブルと共に、もう一度ファーム データを次に示します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-145">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="64d7d-146">ピボットテーブルは、行階層として Farm と **Type** を使用し、データ階層として [ファームで販売されたクレート] と [販売済みクレートの販売済みホールセール] をデータ階層 (合計の既定の集計関数を使用)、および分類をフィルター階層として使用します ([オーガニック] が選択されている場合)。 </span><span class="sxs-lookup"><span data-stu-id="64d7d-146">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span>

![行、データ、およびフィルター階層を持つピボットテーブルの横にある果物の販売データの選択。](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="64d7d-148">このピボットテーブルは、JavaScript API または Excel UI を使用して生成できます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-148">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="64d7d-149">どちらのオプションでも、アドインを介してさらに操作できます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-149">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="64d7d-150">ピボットテーブルの作成</span><span class="sxs-lookup"><span data-stu-id="64d7d-150">Create a PivotTable</span></span>

<span data-ttu-id="64d7d-151">ピボットテーブルには、名前、ソース、および宛先が必要です。</span><span class="sxs-lookup"><span data-stu-id="64d7d-151">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="64d7d-152">ソースには、範囲アドレスまたはテーブル名 (、、または型として渡される) `Range` `string` `Table` を指定できます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-152">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="64d7d-153">宛先は範囲アドレス (a または ) のいずれかとして `Range` 指定 `string` されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-153">The destination is a range address (given as either a `Range` or `string`).</span></span>
<span data-ttu-id="64d7d-154">次のサンプルは、さまざまなピボットテーブル作成手法を示しています。</span><span class="sxs-lookup"><span data-stu-id="64d7d-154">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="64d7d-155">範囲アドレスを使用してピボットテーブルを作成する</span><span class="sxs-lookup"><span data-stu-id="64d7d-155">Create a PivotTable with range addresses</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    return context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="64d7d-156">Range オブジェクトを使用してピボットテーブルを作成する</span><span class="sxs-lookup"><span data-stu-id="64d7d-156">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="64d7d-157">ブック レベルでピボットテーブルを作成する</span><span class="sxs-lookup"><span data-stu-id="64d7d-157">Create a PivotTable at the workbook level</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    return context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="64d7d-158">既存のピボットテーブルを使用する</span><span class="sxs-lookup"><span data-stu-id="64d7d-158">Use an existing PivotTable</span></span>

<span data-ttu-id="64d7d-159">手動で作成されたピボットテーブルには、ブックの PivotTable コレクションまたは個々のワークシートからアクセスすることもできます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-159">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> <span data-ttu-id="64d7d-160">次のコードは、ブックから My Pivot という **名前のピボットテーブル** を取得します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-160">The following code gets a PivotTable named **My Pivot** from the workbook.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="64d7d-161">ピボットテーブルに行と列を追加する</span><span class="sxs-lookup"><span data-stu-id="64d7d-161">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="64d7d-162">行と列は、これらのフィールドの値を中心にデータをピボットします。</span><span class="sxs-lookup"><span data-stu-id="64d7d-162">Rows and columns pivot the data around those fields' values.</span></span>

<span data-ttu-id="64d7d-163">[ファーム] **列を** 追加すると、各ファームの周りのすべての売上がピボットされます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-163">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="64d7d-164">Type 行 **と Classification** **行を追加** すると、販売されたフルーツと、それがオーガニックかどうかに基づいてデータがさらに分類されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-164">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

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

<span data-ttu-id="64d7d-166">行または列のみを含むピボットテーブルを使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-166">You can also have a PivotTable with only rows or columns.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="64d7d-167">ピボットテーブルにデータ階層を追加する</span><span class="sxs-lookup"><span data-stu-id="64d7d-167">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="64d7d-168">データ階層は、ピボットテーブルに行と列に基づいて結合する情報を入力します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-168">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="64d7d-169">ファームで販売されたクレートと **ク** レートの販売済みホールセールのデータ階層を追加すると、各行と列に対してそれらの数値の合計が表示されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-169">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span>

<span data-ttu-id="64d7d-170">この例では **、Farm と** **Type の両方** が行であり、クレート売上をデータとして使用します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-170">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span>

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

## <a name="pivottable-layouts-and-getting-pivoted-data"></a><span data-ttu-id="64d7d-172">ピボットテーブルレイアウトとピボットデータの取得</span><span class="sxs-lookup"><span data-stu-id="64d7d-172">PivotTable layouts and getting pivoted data</span></span>

<span data-ttu-id="64d7d-173">[PivotLayout は](/javascript/api/excel/excel.pivotlayout)、階層とそのデータの配置を定義します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-173">A [PivotLayout](/javascript/api/excel/excel.pivotlayout) defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="64d7d-174">レイアウトにアクセスして、データが格納される範囲を決定します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-174">You access the layout to determine the ranges where data is stored.</span></span>

<span data-ttu-id="64d7d-175">次の図は、ピボットテーブルの範囲に対応するレイアウト関数呼び出しを示しています。</span><span class="sxs-lookup"><span data-stu-id="64d7d-175">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![レイアウトの取得範囲関数によって返されるピボットテーブルのセクションを示す図。](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a><span data-ttu-id="64d7d-177">ピボットテーブルからデータを取得する</span><span class="sxs-lookup"><span data-stu-id="64d7d-177">Get data from the PivotTable</span></span>

<span data-ttu-id="64d7d-178">レイアウトは、ワークシートでのピボットテーブルの表示方法を定義します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-178">The layout defines how the PivotTable is displayed in the worksheet.</span></span> <span data-ttu-id="64d7d-179">つまり、オブジェクト `PivotLayout` はピボットテーブル要素に使用される範囲を制御します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-179">This means the `PivotLayout` object controls the ranges used for PivotTable elements.</span></span> <span data-ttu-id="64d7d-180">ピボットテーブルによって収集および集計されたデータを取得するには、レイアウトによって提供される範囲を使用します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-180">Use the ranges provided by the layout to get data collected and aggregated by the PivotTable.</span></span> <span data-ttu-id="64d7d-181">特に、ピボット `PivotLayout.getDataBodyRange` テーブルによって生成されたデータにアクセスするために使用します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-181">In particular, use `PivotLayout.getDataBodyRange` to access the data produced by the PivotTable.</span></span>

<span data-ttu-id="64d7d-182">次のコードは、レイアウトを実行してピボットテーブル データの最後の行を取得する方法を示しています (前の例では、[ファームで販売されたクレートの合計] 列と [販売済みクレートの合計] 列の両方の総計)。 </span><span class="sxs-lookup"><span data-stu-id="64d7d-182">The following code demonstrates how to get the last row of the PivotTable data by going through the layout (the **Grand Total** of both the **Sum of Crates Sold at Farm** and **Sum of Crates Sold Wholesale** columns in the earlier example).</span></span> <span data-ttu-id="64d7d-183">これらの値は、セル **E30** (ピボットテーブルの外側) に表示される最終的な合計に合わせて合計されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-183">Those values are then summed together for a final total, which is displayed in cell **E30** (outside of the PivotTable).</span></span>

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

### <a name="layout-types"></a><span data-ttu-id="64d7d-184">レイアウトの種類</span><span class="sxs-lookup"><span data-stu-id="64d7d-184">Layout types</span></span>

<span data-ttu-id="64d7d-185">ピボットテーブルには、コンパクト、アウトライン、表形式の 3 つのレイアウト スタイルがあります。</span><span class="sxs-lookup"><span data-stu-id="64d7d-185">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="64d7d-186">前の例では、コンパクトなスタイルを見ていました。</span><span class="sxs-lookup"><span data-stu-id="64d7d-186">We've seen the compact style in the previous examples.</span></span>

<span data-ttu-id="64d7d-187">次の例では、アウトラインスタイルと表形式スタイルをそれぞれ使用します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-187">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="64d7d-188">コード サンプルは、異なるレイアウト間を切り替える方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="64d7d-188">The code sample shows how to cycle between the different layouts.</span></span>

#### <a name="outline-layout"></a><span data-ttu-id="64d7d-189">アウトライン レイアウト</span><span class="sxs-lookup"><span data-stu-id="64d7d-189">Outline layout</span></span>

![アウトライン レイアウトを使用するピボットテーブル。](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a><span data-ttu-id="64d7d-191">表形式のレイアウト</span><span class="sxs-lookup"><span data-stu-id="64d7d-191">Tabular layout</span></span>

![表形式のレイアウトを使用するピボットテーブル。](../images/excel-pivots-tabular-layout.png)

#### <a name="pivotlayout-type-switch-code-sample"></a><span data-ttu-id="64d7d-193">PivotLayout の種類のスイッチ コードのサンプル</span><span class="sxs-lookup"><span data-stu-id="64d7d-193">PivotLayout type switch code sample</span></span>

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

### <a name="other-pivotlayout-functions"></a><span data-ttu-id="64d7d-194">その他の PivotLayout 関数</span><span class="sxs-lookup"><span data-stu-id="64d7d-194">Other PivotLayout functions</span></span>

<span data-ttu-id="64d7d-195">既定では、ピボットテーブルは必要に応じて行と列のサイズを調整します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-195">By default, PivotTables adjust row and column sizes as needed.</span></span> <span data-ttu-id="64d7d-196">これは、ピボットテーブルが更新された場合に実行されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-196">This is done when the PivotTable is refreshed.</span></span> <span data-ttu-id="64d7d-197">`PivotLayout.autoFormat` その動作を指定します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-197">`PivotLayout.autoFormat` specifies that behavior.</span></span> <span data-ttu-id="64d7d-198">アドインによって行われた行または列のサイズの変更は、次の場合も保持 `autoFormat` されます `false` 。</span><span class="sxs-lookup"><span data-stu-id="64d7d-198">Any row or column size changes made by your add-in persist when `autoFormat` is `false`.</span></span> <span data-ttu-id="64d7d-199">さらに、ピボットテーブルの既定の設定では、ピボットテーブル内のカスタム書式 (塗りつぶしやフォントの変更など) が保持されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-199">Additionally, the default settings of a PivotTable keep any custom formatting in the PivotTable (such as fills and font changes).</span></span> <span data-ttu-id="64d7d-200">更新 `PivotLayout.preserveFormatting` 時 `false` に既定の形式を適用する場合に設定します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-200">Set `PivotLayout.preserveFormatting` to `false` to apply the default format when refreshed.</span></span>

<span data-ttu-id="64d7d-201">また `PivotLayout` 、ヘッダーと行の合計設定、空のデータ セルの表示方法、および代替テキスト オプション [も制御](https://support.microsoft.com/topic/add-alternative-text-to-a-shape-picture-chart-smartart-graphic-or-other-object-44989b2a-903c-4d9a-b742-6a75b451c669) します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-201">A `PivotLayout` also controls header and total row settings, how empty data cells are displayed, and [alt text](https://support.microsoft.com/topic/add-alternative-text-to-a-shape-picture-chart-smartart-graphic-or-other-object-44989b2a-903c-4d9a-b742-6a75b451c669) options.</span></span> <span data-ttu-id="64d7d-202">[PivotLayout 参照は](/javascript/api/excel/excel.pivotlayout)、これらの機能の完全な一覧を提供します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-202">The [PivotLayout](/javascript/api/excel/excel.pivotlayout) reference provides a complete list of these features.</span></span>

> [!NOTE]
> <span data-ttu-id="64d7d-203">ここで説明する PivotLayout 機能の一部は、現在パブリック プレビューでのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-203">Some of the PivotLayout functionality mentioned here is currently only available in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="64d7d-204">次のコード サンプルでは、空のデータ セルに文字列を表示し、本文範囲を一貫性のある水平方向の配置に書式設定し、ピボットテーブルが更新された後も書式設定の変更が維持されます `"--"` 。</span><span class="sxs-lookup"><span data-stu-id="64d7d-204">The following code sample makes empty data cells display the string `"--"`, formats the body range to a consistent horizontal alignment, and ensures that the formatting changes remain even after the PivotTable is refreshed.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="64d7d-205">ピボットテーブルの削除</span><span class="sxs-lookup"><span data-stu-id="64d7d-205">Delete a PivotTable</span></span>

<span data-ttu-id="64d7d-206">ピボットテーブルは、その名前を使用して削除されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-206">PivotTables are deleted by using their name.</span></span>

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="filter-a-pivottable"></a><span data-ttu-id="64d7d-207">ピボットテーブルをフィルター処理する</span><span class="sxs-lookup"><span data-stu-id="64d7d-207">Filter a PivotTable</span></span>

<span data-ttu-id="64d7d-208">ピボットテーブル データをフィルター処理する主な方法は、PivotFilters です。</span><span class="sxs-lookup"><span data-stu-id="64d7d-208">The primary method for filtering PivotTable data is with PivotFilters.</span></span> <span data-ttu-id="64d7d-209">スライサーは、柔軟性の低い代替フィルター方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-209">Slicers offer an alternate, less flexible filtering method.</span></span>

<span data-ttu-id="64d7d-210">[PivotFilters は](/javascript/api/excel/excel.pivotfilters) 、ピボットテーブルの 4 つの階層カテゴリ [(フィルター](#hierarchies) 、列、行、値) に基づいてデータをフィルター処理します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-210">[PivotFilters](/javascript/api/excel/excel.pivotfilters) filter data based on a PivotTable's four [hierarchy categories](#hierarchies) (filters, columns, rows, and values).</span></span> <span data-ttu-id="64d7d-211">PivotFilter には 4 つの種類があります。予定表の日付ベースのフィルター処理、文字列解析、数値比較、およびカスタム入力に基づくフィルター処理が可能です。</span><span class="sxs-lookup"><span data-stu-id="64d7d-211">There are four types of PivotFilters, allowing calendar date-based filtering, string parsing, number comparison, and filtering based on a custom input.</span></span>

<span data-ttu-id="64d7d-212">[スライサー](/javascript/api/excel/excel.slicer) は、ピボットテーブルと通常の Excel テーブルの両方に適用できます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-212">[Slicers](/javascript/api/excel/excel.slicer) can be applied to both PivotTables and regular Excel tables.</span></span> <span data-ttu-id="64d7d-213">ピボットテーブルに適用すると、スライサーは [PivotManualFilter](#pivotmanualfilter) のように機能し、カスタム入力に基づいてフィルター処理を許可します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-213">When applied to a PivotTable, slicers function like a [PivotManualFilter](#pivotmanualfilter) and allow filtering based on a custom input.</span></span> <span data-ttu-id="64d7d-214">PivotFilters とは異なり、スライサーには [Excel UI コンポーネントがあります](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)。</span><span class="sxs-lookup"><span data-stu-id="64d7d-214">Unlike PivotFilters, slicers have an [Excel UI component](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d).</span></span> <span data-ttu-id="64d7d-215">クラスを `Slicer` 使用して、この UI コンポーネントを作成し、フィルター処理を管理し、その外観を制御します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-215">With the `Slicer` class, you create this UI component, manage filtering, and control its visual appearance.</span></span>

### <a name="filter-with-pivotfilters"></a><span data-ttu-id="64d7d-216">PivotFilters を使用したフィルター</span><span class="sxs-lookup"><span data-stu-id="64d7d-216">Filter with PivotFilters</span></span>

<span data-ttu-id="64d7d-217">[PivotFilters を使用](/javascript/api/excel/excel.pivotfilters) すると、4 つの階層カテゴリ [(フィルター](#hierarchies) 、列、行、値) に基づいてピボットテーブル データをフィルター処理できます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-217">[PivotFilters](/javascript/api/excel/excel.pivotfilters) allow you to filter PivotTable data based on the four [hierarchy categories](#hierarchies) (filters, columns, rows, and values).</span></span> <span data-ttu-id="64d7d-218">PivotTable オブジェクト モデルでは、PivotField に適用され、それぞれが 1 つ以上 `PivotFilters` の値を割[](/javascript/api/excel/excel.pivotfield) `PivotField` り当てることができます `PivotFilters` 。</span><span class="sxs-lookup"><span data-stu-id="64d7d-218">In the PivotTable object model, `PivotFilters` are applied to a [PivotField](/javascript/api/excel/excel.pivotfield), and each `PivotField` can have one or more assigned `PivotFilters`.</span></span> <span data-ttu-id="64d7d-219">PivotField にピボットフィルターを適用するには、フィールドの対応する [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) を階層カテゴリに割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="64d7d-219">To apply PivotFilters to a PivotField, the field's corresponding [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) must be assigned to a hierarchy category.</span></span>

#### <a name="types-of-pivotfilters"></a><span data-ttu-id="64d7d-220">PivotFilters の種類</span><span class="sxs-lookup"><span data-stu-id="64d7d-220">Types of PivotFilters</span></span>

| <span data-ttu-id="64d7d-221">フィルターの種類</span><span class="sxs-lookup"><span data-stu-id="64d7d-221">Filter type</span></span> | <span data-ttu-id="64d7d-222">フィルターの目的</span><span class="sxs-lookup"><span data-stu-id="64d7d-222">Filter purpose</span></span> | <span data-ttu-id="64d7d-223">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="64d7d-223">Excel JavaScript API reference</span></span> |
|:--- |:--- |:--- |
| <span data-ttu-id="64d7d-224">DateFilter</span><span class="sxs-lookup"><span data-stu-id="64d7d-224">DateFilter</span></span> | <span data-ttu-id="64d7d-225">予定表の日付ベースのフィルター処理。</span><span class="sxs-lookup"><span data-stu-id="64d7d-225">Calendar date-based filtering.</span></span> | [<span data-ttu-id="64d7d-226">PivotDateFilter</span><span class="sxs-lookup"><span data-stu-id="64d7d-226">PivotDateFilter</span></span>](/javascript/api/excel/excel.pivotdatefilter) |
| <span data-ttu-id="64d7d-227">LabelFilter</span><span class="sxs-lookup"><span data-stu-id="64d7d-227">LabelFilter</span></span> | <span data-ttu-id="64d7d-228">テキスト比較フィルター。</span><span class="sxs-lookup"><span data-stu-id="64d7d-228">Text comparison filtering.</span></span> | [<span data-ttu-id="64d7d-229">PivotLabelFilter</span><span class="sxs-lookup"><span data-stu-id="64d7d-229">PivotLabelFilter</span></span>](/javascript/api/excel/excel.pivotlabelfilter) |
| <span data-ttu-id="64d7d-230">ManualFilter</span><span class="sxs-lookup"><span data-stu-id="64d7d-230">ManualFilter</span></span> | <span data-ttu-id="64d7d-231">カスタム入力フィルター。</span><span class="sxs-lookup"><span data-stu-id="64d7d-231">Custom input filtering.</span></span> | [<span data-ttu-id="64d7d-232">PivotManualFilter</span><span class="sxs-lookup"><span data-stu-id="64d7d-232">PivotManualFilter</span></span>](/javascript/api/excel/excel.pivotmanualfilter) |
| <span data-ttu-id="64d7d-233">ValueFilter</span><span class="sxs-lookup"><span data-stu-id="64d7d-233">ValueFilter</span></span> | <span data-ttu-id="64d7d-234">数値比較フィルター。</span><span class="sxs-lookup"><span data-stu-id="64d7d-234">Number comparison filtering.</span></span> | [<span data-ttu-id="64d7d-235">PivotValueFilter</span><span class="sxs-lookup"><span data-stu-id="64d7d-235">PivotValueFilter</span></span>](/javascript/api/excel/excel.pivotvaluefilter) |

#### <a name="create-a-pivotfilter"></a><span data-ttu-id="64d7d-236">ピボットフィルターの作成</span><span class="sxs-lookup"><span data-stu-id="64d7d-236">Create a PivotFilter</span></span>

<span data-ttu-id="64d7d-237">ピボットテーブル データを (a など) でフィルター処理するには、 `Pivot*Filter` `PivotDateFilter` ピボットフィールドにフィルターを [適用します](/javascript/api/excel/excel.pivotfield)。</span><span class="sxs-lookup"><span data-stu-id="64d7d-237">To filter PivotTable data with a `Pivot*Filter` (such as a `PivotDateFilter`), apply the filter to a [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="64d7d-238">次の 4 つのコード サンプルは、4 種類の PivotFilter のそれぞれを使用する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="64d7d-238">The following four code samples show how to use each of the four types of PivotFilters.</span></span>

##### <a name="pivotdatefilter"></a><span data-ttu-id="64d7d-239">PivotDateFilter</span><span class="sxs-lookup"><span data-stu-id="64d7d-239">PivotDateFilter</span></span>

<span data-ttu-id="64d7d-240">最初のコード サンプルでは [、PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) を Date **Updated** PivotField に適用し **、2020-08-01** より前のデータを非表示にしています。</span><span class="sxs-lookup"><span data-stu-id="64d7d-240">The first code sample applies a [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) to the **Date Updated** PivotField, hiding any data prior to **2020-08-01**.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="64d7d-241">そのフィールドの PivotHierarchy が階層カテゴリに割り当てられていない限り、ピボットフィールドに A を `Pivot*Filter` 適用することはできません。</span><span class="sxs-lookup"><span data-stu-id="64d7d-241">A `Pivot*Filter` can't be applied to a PivotField unless that field's PivotHierarchy is assigned to a hierarchy category.</span></span> <span data-ttu-id="64d7d-242">次のコード サンプルでは、ピボットテーブルをフィルター処理に使用する前に、ピボットテーブルのカテゴリに追加 `dateHierarchy` `rowHierarchies` する必要があります。</span><span class="sxs-lookup"><span data-stu-id="64d7d-242">In the following code sample, the `dateHierarchy` must be added to the PivotTable's `rowHierarchies` category before it can be used for filtering.</span></span>

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
> <span data-ttu-id="64d7d-243">次の 3 つのコード スニペットは、完全な呼び出しではなく、フィルター固有の抜粋のみを表示 `Excel.run` します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-243">The following three code snippets only display filter-specific excerpts, instead of full `Excel.run` calls.</span></span>

##### <a name="pivotlabelfilter"></a><span data-ttu-id="64d7d-244">PivotLabelFilter</span><span class="sxs-lookup"><span data-stu-id="64d7d-244">PivotLabelFilter</span></span>

<span data-ttu-id="64d7d-245">2 番目のコード スニペットは、プロパティを使用して文字 L で始まるラベルを除外して [、PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) を Type **PivotField** に適用する方法 `LabelFilterCondition.beginsWith` を **示しています**。</span><span class="sxs-lookup"><span data-stu-id="64d7d-245">The second code snippet demonstrates how to apply a [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) to the **Type** PivotField, using the `LabelFilterCondition.beginsWith` property to exclude labels that start with the letter **L**.</span></span>

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

##### <a name="pivotmanualfilter"></a><span data-ttu-id="64d7d-246">PivotManualFilter</span><span class="sxs-lookup"><span data-stu-id="64d7d-246">PivotManualFilter</span></span>

<span data-ttu-id="64d7d-247">3 番目のコード スニペットは [、PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)を含む手動フィルターを [分類] フィールドに適用し、分類オーガニック を含むデータをフィルター処理 **します**。</span><span class="sxs-lookup"><span data-stu-id="64d7d-247">The third code snippet applies a manual filter with [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) to the the **Classification** field, filtering out data that doesn't include the classification **Organic**.</span></span>

```js
    // Apply a manual filter to include only a specific PivotItem (the string "Organic").
    var filterField = classHierarchy.fields.getItem("Classification");
    var manualFilter = { selectedItems: ["Organic"] };
    filterField.applyFilter({ manualFilter: manualFilter });
```

##### <a name="pivotvaluefilter"></a><span data-ttu-id="64d7d-248">PivotValueFilter</span><span class="sxs-lookup"><span data-stu-id="64d7d-248">PivotValueFilter</span></span>

<span data-ttu-id="64d7d-249">数値を比較するには、最終的なコード スニペットに示すように [、PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)と値フィルターを使用します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-249">To compare numbers, use a value filter with [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter), as shown in the final code snippet.</span></span> <span data-ttu-id="64d7d-250">ファーム ピボットフィールドのデータと、販売されたクレートの合計が値 `PivotValueFilter` **500** を超えるファームのみを含む、クレート販売済みホールセール ピボットフィールドのデータとを比較します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-250">The `PivotValueFilter` compares the data in the **Farm** PivotField to the data in the **Crates Sold Wholesale** PivotField, including only farms whose sum of crates sold exceeds the value **500**.</span></span>

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

#### <a name="remove-pivotfilters"></a><span data-ttu-id="64d7d-251">ピボットフィルターの削除</span><span class="sxs-lookup"><span data-stu-id="64d7d-251">Remove PivotFilters</span></span>

<span data-ttu-id="64d7d-252">すべての PivotFilter を削除するには、次のコード サンプルに示すように、各 PivotField にメソッド `clearAllFilters` を適用します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-252">To remove all PivotFilters, apply the `clearAllFilters` method to each PivotField, as shown in the following code sample.</span></span>

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

### <a name="filter-with-slicers"></a><span data-ttu-id="64d7d-253">スライサーを使用したフィルター</span><span class="sxs-lookup"><span data-stu-id="64d7d-253">Filter with slicers</span></span>

<span data-ttu-id="64d7d-254">[スライサー](/javascript/api/excel/excel.slicer) を使用すると、Excel ピボットテーブルまたはテーブルからデータをフィルター処理できます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-254">[Slicers](/javascript/api/excel/excel.slicer) allow data to be filtered from an Excel PivotTable or table.</span></span> <span data-ttu-id="64d7d-255">スライサーは、指定した列または PivotField の値を使用して、対応する行をフィルター処理します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-255">A slicer uses values from a specified column or PivotField to filter corresponding rows.</span></span> <span data-ttu-id="64d7d-256">これらの値は、 [に SlicerItem](/javascript/api/excel/excel.sliceritem) オブジェクトとして格納されます `Slicer` 。</span><span class="sxs-lookup"><span data-stu-id="64d7d-256">These values are stored as [SlicerItem](/javascript/api/excel/excel.sliceritem) objects in the `Slicer`.</span></span> <span data-ttu-id="64d7d-257">アドインは、(Excel UI を使用して) ユーザーと同様に[、これらのフィルターを調整できます](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)。</span><span class="sxs-lookup"><span data-stu-id="64d7d-257">Your add-in can adjust these filters, as can users ([through the Excel UI](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)).</span></span> <span data-ttu-id="64d7d-258">次のスクリーンショットに示すように、スライサーは図面レイヤーのワークシートの上に配置されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-258">The slicer sits on top of the worksheet in the drawing layer, as shown in the following screenshot.</span></span>

![ピボットテーブル上のデータをスライサー フィルター処理します。](../images/excel-slicer.png)

> [!NOTE]
> <span data-ttu-id="64d7d-260">このセクションで説明する手法は、ピボットテーブルに接続されたスライサーの使い方に焦点を当ててします。</span><span class="sxs-lookup"><span data-stu-id="64d7d-260">The techniques described in this section focus on how to use slicers connected to PivotTables.</span></span> <span data-ttu-id="64d7d-261">同じ手法は、テーブルに接続されたスライサーの使用にも適用されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-261">The same techniques also apply to using slicers connected to tables.</span></span>

#### <a name="create-a-slicer"></a><span data-ttu-id="64d7d-262">スライサーを作成する</span><span class="sxs-lookup"><span data-stu-id="64d7d-262">Create a slicer</span></span>

<span data-ttu-id="64d7d-263">メソッドまたはメソッドを使用して、ブックまたはワークシートにスライサー `Workbook.slicers.add` を作成 `Worksheet.slicers.add` できます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-263">You can create a slicer in a workbook or worksheet by using the `Workbook.slicers.add` method or `Worksheet.slicers.add` method.</span></span> <span data-ttu-id="64d7d-264">指定したオブジェクトまたはオブジェクトの [SlicerCollection](/javascript/api/excel/excel.slicercollection) にスライサーを `Workbook` 追加 `Worksheet` します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-264">Doing so adds a slicer to the [SlicerCollection](/javascript/api/excel/excel.slicercollection) of the specified `Workbook` or `Worksheet` object.</span></span> <span data-ttu-id="64d7d-265">メソッド `SlicerCollection.add` には、次の 3 つのパラメーターがあります。</span><span class="sxs-lookup"><span data-stu-id="64d7d-265">The `SlicerCollection.add` method has three parameters:</span></span>

- <span data-ttu-id="64d7d-266">`slicerSource`: 新しいスライサーが基づくデータ ソース。</span><span class="sxs-lookup"><span data-stu-id="64d7d-266">`slicerSource`: The data source on which the new slicer is based.</span></span> <span data-ttu-id="64d7d-267">名前または ID を表す 、 、または文字列を指定 `PivotTable` `Table` `PivotTable` できます `Table` 。</span><span class="sxs-lookup"><span data-stu-id="64d7d-267">It can be a `PivotTable`, `Table`, or string representing the name or ID of a `PivotTable` or `Table`.</span></span>
- <span data-ttu-id="64d7d-268">`sourceField`: フィルター処理するデータ ソースのフィールド。</span><span class="sxs-lookup"><span data-stu-id="64d7d-268">`sourceField`: The field in the data source by which to filter.</span></span> <span data-ttu-id="64d7d-269">名前または ID を表す 、 、または文字列を指定 `PivotField` `TableColumn` `PivotField` できます `TableColumn` 。</span><span class="sxs-lookup"><span data-stu-id="64d7d-269">It can be a `PivotField`, `TableColumn`, or string representing the name or ID of a `PivotField` or `TableColumn`.</span></span>
- <span data-ttu-id="64d7d-270">`slicerDestination`: 新しいスライサーが作成されるワークシート。</span><span class="sxs-lookup"><span data-stu-id="64d7d-270">`slicerDestination`: The worksheet where the new slicer will be created.</span></span> <span data-ttu-id="64d7d-271">オブジェクト、または `Worksheet` . の名前または ID を指定できます `Worksheet` 。</span><span class="sxs-lookup"><span data-stu-id="64d7d-271">It can be a `Worksheet` object or the name or ID of a `Worksheet`.</span></span> <span data-ttu-id="64d7d-272">を使用してアクセスする場合 `SlicerCollection` 、このパラメーターは不要です `Worksheet.slicers` 。</span><span class="sxs-lookup"><span data-stu-id="64d7d-272">This parameter is unnecessary when the `SlicerCollection` is accessed through `Worksheet.slicers`.</span></span> <span data-ttu-id="64d7d-273">この場合、コレクションのワークシートが移動先として使用されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-273">In this case, the collection's worksheet is used as the destination.</span></span>

<span data-ttu-id="64d7d-274">次のコード サンプルでは、ピボット ワークシートに新しいスライサー **を追加** します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-274">The following code sample adds a new slicer to the **Pivot** worksheet.</span></span> <span data-ttu-id="64d7d-275">スライサーのソースは **、Farm Sales ピボット** テーブルであり、Type データを使用して **フィルター処理** します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-275">The slicer's source is the **Farm Sales** PivotTable and filters using the **Type** data.</span></span> <span data-ttu-id="64d7d-276">スライサーは、将来の参照 **のために、Fruit Slicer という** 名前も付けます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-276">The slicer is also named **Fruit Slicer** for future reference.</span></span>

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

#### <a name="filter-items-with-a-slicer"></a><span data-ttu-id="64d7d-277">スライサーを使用してアイテムをフィルター処理する</span><span class="sxs-lookup"><span data-stu-id="64d7d-277">Filter items with a slicer</span></span>

<span data-ttu-id="64d7d-278">スライサーはピボットテーブルにフィルターを適用し、. `sourceField`</span><span class="sxs-lookup"><span data-stu-id="64d7d-278">The slicer filters the PivotTable with items from the `sourceField`.</span></span> <span data-ttu-id="64d7d-279">この `Slicer.selectItems` メソッドは、スライサーに残るアイテムを設定します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-279">The `Slicer.selectItems` method sets the items that remain in the slicer.</span></span> <span data-ttu-id="64d7d-280">これらのアイテムは、アイテムのキーを表す `string[]` 、 としてメソッドに渡されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-280">These items are passed to the method as a `string[]`, representing the keys of the items.</span></span> <span data-ttu-id="64d7d-281">これらのアイテムを含む行は、ピボットテーブルの集計に残ります。</span><span class="sxs-lookup"><span data-stu-id="64d7d-281">Any rows containing those items remain in the PivotTable's aggregation.</span></span> <span data-ttu-id="64d7d-282">以降の呼び `selectItems` 出しでは、リストをそれらの呼び出しで指定されたキーに設定します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-282">Subsequent calls to `selectItems` set the list to the keys specified in those calls.</span></span>

> [!NOTE]
> <span data-ttu-id="64d7d-283">データ ソースに含めされていないアイテムが渡された場合は `Slicer.selectItems` 、 `InvalidArgument` エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-283">If `Slicer.selectItems` is passed an item that's not in the data source, an `InvalidArgument` error is thrown.</span></span> <span data-ttu-id="64d7d-284">コンテンツは `Slicer.slicerItems` [、SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)であるプロパティを通じて確認できます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-284">The contents can be verified through the `Slicer.slicerItems` property, which is a [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).</span></span>

<span data-ttu-id="64d7d-285">次のコード サンプルは、スライサーで選択されている 3つの項目を示 **しています。**</span><span class="sxs-lookup"><span data-stu-id="64d7d-285">The following code sample shows three items being selected for the slicer: **Lemon**, **Lime**, and **Orange**.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

<span data-ttu-id="64d7d-286">スライサーからすべてのフィルターを削除するには、次のサンプルに示 `Slicer.clearFilters` すように、メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-286">To remove all filters from the slicer, use the `Slicer.clearFilters` method, as shown in the following sample.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

#### <a name="style-and-format-a-slicer"></a><span data-ttu-id="64d7d-287">スライサーのスタイルと書式設定</span><span class="sxs-lookup"><span data-stu-id="64d7d-287">Style and format a slicer</span></span>

<span data-ttu-id="64d7d-288">アドインは、プロパティを使用してスライサーの表示設定を調整 `Slicer` できます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-288">You add-in can adjust a slicer's display settings through `Slicer` properties.</span></span> <span data-ttu-id="64d7d-289">次のコード サンプルでは、スタイルを **SlicerStyleLight6** に設定し、スライサーの上部にあるテキストを **[フルーツ** の種類] に設定し、スライサーを描画レイヤーの **位置 (395、 15)** に配置し、スライサーのサイズを **135x150** ピクセルに設定します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-289">The following code sample sets the style to **SlicerStyleLight6**, sets the text at the top of the slicer to **Fruit Types**, places the slicer at the position **(395, 15)** on the drawing layer, and sets the slicer's size to **135x150** pixels.</span></span>

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

#### <a name="delete-a-slicer"></a><span data-ttu-id="64d7d-290">スライサーを削除する</span><span class="sxs-lookup"><span data-stu-id="64d7d-290">Delete a slicer</span></span>

<span data-ttu-id="64d7d-291">スライサーを削除するには、メソッドを呼び出 `Slicer.delete` します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-291">To delete a slicer, call the `Slicer.delete` method.</span></span> <span data-ttu-id="64d7d-292">次のコード サンプルでは、現在のワークシートから最初のスライサーを削除します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-292">The following code sample deletes the first slicer from the current worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a><span data-ttu-id="64d7d-293">変更集約関数</span><span class="sxs-lookup"><span data-stu-id="64d7d-293">Change aggregation function</span></span>

<span data-ttu-id="64d7d-294">データ階層の値は集計されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-294">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="64d7d-295">数値のデータセットの場合、これは既定では合計です。</span><span class="sxs-lookup"><span data-stu-id="64d7d-295">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="64d7d-296">この `summarizeBy` プロパティは、AggregationFunction 型に基づいて [この動作を定義](/javascript/api/excel/excel.aggregationfunction) します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-296">The `summarizeBy` property defines this behavior based on an [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) type.</span></span>

<span data-ttu-id="64d7d-297">現在サポートされている集計関数の種類は `Sum` `Count` 、、、( `Average` `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP` `Automatic` 既定値) です。</span><span class="sxs-lookup"><span data-stu-id="64d7d-297">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="64d7d-298">次のコード サンプルでは、集計をデータの平均に変更します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-298">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="64d7d-299">ShowAsRule を使用して計算を変更する</span><span class="sxs-lookup"><span data-stu-id="64d7d-299">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="64d7d-300">ピボットテーブルは、既定では、行階層と列階層のデータを個別に集計します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-300">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="64d7d-301">[ShowAsRule は](/javascript/api/excel/excel.showasrule)、ピボットテーブル内の他のアイテムに基づいてデータ階層を出力値に変更します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-301">A [ShowAsRule](/javascript/api/excel/excel.showasrule) changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="64d7d-302">オブジェクト `ShowAsRule` には、次の 3 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="64d7d-302">The `ShowAsRule` object has three properties:</span></span>

- <span data-ttu-id="64d7d-303">`calculation`: データ階層に適用する相対計算の種類 (既定値は `none` ) です。</span><span class="sxs-lookup"><span data-stu-id="64d7d-303">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
- <span data-ttu-id="64d7d-304">`baseField`: 計算が適用される前の基本データを含む階層内の[PivotField。](/javascript/api/excel/excel.pivotfield)</span><span class="sxs-lookup"><span data-stu-id="64d7d-304">`baseField`: The [PivotField](/javascript/api/excel/excel.pivotfield) in the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="64d7d-305">Excel PivotTables には、階層からフィールドへの 1 対 1 のマッピングが含まれていますので、同じ名前を使用して階層とフィールドの両方にアクセスします。</span><span class="sxs-lookup"><span data-stu-id="64d7d-305">Since Excel PivotTables have a one-to-one mapping of hierarchy to field, you'll use the same name to access both the hierarchy and the field.</span></span>
- <span data-ttu-id="64d7d-306">`baseItem`: 計算の種類に基づいて基本フィールドの値と比較される個々の[PivotItem。](/javascript/api/excel/excel.pivotitem)</span><span class="sxs-lookup"><span data-stu-id="64d7d-306">`baseItem`: The individual [PivotItem](/javascript/api/excel/excel.pivotitem) compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="64d7d-307">すべての計算でこのフィールドが必要な場合があります。</span><span class="sxs-lookup"><span data-stu-id="64d7d-307">Not all calculations require this field.</span></span>

<span data-ttu-id="64d7d-308">次の使用例は、ファームデータ階層で販売されたクレートの合計の計算を、列の合計に対する割合に設定します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-308">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span>
<span data-ttu-id="64d7d-309">この場合も、粒度をフルーツの種類レベルまで拡張する必要があります。そのため **、Type** 行階層とその基になるフィールドを使用します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-309">We still want the granularity to extend to the fruit type level, so we'll use the **Type** row hierarchy and its underlying field.</span></span>
<span data-ttu-id="64d7d-310">この例では、 **最初** の行階層として Farm も含まれています。そのため、ファームの合計エントリには、各ファームが生成する割合も表示されます。</span><span class="sxs-lookup"><span data-stu-id="64d7d-310">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

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

<span data-ttu-id="64d7d-312">前の使用例は、個々の行階層のフィールドを基準として、列に計算を設定します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-312">The previous example set the calculation to the column, relative to the field of an individual row hierarchy.</span></span> <span data-ttu-id="64d7d-313">計算が個々のアイテムに関連する場合は、プロパティを使用 `baseItem` します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-313">When the calculation relates to an individual item, use the `baseItem` property.</span></span>

<span data-ttu-id="64d7d-314">次の例は、計算を示 `differenceFrom` しています。</span><span class="sxs-lookup"><span data-stu-id="64d7d-314">The following example shows the `differenceFrom` calculation.</span></span> <span data-ttu-id="64d7d-315">ファームクレート販売データ階層エントリの違いを **、A Farms** のエントリと相対的に表示します。</span><span class="sxs-lookup"><span data-stu-id="64d7d-315">It displays the difference of the farm crate sales data hierarchy entries relative to those of **A Farms**.</span></span>
<span data-ttu-id="64d7d-316">is Farm です。したがって、他のファーム間の違い、および同様のフルーツの種類ごとに内訳が表示されます `baseField` (この例では **、Type** も行階層です)。</span><span class="sxs-lookup"><span data-stu-id="64d7d-316">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

!["A Farms" と他のファームの果物販売の違いを示すピボットテーブル。](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="change-hierarchy-names"></a><span data-ttu-id="64d7d-320">階層名の変更</span><span class="sxs-lookup"><span data-stu-id="64d7d-320">Change hierarchy names</span></span>

<span data-ttu-id="64d7d-321">階層フィールドは編集可能です。</span><span class="sxs-lookup"><span data-stu-id="64d7d-321">Hierarchy fields are editable.</span></span> <span data-ttu-id="64d7d-322">次のコードは、2 つのデータ階層の表示名を変更する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="64d7d-322">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="64d7d-323">関連項目</span><span class="sxs-lookup"><span data-stu-id="64d7d-323">See also</span></span>

- [<span data-ttu-id="64d7d-324">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="64d7d-324">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="64d7d-325">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="64d7d-325">Excel JavaScript API Reference</span></span>](/javascript/api/excel)
