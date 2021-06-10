---
title: Excel JavaScript API を使用してブックを操作する
description: JavaScript API を使用してブックまたはアプリケーション レベルの機能で一般的なタスクを実行するExcel説明します。
ms.date: 06/07/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 48ceb882a7beea3fa3ca08216f3ee1dd82ba4fa9
ms.sourcegitcommit: 5a151d4df81e5640363774406d0f329d6a0d3db8
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/09/2021
ms.locfileid: "52853984"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="b9307-103">Excel JavaScript API を使用してブックを操作する</span><span class="sxs-lookup"><span data-stu-id="b9307-103">Work with workbooks using the Excel JavaScript API</span></span>

<span data-ttu-id="b9307-104">この記事では、Excel JavaScript API を使用して、ブックでタスクを実行する方法のコード サンプルを示しています。</span><span class="sxs-lookup"><span data-stu-id="b9307-104">This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API.</span></span> <span data-ttu-id="b9307-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Workbook` [、「Workbook オブジェクト (JavaScript API for Excel)」を参照してください](/javascript/api/excel/excel.workbook)。</span><span class="sxs-lookup"><span data-stu-id="b9307-105">For the complete list of properties and methods that the `Workbook` object supports, see [Workbook Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="b9307-106">この記事では、[Application](/javascript/api/excel/excel.application) オブジェクトを使用して実行するブック レベルのアクションについても説明します。</span><span class="sxs-lookup"><span data-stu-id="b9307-106">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="b9307-107">Workbook オブジェクトは、Excel を操作するアドインのエントリ ポイントです。</span><span class="sxs-lookup"><span data-stu-id="b9307-107">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="b9307-108">このオブジェクトは、Excel データのアクセスや変更に使用するワークシート、テーブル、ピボットテーブル、その他のコレクションを保持します。</span><span class="sxs-lookup"><span data-stu-id="b9307-108">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="b9307-109">[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) オブジェクトは、個々のワークシートを使用して、ブックのすべてのデータにアドインからアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="b9307-109">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through individual worksheets.</span></span> <span data-ttu-id="b9307-110">具体的には、アドインからワークシートの追加、ワークシート間の移動、ワークシート イベントへのハンドラーの割り当てができます。</span><span class="sxs-lookup"><span data-stu-id="b9307-110">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="b9307-111">ワークシートへのアクセスと編集の方法については、「[Excel JavaScript API を使用してワークシートを操作する](excel-add-ins-worksheets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b9307-111">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="b9307-112">アクティブ セルまたは選択した範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="b9307-112">Get the active cell or selected range</span></span>

<span data-ttu-id="b9307-113">Workbook オブジェクトには、ユーザーまたはアドインが選択したセルの範囲を取得する 2 つのメソッド `getActiveCell()` と `getSelectedRange()` があります。</span><span class="sxs-lookup"><span data-stu-id="b9307-113">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="b9307-114">`getActiveCell()` はブックからアクティブ セルを [Range オブジェクト](/javascript/api/excel/excel.range)として取得します。</span><span class="sxs-lookup"><span data-stu-id="b9307-114">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="b9307-115">次の例では、`getActiveCell()` を呼び出し、コンソールにセルのアドレスを表示します。</span><span class="sxs-lookup"><span data-stu-id="b9307-115">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="b9307-116">`getSelectedRange()` メソッドは現在選択されている単一の範囲を返します。</span><span class="sxs-lookup"><span data-stu-id="b9307-116">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="b9307-117">複数の範囲が選択されている場合は、InvalidSelection エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="b9307-117">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="b9307-118">次の例では、`getSelectedRange()` を呼び出し、範囲の塗りつぶし色を黄色に設定します。</span><span class="sxs-lookup"><span data-stu-id="b9307-118">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="b9307-119">ブックを作成する</span><span class="sxs-lookup"><span data-stu-id="b9307-119">Create a workbook</span></span>

<span data-ttu-id="b9307-120">アドインでは、アドインが現在実行されている Excel のインスタンスとは異なる新しいブックを作成できます。</span><span class="sxs-lookup"><span data-stu-id="b9307-120">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="b9307-121">Excel オブジェクトには、この目的の `createWorkbook` メソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="b9307-121">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="b9307-122">このメソッドが呼び出されると、新しいブックが Excel の新しいインスタンスですぐに開いて表示されます。</span><span class="sxs-lookup"><span data-stu-id="b9307-122">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="b9307-123">アドインは前のブックで開いて実行されたままになります。</span><span class="sxs-lookup"><span data-stu-id="b9307-123">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="b9307-124">`createWorkbook` メソッドは既存のブックのコピーの作成もできます。</span><span class="sxs-lookup"><span data-stu-id="b9307-124">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="b9307-125">このメソッドは、オプションのパラメーターとして .xlsx ファイルの base64 エンコード文字列表現を受け取ります。</span><span class="sxs-lookup"><span data-stu-id="b9307-125">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="b9307-126">文字列の引数は有効な .xlsx ファイルと見なされ、作成されるブックはそのファイルのコピーになります。</span><span class="sxs-lookup"><span data-stu-id="b9307-126">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="b9307-127">ファイルスライスを使用すると、アドインの現在のブックを base64 エンコード文字列 [として取得できます](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)。</span><span class="sxs-lookup"><span data-stu-id="b9307-127">You can get your add-in's current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="b9307-128">次の例に示すように、[FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) クラスを使用して、ファイルを必要な base64 エンコード文字列に変換できます。</span><span class="sxs-lookup"><span data-stu-id="b9307-128">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

```js
// Retrieve the external workbook file and set up a `FileReader` object. 
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // Remove the metadata before the base64-encoded string.
        var startIndex = reader.result.toString().indexOf("base64,");
        var externalWorkbook = reader.result.toString().substr(startIndex + 7);

        Excel.createWorkbook(externalWorkbook);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// Read the file as a data URL so we can parse the base64-encoded string.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one-preview"></a><span data-ttu-id="b9307-129">既存のブックのコピーを現在のブックに挿入する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="b9307-129">Insert a copy of an existing workbook into the current one (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="b9307-130">この `Workbook.insertWorksheetsFromBase64` メソッドは現在、パブリック プレビューでのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="b9307-130">The `Workbook.insertWorksheetsFromBase64` method is currently only available in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

<span data-ttu-id="b9307-131">前の例は、既存のブックから作成された新しいブックを示しています。</span><span class="sxs-lookup"><span data-stu-id="b9307-131">The previous example shows a new workbook being created from an existing workbook.</span></span> <span data-ttu-id="b9307-132">既存のブックの一部またはすべてを、アドインに関連付けられているブックにコピーすることもできます。</span><span class="sxs-lookup"><span data-stu-id="b9307-132">You can also copy some or all of an existing workbook into the one currently associated with your add-in.</span></span> <span data-ttu-id="b9307-133">ブック [には](/javascript/api/excel/excel.workbook) 、ターゲット ブックのワークシートのコピーを自体に `insertWorksheetsFromBase64` 挿入するメソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="b9307-133">A [Workbook](/javascript/api/excel/excel.workbook) has the `insertWorksheetsFromBase64` method to insert copies of the target workbook's worksheets into itself.</span></span> <span data-ttu-id="b9307-134">他のブックのファイルは、呼び出しと同様に、base64 エンコードされた文字列として渡 `Excel.createWorkbook` されます。</span><span class="sxs-lookup"><span data-stu-id="b9307-134">The other workbook's file is passed as a base64-encoded string, just like the `Excel.createWorkbook` call.</span></span> 

```TypeScript
insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions): OfficeExtension.ClientResult<string[]>;
```

> [!IMPORTANT]
> <span data-ttu-id="b9307-135">この `insertWorksheetsFromBase64` メソッドは、Excel Mac、Windows Web 上でサポートされています。</span><span class="sxs-lookup"><span data-stu-id="b9307-135">The `insertWorksheetsFromBase64` method is supported for Excel on Windows, Mac, and the web.</span></span> <span data-ttu-id="b9307-136">iOS ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b9307-136">It's not supported for iOS.</span></span> <span data-ttu-id="b9307-137">さらに、このExcel on the webピボットテーブル、グラフ、コメント、またはスライサー要素を持つソース ワークシートはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b9307-137">Additionally, in Excel on the web this method doesn't support source worksheets with PivotTable, Chart, Comment, or Slicer elements.</span></span> <span data-ttu-id="b9307-138">これらのオブジェクトが存在する場合、 `insertWorksheetsFromBase64` メソッドはエラーを返 `UnsupportedFeature` Excel on the web。</span><span class="sxs-lookup"><span data-stu-id="b9307-138">If those objects are present, the `insertWorksheetsFromBase64` method returns the `UnsupportedFeature` error in Excel on the web.</span></span> 

<span data-ttu-id="b9307-139">次のコード サンプルは、別のブックから現在のブックにワークシートを挿入する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="b9307-139">The following code sample shows how to insert worksheets from another workbook into the current workbook.</span></span> <span data-ttu-id="b9307-140">このコード サンプルでは、まずオブジェクトを使用してブック ファイルを処理し、base64 エンコードされた文字列を抽出し、次にこの base64 エンコードされた文字列を現在のブック [`FileReader`](https://developer.mozilla.org/docs/Web/API/FileReader) に挿入します。</span><span class="sxs-lookup"><span data-stu-id="b9307-140">This code sample first processes a workbook file with a [`FileReader`](https://developer.mozilla.org/docs/Web/API/FileReader) object and extracts a base64-encoded string, and then it inserts this base64-encoded string into the current workbook.</span></span> <span data-ttu-id="b9307-141">新しいワークシートは、Sheet1 という名前のワークシートの後 **に挿入されます**。</span><span class="sxs-lookup"><span data-stu-id="b9307-141">The new worksheets are inserted after the worksheet named **Sheet1**.</span></span> <span data-ttu-id="b9307-142">`[]` [InsertWorksheetOptions.sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetNamesToInsert)プロパティのパラメーターとして渡されます。</span><span class="sxs-lookup"><span data-stu-id="b9307-142">Note that `[]` is passed as the parameter for the [InsertWorksheetOptions.sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetNamesToInsert) property.</span></span> <span data-ttu-id="b9307-143">つまり、ターゲット ブックのすべてのワークシートが現在のブックに挿入されます。</span><span class="sxs-lookup"><span data-stu-id="b9307-143">This means that all the worksheets from the target workbook are inserted into the current workbook.</span></span>

```js
// Retrieve the external workbook file and set up a `FileReader` object. 
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // Remove the metadata before the base64-encoded string.
        var startIndex = reader.result.toString().indexOf("base64,");
        var externalWorkbook = reader.result.toString().substr(startIndex + 7);
            
        // Retrieve the current workbook.
        var workbook = context.workbook;
            
        // Set up the insert options. 
        var options = { 
            sheetNamesToInsert: [], // Insert all the worksheets from the source workbook.
            positionType: Excel.WorksheetPositionType.after, // Insert after the `relativeTo` sheet.
            relativeTo: "Sheet1" // The sheet relative to which the other worksheets will be inserted. Used with `positionType`.
        }; 
            
         // Insert the new worksheets into the current workbook.
         workbook.insertWorksheetsFromBase64(externalWorkbook, options);
         return context.sync();
    });
};

// Read the file as a data URL so we can parse the base64-encoded string.
reader.readAsDataURL(myFile.files[0]);
```

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="b9307-144">ブックのシート構成を保護する</span><span class="sxs-lookup"><span data-stu-id="b9307-144">Protect the workbook's structure</span></span>

<span data-ttu-id="b9307-145">アドインでは、ブックのシート構成を編集するユーザーの機能を制御できます。</span><span class="sxs-lookup"><span data-stu-id="b9307-145">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="b9307-146">Workbook オブジェクトの `protection` プロパティは [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) オブジェクトであり、`protect()` メソッドを備えています。</span><span class="sxs-lookup"><span data-stu-id="b9307-146">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="b9307-147">次の例では、ブックのシート構成の保護を切り替える基本的なシナリオを示します。</span><span class="sxs-lookup"><span data-stu-id="b9307-147">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span>

```js
Excel.run(function (context) {
    var workbook = context.workbook;
    workbook.load("protection/protected");

    return context.sync().then(function() {
        if (!workbook.protection.protected) {
            workbook.protection.protect();
        }
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="b9307-148">`protect` メソッドは、オプションの文字列パラメーターを受け取ります。</span><span class="sxs-lookup"><span data-stu-id="b9307-148">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="b9307-149">この文字列は、ユーザーが保護をバイパスしてブックのシート構成を変更するために必要なパスワードを表します。</span><span class="sxs-lookup"><span data-stu-id="b9307-149">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="b9307-150">保護は、不必要なデータ編集をできないようにするため、ワークシート レベルで設定することもできます。</span><span class="sxs-lookup"><span data-stu-id="b9307-150">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="b9307-151">詳細については、「[Excel JavaScript API を使用してワークシートを操作する](excel-add-ins-worksheets.md#data-protection)」の **データの保護** のセクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="b9307-151">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE]
> <span data-ttu-id="b9307-152">Excel のブックの保護の詳細については、「[ブックを保護する](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b9307-152">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="b9307-153">ドキュメント プロパティへのアクセス</span><span class="sxs-lookup"><span data-stu-id="b9307-153">Access document properties</span></span>

<span data-ttu-id="b9307-154">Workbook オブジェクトは、[ドキュメント プロパティ](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75)と呼ばれる Office ファイルのメタデータにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="b9307-154">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="b9307-155">Workbook オブジェクトの `properties` プロパティは、これらのメタデータ値を含む [DocumentProperties](/javascript/api/excel/excel.documentproperties) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="b9307-155">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="b9307-156">次の例は、プロパティを設定する方法を示 `author` しています。</span><span class="sxs-lookup"><span data-stu-id="b9307-156">The following example shows how to set the `author` property.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="custom-properties"></a><span data-ttu-id="b9307-157">カスタム プロパティ</span><span class="sxs-lookup"><span data-stu-id="b9307-157">Custom properties</span></span>

<span data-ttu-id="b9307-158">カスタム プロパティを定義することもできます。</span><span class="sxs-lookup"><span data-stu-id="b9307-158">You can also define custom properties.</span></span> <span data-ttu-id="b9307-159">DocumentProperties オブジェクトには `custom` プロパティが含まれていて、ユーザー定義プロパティのキー値のペアのコレクションを表します。</span><span class="sxs-lookup"><span data-stu-id="b9307-159">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="b9307-160">次の例では、"Hello" という値を持つ **Introduction** という名前のカスタム プロパティを作成し、それを取得する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="b9307-160">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

```js
Excel.run(function (context) {
    var customDocProperties = context.workbook.properties.custom;
    customDocProperties.add("Introduction", "Hello");
    return context.sync();
}).catch(errorHandlerFunction);

[...]

Excel.run(function (context) {
    var customDocProperties = context.workbook.properties.custom;
    var customProperty = customDocProperties.getItem("Introduction");
    customProperty.load(["key, value"]);

    return context.sync().then(function() {
        console.log("Custom key  : " + customProperty.key); // "Introduction"
        console.log("Custom value : " + customProperty.value); // "Hello"
    });
}).catch(errorHandlerFunction);
```

#### <a name="worksheet-level-custom-properties"></a><span data-ttu-id="b9307-161">ワークシート レベルのカスタム プロパティ</span><span class="sxs-lookup"><span data-stu-id="b9307-161">Worksheet-level custom properties</span></span>

<span data-ttu-id="b9307-162">カスタム プロパティは、ワークシート レベルで設定することもできます。</span><span class="sxs-lookup"><span data-stu-id="b9307-162">Custom properties can also be set at the worksheet level.</span></span> <span data-ttu-id="b9307-163">これらはドキュメント レベルのカスタム プロパティに似ていますが、異なるワークシートで同じキーを繰り返す場合があります。</span><span class="sxs-lookup"><span data-stu-id="b9307-163">These are similar to document-level custom properties, except that the same key can be repeated across different worksheets.</span></span> <span data-ttu-id="b9307-164">次の例は、現在のワークシートに値 **"Alpha"** を指定して WorksheetGroup という名前のカスタム プロパティを作成し、それを取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="b9307-164">The following example shows how to create a custom property named **WorksheetGroup** with the value "Alpha" on the current worksheet, then retrieve it.</span></span>

```js
Excel.run(function (context) {
    // Add the custom property.
    var customWorksheetProperties = context.workbook.worksheets.getActiveWorksheet().customProperties;
    customWorksheetProperties.add("WorksheetGroup", "Alpha");

    return context.sync();
}).catch(errorHandlerFunction);

[...]

Excel.run(function (context) {
    // Load the keys and values of all custom properties in the current worksheet.
    var worksheet = context.workbook.worksheets.getActiveWorksheet();
    worksheet.load("name");

    var customWorksheetProperties = worksheet.customProperties;
    var customWorksheetProperty = customWorksheetProperties.getItem("WorksheetGroup");
    customWorksheetProperty.load(["key", "value"]);

    return context.sync().then(function() {
        // Log the WorksheetGroup custom property to the console.
        console.log(worksheet.name + ": " + customWorksheetProperty.key); // "WorksheetGroup"
        console.log("  Custom value : " + customWorksheetProperty.value); // "Alpha"
    });
}).catch(errorHandlerFunction);
```

## <a name="access-document-settings"></a><span data-ttu-id="b9307-165">ドキュメント設定へのアクセス</span><span class="sxs-lookup"><span data-stu-id="b9307-165">Access document settings</span></span>

<span data-ttu-id="b9307-166">ブックの設定は、カスタム プロパティのコレクションに似ています。</span><span class="sxs-lookup"><span data-stu-id="b9307-166">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="b9307-167">設定は 1 つの Excel ファイルとアドインのペアリングに固有であるのに対して、プロパティはファイルに接続しているだけである点が異なります。</span><span class="sxs-lookup"><span data-stu-id="b9307-167">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="b9307-168">次の例は、設定を作成してアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="b9307-168">The following example shows how to create and access a setting.</span></span>

```js
Excel.run(function (context) {
    var settings = context.workbook.settings;
    settings.add("NeedsReview", true);
    var needsReview = settings.getItem("NeedsReview");
    needsReview.load("value");

    return context.sync().then(function() {
        console.log("Workbook needs review : " + needsReview.value);
    });
}).catch(errorHandlerFunction);
```

## <a name="access-application-culture-settings"></a><span data-ttu-id="b9307-169">アプリケーション カルチャの設定にアクセスする</span><span class="sxs-lookup"><span data-stu-id="b9307-169">Access application culture settings</span></span>

<span data-ttu-id="b9307-170">ブックには、特定のデータの表示方法に影響する言語とカルチャの設定があります。</span><span class="sxs-lookup"><span data-stu-id="b9307-170">A workbook has language and culture settings that affect how certain data is displayed.</span></span> <span data-ttu-id="b9307-171">これらの設定は、アドインのユーザーが異なる言語やカルチャ間でブックを共有している場合にデータをローカライズするのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="b9307-171">These settings can help localize data when your add-in's users are sharing workbooks across different languages and cultures.</span></span> <span data-ttu-id="b9307-172">アドインでは、文字列解析を使用して、システム カルチャ設定に基づいて数値、日付、時刻の形式をローカライズし、各ユーザーが独自のカルチャの形式でデータを表示できます。</span><span class="sxs-lookup"><span data-stu-id="b9307-172">Your add-in can use string parsing to localize the format of numbers, dates, and times based on the system culture settings so that each user sees data in their own culture's format.</span></span>

<span data-ttu-id="b9307-173">`Application.cultureInfo` システム カルチャ設定を [CultureInfo オブジェクトとして定義](/javascript/api/excel/excel.cultureinfo) します。</span><span class="sxs-lookup"><span data-stu-id="b9307-173">`Application.cultureInfo` defines the system culture settings as a [CultureInfo](/javascript/api/excel/excel.cultureinfo) object.</span></span> <span data-ttu-id="b9307-174">これには、数値の小数点記号や日付形式のような設定が含まれる。</span><span class="sxs-lookup"><span data-stu-id="b9307-174">This contains settings like the numerical decimal separator or the date format.</span></span>

<span data-ttu-id="b9307-175">一部のカルチャ設定は[、UI を使用Excelできます](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e)。</span><span class="sxs-lookup"><span data-stu-id="b9307-175">Some culture settings can be [changed through the Excel UI](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e).</span></span> <span data-ttu-id="b9307-176">システム設定はオブジェクトに保持 `CultureInfo` されます。</span><span class="sxs-lookup"><span data-stu-id="b9307-176">The system settings are preserved in the `CultureInfo` object.</span></span> <span data-ttu-id="b9307-177">ローカルの変更は、アプリケーション レベル [の](/javascript/api/excel/excel.application)プロパティ (など) として保持されます `Application.decimalSeparator` 。</span><span class="sxs-lookup"><span data-stu-id="b9307-177">Any local changes are kept as [Application](/javascript/api/excel/excel.application)-level properties, such as `Application.decimalSeparator`.</span></span>

<span data-ttu-id="b9307-178">次のサンプルでは、数値文字列の小数点記号を ',' からシステム設定で使用される文字に変更します。</span><span class="sxs-lookup"><span data-stu-id="b9307-178">The following sample changes the decimal separator character of a numerical string from a ',' to the character used by the system settings.</span></span>

```js
// This will convert a number like "14,37" to "14.37"
// (assuming the system decimal separator is ".").
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var decimalSource = sheet.getRange("B2");
    decimalSource.load("values");
    context.application.cultureInfo.numberFormat.load("numberDecimalSeparator");

    return context.sync().then(function() {
        var systemDecimalSeparator =
            context.application.cultureInfo.numberFormat.numberDecimalSeparator;
        var oldDecimalString = decimalSource.values[0][0];

        // This assumes the input column is standardized to use "," as the decimal separator.
        var newDecimalString = oldDecimalString.replace(",", systemDecimalSeparator);

        var resultRange = sheet.getRange("C2");
        resultRange.values = [[newDecimalString]];
        resultRange.format.autofitColumns();
        return context.sync();
    });
});
```

## <a name="add-custom-xml-data-to-the-workbook"></a><span data-ttu-id="b9307-179">カスタム XML データをブックに追加する</span><span class="sxs-lookup"><span data-stu-id="b9307-179">Add custom XML data to the workbook</span></span>

<span data-ttu-id="b9307-180">Excel の Open XML **.xlsx** ファイル形式を使用すると、アドインでカスタム XML データをブックに埋め込むことができます。</span><span class="sxs-lookup"><span data-stu-id="b9307-180">Excel's Open XML **.xlsx** file format lets your add-in embed custom XML data in the workbook.</span></span> <span data-ttu-id="b9307-181">このデータは、アドインに関係なく、ブックで保持されます。</span><span class="sxs-lookup"><span data-stu-id="b9307-181">This data persists with the workbook, independent of the add-in.</span></span>

<span data-ttu-id="b9307-182">ブックには [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection) が含まれます。これは [CustomXmlParts](/javascript/api/excel/excel.customxmlpart) のリストです。</span><span class="sxs-lookup"><span data-stu-id="b9307-182">A workbook contains a [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), which is a list of [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span></span> <span data-ttu-id="b9307-183">これにより、XML 文字列と対応する一意の ID へのアクセスが提供されます。</span><span class="sxs-lookup"><span data-stu-id="b9307-183">These give access to the XML strings and a corresponding unique ID.</span></span> <span data-ttu-id="b9307-184">これらの ID を設定として保管することにより、アドインはセッション間で XML パーツへのキーを保持できます。</span><span class="sxs-lookup"><span data-stu-id="b9307-184">By storing these IDs as settings, your add-in can maintain the keys to its XML parts between sessions.</span></span>

<span data-ttu-id="b9307-185">以下のサンプルは、カスタム XML パーツを使用する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="b9307-185">The following samples show how to use custom XML parts.</span></span> <span data-ttu-id="b9307-186">最初のコード ブロックは、XML データをドキュメントに埋め込む方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="b9307-186">The first code block demonstrates how to embed XML data in the document.</span></span> <span data-ttu-id="b9307-187">レビュー担当者のリストを保管してから、ブックの設定を使用して XML の `id` を保存して、後から取得できるようにします。</span><span class="sxs-lookup"><span data-stu-id="b9307-187">It stores a list of reviewers, then uses the workbook's settings to save the XML's `id` for future retrieval.</span></span> <span data-ttu-id="b9307-188">2 番目のブロックでは、後からその XML にアクセスする方法が示されています。</span><span class="sxs-lookup"><span data-stu-id="b9307-188">The second block shows how to access that XML later.</span></span> <span data-ttu-id="b9307-189">"ContosoReviewXmlPartId" 設定がロードされ、ブックの `customXmlParts` に渡されます。</span><span class="sxs-lookup"><span data-stu-id="b9307-189">The "ContosoReviewXmlPartId" setting is loaded and passed to the workbook's `customXmlParts`.</span></span> <span data-ttu-id="b9307-190">それから、XML データがコンソールに出力されます。</span><span class="sxs-lookup"><span data-stu-id="b9307-190">The XML data is then printed to the console.</span></span>

```js
Excel.run(async (context) => {
    // Add reviewer data to the document as XML
    var originalXml = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    var customXmlPart = context.workbook.customXmlParts.add(originalXml);
    customXmlPart.load("id");

    return context.sync().then(function() {
        // Store the XML part's ID in a setting
        var settings = context.workbook.settings;
        settings.add("ContosoReviewXmlPartId", customXmlPart.id);
    });
}).catch(errorHandlerFunction);
```

```js
Excel.run(async (context) => {
    // Retrieve the XML part's id from the setting
    var settings = context.workbook.settings;
    var xmlPartIDSetting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");

    return context.sync().then(function () {
        if (xmlPartIDSetting.value) {
            var customXmlPart = context.workbook.customXmlParts.getItem(xmlPartIDSetting.value);
            var xmlBlob = customXmlPart.getXml();

            return context.sync().then(function () {
                // Add spaces to make more human readable in the console
                var readableXML = xmlBlob.value.replace(/></g, "> <");
                console.log(readableXML);
            });
        }
    });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="b9307-191">`CustomXMLPart.namespaceUri` にデータが入れられるのは、トップレベルのカスタム XML 要素に `xmlns` 属性が含まれている場合に限ります。</span><span class="sxs-lookup"><span data-stu-id="b9307-191">`CustomXMLPart.namespaceUri` is only populated if the top-level custom XML element contains the `xmlns` attribute.</span></span>

## <a name="control-calculation-behavior"></a><span data-ttu-id="b9307-192">計算の動作を制御する</span><span class="sxs-lookup"><span data-stu-id="b9307-192">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="b9307-193">計算モードを設定する</span><span class="sxs-lookup"><span data-stu-id="b9307-193">Set calculation mode</span></span>

<span data-ttu-id="b9307-194">既定では、Excel は、参照されているセルが変更されたときに数式の結果を再計算します。</span><span class="sxs-lookup"><span data-stu-id="b9307-194">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="b9307-195">この計算の動作を調整すると、アドインのパフォーマンス向上に役立つ場合があります。</span><span class="sxs-lookup"><span data-stu-id="b9307-195">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="b9307-196">Application オブジェクトには、`CalculationMode` 型のプロパティ `calculationMode` があります。</span><span class="sxs-lookup"><span data-stu-id="b9307-196">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="b9307-197">次のいずれかの値を設定できます。</span><span class="sxs-lookup"><span data-stu-id="b9307-197">It can be set to the following values:</span></span>

- <span data-ttu-id="b9307-198">`automatic`: 既定の再計算動作。関連するデータが変更されるたびに、Excel は新しい数式の結果を計算します。</span><span class="sxs-lookup"><span data-stu-id="b9307-198">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
- <span data-ttu-id="b9307-199">`automaticExceptTables`: `automatic` と同様ですが、テーブル内の値に加えた変更は無視されます。</span><span class="sxs-lookup"><span data-stu-id="b9307-199">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
- <span data-ttu-id="b9307-200">`manual`: ユーザーまたはアドインが要求した場合にのみ計算します。</span><span class="sxs-lookup"><span data-stu-id="b9307-200">`manual`: Calculations only occur when the user or add-in requests them.</span></span>

### <a name="set-calculation-type"></a><span data-ttu-id="b9307-201">計算タイプを設定する</span><span class="sxs-lookup"><span data-stu-id="b9307-201">Set calculation type</span></span>

<span data-ttu-id="b9307-202">[Application](/javascript/api/excel/excel.application) オブジェクトは、強制的に即時再計算する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="b9307-202">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="b9307-203">`Application.calculate(calculationType)` は、指定した `calculationType` に基づいて手動再計算を開始します。</span><span class="sxs-lookup"><span data-stu-id="b9307-203">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="b9307-204">次の値を指定できます。</span><span class="sxs-lookup"><span data-stu-id="b9307-204">The following values can be specified:</span></span>

- <span data-ttu-id="b9307-205">`full`: 最後に再計算されてから変更されたかどうかに関係なく、開いているすべてのブックのすべての数式を再計算します。</span><span class="sxs-lookup"><span data-stu-id="b9307-205">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="b9307-206">`fullRebuild`: 最後に再計算されてから変更されたかどうかに関係なく、依存関係のある数式を確認してから、開いているすべてのブックのすべての数式を再計算します。</span><span class="sxs-lookup"><span data-stu-id="b9307-206">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="b9307-207">`recalculate`: すべてのアクティブなブックで、最後に計算されてから変更された数式 (またはプログラムで再計算用にマークされている数式)、およびそれに依存する数式を再計算します。</span><span class="sxs-lookup"><span data-stu-id="b9307-207">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>

> [!NOTE]
> <span data-ttu-id="b9307-208">再計算の詳細については、「[数式の再計算、反復計算、または精度を変更する](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b9307-208">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="b9307-209">計算を一時的に中断する</span><span class="sxs-lookup"><span data-stu-id="b9307-209">Temporarily suspend calculations</span></span>

<span data-ttu-id="b9307-210">Excel API では、アドインから `RequestContext.sync()` を呼び出すまで計算をオフにすることもできます。</span><span class="sxs-lookup"><span data-stu-id="b9307-210">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="b9307-211">これは、`suspendApiCalculationUntilNextSync()` で実行できます。</span><span class="sxs-lookup"><span data-stu-id="b9307-211">This is done with `suspendApiCalculationUntilNextSync()`.</span></span> <span data-ttu-id="b9307-212">このメソッドは、アドインから大きな範囲を編集し、複数の編集の間でデータにアクセスする必要がない場合に使用します。</span><span class="sxs-lookup"><span data-stu-id="b9307-212">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="detect-workbook-activation-preview"></a><span data-ttu-id="b9307-213">ブックのアクティブ化を検出する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="b9307-213">Detect workbook activation (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="b9307-214">現在 `Workbook.onActivated` 、イベントはパブリック プレビューでのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="b9307-214">The `Workbook.onActivated` event is currently only available in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

<span data-ttu-id="b9307-215">アドインは、ブックがアクティブ化された場合に検出できます。</span><span class="sxs-lookup"><span data-stu-id="b9307-215">Your add-in can detect when a workbook is activated.</span></span> <span data-ttu-id="b9307-216">ユーザーが別 *のブック、* 別のアプリケーション、または (Excel on the web) Web ブラウザーの別のタブにフォーカスを切り替え、ブックが非アクティブになります。</span><span class="sxs-lookup"><span data-stu-id="b9307-216">A workbook becomes *inactive* when the user switches focus to another workbook, to another application, or (in Excel on the web) to another tab of the web browser.</span></span> <span data-ttu-id="b9307-217">ブックは、 *ユーザーが* ブックにフォーカスを返すときにアクティブ化されます。</span><span class="sxs-lookup"><span data-stu-id="b9307-217">A workbook is *activated* when the user returns focus to the workbook.</span></span> <span data-ttu-id="b9307-218">ブックのアクティブ化によって、ブック データの更新など、アドイン内のコールバック関数をトリガーできます。</span><span class="sxs-lookup"><span data-stu-id="b9307-218">The workbook activation can trigger callback functions in your add-in, such as refreshing workbook data.</span></span>

<span data-ttu-id="b9307-219">ブックがアクティブ化された場合を検出[](excel-add-ins-events.md#register-an-event-handler)するには、ブックの[onActivated イベントのイベント ハンドラー](/javascript/api/excel/excel.workbook#onActivated)を登録します。</span><span class="sxs-lookup"><span data-stu-id="b9307-219">To detect when a workbook is activated, [register an event handler](excel-add-ins-events.md#register-an-event-handler) for the [onActivated](/javascript/api/excel/excel.workbook#onActivated) event of a workbook.</span></span> <span data-ttu-id="b9307-220">イベントのイベント ハンドラーは `onActivated` 、イベントが発生すると [WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs) オブジェクトを受け取る。</span><span class="sxs-lookup"><span data-stu-id="b9307-220">Event handlers for the `onActivated` event receive a [WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs) object when the event fires.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b9307-221">ブック `onActivated` を開いた場合、イベントは検出されません。</span><span class="sxs-lookup"><span data-stu-id="b9307-221">The `onActivated` event doesn't detect when a workbook is opened.</span></span> <span data-ttu-id="b9307-222">このイベントは、ユーザーがフォーカスを既に開いているブックに戻す場合にのみ検出されます。</span><span class="sxs-lookup"><span data-stu-id="b9307-222">This event only detects when a user switches focus back to an already open workbook.</span></span>

<span data-ttu-id="b9307-223">次のコード サンプルは、イベント ハンドラーを登録し、 `onActivated` コールバック関数を設定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="b9307-223">The following code sample shows how to register the `onActivated` event handler and set up a callback function.</span></span>

```js
Excel.run(function (context) {
    // Retrieve the workbook.
    var workbook = context.workbook;

    // Register the workbook activated event handler.
    workbook.onActivated.add(workbookActivated);

    return context.sync();
});

function workbookActivated(event) {
    Excel.run(function (context) {
        // Retrieve the workbook and load the name.
        var workbook = context.workbook;
        workbook.load("name");
        
        return context.sync().then(function () {
            // Callback function for when the workbook is activated.
            console.log(`The workbook ${workbook.name} was activated.`);
        });
    });
}
```

## <a name="save-the-workbook"></a><span data-ttu-id="b9307-224">ブックを保存する</span><span class="sxs-lookup"><span data-stu-id="b9307-224">Save the workbook</span></span>

<span data-ttu-id="b9307-225">`Workbook.save` は、ブックを永続記憶装置に保存します。</span><span class="sxs-lookup"><span data-stu-id="b9307-225">`Workbook.save` saves the workbook to persistent storage.</span></span> <span data-ttu-id="b9307-226">`save` メソッドにはオプションの `saveBehavior` パラメーターを 1 つ指定できます。値は次のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="b9307-226">The `save` method takes a single, optional `saveBehavior` parameter that can be one of the following values:</span></span>

- <span data-ttu-id="b9307-227">`Excel.SaveBehavior.save` (既定値): ファイル名や保存場所を指定するようにユーザーに促すダイアログは表示されず、そのままファイルが保存されます。</span><span class="sxs-lookup"><span data-stu-id="b9307-227">`Excel.SaveBehavior.save` (default): The file is saved without prompting the user to specify file name and save location.</span></span> <span data-ttu-id="b9307-228">ファイルが以前に保存されていない場合は、既定の場所に保存されます。</span><span class="sxs-lookup"><span data-stu-id="b9307-228">If the file has not been saved previously, it's saved to the default location.</span></span> <span data-ttu-id="b9307-229">ファイルが以前に保存されている場合は、同じ場所に保存されます。</span><span class="sxs-lookup"><span data-stu-id="b9307-229">If the file has been saved previously, it's saved to the same location.</span></span>
- <span data-ttu-id="b9307-230">`Excel.SaveBehavior.prompt`: ファイルが以前に保存されていない場合は、ファイル名や保存場所を指定するようにユーザーに促すダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="b9307-230">`Excel.SaveBehavior.prompt`: If file has not been saved previously, the user will be prompted to specify file name and save location.</span></span> <span data-ttu-id="b9307-231">ファイルが以前に保存されている場合、ファイルは同じ場所に保存され、ダイアログは表示されません。</span><span class="sxs-lookup"><span data-stu-id="b9307-231">If the file has been saved previously, it will be saved to the same location and the user will not be prompted.</span></span>

> [!CAUTION]
> <span data-ttu-id="b9307-232">保存を促すダイアログが表示されたのにユーザーがその操作をキャンセルすると、`save` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="b9307-232">If the user is prompted to save and cancels the operation, `save` throws an exception.</span></span>

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook"></a><span data-ttu-id="b9307-233">ブックを閉じる</span><span class="sxs-lookup"><span data-stu-id="b9307-233">Close the workbook</span></span>

<span data-ttu-id="b9307-234">`Workbook.close` は、ブックとそのブックに関連付けられているアドインを終了します (Excel アプリケーションは開いたまま)。</span><span class="sxs-lookup"><span data-stu-id="b9307-234">`Workbook.close` closes the workbook, along with add-ins that are associated with the workbook (the Excel application remains open).</span></span> <span data-ttu-id="b9307-235">`close` メソッドにはオプションの `closeBehavior` パラメーターを 1 つ指定できます。値は次のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="b9307-235">The `close` method takes a single, optional `closeBehavior` parameter that can be one of the following values:</span></span>

- <span data-ttu-id="b9307-236">`Excel.CloseBehavior.save` (既定値): ファイルは閉じる前に保存されます。</span><span class="sxs-lookup"><span data-stu-id="b9307-236">`Excel.CloseBehavior.save` (default): The file is saved before closing.</span></span> <span data-ttu-id="b9307-237">そのファイルが以前に保存されていない場合は、ファイル名や保存場所を指定するようにユーザーに促すダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="b9307-237">If the file has not been saved previously, the user will be prompted to specify file name and save location.</span></span>
- <span data-ttu-id="b9307-238">`Excel.CloseBehavior.skipSave`: ファイルはそのまま閉じられ、保存されません。</span><span class="sxs-lookup"><span data-stu-id="b9307-238">`Excel.CloseBehavior.skipSave`: The file is immediately closed, without saving.</span></span> <span data-ttu-id="b9307-239">未保存の変更は失われます。</span><span class="sxs-lookup"><span data-stu-id="b9307-239">Any unsaved changes will be lost.</span></span>

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a><span data-ttu-id="b9307-240">関連項目</span><span class="sxs-lookup"><span data-stu-id="b9307-240">See also</span></span>

- [<span data-ttu-id="b9307-241">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="b9307-241">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="b9307-242">Excel JavaScript API を使用してワークシートを操作する</span><span class="sxs-lookup"><span data-stu-id="b9307-242">Work with worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
