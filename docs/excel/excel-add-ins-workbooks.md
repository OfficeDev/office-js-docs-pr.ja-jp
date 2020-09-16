---
title: Excel JavaScript API を使用してブックを操作する
description: Excel JavaScript API を使用して、ブックまたはアプリケーションレベルの機能を使用して一般的なタスクを実行する方法を示すコードサンプルです。
ms.date: 08/24/2020
localization_priority: Normal
ms.openlocfilehash: f0af6cc889a110406d987664575a6f3d1b30aa7b
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819505"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="8e916-103">Excel JavaScript API を使用してブックを操作する</span><span class="sxs-lookup"><span data-stu-id="8e916-103">Work with workbooks using the Excel JavaScript API</span></span>

<span data-ttu-id="8e916-104">この記事では、Excel JavaScript API を使用して、ブックでタスクを実行する方法のコード サンプルを示しています。</span><span class="sxs-lookup"><span data-stu-id="8e916-104">This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API.</span></span> <span data-ttu-id="8e916-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧につい `Workbook` ては、「 [Workbook オブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.workbook)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8e916-105">For the complete list of properties and methods that the `Workbook` object supports, see [Workbook Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="8e916-106">この記事では、[Application](/javascript/api/excel/excel.application) オブジェクトを使用して実行するブック レベルのアクションについても説明します。</span><span class="sxs-lookup"><span data-stu-id="8e916-106">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="8e916-107">Workbook オブジェクトは、Excel を操作するアドインのエントリ ポイントです。</span><span class="sxs-lookup"><span data-stu-id="8e916-107">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="8e916-108">このオブジェクトは、Excel データのアクセスや変更に使用するワークシート、テーブル、ピボットテーブル、その他のコレクションを保持します。</span><span class="sxs-lookup"><span data-stu-id="8e916-108">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="8e916-109">[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) オブジェクトは、個々のワークシートを使用して、ブックのすべてのデータにアドインからアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="8e916-109">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through individual worksheets.</span></span> <span data-ttu-id="8e916-110">具体的には、アドインからワークシートの追加、ワークシート間の移動、ワークシート イベントへのハンドラーの割り当てができます。</span><span class="sxs-lookup"><span data-stu-id="8e916-110">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="8e916-111">ワークシートへのアクセスと編集の方法については、「[Excel JavaScript API を使用してワークシートを操作する](excel-add-ins-worksheets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8e916-111">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="8e916-112">アクティブ セルまたは選択した範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="8e916-112">Get the active cell or selected range</span></span>

<span data-ttu-id="8e916-113">Workbook オブジェクトには、ユーザーまたはアドインが選択したセルの範囲を取得する 2 つのメソッド `getActiveCell()` と `getSelectedRange()` があります。</span><span class="sxs-lookup"><span data-stu-id="8e916-113">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="8e916-114">`getActiveCell()` はブックからアクティブ セルを [Range オブジェクト](/javascript/api/excel/excel.range)として取得します。</span><span class="sxs-lookup"><span data-stu-id="8e916-114">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="8e916-115">次の例では、`getActiveCell()` を呼び出し、コンソールにセルのアドレスを表示します。</span><span class="sxs-lookup"><span data-stu-id="8e916-115">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="8e916-116">`getSelectedRange()` メソッドは現在選択されている単一の範囲を返します。</span><span class="sxs-lookup"><span data-stu-id="8e916-116">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="8e916-117">複数の範囲が選択されている場合は、InvalidSelection エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="8e916-117">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="8e916-118">次の例では、`getSelectedRange()` を呼び出し、範囲の塗りつぶし色を黄色に設定します。</span><span class="sxs-lookup"><span data-stu-id="8e916-118">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="8e916-119">ブックを作成する</span><span class="sxs-lookup"><span data-stu-id="8e916-119">Create a workbook</span></span>

<span data-ttu-id="8e916-120">アドインでは、アドインが現在実行されている Excel のインスタンスとは異なる新しいブックを作成できます。</span><span class="sxs-lookup"><span data-stu-id="8e916-120">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="8e916-121">Excel オブジェクトには、この目的の `createWorkbook` メソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="8e916-121">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="8e916-122">このメソッドが呼び出されると、新しいブックが Excel の新しいインスタンスですぐに開いて表示されます。</span><span class="sxs-lookup"><span data-stu-id="8e916-122">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="8e916-123">アドインは前のブックで開いて実行されたままになります。</span><span class="sxs-lookup"><span data-stu-id="8e916-123">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="8e916-124">`createWorkbook` メソッドは既存のブックのコピーの作成もできます。</span><span class="sxs-lookup"><span data-stu-id="8e916-124">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="8e916-125">このメソッドは、オプションのパラメーターとして .xlsx ファイルの base64 エンコード文字列表現を受け取ります。</span><span class="sxs-lookup"><span data-stu-id="8e916-125">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="8e916-126">文字列の引数は有効な .xlsx ファイルと見なされ、作成されるブックはそのファイルのコピーになります。</span><span class="sxs-lookup"><span data-stu-id="8e916-126">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="8e916-127">[ファイルスライシング](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)を使用して、アドインの現在のブックを base64 でエンコードされた文字列として取得できます。</span><span class="sxs-lookup"><span data-stu-id="8e916-127">You can get your add-in's current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="8e916-128">次の例に示すように、[FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) クラスを使用して、ファイルを必要な base64 エンコード文字列に変換できます。</span><span class="sxs-lookup"><span data-stu-id="8e916-128">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // strip off the metadata before the base64-encoded string
        var startIndex = reader.result.toString().indexOf("base64,");
        var workbookContents = reader.result.toString().substr(startIndex + 7);

        Excel.createWorkbook(workbookContents);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one-preview"></a><span data-ttu-id="8e916-129">既存のブックのコピーを現在のブックに挿入する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="8e916-129">Insert a copy of an existing workbook into the current one (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="8e916-130">`WorksheetCollection.addFromBase64` メソッドは現在、パブリック プレビューでのみ使用でき、Windows および Mac 上の Office でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="8e916-130">The `WorksheetCollection.addFromBase64` method is currently only available in public preview and only for Office on Windows and Mac.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="8e916-131">前の例は、既存のブックから作成された新しいブックを示しています。</span><span class="sxs-lookup"><span data-stu-id="8e916-131">The previous example shows a new workbook being created from an existing workbook.</span></span> <span data-ttu-id="8e916-132">既存のブックの一部またはすべてを、アドインに関連付けられているブックにコピーすることもできます。</span><span class="sxs-lookup"><span data-stu-id="8e916-132">You can also copy some or all of an existing workbook into the one currently associated with your add-in.</span></span> <span data-ttu-id="8e916-133">ブックの [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) にある `addFromBase64` メソッドは、対象のブックのワークシートのコピーを現在のブックに挿入します。</span><span class="sxs-lookup"><span data-stu-id="8e916-133">A workbook's [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) has the `addFromBase64` method to insert copies of the target workbook's worksheets into itself.</span></span> <span data-ttu-id="8e916-134">他のブックのファイルは、`Excel.createWorkbook` 呼び出しの場合と同様に、base64 エンコード文字列として渡されます。</span><span class="sxs-lookup"><span data-stu-id="8e916-134">The other workbook's file is passed as base64-encoded string, just like the `Excel.createWorkbook` call.</span></span>

```TypeScript
addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet | string): OfficeExtension.ClientResult<string[]>;
```

<span data-ttu-id="8e916-135">次の例では、ブックのワークシートが現在のブックのアクティブ ワークシートの直後に挿入されています。</span><span class="sxs-lookup"><span data-stu-id="8e916-135">The following example shows a workbook's worksheets being inserted in the current workbook, directly after the active worksheet.</span></span> <span data-ttu-id="8e916-136">`null` が `sheetNamesToInsert?: string[]` パラメーターに渡されている点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="8e916-136">Note that `null` is passed for the `sheetNamesToInsert?: string[]` parameter.</span></span> <span data-ttu-id="8e916-137">つまり、すべてのワークシートが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="8e916-137">This means all the worksheets are being inserted.</span></span>

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // strip off the metadata before the base64-encoded string
        var startIndex = reader.result.toString().indexOf("base64,");
        var workbookContents = reader.result.toString().substr(startIndex + 7);

        var sheets = context.workbook.worksheets;
        sheets.addFromBase64(
            workbookContents,
            null, // get all the worksheets
            Excel.WorksheetPositionType.after, // insert them after the worksheet specified by the next parameter
            sheets.getActiveWorksheet() // insert them after the active worksheet
        );
        return context.sync();
    });
};

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="8e916-138">ブックのシート構成を保護する</span><span class="sxs-lookup"><span data-stu-id="8e916-138">Protect the workbook's structure</span></span>

<span data-ttu-id="8e916-139">アドインでは、ブックのシート構成を編集するユーザーの機能を制御できます。</span><span class="sxs-lookup"><span data-stu-id="8e916-139">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="8e916-140">Workbook オブジェクトの `protection` プロパティは [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) オブジェクトであり、`protect()` メソッドを備えています。</span><span class="sxs-lookup"><span data-stu-id="8e916-140">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="8e916-141">次の例では、ブックのシート構成の保護を切り替える基本的なシナリオを示します。</span><span class="sxs-lookup"><span data-stu-id="8e916-141">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span>

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

<span data-ttu-id="8e916-142">`protect` メソッドは、オプションの文字列パラメーターを受け取ります。</span><span class="sxs-lookup"><span data-stu-id="8e916-142">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="8e916-143">この文字列は、ユーザーが保護をバイパスしてブックのシート構成を変更するために必要なパスワードを表します。</span><span class="sxs-lookup"><span data-stu-id="8e916-143">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="8e916-144">保護は、不必要なデータ編集をできないようにするため、ワークシート レベルで設定することもできます。</span><span class="sxs-lookup"><span data-stu-id="8e916-144">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="8e916-145">詳細については、「[Excel JavaScript API を使用してワークシートを操作する](excel-add-ins-worksheets.md#data-protection)」の**データの保護**のセクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="8e916-145">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE]
> <span data-ttu-id="8e916-146">Excel のブックの保護の詳細については、「[ブックを保護する](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8e916-146">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="8e916-147">ドキュメント プロパティへのアクセス</span><span class="sxs-lookup"><span data-stu-id="8e916-147">Access document properties</span></span>

<span data-ttu-id="8e916-148">Workbook オブジェクトは、[ドキュメント プロパティ](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75)と呼ばれる Office ファイルのメタデータにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="8e916-148">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="8e916-149">Workbook オブジェクトの `properties` プロパティは、これらのメタデータ値を含む [DocumentProperties](/javascript/api/excel/excel.documentproperties) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8e916-149">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="8e916-150">次の例は、プロパティを設定する方法を示して `author` います。</span><span class="sxs-lookup"><span data-stu-id="8e916-150">The following example shows how to set the `author` property.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="custom-properties"></a><span data-ttu-id="8e916-151">カスタム プロパティ</span><span class="sxs-lookup"><span data-stu-id="8e916-151">Custom properties</span></span>

<span data-ttu-id="8e916-152">カスタム プロパティを定義することもできます。</span><span class="sxs-lookup"><span data-stu-id="8e916-152">You can also define custom properties.</span></span> <span data-ttu-id="8e916-153">DocumentProperties オブジェクトには `custom` プロパティが含まれていて、ユーザー定義プロパティのキー値のペアのコレクションを表します。</span><span class="sxs-lookup"><span data-stu-id="8e916-153">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="8e916-154">次の例では、"Hello" という値を持つ **Introduction** という名前のカスタム プロパティを作成し、それを取得する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="8e916-154">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

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

#### <a name="worksheet-level-custom-properties"></a><span data-ttu-id="8e916-155">ワークシートレベルのカスタムプロパティ</span><span class="sxs-lookup"><span data-stu-id="8e916-155">Worksheet-level custom properties</span></span>

<span data-ttu-id="8e916-156">カスタムプロパティは、ワークシートレベルで設定することもできます。</span><span class="sxs-lookup"><span data-stu-id="8e916-156">Custom properties can also be set at the worksheet level.</span></span> <span data-ttu-id="8e916-157">これらはドキュメントレベルのカスタムプロパティに似ていますが、異なるワークシート間で同じキーを繰り返すことができる点が異なります。</span><span class="sxs-lookup"><span data-stu-id="8e916-157">These are similar to document-level custom properties, except that the same key can be repeated across different worksheets.</span></span> <span data-ttu-id="8e916-158">次の例は、現在のワークシートで "α" という値を持つ、"worksheet **group** " という名前のカスタムプロパティを作成し、それを取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8e916-158">The following example shows how to create a custom property named **WorksheetGroup** with the value "Alpha" on the current worksheet, then retrieve it.</span></span>

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

## <a name="access-document-settings"></a><span data-ttu-id="8e916-159">ドキュメント設定へのアクセス</span><span class="sxs-lookup"><span data-stu-id="8e916-159">Access document settings</span></span>

<span data-ttu-id="8e916-160">ブックの設定は、カスタム プロパティのコレクションに似ています。</span><span class="sxs-lookup"><span data-stu-id="8e916-160">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="8e916-161">設定は 1 つの Excel ファイルとアドインのペアリングに固有であるのに対して、プロパティはファイルに接続しているだけである点が異なります。</span><span class="sxs-lookup"><span data-stu-id="8e916-161">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="8e916-162">次の例は、設定を作成してアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8e916-162">The following example shows how to create and access a setting.</span></span>

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

## <a name="access-application-culture-settings"></a><span data-ttu-id="8e916-163">Access アプリケーションのカルチャ設定</span><span class="sxs-lookup"><span data-stu-id="8e916-163">Access application culture settings</span></span>

<span data-ttu-id="8e916-164">ブックには、特定のデータの表示方法に影響する言語とカルチャの設定が含まれています。</span><span class="sxs-lookup"><span data-stu-id="8e916-164">A workbook has language and culture settings that affect how certain data is displayed.</span></span> <span data-ttu-id="8e916-165">これらの設定は、アドインのユーザーが異なる言語とカルチャでブックを共有している場合に、データをローカライズするのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="8e916-165">These settings can help localize data when your add-in's users are sharing workbooks across different languages and cultures.</span></span> <span data-ttu-id="8e916-166">アドインでは、文字列の解析を使用して、各ユーザーが独自のカルチャの形式でデータを表示できるように、システムのカルチャ設定に基づいて数値、日付、時刻の形式をローカライズできます。</span><span class="sxs-lookup"><span data-stu-id="8e916-166">Your add-in can use string parsing to localize the format of numbers, dates, and times based on the system culture settings so that each user sees data in their own culture's format.</span></span>

<span data-ttu-id="8e916-167">`Application.cultureInfo` システムのカルチャ設定を [CultureInfo](/javascript/api/excel/excel.cultureinfo) オブジェクトとして定義します。</span><span class="sxs-lookup"><span data-stu-id="8e916-167">`Application.cultureInfo` defines the system culture settings as a [CultureInfo](/javascript/api/excel/excel.cultureinfo) object.</span></span> <span data-ttu-id="8e916-168">これには、数値の小数点の記号や日付の形式などの設定が含まれます。</span><span class="sxs-lookup"><span data-stu-id="8e916-168">This contains settings like the numerical decimal separator or the date format.</span></span>

<span data-ttu-id="8e916-169">一部のカルチャ設定は [、EXCEL UI を使用して変更](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e)できます。</span><span class="sxs-lookup"><span data-stu-id="8e916-169">Some culture settings can be [changed through the Excel UI](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e).</span></span> <span data-ttu-id="8e916-170">システム設定は、オブジェクトに保持され `CultureInfo` ます。</span><span class="sxs-lookup"><span data-stu-id="8e916-170">The system settings are preserved in the `CultureInfo` object.</span></span> <span data-ttu-id="8e916-171">ローカルの変更は、など、 [アプリケーション](/javascript/api/excel/excel.application)レベルのプロパティとして保持され `Application.decimalSeparator` ます。</span><span class="sxs-lookup"><span data-stu-id="8e916-171">Any local changes are kept as [Application](/javascript/api/excel/excel.application)-level properties, such as `Application.decimalSeparator`.</span></span>

<span data-ttu-id="8e916-172">次の例では、"," から、システム設定で使用される文字への数値文字列の小数点の区切り文字を変更します。</span><span class="sxs-lookup"><span data-stu-id="8e916-172">The following sample changes the decimal separator character of a numerical string from a ',' to the character used by the system settings.</span></span>

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

## <a name="add-custom-xml-data-to-the-workbook"></a><span data-ttu-id="8e916-173">カスタム XML データをブックに追加する</span><span class="sxs-lookup"><span data-stu-id="8e916-173">Add custom XML data to the workbook</span></span>

<span data-ttu-id="8e916-174">Excel の Open XML **.xlsx** ファイル形式を使用すると、アドインでカスタム XML データをブックに埋め込むことができます。</span><span class="sxs-lookup"><span data-stu-id="8e916-174">Excel's Open XML **.xlsx** file format lets your add-in embed custom XML data in the workbook.</span></span> <span data-ttu-id="8e916-175">このデータは、アドインに関係なく、ブックで保持されます。</span><span class="sxs-lookup"><span data-stu-id="8e916-175">This data persists with the workbook, independent of the add-in.</span></span>

<span data-ttu-id="8e916-176">ブックには [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection) が含まれます。これは [CustomXmlParts](/javascript/api/excel/excel.customxmlpart) のリストです。</span><span class="sxs-lookup"><span data-stu-id="8e916-176">A workbook contains a [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), which is a list of [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span></span> <span data-ttu-id="8e916-177">これにより、XML 文字列と対応する一意の ID へのアクセスが提供されます。</span><span class="sxs-lookup"><span data-stu-id="8e916-177">These give access to the XML strings and a corresponding unique ID.</span></span> <span data-ttu-id="8e916-178">これらの ID を設定として保管することにより、アドインはセッション間で XML パーツへのキーを保持できます。</span><span class="sxs-lookup"><span data-stu-id="8e916-178">By storing these IDs as settings, your add-in can maintain the keys to its XML parts between sessions.</span></span>

<span data-ttu-id="8e916-179">以下のサンプルは、カスタム XML パーツを使用する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8e916-179">The following samples show how to use custom XML parts.</span></span> <span data-ttu-id="8e916-180">最初のコード ブロックは、XML データをドキュメントに埋め込む方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8e916-180">The first code block demonstrates how to embed XML data in the document.</span></span> <span data-ttu-id="8e916-181">レビュー担当者のリストを保管してから、ブックの設定を使用して XML の `id` を保存して、後から取得できるようにします。</span><span class="sxs-lookup"><span data-stu-id="8e916-181">It stores a list of reviewers, then uses the workbook's settings to save the XML's `id` for future retrieval.</span></span> <span data-ttu-id="8e916-182">2 番目のブロックでは、後からその XML にアクセスする方法が示されています。</span><span class="sxs-lookup"><span data-stu-id="8e916-182">The second block shows how to access that XML later.</span></span> <span data-ttu-id="8e916-183">"ContosoReviewXmlPartId" 設定がロードされ、ブックの `customXmlParts` に渡されます。</span><span class="sxs-lookup"><span data-stu-id="8e916-183">The "ContosoReviewXmlPartId" setting is loaded and passed to the workbook's `customXmlParts`.</span></span> <span data-ttu-id="8e916-184">それから、XML データがコンソールに出力されます。</span><span class="sxs-lookup"><span data-stu-id="8e916-184">The XML data is then printed to the console.</span></span>

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
> <span data-ttu-id="8e916-185">`CustomXMLPart.namespaceUri` にデータが入れられるのは、トップレベルのカスタム XML 要素に `xmlns` 属性が含まれている場合に限ります。</span><span class="sxs-lookup"><span data-stu-id="8e916-185">`CustomXMLPart.namespaceUri` is only populated if the top-level custom XML element contains the `xmlns` attribute.</span></span>

## <a name="control-calculation-behavior"></a><span data-ttu-id="8e916-186">計算の動作を制御する</span><span class="sxs-lookup"><span data-stu-id="8e916-186">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="8e916-187">計算モードを設定する</span><span class="sxs-lookup"><span data-stu-id="8e916-187">Set calculation mode</span></span>

<span data-ttu-id="8e916-188">既定では、Excel は、参照されているセルが変更されたときに数式の結果を再計算します。</span><span class="sxs-lookup"><span data-stu-id="8e916-188">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="8e916-189">この計算の動作を調整すると、アドインのパフォーマンス向上に役立つ場合があります。</span><span class="sxs-lookup"><span data-stu-id="8e916-189">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="8e916-190">Application オブジェクトには、`CalculationMode` 型のプロパティ `calculationMode` があります。</span><span class="sxs-lookup"><span data-stu-id="8e916-190">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="8e916-191">次のいずれかの値を設定できます。</span><span class="sxs-lookup"><span data-stu-id="8e916-191">It can be set to the following values:</span></span>

- <span data-ttu-id="8e916-192">`automatic`: 既定の再計算動作。関連するデータが変更されるたびに、Excel は新しい数式の結果を計算します。</span><span class="sxs-lookup"><span data-stu-id="8e916-192">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
- <span data-ttu-id="8e916-193">`automaticExceptTables`: `automatic` と同様ですが、テーブル内の値に加えた変更は無視されます。</span><span class="sxs-lookup"><span data-stu-id="8e916-193">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
- <span data-ttu-id="8e916-194">`manual`: ユーザーまたはアドインが要求した場合にのみ計算します。</span><span class="sxs-lookup"><span data-stu-id="8e916-194">`manual`: Calculations only occur when the user or add-in requests them.</span></span>

### <a name="set-calculation-type"></a><span data-ttu-id="8e916-195">計算タイプを設定する</span><span class="sxs-lookup"><span data-stu-id="8e916-195">Set calculation type</span></span>

<span data-ttu-id="8e916-196">[Application](/javascript/api/excel/excel.application) オブジェクトは、強制的に即時再計算する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="8e916-196">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="8e916-197">`Application.calculate(calculationType)` は、指定した `calculationType` に基づいて手動再計算を開始します。</span><span class="sxs-lookup"><span data-stu-id="8e916-197">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="8e916-198">次の値を指定できます。</span><span class="sxs-lookup"><span data-stu-id="8e916-198">The following values can be specified:</span></span>

- <span data-ttu-id="8e916-199">`full`: 最後に再計算されてから変更されたかどうかに関係なく、開いているすべてのブックのすべての数式を再計算します。</span><span class="sxs-lookup"><span data-stu-id="8e916-199">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="8e916-200">`fullRebuild`: 最後に再計算されてから変更されたかどうかに関係なく、依存関係のある数式を確認してから、開いているすべてのブックのすべての数式を再計算します。</span><span class="sxs-lookup"><span data-stu-id="8e916-200">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="8e916-201">`recalculate`: すべてのアクティブなブックで、最後に計算されてから変更された数式 (またはプログラムで再計算用にマークされている数式)、およびそれに依存する数式を再計算します。</span><span class="sxs-lookup"><span data-stu-id="8e916-201">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>

> [!NOTE]
> <span data-ttu-id="8e916-202">再計算の詳細については、「[数式の再計算、反復計算、または精度を変更する](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8e916-202">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="8e916-203">計算を一時的に中断する</span><span class="sxs-lookup"><span data-stu-id="8e916-203">Temporarily suspend calculations</span></span>

<span data-ttu-id="8e916-204">Excel API では、アドインから `RequestContext.sync()` を呼び出すまで計算をオフにすることもできます。</span><span class="sxs-lookup"><span data-stu-id="8e916-204">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="8e916-205">これは、`suspendApiCalculationUntilNextSync()` で実行できます。</span><span class="sxs-lookup"><span data-stu-id="8e916-205">This is done with `suspendApiCalculationUntilNextSync()`.</span></span> <span data-ttu-id="8e916-206">このメソッドは、アドインから大きな範囲を編集し、複数の編集の間でデータにアクセスする必要がない場合に使用します。</span><span class="sxs-lookup"><span data-stu-id="8e916-206">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="save-the-workbook"></a><span data-ttu-id="8e916-207">ブックを保存する</span><span class="sxs-lookup"><span data-stu-id="8e916-207">Save the workbook</span></span>

<span data-ttu-id="8e916-208">`Workbook.save` は、ブックを永続記憶装置に保存します。</span><span class="sxs-lookup"><span data-stu-id="8e916-208">`Workbook.save` saves the workbook to persistent storage.</span></span> <span data-ttu-id="8e916-209">`save` メソッドにはオプションの `saveBehavior` パラメーターを 1 つ指定できます。値は次のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="8e916-209">The `save` method takes a single, optional `saveBehavior` parameter that can be one of the following values:</span></span>

- <span data-ttu-id="8e916-210">`Excel.SaveBehavior.save` (既定値): ファイル名や保存場所を指定するようにユーザーに促すダイアログは表示されず、そのままファイルが保存されます。</span><span class="sxs-lookup"><span data-stu-id="8e916-210">`Excel.SaveBehavior.save` (default): The file is saved without prompting the user to specify file name and save location.</span></span> <span data-ttu-id="8e916-211">ファイルが以前に保存されていない場合は、既定の場所に保存されます。</span><span class="sxs-lookup"><span data-stu-id="8e916-211">If the file has not been saved previously, it's saved to the default location.</span></span> <span data-ttu-id="8e916-212">ファイルが以前に保存されている場合は、同じ場所に保存されます。</span><span class="sxs-lookup"><span data-stu-id="8e916-212">If the file has been saved previously, it's saved to the same location.</span></span>
- <span data-ttu-id="8e916-213">`Excel.SaveBehavior.prompt`: ファイルが以前に保存されていない場合は、ファイル名や保存場所を指定するようにユーザーに促すダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8e916-213">`Excel.SaveBehavior.prompt`: If file has not been saved previously, the user will be prompted to specify file name and save location.</span></span> <span data-ttu-id="8e916-214">ファイルが以前に保存されている場合、ファイルは同じ場所に保存され、ダイアログは表示されません。</span><span class="sxs-lookup"><span data-stu-id="8e916-214">If the file has been saved previously, it will be saved to the same location and the user will not be prompted.</span></span>

> [!CAUTION]
> <span data-ttu-id="8e916-215">保存を促すダイアログが表示されたのにユーザーがその操作をキャンセルすると、`save` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="8e916-215">If the user is prompted to save and cancels the operation, `save` throws an exception.</span></span>

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook"></a><span data-ttu-id="8e916-216">ブックを閉じる</span><span class="sxs-lookup"><span data-stu-id="8e916-216">Close the workbook</span></span>

<span data-ttu-id="8e916-217">`Workbook.close` は、ブックとそのブックに関連付けられているアドインを終了します (Excel アプリケーションは開いたまま)。</span><span class="sxs-lookup"><span data-stu-id="8e916-217">`Workbook.close` closes the workbook, along with add-ins that are associated with the workbook (the Excel application remains open).</span></span> <span data-ttu-id="8e916-218">`close` メソッドにはオプションの `closeBehavior` パラメーターを 1 つ指定できます。値は次のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="8e916-218">The `close` method takes a single, optional `closeBehavior` parameter that can be one of the following values:</span></span>

- <span data-ttu-id="8e916-219">`Excel.CloseBehavior.save` (既定値): ファイルは閉じる前に保存されます。</span><span class="sxs-lookup"><span data-stu-id="8e916-219">`Excel.CloseBehavior.save` (default): The file is saved before closing.</span></span> <span data-ttu-id="8e916-220">そのファイルが以前に保存されていない場合は、ファイル名や保存場所を指定するようにユーザーに促すダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8e916-220">If the file has not been saved previously, the user will be prompted to specify file name and save location.</span></span>
- <span data-ttu-id="8e916-221">`Excel.CloseBehavior.skipSave`: ファイルはそのまま閉じられ、保存されません。</span><span class="sxs-lookup"><span data-stu-id="8e916-221">`Excel.CloseBehavior.skipSave`: The file is immediately closed, without saving.</span></span> <span data-ttu-id="8e916-222">未保存の変更は失われます。</span><span class="sxs-lookup"><span data-stu-id="8e916-222">Any unsaved changes will be lost.</span></span>

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a><span data-ttu-id="8e916-223">関連項目</span><span class="sxs-lookup"><span data-stu-id="8e916-223">See also</span></span>

- [<span data-ttu-id="8e916-224">Office アドインでの Excel JavaScript オブジェクトモデル</span><span class="sxs-lookup"><span data-stu-id="8e916-224">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="8e916-225">Excel JavaScript API を使用してワークシートを操作する</span><span class="sxs-lookup"><span data-stu-id="8e916-225">Work with worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
- [<span data-ttu-id="8e916-226">Excel JavaScript API を使用して範囲を操作する</span><span class="sxs-lookup"><span data-stu-id="8e916-226">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
