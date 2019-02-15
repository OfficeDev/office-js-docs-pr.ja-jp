---
title: Excel JavaScript API を使用してブックを操作する
description: ''
ms.date: 01/07/2019
localization_priority: Priority
ms.openlocfilehash: 9e88809ea7174df972dfc31110e8370a5294fb6c
ms.sourcegitcommit: 70ef38a290c18a1d1a380fd02b263470207a5dc6
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/15/2019
ms.locfileid: "30052757"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="924d6-102">Excel JavaScript API を使用してブックを操作する</span><span class="sxs-lookup"><span data-stu-id="924d6-102">Work with workbooks using the Excel JavaScript API</span></span>

<span data-ttu-id="924d6-103">この記事では、Excel JavaScript API を使用して、ブックでタスクを実行する方法のコード サンプルを示しています。</span><span class="sxs-lookup"><span data-stu-id="924d6-103">This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API.</span></span> <span data-ttu-id="924d6-104">**Workbook** オブジェクトがサポートするプロパティとメソッドの完全な一覧については、「[Workbook オブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.workbook)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="924d6-104">For the complete list of properties and methods that the **Workbook** object supports, see [Workbook Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="924d6-105">この記事では、[Application](/javascript/api/excel/excel.application) オブジェクトを使用して実行するブック レベルのアクションについても説明します。</span><span class="sxs-lookup"><span data-stu-id="924d6-105">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="924d6-106">Workbook オブジェクトは、Excel を操作するアドインのエントリ ポイントです。</span><span class="sxs-lookup"><span data-stu-id="924d6-106">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="924d6-107">このオブジェクトは、Excel データのアクセスや変更に使用するワークシート、テーブル、ピボットテーブル、その他のコレクションを保持します。</span><span class="sxs-lookup"><span data-stu-id="924d6-107">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="924d6-108">[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) オブジェクトは、個々のワークシートを使用して、ブックのすべてのデータにアドインからアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="924d6-108">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through individual worksheets.</span></span> <span data-ttu-id="924d6-109">具体的には、アドインからワークシートの追加、ワークシート間の移動、ワークシート イベントへのハンドラーの割り当てができます。</span><span class="sxs-lookup"><span data-stu-id="924d6-109">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="924d6-110">ワークシートへのアクセスと編集の方法については、「[Excel JavaScript API を使用してワークシートを操作する](excel-add-ins-worksheets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="924d6-110">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="924d6-111">アクティブ セルまたは選択した範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="924d6-111">Get the active cell or selected range</span></span>

<span data-ttu-id="924d6-112">Workbook オブジェクトには、ユーザーまたはアドインが選択したセルの範囲を取得する 2 つのメソッド `getActiveCell()` と `getSelectedRange()` があります。</span><span class="sxs-lookup"><span data-stu-id="924d6-112">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="924d6-113">`getActiveCell()` はブックからアクティブ セルを [Range オブジェクト](/javascript/api/excel/excel.range)として取得します。</span><span class="sxs-lookup"><span data-stu-id="924d6-113">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="924d6-114">次の例では、`getActiveCell()` を呼び出し、コンソールにセルのアドレスを表示します。</span><span class="sxs-lookup"><span data-stu-id="924d6-114">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="924d6-115">`getSelectedRange()` メソッドは現在選択されている単一の範囲を返します。</span><span class="sxs-lookup"><span data-stu-id="924d6-115">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="924d6-116">複数の範囲が選択されている場合は、InvalidSelection エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="924d6-116">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="924d6-117">次の例では、`getSelectedRange()` を呼び出し、範囲の塗りつぶし色を黄色に設定します。</span><span class="sxs-lookup"><span data-stu-id="924d6-117">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="924d6-118">ブックを作成する</span><span class="sxs-lookup"><span data-stu-id="924d6-118">Create a workbook</span></span>

<span data-ttu-id="924d6-119">アドインでは、アドインが現在実行されている Excel のインスタンスとは異なる新しいブックを作成できます。</span><span class="sxs-lookup"><span data-stu-id="924d6-119">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="924d6-120">Excel オブジェクトには、この目的の `createWorkbook` メソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="924d6-120">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="924d6-121">このメソッドが呼び出されると、新しいブックが Excel の新しいインスタンスですぐに開いて表示されます。</span><span class="sxs-lookup"><span data-stu-id="924d6-121">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="924d6-122">アドインは前のブックで開いて実行されたままになります。</span><span class="sxs-lookup"><span data-stu-id="924d6-122">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="924d6-123">`createWorkbook` メソッドは既存のブックのコピーの作成もできます。</span><span class="sxs-lookup"><span data-stu-id="924d6-123">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="924d6-124">このメソッドは、オプションのパラメーターとして .xlsx ファイルの base64 エンコード文字列表現を受け取ります。</span><span class="sxs-lookup"><span data-stu-id="924d6-124">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="924d6-125">文字列の引数は有効な .xlsx ファイルと見なされ、作成されるブックはそのファイルのコピーになります。</span><span class="sxs-lookup"><span data-stu-id="924d6-125">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="924d6-126">[ファイルのスライス](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)を使用して、アドインの現在のブックを base64 エンコード文字列として取得できます。</span><span class="sxs-lookup"><span data-stu-id="924d6-126">You can get your add-in’s current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="924d6-127">次の例に示すように、[FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) クラスを使用して、ファイルを必要な base64 エンコード文字列に変換できます。</span><span class="sxs-lookup"><span data-stu-id="924d6-127">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // strip off the metadata before the base64-encoded string
        var startIndex = event.target.result.indexOf("base64,");
        var workbookContents = event.target.result.substr(startIndex + 7);

        Excel.createWorkbook(workbookContents);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one"></a><span data-ttu-id="924d6-128">既存のブックのコピーを現在のブックに挿入する</span><span class="sxs-lookup"><span data-stu-id="924d6-128">Insert a copy of an existing workbook into the current one</span></span>

> [!NOTE]
> <span data-ttu-id="924d6-129">現在、`WorksheetCollection.addFromBase64` 関数は、パブリック プレビュー (ベータ版) でのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="924d6-129">The `WorksheetCollection.addFromBase64` function is currently available only in public preview (beta).</span></span> <span data-ttu-id="924d6-130">この機能を使用するには、Office.js CDN のベータ版のライブラリを使用する必要があります: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="924d6-130">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="924d6-131">TypeScript を使用している場合、または IntelliSense に TypeScript 型定義ファイルを使用するコード エディターを使用している場合は、https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts を使用してください。</span><span class="sxs-lookup"><span data-stu-id="924d6-131">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="924d6-132">前の例は、既存のブックから作成された新しいブックを示しています。</span><span class="sxs-lookup"><span data-stu-id="924d6-132">The previous example shows a new workbook being created from an existing workbook.</span></span> <span data-ttu-id="924d6-133">既存のブックの一部またはすべてを、アドインに関連付けられているブックにコピーすることもできます。</span><span class="sxs-lookup"><span data-stu-id="924d6-133">You can also copy some or all of an existing workbook into the one currently associated with your add-in.</span></span> <span data-ttu-id="924d6-134">ブックの [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) にある `addFromBase64` メソッドは、対象のブックのワークシートのコピーを現在のブックに挿入します。</span><span class="sxs-lookup"><span data-stu-id="924d6-134">A workbook's [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) has the `addFromBase64` method to insert copies of the target workbook's worksheets into itself.</span></span> <span data-ttu-id="924d6-135">他のブックのファイルは、`Excel.createWorkbook` 呼び出しの場合と同様に、base64 エンコード文字列として渡されます。</span><span class="sxs-lookup"><span data-stu-id="924d6-135">The other workbook's file is passed as base64-encoded string, just like the `Excel.createWorkbook` call.</span></span>

```TypeScript
addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet | string): OfficeExtension.ClientResult<string[]>;
```

<span data-ttu-id="924d6-136">次の例では、ブックのワークシートが現在のブックのアクティブ ワークシートの直後に挿入されています。</span><span class="sxs-lookup"><span data-stu-id="924d6-136">The following example shows a workbook's worksheets being inserted in the current workbook, directly after the active worksheet.</span></span> <span data-ttu-id="924d6-137">`null` が `sheetNamesToInsert?: string[]` パラメーターに渡されている点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="924d6-137">Note that `null` is passed for the `sheetNamesToInsert?: string[]` parameter.</span></span> <span data-ttu-id="924d6-138">つまり、すべてのワークシートが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="924d6-138">This means all the worksheets are being inserted.</span></span>

```js
var myFile = <HTMLInputElement>document.getElementById("file");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // strip off the metadata before the base64-encoded string
        var startIndex = (<string>(<FileReader>event.target).result).indexOf("base64,");
        var workbookContents = (<string>(<FileReader>event.target).result).substr(startIndex + 7);

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

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="924d6-139">ブックのシート構成を保護する</span><span class="sxs-lookup"><span data-stu-id="924d6-139">Protect the workbook's structure</span></span>

<span data-ttu-id="924d6-140">アドインでは、ブックのシート構成を編集するユーザーの機能を制御できます。</span><span class="sxs-lookup"><span data-stu-id="924d6-140">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="924d6-141">Workbook オブジェクトの `protection` プロパティは [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) オブジェクトであり、`protect()` メソッドを備えています。</span><span class="sxs-lookup"><span data-stu-id="924d6-141">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="924d6-142">次の例では、ブックのシート構成の保護を切り替える基本的なシナリオを示します。</span><span class="sxs-lookup"><span data-stu-id="924d6-142">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span>

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

<span data-ttu-id="924d6-143">`protect` メソッドは、オプションの文字列パラメーターを受け取ります。</span><span class="sxs-lookup"><span data-stu-id="924d6-143">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="924d6-144">この文字列は、ユーザーが保護をバイパスしてブックのシート構成を変更するために必要なパスワードを表します。</span><span class="sxs-lookup"><span data-stu-id="924d6-144">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="924d6-145">保護は、不必要なデータ編集をできないようにするため、ワークシート レベルで設定することもできます。</span><span class="sxs-lookup"><span data-stu-id="924d6-145">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="924d6-146">詳細については、「[Excel JavaScript API を使用してワークシートを操作する](excel-add-ins-worksheets.md#data-protection)」の**データの保護**のセクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="924d6-146">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE]
> <span data-ttu-id="924d6-147">Excel のブックの保護の詳細については、「[ブックを保護する](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="924d6-147">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="924d6-148">ドキュメント プロパティへのアクセス</span><span class="sxs-lookup"><span data-stu-id="924d6-148">Access document properties</span></span>

<span data-ttu-id="924d6-149">Workbook オブジェクトは、[ドキュメント プロパティ](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75)と呼ばれる Office ファイルのメタデータにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="924d6-149">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="924d6-150">Workbook オブジェクトの `properties` プロパティは、これらのメタデータ値を含む [DocumentProperties](/javascript/api/excel/excel.documentproperties) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="924d6-150">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="924d6-151">次のコード例では、**author** プロパティの設定方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="924d6-151">The following example shows how to set the **author** property.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="924d6-152">カスタム プロパティを定義することもできます。</span><span class="sxs-lookup"><span data-stu-id="924d6-152">You can also define custom properties.</span></span> <span data-ttu-id="924d6-153">DocumentProperties オブジェクトには `custom` プロパティが含まれていて、ユーザー定義プロパティのキー値のペアのコレクションを表します。</span><span class="sxs-lookup"><span data-stu-id="924d6-153">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="924d6-154">次の例では、"Hello" という値を持つ **Introduction** という名前のカスタム プロパティを作成し、それを取得する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="924d6-154">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

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
    customProperty.load("key, value");

    return context.sync().then(function() {
        console.log("Custom key  : " + customProperty.key); // "Introduction"
        console.log("Custom value : " + customProperty.value); // "Hello"
    });
}).catch(errorHandlerFunction);
```

## <a name="access-document-settings"></a><span data-ttu-id="924d6-155">ドキュメント設定へのアクセス</span><span class="sxs-lookup"><span data-stu-id="924d6-155">Access document settings</span></span>

<span data-ttu-id="924d6-156">ブックの設定は、カスタム プロパティのコレクションに似ています。</span><span class="sxs-lookup"><span data-stu-id="924d6-156">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="924d6-157">設定は 1 つの Excel ファイルとアドインのペアリングに固有であるのに対して、プロパティはファイルに接続しているだけである点が異なります。</span><span class="sxs-lookup"><span data-stu-id="924d6-157">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="924d6-158">次の例は、設定を作成してアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="924d6-158">The following example shows how to create and access a setting.</span></span>

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

## <a name="add-custom-xml-data-to-the-workbook"></a><span data-ttu-id="924d6-159">カスタム XML データをブックに追加する</span><span class="sxs-lookup"><span data-stu-id="924d6-159">Add custom XML data to the workbook</span></span>

<span data-ttu-id="924d6-160">Excel の Open XML **.xlsx** ファイル形式を使用すると、アドインでカスタム XML データをブックに埋め込むことができます。</span><span class="sxs-lookup"><span data-stu-id="924d6-160">Excel's Open XML **.xlsx** file format lets your add-in embed custom XML data in the workbook.</span></span> <span data-ttu-id="924d6-161">このデータは、アドインに関係なく、ブックで保持されます。</span><span class="sxs-lookup"><span data-stu-id="924d6-161">This data persists with the workbook, independent of the add-in.</span></span>

<span data-ttu-id="924d6-162">ブックには [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection) が含まれます。これは [CustomXmlParts](/javascript/api/excel/excel.customxmlpart) のリストです。</span><span class="sxs-lookup"><span data-stu-id="924d6-162">A workbook contains a [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), which is a list of [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span></span> <span data-ttu-id="924d6-163">これにより、XML 文字列と対応する一意の ID へのアクセスが提供されます。</span><span class="sxs-lookup"><span data-stu-id="924d6-163">These give access to the XML strings and a corresponding unique ID.</span></span> <span data-ttu-id="924d6-164">これらの ID を設定として保管することにより、アドインはセッション間で XML パーツへのキーを保持できます。</span><span class="sxs-lookup"><span data-stu-id="924d6-164">By storing these IDs as settings, your add-in can maintain the keys to its XML parts between sessions.</span></span>

<span data-ttu-id="924d6-165">以下のサンプルは、カスタム XML パーツを使用する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="924d6-165">The following samples show how to use custom XML parts.</span></span> <span data-ttu-id="924d6-166">最初のコード ブロックは、XML データをドキュメントに埋め込む方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="924d6-166">The first code block demonstrates how to embed XML data in the document.</span></span> <span data-ttu-id="924d6-167">レビュー担当者のリストを保管してから、ブックの設定を使用して XML の `id` を保存して、後から取得できるようにします。</span><span class="sxs-lookup"><span data-stu-id="924d6-167">It stores a list of reviewers, then uses the workbook's settings to save the XML's `id` for future retrieval.</span></span> <span data-ttu-id="924d6-168">2 番目のブロックでは、後からその XML にアクセスする方法が示されています。</span><span class="sxs-lookup"><span data-stu-id="924d6-168">The second block shows how to access that XML later.</span></span> <span data-ttu-id="924d6-169">"ContosoReviewXmlPartId" 設定がロードされ、ブックの `customXmlParts` に渡されます。</span><span class="sxs-lookup"><span data-stu-id="924d6-169">The "ContosoReviewXmlPartId" setting is loaded and passed to the workbook's `customXmlParts`.</span></span> <span data-ttu-id="924d6-170">それから、XML データがコンソールに出力されます。</span><span class="sxs-lookup"><span data-stu-id="924d6-170">The XML data is then printed to the console.</span></span>

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
> <span data-ttu-id="924d6-171">`CustomXMLPart.namespaceUri` にデータが入れられるのは、トップレベルのカスタム XML 要素に `xmlns` 属性が含まれている場合に限ります。</span><span class="sxs-lookup"><span data-stu-id="924d6-171">`CustomXMLPart.namespaceUri` is only populated if the top-level custom XML element contains the `xmlns` attribute.</span></span>

## <a name="control-calculation-behavior"></a><span data-ttu-id="924d6-172">計算の動作を制御する</span><span class="sxs-lookup"><span data-stu-id="924d6-172">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="924d6-173">計算モードを設定する</span><span class="sxs-lookup"><span data-stu-id="924d6-173">Set calculation mode</span></span>

<span data-ttu-id="924d6-174">既定では、Excel は、参照されているセルが変更されたときに数式の結果を再計算します。</span><span class="sxs-lookup"><span data-stu-id="924d6-174">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="924d6-175">この計算の動作を調整すると、アドインのパフォーマンス向上に役立つ場合があります。</span><span class="sxs-lookup"><span data-stu-id="924d6-175">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="924d6-176">Application オブジェクトには、`CalculationMode` 型のプロパティ `calculationMode` があります。</span><span class="sxs-lookup"><span data-stu-id="924d6-176">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="924d6-177">次のいずれかの値を設定できます。</span><span class="sxs-lookup"><span data-stu-id="924d6-177">It can be set to the following values:</span></span>

- <span data-ttu-id="924d6-178">`automatic`: 既定の再計算動作。関連するデータが変更されるたびに、Excel は新しい数式の結果を計算します。</span><span class="sxs-lookup"><span data-stu-id="924d6-178">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
- <span data-ttu-id="924d6-179">`automaticExceptTables`: `automatic` と同様ですが、テーブル内の値に加えた変更は無視されます。</span><span class="sxs-lookup"><span data-stu-id="924d6-179">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
- <span data-ttu-id="924d6-180">`manual`: ユーザーまたはアドインが要求した場合にのみ計算します。</span><span class="sxs-lookup"><span data-stu-id="924d6-180">`manual`: Calculations only occur when the user or add-in requests them.</span></span>

### <a name="set-calculation-type"></a><span data-ttu-id="924d6-181">計算タイプを設定する</span><span class="sxs-lookup"><span data-stu-id="924d6-181">Set calculation type</span></span>

<span data-ttu-id="924d6-182">[Application](/javascript/api/excel/excel.application) オブジェクトは、強制的に即時再計算する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="924d6-182">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="924d6-183">`Application.calculate(calculationType)` は、指定した `calculationType` に基づいて手動再計算を開始します。</span><span class="sxs-lookup"><span data-stu-id="924d6-183">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="924d6-184">次の値を指定できます。</span><span class="sxs-lookup"><span data-stu-id="924d6-184">The following values can be specified:</span></span>

- <span data-ttu-id="924d6-185">`full`: 最後に再計算されてから変更されたかどうかに関係なく、開いているすべてのブックのすべての数式を再計算します。</span><span class="sxs-lookup"><span data-stu-id="924d6-185">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="924d6-186">`fullRebuild`: 最後に再計算されてから変更されたかどうかに関係なく、依存関係のある数式を確認してから、開いているすべてのブックのすべての数式を再計算します。</span><span class="sxs-lookup"><span data-stu-id="924d6-186">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="924d6-187">`recalculate`: すべてのアクティブなブックで、最後に計算されてから変更された数式 (またはプログラムで再計算用にマークされている数式)、およびそれに依存する数式を再計算します。</span><span class="sxs-lookup"><span data-stu-id="924d6-187">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>

> [!NOTE]
> <span data-ttu-id="924d6-188">再計算の詳細については、「[数式の再計算、反復計算、または精度を変更する](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="924d6-188">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="924d6-189">計算を一時的に中断する</span><span class="sxs-lookup"><span data-stu-id="924d6-189">Temporarily suspend calculations</span></span>

<span data-ttu-id="924d6-190">Excel API では、アドインから `RequestContext.sync()` を呼び出すまで計算をオフにすることもできます。</span><span class="sxs-lookup"><span data-stu-id="924d6-190">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="924d6-191">これは、`suspendApiCalculationUntilNextSync()` で実行できます。</span><span class="sxs-lookup"><span data-stu-id="924d6-191">This is done with `suspendApiCalculationUntilNextSync()`.</span></span> <span data-ttu-id="924d6-192">このメソッドは、アドインから大きな範囲を編集し、複数の編集の間でデータにアクセスする必要がない場合に使用します。</span><span class="sxs-lookup"><span data-stu-id="924d6-192">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="see-also"></a><span data-ttu-id="924d6-193">関連項目</span><span class="sxs-lookup"><span data-stu-id="924d6-193">See also</span></span>

- [<span data-ttu-id="924d6-194">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="924d6-194">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="924d6-195">Excel JavaScript API を使用してワークシートを操作する</span><span class="sxs-lookup"><span data-stu-id="924d6-195">Work with worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
- [<span data-ttu-id="924d6-196">Excel JavaScript API を使用して範囲を操作する</span><span class="sxs-lookup"><span data-stu-id="924d6-196">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
