---
title: Excel JavaScript API を使用してブックを操作する
description: ''
ms.date: 09/03/2019
localization_priority: Priority
ms.openlocfilehash: eb2313203e770e173d4db12d2bbc03048a08acaa
ms.sourcegitcommit: 78998a9f0ebb81c4dd2b77574148b16fe6725cfc
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/03/2019
ms.locfileid: "36715621"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="8f090-102">Excel JavaScript API を使用してブックを操作する</span><span class="sxs-lookup"><span data-stu-id="8f090-102">Work with workbooks using the Excel JavaScript API</span></span>

<span data-ttu-id="8f090-103">この記事では、Excel JavaScript API を使用して、ブックでタスクを実行する方法のコード サンプルを示しています。</span><span class="sxs-lookup"><span data-stu-id="8f090-103">This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API.</span></span> <span data-ttu-id="8f090-104">**Workbook** オブジェクトがサポートするプロパティとメソッドの完全な一覧については、「[Workbook オブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.workbook)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8f090-104">For the complete list of properties and methods that the **Workbook** object supports, see [Workbook Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="8f090-105">この記事では、[Application](/javascript/api/excel/excel.application) オブジェクトを使用して実行するブック レベルのアクションについても説明します。</span><span class="sxs-lookup"><span data-stu-id="8f090-105">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="8f090-106">Workbook オブジェクトは、Excel を操作するアドインのエントリ ポイントです。</span><span class="sxs-lookup"><span data-stu-id="8f090-106">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="8f090-107">このオブジェクトは、Excel データのアクセスや変更に使用するワークシート、テーブル、ピボットテーブル、その他のコレクションを保持します。</span><span class="sxs-lookup"><span data-stu-id="8f090-107">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="8f090-108">[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) オブジェクトは、個々のワークシートを使用して、ブックのすべてのデータにアドインからアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="8f090-108">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through individual worksheets.</span></span> <span data-ttu-id="8f090-109">具体的には、アドインからワークシートの追加、ワークシート間の移動、ワークシート イベントへのハンドラーの割り当てができます。</span><span class="sxs-lookup"><span data-stu-id="8f090-109">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="8f090-110">ワークシートへのアクセスと編集の方法については、「[Excel JavaScript API を使用してワークシートを操作する](excel-add-ins-worksheets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8f090-110">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="8f090-111">アクティブ セルまたは選択した範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="8f090-111">Get the active cell or selected range</span></span>

<span data-ttu-id="8f090-112">Workbook オブジェクトには、ユーザーまたはアドインが選択したセルの範囲を取得する 2 つのメソッド `getActiveCell()` と `getSelectedRange()` があります。</span><span class="sxs-lookup"><span data-stu-id="8f090-112">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="8f090-113">`getActiveCell()` はブックからアクティブ セルを [Range オブジェクト](/javascript/api/excel/excel.range)として取得します。</span><span class="sxs-lookup"><span data-stu-id="8f090-113">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="8f090-114">次の例では、`getActiveCell()` を呼び出し、コンソールにセルのアドレスを表示します。</span><span class="sxs-lookup"><span data-stu-id="8f090-114">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="8f090-115">`getSelectedRange()` メソッドは現在選択されている単一の範囲を返します。</span><span class="sxs-lookup"><span data-stu-id="8f090-115">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="8f090-116">複数の範囲が選択されている場合は、InvalidSelection エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="8f090-116">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="8f090-117">次の例では、`getSelectedRange()` を呼び出し、範囲の塗りつぶし色を黄色に設定します。</span><span class="sxs-lookup"><span data-stu-id="8f090-117">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="8f090-118">ブックを作成する</span><span class="sxs-lookup"><span data-stu-id="8f090-118">Create a workbook</span></span>

<span data-ttu-id="8f090-119">アドインでは、アドインが現在実行されている Excel のインスタンスとは異なる新しいブックを作成できます。</span><span class="sxs-lookup"><span data-stu-id="8f090-119">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="8f090-120">Excel オブジェクトには、この目的の `createWorkbook` メソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="8f090-120">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="8f090-121">このメソッドが呼び出されると、新しいブックが Excel の新しいインスタンスですぐに開いて表示されます。</span><span class="sxs-lookup"><span data-stu-id="8f090-121">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="8f090-122">アドインは前のブックで開いて実行されたままになります。</span><span class="sxs-lookup"><span data-stu-id="8f090-122">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="8f090-123">`createWorkbook` メソッドは既存のブックのコピーの作成もできます。</span><span class="sxs-lookup"><span data-stu-id="8f090-123">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="8f090-124">このメソッドは、オプションのパラメーターとして .xlsx ファイルの base64 エンコード文字列表現を受け取ります。</span><span class="sxs-lookup"><span data-stu-id="8f090-124">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="8f090-125">文字列の引数は有効な .xlsx ファイルと見なされ、作成されるブックはそのファイルのコピーになります。</span><span class="sxs-lookup"><span data-stu-id="8f090-125">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="8f090-126">[ファイルのスライス](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)を使用して、アドインの現在のブックを base64 エンコード文字列として取得できます。</span><span class="sxs-lookup"><span data-stu-id="8f090-126">You can get your add-in’s current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="8f090-127">次の例に示すように、[FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) クラスを使用して、ファイルを必要な base64 エンコード文字列に変換できます。</span><span class="sxs-lookup"><span data-stu-id="8f090-127">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

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

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one-preview"></a><span data-ttu-id="8f090-128">既存のブックのコピーを現在のブックに挿入する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="8f090-128">Insert a copy of an existing workbook into the current one</span></span>

> [!NOTE]
> <span data-ttu-id="8f090-129">`WorksheetCollection.addFromBase64` メソッドは、現在パブリック プレビューでのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="8f090-129">The `WorksheetCollection.addFromBase64` method described in this article is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="8f090-130">前の例は、既存のブックから作成された新しいブックを示しています。</span><span class="sxs-lookup"><span data-stu-id="8f090-130">The previous example shows a new workbook being created from an existing workbook.</span></span> <span data-ttu-id="8f090-131">既存のブックの一部またはすべてを、アドインに関連付けられているブックにコピーすることもできます。</span><span class="sxs-lookup"><span data-stu-id="8f090-131">You can also copy some or all of an existing workbook into the one currently associated with your add-in.</span></span> <span data-ttu-id="8f090-132">ブックの [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) にある `addFromBase64` メソッドは、対象のブックのワークシートのコピーを現在のブックに挿入します。</span><span class="sxs-lookup"><span data-stu-id="8f090-132">A workbook's [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) has the `addFromBase64` method to insert copies of the target workbook's worksheets into itself.</span></span> <span data-ttu-id="8f090-133">他のブックのファイルは、`Excel.createWorkbook` 呼び出しの場合と同様に、base64 エンコード文字列として渡されます。</span><span class="sxs-lookup"><span data-stu-id="8f090-133">The other workbook's file is passed as base64-encoded string, just like the `Excel.createWorkbook` call.</span></span>

```TypeScript
addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet | string): OfficeExtension.ClientResult<string[]>;
```

<span data-ttu-id="8f090-134">次の例では、ブックのワークシートが現在のブックのアクティブ ワークシートの直後に挿入されています。</span><span class="sxs-lookup"><span data-stu-id="8f090-134">The following example shows a workbook's worksheets being inserted in the current workbook, directly after the active worksheet.</span></span> <span data-ttu-id="8f090-135">`null` が `sheetNamesToInsert?: string[]` パラメーターに渡されている点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="8f090-135">Note that `null` is passed for the `sheetNamesToInsert?: string[]` parameter.</span></span> <span data-ttu-id="8f090-136">つまり、すべてのワークシートが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="8f090-136">This means all the worksheets are being inserted.</span></span>

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

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="8f090-137">ブックのシート構成を保護する</span><span class="sxs-lookup"><span data-stu-id="8f090-137">Protect the workbook's structure</span></span>

<span data-ttu-id="8f090-138">アドインでは、ブックのシート構成を編集するユーザーの機能を制御できます。</span><span class="sxs-lookup"><span data-stu-id="8f090-138">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="8f090-139">Workbook オブジェクトの `protection` プロパティは [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) オブジェクトであり、`protect()` メソッドを備えています。</span><span class="sxs-lookup"><span data-stu-id="8f090-139">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="8f090-140">次の例では、ブックのシート構成の保護を切り替える基本的なシナリオを示します。</span><span class="sxs-lookup"><span data-stu-id="8f090-140">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span>

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

<span data-ttu-id="8f090-141">`protect` メソッドは、オプションの文字列パラメーターを受け取ります。</span><span class="sxs-lookup"><span data-stu-id="8f090-141">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="8f090-142">この文字列は、ユーザーが保護をバイパスしてブックのシート構成を変更するために必要なパスワードを表します。</span><span class="sxs-lookup"><span data-stu-id="8f090-142">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="8f090-143">保護は、不必要なデータ編集をできないようにするため、ワークシート レベルで設定することもできます。</span><span class="sxs-lookup"><span data-stu-id="8f090-143">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="8f090-144">詳細については、「[Excel JavaScript API を使用してワークシートを操作する](excel-add-ins-worksheets.md#data-protection)」の**データの保護**のセクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="8f090-144">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE]
> <span data-ttu-id="8f090-145">Excel のブックの保護の詳細については、「[ブックを保護する](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8f090-145">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="8f090-146">ドキュメント プロパティへのアクセス</span><span class="sxs-lookup"><span data-stu-id="8f090-146">Access document properties</span></span>

<span data-ttu-id="8f090-147">Workbook オブジェクトは、[ドキュメント プロパティ](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75)と呼ばれる Office ファイルのメタデータにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="8f090-147">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="8f090-148">Workbook オブジェクトの `properties` プロパティは、これらのメタデータ値を含む [DocumentProperties](/javascript/api/excel/excel.documentproperties) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8f090-148">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="8f090-149">次のコード例では、**author** プロパティの設定方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8f090-149">The following example shows how to set the **author** property.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="8f090-150">カスタム プロパティを定義することもできます。</span><span class="sxs-lookup"><span data-stu-id="8f090-150">You can also define custom properties.</span></span> <span data-ttu-id="8f090-151">DocumentProperties オブジェクトには `custom` プロパティが含まれていて、ユーザー定義プロパティのキー値のペアのコレクションを表します。</span><span class="sxs-lookup"><span data-stu-id="8f090-151">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="8f090-152">次の例では、"Hello" という値を持つ **Introduction** という名前のカスタム プロパティを作成し、それを取得する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="8f090-152">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

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

## <a name="access-document-settings"></a><span data-ttu-id="8f090-153">ドキュメント設定へのアクセス</span><span class="sxs-lookup"><span data-stu-id="8f090-153">Access document settings</span></span>

<span data-ttu-id="8f090-154">ブックの設定は、カスタム プロパティのコレクションに似ています。</span><span class="sxs-lookup"><span data-stu-id="8f090-154">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="8f090-155">設定は 1 つの Excel ファイルとアドインのペアリングに固有であるのに対して、プロパティはファイルに接続しているだけである点が異なります。</span><span class="sxs-lookup"><span data-stu-id="8f090-155">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="8f090-156">次の例は、設定を作成してアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8f090-156">The following example shows how to create and access a setting.</span></span>

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

## <a name="add-custom-xml-data-to-the-workbook"></a><span data-ttu-id="8f090-157">カスタム XML データをブックに追加する</span><span class="sxs-lookup"><span data-stu-id="8f090-157">Add custom XML data to the workbook</span></span>

<span data-ttu-id="8f090-158">Excel の Open XML **.xlsx** ファイル形式を使用すると、アドインでカスタム XML データをブックに埋め込むことができます。</span><span class="sxs-lookup"><span data-stu-id="8f090-158">Excel's Open XML **.xlsx** file format lets your add-in embed custom XML data in the workbook.</span></span> <span data-ttu-id="8f090-159">このデータは、アドインに関係なく、ブックで保持されます。</span><span class="sxs-lookup"><span data-stu-id="8f090-159">This data persists with the workbook, independent of the add-in.</span></span>

<span data-ttu-id="8f090-160">ブックには [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection) が含まれます。これは [CustomXmlParts](/javascript/api/excel/excel.customxmlpart) のリストです。</span><span class="sxs-lookup"><span data-stu-id="8f090-160">A workbook contains a [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), which is a list of [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span></span> <span data-ttu-id="8f090-161">これにより、XML 文字列と対応する一意の ID へのアクセスが提供されます。</span><span class="sxs-lookup"><span data-stu-id="8f090-161">These give access to the XML strings and a corresponding unique ID.</span></span> <span data-ttu-id="8f090-162">これらの ID を設定として保管することにより、アドインはセッション間で XML パーツへのキーを保持できます。</span><span class="sxs-lookup"><span data-stu-id="8f090-162">By storing these IDs as settings, your add-in can maintain the keys to its XML parts between sessions.</span></span>

<span data-ttu-id="8f090-163">以下のサンプルは、カスタム XML パーツを使用する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8f090-163">The following samples show how to use custom XML parts.</span></span> <span data-ttu-id="8f090-164">最初のコード ブロックは、XML データをドキュメントに埋め込む方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8f090-164">The first code block demonstrates how to embed XML data in the document.</span></span> <span data-ttu-id="8f090-165">レビュー担当者のリストを保管してから、ブックの設定を使用して XML の `id` を保存して、後から取得できるようにします。</span><span class="sxs-lookup"><span data-stu-id="8f090-165">It stores a list of reviewers, then uses the workbook's settings to save the XML's `id` for future retrieval.</span></span> <span data-ttu-id="8f090-166">2 番目のブロックでは、後からその XML にアクセスする方法が示されています。</span><span class="sxs-lookup"><span data-stu-id="8f090-166">The second block shows how to access that XML later.</span></span> <span data-ttu-id="8f090-167">"ContosoReviewXmlPartId" 設定がロードされ、ブックの `customXmlParts` に渡されます。</span><span class="sxs-lookup"><span data-stu-id="8f090-167">The "ContosoReviewXmlPartId" setting is loaded and passed to the workbook's `customXmlParts`.</span></span> <span data-ttu-id="8f090-168">それから、XML データがコンソールに出力されます。</span><span class="sxs-lookup"><span data-stu-id="8f090-168">The XML data is then printed to the console.</span></span>

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
> <span data-ttu-id="8f090-169">`CustomXMLPart.namespaceUri` にデータが入れられるのは、トップレベルのカスタム XML 要素に `xmlns` 属性が含まれている場合に限ります。</span><span class="sxs-lookup"><span data-stu-id="8f090-169">`CustomXMLPart.namespaceUri` is only populated if the top-level custom XML element contains the `xmlns` attribute.</span></span>

## <a name="control-calculation-behavior"></a><span data-ttu-id="8f090-170">計算の動作を制御する</span><span class="sxs-lookup"><span data-stu-id="8f090-170">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="8f090-171">計算モードを設定する</span><span class="sxs-lookup"><span data-stu-id="8f090-171">Set calculation mode</span></span>

<span data-ttu-id="8f090-172">既定では、Excel は、参照されているセルが変更されたときに数式の結果を再計算します。</span><span class="sxs-lookup"><span data-stu-id="8f090-172">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="8f090-173">この計算の動作を調整すると、アドインのパフォーマンス向上に役立つ場合があります。</span><span class="sxs-lookup"><span data-stu-id="8f090-173">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="8f090-174">Application オブジェクトには、`CalculationMode` 型のプロパティ `calculationMode` があります。</span><span class="sxs-lookup"><span data-stu-id="8f090-174">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="8f090-175">次のいずれかの値を設定できます。</span><span class="sxs-lookup"><span data-stu-id="8f090-175">It can be set to the following values:</span></span>

- <span data-ttu-id="8f090-176">`automatic`: 既定の再計算動作。関連するデータが変更されるたびに、Excel は新しい数式の結果を計算します。</span><span class="sxs-lookup"><span data-stu-id="8f090-176">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
- <span data-ttu-id="8f090-177">`automaticExceptTables`: `automatic` と同様ですが、テーブル内の値に加えた変更は無視されます。</span><span class="sxs-lookup"><span data-stu-id="8f090-177">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
- <span data-ttu-id="8f090-178">`manual`: ユーザーまたはアドインが要求した場合にのみ計算します。</span><span class="sxs-lookup"><span data-stu-id="8f090-178">`manual`: Calculations only occur when the user or add-in requests them.</span></span>

### <a name="set-calculation-type"></a><span data-ttu-id="8f090-179">計算タイプを設定する</span><span class="sxs-lookup"><span data-stu-id="8f090-179">Set calculation type</span></span>

<span data-ttu-id="8f090-180">[Application](/javascript/api/excel/excel.application) オブジェクトは、強制的に即時再計算する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="8f090-180">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="8f090-181">`Application.calculate(calculationType)` は、指定した `calculationType` に基づいて手動再計算を開始します。</span><span class="sxs-lookup"><span data-stu-id="8f090-181">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="8f090-182">次の値を指定できます。</span><span class="sxs-lookup"><span data-stu-id="8f090-182">The following values can be specified:</span></span>

- <span data-ttu-id="8f090-183">`full`: 最後に再計算されてから変更されたかどうかに関係なく、開いているすべてのブックのすべての数式を再計算します。</span><span class="sxs-lookup"><span data-stu-id="8f090-183">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="8f090-184">`fullRebuild`: 最後に再計算されてから変更されたかどうかに関係なく、依存関係のある数式を確認してから、開いているすべてのブックのすべての数式を再計算します。</span><span class="sxs-lookup"><span data-stu-id="8f090-184">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="8f090-185">`recalculate`: すべてのアクティブなブックで、最後に計算されてから変更された数式 (またはプログラムで再計算用にマークされている数式)、およびそれに依存する数式を再計算します。</span><span class="sxs-lookup"><span data-stu-id="8f090-185">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>

> [!NOTE]
> <span data-ttu-id="8f090-186">再計算の詳細については、「[数式の再計算、反復計算、または精度を変更する](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8f090-186">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="8f090-187">計算を一時的に中断する</span><span class="sxs-lookup"><span data-stu-id="8f090-187">Temporarily suspend calculations</span></span>

<span data-ttu-id="8f090-188">Excel API では、アドインから `RequestContext.sync()` を呼び出すまで計算をオフにすることもできます。</span><span class="sxs-lookup"><span data-stu-id="8f090-188">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="8f090-189">これは、`suspendApiCalculationUntilNextSync()` で実行できます。</span><span class="sxs-lookup"><span data-stu-id="8f090-189">This is done with `suspendApiCalculationUntilNextSync()`.</span></span> <span data-ttu-id="8f090-190">このメソッドは、アドインから大きな範囲を編集し、複数の編集の間でデータにアクセスする必要がない場合に使用します。</span><span class="sxs-lookup"><span data-stu-id="8f090-190">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="comments-preview"></a><span data-ttu-id="8f090-191">コメント (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="8f090-191">Comments (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="8f090-192">コメント API は現在、パブリック プレビューでのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="8f090-192">The following events are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="8f090-193">ブック内のすべての[コメント](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)は、`Workbook.comments` プロパティによって追跡されます。</span><span class="sxs-lookup"><span data-stu-id="8f090-193">All [comments](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) within a workbook are tracked by the `Workbook.comments` property.</span></span> <span data-ttu-id="8f090-194">これには、ユーザーによって作成されたコメントだけでなく、アドインによって作成されたコメントも含まれます。</span><span class="sxs-lookup"><span data-stu-id="8f090-194">This includes comments created by users and also comments created by your add-in.</span></span> <span data-ttu-id="8f090-195">`Workbook.comments` プロパティは、[Comment](/javascript/api/excel/excel.comment) オブジェクトのコレクションを含む [CommentCollection](/javascript/api/excel/excel.commentcollection) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8f090-195">The `Workbook.comments` property is a [CommentCollection](/javascript/api/excel/excel.commentcollection) object that contains a collection of [Comment](/javascript/api/excel/excel.comment) objects.</span></span>

<span data-ttu-id="8f090-196">コメントをブックに追加するには、`CommentCollection.add` メソッドを使用して、コメントのテキストを文字列として渡し、コメントを追加するセルを文字列または [Range](/javascript/api/excel/excel.range) オブジェクトのいずれかとして渡します。</span><span class="sxs-lookup"><span data-stu-id="8f090-196">To add comments to a workbook, use the `CommentCollection.add` method, passing in the comment's text, as a string, and the cell where the comment will be added, as either a string or [Range](/javascript/api/excel/excel.range) object.</span></span> <span data-ttu-id="8f090-197">次のコード例は、コメントをセル **A2** に追加します。</span><span class="sxs-lookup"><span data-stu-id="8f090-197">The following code sample adds a comment to cell **A2**.</span></span>

```js
Excel.run(function (context) {
    var comments = context.workbook.comments;

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `Comment.add`.
    comments.add("TODO: add data.", "A2");
    return context.sync();
});
```

<span data-ttu-id="8f090-198">各コメントには、作成者や作成日などの作成に関するメタデータが含まれています。</span><span class="sxs-lookup"><span data-stu-id="8f090-198">Each comment contains metadata about its creation, such as the author and creation date.</span></span> <span data-ttu-id="8f090-199">アドインによって作成されたコメントは、現在のユーザーによって作成されたものと見なされます。</span><span class="sxs-lookup"><span data-stu-id="8f090-199">Comments created by your add-in are considered to be authored by the current user.</span></span> <span data-ttu-id="8f090-200">次のサンプルは、**A2** に作成者のメール、作成者の名前、コメントの作成日を表示する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8f090-200">The following sample shows how to display the author's email, author's name, and creation date of a comment at **A2**.</span></span>

```js
Excel.run(function (context) {
    // Get the comment at cell A2.
    var comment = context.workbook.comments.getItemByCell("Comments!A2");
    comment.load(["authorEmail", "authorName", "creationDate"]);
    return context.sync().then(function () {
        console.log(`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`);
    });
});
```

<span data-ttu-id="8f090-201">各コメントには、0 個以上の返信が含まれます。</span><span class="sxs-lookup"><span data-stu-id="8f090-201">Each comment contains zero or more replies.</span></span> <span data-ttu-id="8f090-202">`Comment` オブジェクトには `replies` プロパティがあり、これは [CommentReply](/javascript/api/excel/excel.commentreply) オブジェクトを含む [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) です。</span><span class="sxs-lookup"><span data-stu-id="8f090-202">`Comment` objects have a `replies` property, which is a [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) that contains [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="8f090-203">コメントに返信を追加するには、`CommentReplyCollection.add` メソッドを使用して、返信のテキストを渡します。</span><span class="sxs-lookup"><span data-stu-id="8f090-203">To add a reply to a comment, use the `CommentReplyCollection.add` method, passing in the text of the reply.</span></span> <span data-ttu-id="8f090-204">返信は、追加された順に表示されます。</span><span class="sxs-lookup"><span data-stu-id="8f090-204">Replies are displayed in the order they are added.</span></span> <span data-ttu-id="8f090-205">次のコード サンプルは、ブックの最初のコメントに返信を追加します。</span><span class="sxs-lookup"><span data-stu-id="8f090-205">The following code sample adds a data series to the first chart in the worksheet.</span></span>

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

<span data-ttu-id="8f090-206">コメントまたはコメントの返信を編集するには、その `Comment.content` プロパティまたは `CommentReply.content` プロパティを設定します。</span><span class="sxs-lookup"><span data-stu-id="8f090-206">To edit a comment or comment reply, set its `Comment.content` property or `CommentReply.content` property.</span></span> <span data-ttu-id="8f090-207">コメントまたはコメントの返信を削除するには、`Comment.delete` メソッドまたは `CommentReply.delete` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="8f090-207">To delete a comment or comment reply, use the `Comment.delete` method or `CommentReply.delete` method.</span></span> <span data-ttu-id="8f090-208">コメントを削除すると、そのコメントに関連付けられている返信もすべて削除されます。</span><span class="sxs-lookup"><span data-stu-id="8f090-208">Deleting a comment also deletes all the replies associated with that comment.</span></span>

> [!TIP]
> <span data-ttu-id="8f090-209">コメントは、同じ手法を使用して[ワークシート](/javascript/api/excel/excel.worksheet) レベルでも管理できます。</span><span class="sxs-lookup"><span data-stu-id="8f090-209">Comments can also be managed at the [Worksheet](/javascript/api/excel/excel.worksheet) level using the same techniques.</span></span>

## <a name="save-the-workbook-preview"></a><span data-ttu-id="8f090-210">ブックを保存する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="8f090-210">Save the workbook</span></span>

> [!NOTE]
> <span data-ttu-id="8f090-211">`Workbook.save` メソッドは、現在パブリック プレビューでのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="8f090-211">The `Workbook.save` method described in this article is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="8f090-212">`Workbook.save` は、ブックを永続記憶装置に保存します。</span><span class="sxs-lookup"><span data-stu-id="8f090-212">`Workbook.save` saves the workbook to persistent storage .</span></span> <span data-ttu-id="8f090-213">`save` メソッドにはオプションの `saveBehavior` パラメーターを 1 つ指定できます。値は次のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="8f090-213">The `save` method takes a single, optional parameter that can be one of the following values:</span></span>

- <span data-ttu-id="8f090-214">`Excel.SaveBehavior.save` (既定値): ファイル名や保存場所を指定するようにユーザーに促すダイアログは表示されず、そのままファイルが保存されます。</span><span class="sxs-lookup"><span data-stu-id="8f090-214">`Excel.SaveBehavior.save` (default): The file is saved without prompting the user to specify file name and save location.</span></span> <span data-ttu-id="8f090-215">ファイルが以前に保存されていない場合は、既定の場所に保存されます。</span><span class="sxs-lookup"><span data-stu-id="8f090-215">If the file has not been saved previously, it's saved to the default location.</span></span> <span data-ttu-id="8f090-216">ファイルが以前に保存されている場合は、同じ場所に保存されます。</span><span class="sxs-lookup"><span data-stu-id="8f090-216">If the file has been saved previously, it's saved to the same location.</span></span>
- <span data-ttu-id="8f090-217">`Excel.SaveBehavior.prompt`: ファイルが以前に保存されていない場合は、ファイル名や保存場所を指定するようにユーザーに促すダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8f090-217">`Excel.SaveBehavior.prompt`: If file has not been saved previously, the user will be prompted to specify file name and save location.</span></span> <span data-ttu-id="8f090-218">ファイルが以前に保存されている場合、ファイルは同じ場所に保存され、ダイアログは表示されません。</span><span class="sxs-lookup"><span data-stu-id="8f090-218">If the file has been saved previously, it will be saved to the same location and the user will not be prompted.</span></span>

> [!CAUTION]
> <span data-ttu-id="8f090-219">保存を促すダイアログが表示されたのにユーザーがその操作をキャンセルすると、`save` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="8f090-219">If the user is prompted to save and cancels the operation, `save` throws an exception.</span></span>

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook-preview"></a><span data-ttu-id="8f090-220">ブックを閉じる (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="8f090-220">Close the workbook</span></span>

> [!NOTE]
> <span data-ttu-id="8f090-221">`Workbook.close` メソッドは、現在パブリック プレビューでのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="8f090-221">The `Workbook.close` method described in this article is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="8f090-222">`Workbook.close` は、ブックとそのブックに関連付けられているアドインを終了します (Excel アプリケーションは開いたまま)。</span><span class="sxs-lookup"><span data-stu-id="8f090-222">`Workbook.close` closes the workbook, along with add-ins that are associated with the workbook (the Excel application remains open).</span></span> <span data-ttu-id="8f090-223">`close` メソッドにはオプションの `closeBehavior` パラメーターを 1 つ指定できます。値は次のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="8f090-223">The `close` method takes a single, optional parameter that can be one of the following values:</span></span>

- <span data-ttu-id="8f090-224">`Excel.CloseBehavior.save` (既定値): ファイルは閉じる前に保存されます。</span><span class="sxs-lookup"><span data-stu-id="8f090-224">`Excel.CloseBehavior.save` (default): The file is saved before closing.</span></span> <span data-ttu-id="8f090-225">そのファイルが以前に保存されていない場合は、ファイル名や保存場所を指定するようにユーザーに促すダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8f090-225">If the file has not been saved previously, the user will be prompted to specify file name and save location.</span></span>
- <span data-ttu-id="8f090-226">`Excel.CloseBehavior.skipSave`: ファイルはそのまま閉じられ、保存されません。</span><span class="sxs-lookup"><span data-stu-id="8f090-226">`Excel.CloseBehavior.skipSave`: The file is immediately closed, without saving.</span></span> <span data-ttu-id="8f090-227">未保存の変更は失われます。</span><span class="sxs-lookup"><span data-stu-id="8f090-227">Any unsaved changes will be lost.</span></span>

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a><span data-ttu-id="8f090-228">関連項目</span><span class="sxs-lookup"><span data-stu-id="8f090-228">See also</span></span>

- [<span data-ttu-id="8f090-229">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="8f090-229">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="8f090-230">Excel JavaScript API を使用してワークシートを操作する</span><span class="sxs-lookup"><span data-stu-id="8f090-230">Work with worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
- [<span data-ttu-id="8f090-231">Excel JavaScript API を使用して範囲を操作する</span><span class="sxs-lookup"><span data-stu-id="8f090-231">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
