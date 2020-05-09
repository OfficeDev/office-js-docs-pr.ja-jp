---
title: Excel JavaScript API を使用してブックを操作する
description: Excel JavaScript API を使用して、ブックまたはアプリケーションレベルの機能を使用して一般的なタスクを実行する方法を示すコードサンプルです。
ms.date: 05/06/2020
localization_priority: Normal
ms.openlocfilehash: 4fec6a217a2764eaf664463943ca384b3a2d847b
ms.sourcegitcommit: 735bf94ac3c838f580a992e7ef074dbc8be2b0ea
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/08/2020
ms.locfileid: "44170766"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="36400-103">Excel JavaScript API を使用してブックを操作する</span><span class="sxs-lookup"><span data-stu-id="36400-103">Work with workbooks using the Excel JavaScript API</span></span>

<span data-ttu-id="36400-104">この記事では、Excel JavaScript API を使用して、ブックでタスクを実行する方法のコード サンプルを示しています。</span><span class="sxs-lookup"><span data-stu-id="36400-104">This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API.</span></span> <span data-ttu-id="36400-105">オブジェクトが`Workbook`サポートするプロパティとメソッドの完全な一覧については、「 [Workbook オブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.workbook)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="36400-105">For the complete list of properties and methods that the `Workbook` object supports, see [Workbook Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="36400-106">この記事では、[Application](/javascript/api/excel/excel.application) オブジェクトを使用して実行するブック レベルのアクションについても説明します。</span><span class="sxs-lookup"><span data-stu-id="36400-106">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="36400-107">Workbook オブジェクトは、Excel を操作するアドインのエントリ ポイントです。</span><span class="sxs-lookup"><span data-stu-id="36400-107">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="36400-108">このオブジェクトは、Excel データのアクセスや変更に使用するワークシート、テーブル、ピボットテーブル、その他のコレクションを保持します。</span><span class="sxs-lookup"><span data-stu-id="36400-108">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="36400-109">[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) オブジェクトは、個々のワークシートを使用して、ブックのすべてのデータにアドインからアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="36400-109">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through individual worksheets.</span></span> <span data-ttu-id="36400-110">具体的には、アドインからワークシートの追加、ワークシート間の移動、ワークシート イベントへのハンドラーの割り当てができます。</span><span class="sxs-lookup"><span data-stu-id="36400-110">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="36400-111">ワークシートへのアクセスと編集の方法については、「[Excel JavaScript API を使用してワークシートを操作する](excel-add-ins-worksheets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="36400-111">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="36400-112">アクティブ セルまたは選択した範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="36400-112">Get the active cell or selected range</span></span>

<span data-ttu-id="36400-113">Workbook オブジェクトには、ユーザーまたはアドインが選択したセルの範囲を取得する 2 つのメソッド `getActiveCell()` と `getSelectedRange()` があります。</span><span class="sxs-lookup"><span data-stu-id="36400-113">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="36400-114">`getActiveCell()` はブックからアクティブ セルを [Range オブジェクト](/javascript/api/excel/excel.range)として取得します。</span><span class="sxs-lookup"><span data-stu-id="36400-114">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="36400-115">次の例では、`getActiveCell()` を呼び出し、コンソールにセルのアドレスを表示します。</span><span class="sxs-lookup"><span data-stu-id="36400-115">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="36400-116">`getSelectedRange()` メソッドは現在選択されている単一の範囲を返します。</span><span class="sxs-lookup"><span data-stu-id="36400-116">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="36400-117">複数の範囲が選択されている場合は、InvalidSelection エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="36400-117">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="36400-118">次の例では、`getSelectedRange()` を呼び出し、範囲の塗りつぶし色を黄色に設定します。</span><span class="sxs-lookup"><span data-stu-id="36400-118">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="36400-119">ブックを作成する</span><span class="sxs-lookup"><span data-stu-id="36400-119">Create a workbook</span></span>

<span data-ttu-id="36400-120">アドインでは、アドインが現在実行されている Excel のインスタンスとは異なる新しいブックを作成できます。</span><span class="sxs-lookup"><span data-stu-id="36400-120">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="36400-121">Excel オブジェクトには、この目的の `createWorkbook` メソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="36400-121">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="36400-122">このメソッドが呼び出されると、新しいブックが Excel の新しいインスタンスですぐに開いて表示されます。</span><span class="sxs-lookup"><span data-stu-id="36400-122">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="36400-123">アドインは前のブックで開いて実行されたままになります。</span><span class="sxs-lookup"><span data-stu-id="36400-123">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="36400-124">`createWorkbook` メソッドは既存のブックのコピーの作成もできます。</span><span class="sxs-lookup"><span data-stu-id="36400-124">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="36400-125">このメソッドは、オプションのパラメーターとして .xlsx ファイルの base64 エンコード文字列表現を受け取ります。</span><span class="sxs-lookup"><span data-stu-id="36400-125">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="36400-126">文字列の引数は有効な .xlsx ファイルと見なされ、作成されるブックはそのファイルのコピーになります。</span><span class="sxs-lookup"><span data-stu-id="36400-126">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="36400-127">[ファイルスライシング](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)を使用して、アドインの現在のブックを base64 でエンコードされた文字列として取得できます。</span><span class="sxs-lookup"><span data-stu-id="36400-127">You can get your add-in's current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="36400-128">次の例に示すように、[FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) クラスを使用して、ファイルを必要な base64 エンコード文字列に変換できます。</span><span class="sxs-lookup"><span data-stu-id="36400-128">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

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

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one-preview"></a><span data-ttu-id="36400-129">既存のブックのコピーを現在のブックに挿入する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="36400-129">Insert a copy of an existing workbook into the current one (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="36400-130">`WorksheetCollection.addFromBase64` メソッドは現在、パブリック プレビューでのみ使用でき、Windows および Mac 上の Office でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="36400-130">The `WorksheetCollection.addFromBase64` method is currently only available in public preview and only for Office on Windows and Mac.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="36400-131">前の例は、既存のブックから作成された新しいブックを示しています。</span><span class="sxs-lookup"><span data-stu-id="36400-131">The previous example shows a new workbook being created from an existing workbook.</span></span> <span data-ttu-id="36400-132">既存のブックの一部またはすべてを、アドインに関連付けられているブックにコピーすることもできます。</span><span class="sxs-lookup"><span data-stu-id="36400-132">You can also copy some or all of an existing workbook into the one currently associated with your add-in.</span></span> <span data-ttu-id="36400-133">ブックの [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) にある `addFromBase64` メソッドは、対象のブックのワークシートのコピーを現在のブックに挿入します。</span><span class="sxs-lookup"><span data-stu-id="36400-133">A workbook's [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) has the `addFromBase64` method to insert copies of the target workbook's worksheets into itself.</span></span> <span data-ttu-id="36400-134">他のブックのファイルは、`Excel.createWorkbook` 呼び出しの場合と同様に、base64 エンコード文字列として渡されます。</span><span class="sxs-lookup"><span data-stu-id="36400-134">The other workbook's file is passed as base64-encoded string, just like the `Excel.createWorkbook` call.</span></span>

```TypeScript
addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet | string): OfficeExtension.ClientResult<string[]>;
```

<span data-ttu-id="36400-135">次の例では、ブックのワークシートが現在のブックのアクティブ ワークシートの直後に挿入されています。</span><span class="sxs-lookup"><span data-stu-id="36400-135">The following example shows a workbook's worksheets being inserted in the current workbook, directly after the active worksheet.</span></span> <span data-ttu-id="36400-136">`null` が `sheetNamesToInsert?: string[]` パラメーターに渡されている点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="36400-136">Note that `null` is passed for the `sheetNamesToInsert?: string[]` parameter.</span></span> <span data-ttu-id="36400-137">つまり、すべてのワークシートが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="36400-137">This means all the worksheets are being inserted.</span></span>

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

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="36400-138">ブックのシート構成を保護する</span><span class="sxs-lookup"><span data-stu-id="36400-138">Protect the workbook's structure</span></span>

<span data-ttu-id="36400-139">アドインでは、ブックのシート構成を編集するユーザーの機能を制御できます。</span><span class="sxs-lookup"><span data-stu-id="36400-139">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="36400-140">Workbook オブジェクトの `protection` プロパティは [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) オブジェクトであり、`protect()` メソッドを備えています。</span><span class="sxs-lookup"><span data-stu-id="36400-140">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="36400-141">次の例では、ブックのシート構成の保護を切り替える基本的なシナリオを示します。</span><span class="sxs-lookup"><span data-stu-id="36400-141">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span>

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

<span data-ttu-id="36400-142">`protect` メソッドは、オプションの文字列パラメーターを受け取ります。</span><span class="sxs-lookup"><span data-stu-id="36400-142">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="36400-143">この文字列は、ユーザーが保護をバイパスしてブックのシート構成を変更するために必要なパスワードを表します。</span><span class="sxs-lookup"><span data-stu-id="36400-143">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="36400-144">保護は、不必要なデータ編集をできないようにするため、ワークシート レベルで設定することもできます。</span><span class="sxs-lookup"><span data-stu-id="36400-144">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="36400-145">詳細については、「[Excel JavaScript API を使用してワークシートを操作する](excel-add-ins-worksheets.md#data-protection)」の**データの保護**のセクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="36400-145">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE]
> <span data-ttu-id="36400-146">Excel のブックの保護の詳細については、「[ブックを保護する](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="36400-146">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="36400-147">ドキュメント プロパティへのアクセス</span><span class="sxs-lookup"><span data-stu-id="36400-147">Access document properties</span></span>

<span data-ttu-id="36400-148">Workbook オブジェクトは、[ドキュメント プロパティ](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75)と呼ばれる Office ファイルのメタデータにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="36400-148">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="36400-149">Workbook オブジェクトの `properties` プロパティは、これらのメタデータ値を含む [DocumentProperties](/javascript/api/excel/excel.documentproperties) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="36400-149">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="36400-150">次の例は、 `author`プロパティを設定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="36400-150">The following example shows how to set the `author` property.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="36400-151">カスタム プロパティを定義することもできます。</span><span class="sxs-lookup"><span data-stu-id="36400-151">You can also define custom properties.</span></span> <span data-ttu-id="36400-152">DocumentProperties オブジェクトには `custom` プロパティが含まれていて、ユーザー定義プロパティのキー値のペアのコレクションを表します。</span><span class="sxs-lookup"><span data-stu-id="36400-152">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="36400-153">次の例では、"Hello" という値を持つ **Introduction** という名前のカスタム プロパティを作成し、それを取得する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="36400-153">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

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

## <a name="access-document-settings"></a><span data-ttu-id="36400-154">ドキュメント設定へのアクセス</span><span class="sxs-lookup"><span data-stu-id="36400-154">Access document settings</span></span>

<span data-ttu-id="36400-155">ブックの設定は、カスタム プロパティのコレクションに似ています。</span><span class="sxs-lookup"><span data-stu-id="36400-155">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="36400-156">設定は 1 つの Excel ファイルとアドインのペアリングに固有であるのに対して、プロパティはファイルに接続しているだけである点が異なります。</span><span class="sxs-lookup"><span data-stu-id="36400-156">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="36400-157">次の例は、設定を作成してアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="36400-157">The following example shows how to create and access a setting.</span></span>

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

## <a name="access-application-culture-settings"></a><span data-ttu-id="36400-158">Access アプリケーションのカルチャ設定</span><span class="sxs-lookup"><span data-stu-id="36400-158">Access application culture settings</span></span>

<span data-ttu-id="36400-159">ブックには、特定のデータの表示方法に影響する言語とカルチャの設定が含まれています。</span><span class="sxs-lookup"><span data-stu-id="36400-159">A workbook has language and culture settings that affect how certain data is displayed.</span></span> <span data-ttu-id="36400-160">これらの設定は、アドインのユーザーが異なる言語とカルチャでブックを共有している場合に、データをローカライズするのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="36400-160">These settings can help localize data when your add-in's users are sharing workbooks across different languages and cultures.</span></span> <span data-ttu-id="36400-161">アドインでは、文字列の解析を使用して、各ユーザーが独自のカルチャの形式でデータを表示できるように、システムのカルチャ設定に基づいて数値、日付、時刻の形式をローカライズできます。</span><span class="sxs-lookup"><span data-stu-id="36400-161">Your add-in can use string parsing to localize the format of numbers, dates, and times based on the system culture settings so that each user sees data in their own culture's format.</span></span>

<span data-ttu-id="36400-162">`Application.cultureInfo`システムのカルチャ設定を[CultureInfo](/javascript/api/excel/excel.cultureinfo)オブジェクトとして定義します。</span><span class="sxs-lookup"><span data-stu-id="36400-162">`Application.cultureInfo` defines the system culture settings as a [CultureInfo](/javascript/api/excel/excel.cultureinfo) object.</span></span> <span data-ttu-id="36400-163">これには、数値の小数点の記号や日付の形式などの設定が含まれます。</span><span class="sxs-lookup"><span data-stu-id="36400-163">This contains settings like the numerical decimal separator or the date format.</span></span>

<span data-ttu-id="36400-164">一部のカルチャ設定は[、EXCEL UI を使用して変更](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e)できます。</span><span class="sxs-lookup"><span data-stu-id="36400-164">Some culture settings can be [changed through the Excel UI](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e).</span></span> <span data-ttu-id="36400-165">システム設定は、 `CultureInfo`オブジェクトに保持されます。</span><span class="sxs-lookup"><span data-stu-id="36400-165">The system settings are preserved in the `CultureInfo` object.</span></span> <span data-ttu-id="36400-166">ローカルの変更は、など、[アプリケーション](/javascript/api/excel/excel.application)レベルのプロパティとし`Application.decimalSeparator`て保持されます。</span><span class="sxs-lookup"><span data-stu-id="36400-166">Any local changes are kept as [Application](/javascript/api/excel/excel.application)-level properties, such as `Application.decimalSeparator`.</span></span>

<span data-ttu-id="36400-167">次の例では、"," から、システム設定で使用される文字への数値文字列の小数点の区切り文字を変更します。</span><span class="sxs-lookup"><span data-stu-id="36400-167">The following sample changes the decimal separator character of a numerical string from a ',' to the character used by the system settings.</span></span>

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

## <a name="add-custom-xml-data-to-the-workbook"></a><span data-ttu-id="36400-168">カスタム XML データをブックに追加する</span><span class="sxs-lookup"><span data-stu-id="36400-168">Add custom XML data to the workbook</span></span>

<span data-ttu-id="36400-169">Excel の Open XML **.xlsx** ファイル形式を使用すると、アドインでカスタム XML データをブックに埋め込むことができます。</span><span class="sxs-lookup"><span data-stu-id="36400-169">Excel's Open XML **.xlsx** file format lets your add-in embed custom XML data in the workbook.</span></span> <span data-ttu-id="36400-170">このデータは、アドインに関係なく、ブックで保持されます。</span><span class="sxs-lookup"><span data-stu-id="36400-170">This data persists with the workbook, independent of the add-in.</span></span>

<span data-ttu-id="36400-171">ブックには [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection) が含まれます。これは [CustomXmlParts](/javascript/api/excel/excel.customxmlpart) のリストです。</span><span class="sxs-lookup"><span data-stu-id="36400-171">A workbook contains a [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), which is a list of [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span></span> <span data-ttu-id="36400-172">これにより、XML 文字列と対応する一意の ID へのアクセスが提供されます。</span><span class="sxs-lookup"><span data-stu-id="36400-172">These give access to the XML strings and a corresponding unique ID.</span></span> <span data-ttu-id="36400-173">これらの ID を設定として保管することにより、アドインはセッション間で XML パーツへのキーを保持できます。</span><span class="sxs-lookup"><span data-stu-id="36400-173">By storing these IDs as settings, your add-in can maintain the keys to its XML parts between sessions.</span></span>

<span data-ttu-id="36400-174">以下のサンプルは、カスタム XML パーツを使用する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="36400-174">The following samples show how to use custom XML parts.</span></span> <span data-ttu-id="36400-175">最初のコード ブロックは、XML データをドキュメントに埋め込む方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="36400-175">The first code block demonstrates how to embed XML data in the document.</span></span> <span data-ttu-id="36400-176">レビュー担当者のリストを保管してから、ブックの設定を使用して XML の `id` を保存して、後から取得できるようにします。</span><span class="sxs-lookup"><span data-stu-id="36400-176">It stores a list of reviewers, then uses the workbook's settings to save the XML's `id` for future retrieval.</span></span> <span data-ttu-id="36400-177">2 番目のブロックでは、後からその XML にアクセスする方法が示されています。</span><span class="sxs-lookup"><span data-stu-id="36400-177">The second block shows how to access that XML later.</span></span> <span data-ttu-id="36400-178">"ContosoReviewXmlPartId" 設定がロードされ、ブックの `customXmlParts` に渡されます。</span><span class="sxs-lookup"><span data-stu-id="36400-178">The "ContosoReviewXmlPartId" setting is loaded and passed to the workbook's `customXmlParts`.</span></span> <span data-ttu-id="36400-179">それから、XML データがコンソールに出力されます。</span><span class="sxs-lookup"><span data-stu-id="36400-179">The XML data is then printed to the console.</span></span>

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
> <span data-ttu-id="36400-180">`CustomXMLPart.namespaceUri` にデータが入れられるのは、トップレベルのカスタム XML 要素に `xmlns` 属性が含まれている場合に限ります。</span><span class="sxs-lookup"><span data-stu-id="36400-180">`CustomXMLPart.namespaceUri` is only populated if the top-level custom XML element contains the `xmlns` attribute.</span></span>

## <a name="control-calculation-behavior"></a><span data-ttu-id="36400-181">計算の動作を制御する</span><span class="sxs-lookup"><span data-stu-id="36400-181">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="36400-182">計算モードを設定する</span><span class="sxs-lookup"><span data-stu-id="36400-182">Set calculation mode</span></span>

<span data-ttu-id="36400-183">既定では、Excel は、参照されているセルが変更されたときに数式の結果を再計算します。</span><span class="sxs-lookup"><span data-stu-id="36400-183">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="36400-184">この計算の動作を調整すると、アドインのパフォーマンス向上に役立つ場合があります。</span><span class="sxs-lookup"><span data-stu-id="36400-184">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="36400-185">Application オブジェクトには、`CalculationMode` 型のプロパティ `calculationMode` があります。</span><span class="sxs-lookup"><span data-stu-id="36400-185">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="36400-186">次のいずれかの値を設定できます。</span><span class="sxs-lookup"><span data-stu-id="36400-186">It can be set to the following values:</span></span>

- <span data-ttu-id="36400-187">`automatic`: 既定の再計算動作。関連するデータが変更されるたびに、Excel は新しい数式の結果を計算します。</span><span class="sxs-lookup"><span data-stu-id="36400-187">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
- <span data-ttu-id="36400-188">`automaticExceptTables`: `automatic` と同様ですが、テーブル内の値に加えた変更は無視されます。</span><span class="sxs-lookup"><span data-stu-id="36400-188">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
- <span data-ttu-id="36400-189">`manual`: ユーザーまたはアドインが要求した場合にのみ計算します。</span><span class="sxs-lookup"><span data-stu-id="36400-189">`manual`: Calculations only occur when the user or add-in requests them.</span></span>

### <a name="set-calculation-type"></a><span data-ttu-id="36400-190">計算タイプを設定する</span><span class="sxs-lookup"><span data-stu-id="36400-190">Set calculation type</span></span>

<span data-ttu-id="36400-191">[Application](/javascript/api/excel/excel.application) オブジェクトは、強制的に即時再計算する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="36400-191">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="36400-192">`Application.calculate(calculationType)` は、指定した `calculationType` に基づいて手動再計算を開始します。</span><span class="sxs-lookup"><span data-stu-id="36400-192">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="36400-193">次の値を指定できます。</span><span class="sxs-lookup"><span data-stu-id="36400-193">The following values can be specified:</span></span>

- <span data-ttu-id="36400-194">`full`: 最後に再計算されてから変更されたかどうかに関係なく、開いているすべてのブックのすべての数式を再計算します。</span><span class="sxs-lookup"><span data-stu-id="36400-194">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="36400-195">`fullRebuild`: 最後に再計算されてから変更されたかどうかに関係なく、依存関係のある数式を確認してから、開いているすべてのブックのすべての数式を再計算します。</span><span class="sxs-lookup"><span data-stu-id="36400-195">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="36400-196">`recalculate`: すべてのアクティブなブックで、最後に計算されてから変更された数式 (またはプログラムで再計算用にマークされている数式)、およびそれに依存する数式を再計算します。</span><span class="sxs-lookup"><span data-stu-id="36400-196">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>

> [!NOTE]
> <span data-ttu-id="36400-197">再計算の詳細については、「[数式の再計算、反復計算、または精度を変更する](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="36400-197">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="36400-198">計算を一時的に中断する</span><span class="sxs-lookup"><span data-stu-id="36400-198">Temporarily suspend calculations</span></span>

<span data-ttu-id="36400-199">Excel API では、アドインから `RequestContext.sync()` を呼び出すまで計算をオフにすることもできます。</span><span class="sxs-lookup"><span data-stu-id="36400-199">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="36400-200">これは、`suspendApiCalculationUntilNextSync()` で実行できます。</span><span class="sxs-lookup"><span data-stu-id="36400-200">This is done with `suspendApiCalculationUntilNextSync()`.</span></span> <span data-ttu-id="36400-201">このメソッドは、アドインから大きな範囲を編集し、複数の編集の間でデータにアクセスする必要がない場合に使用します。</span><span class="sxs-lookup"><span data-stu-id="36400-201">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="save-the-workbook"></a><span data-ttu-id="36400-202">ブックを保存する</span><span class="sxs-lookup"><span data-stu-id="36400-202">Save the workbook</span></span>

<span data-ttu-id="36400-203">`Workbook.save` は、ブックを永続記憶装置に保存します。</span><span class="sxs-lookup"><span data-stu-id="36400-203">`Workbook.save` saves the workbook to persistent storage.</span></span> <span data-ttu-id="36400-204">`save` メソッドにはオプションの `saveBehavior` パラメーターを 1 つ指定できます。値は次のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="36400-204">The `save` method takes a single, optional `saveBehavior` parameter that can be one of the following values:</span></span>

- <span data-ttu-id="36400-205">`Excel.SaveBehavior.save` (既定値): ファイル名や保存場所を指定するようにユーザーに促すダイアログは表示されず、そのままファイルが保存されます。</span><span class="sxs-lookup"><span data-stu-id="36400-205">`Excel.SaveBehavior.save` (default): The file is saved without prompting the user to specify file name and save location.</span></span> <span data-ttu-id="36400-206">ファイルが以前に保存されていない場合は、既定の場所に保存されます。</span><span class="sxs-lookup"><span data-stu-id="36400-206">If the file has not been saved previously, it's saved to the default location.</span></span> <span data-ttu-id="36400-207">ファイルが以前に保存されている場合は、同じ場所に保存されます。</span><span class="sxs-lookup"><span data-stu-id="36400-207">If the file has been saved previously, it's saved to the same location.</span></span>
- <span data-ttu-id="36400-208">`Excel.SaveBehavior.prompt`: ファイルが以前に保存されていない場合は、ファイル名や保存場所を指定するようにユーザーに促すダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="36400-208">`Excel.SaveBehavior.prompt`: If file has not been saved previously, the user will be prompted to specify file name and save location.</span></span> <span data-ttu-id="36400-209">ファイルが以前に保存されている場合、ファイルは同じ場所に保存され、ダイアログは表示されません。</span><span class="sxs-lookup"><span data-stu-id="36400-209">If the file has been saved previously, it will be saved to the same location and the user will not be prompted.</span></span>

> [!CAUTION]
> <span data-ttu-id="36400-210">保存を促すダイアログが表示されたのにユーザーがその操作をキャンセルすると、`save` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="36400-210">If the user is prompted to save and cancels the operation, `save` throws an exception.</span></span>

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook"></a><span data-ttu-id="36400-211">ブックを閉じる</span><span class="sxs-lookup"><span data-stu-id="36400-211">Close the workbook</span></span>

<span data-ttu-id="36400-212">`Workbook.close` は、ブックとそのブックに関連付けられているアドインを終了します (Excel アプリケーションは開いたまま)。</span><span class="sxs-lookup"><span data-stu-id="36400-212">`Workbook.close` closes the workbook, along with add-ins that are associated with the workbook (the Excel application remains open).</span></span> <span data-ttu-id="36400-213">`close` メソッドにはオプションの `closeBehavior` パラメーターを 1 つ指定できます。値は次のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="36400-213">The `close` method takes a single, optional `closeBehavior` parameter that can be one of the following values:</span></span>

- <span data-ttu-id="36400-214">`Excel.CloseBehavior.save` (既定値): ファイルは閉じる前に保存されます。</span><span class="sxs-lookup"><span data-stu-id="36400-214">`Excel.CloseBehavior.save` (default): The file is saved before closing.</span></span> <span data-ttu-id="36400-215">そのファイルが以前に保存されていない場合は、ファイル名や保存場所を指定するようにユーザーに促すダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="36400-215">If the file has not been saved previously, the user will be prompted to specify file name and save location.</span></span>
- <span data-ttu-id="36400-216">`Excel.CloseBehavior.skipSave`: ファイルはそのまま閉じられ、保存されません。</span><span class="sxs-lookup"><span data-stu-id="36400-216">`Excel.CloseBehavior.skipSave`: The file is immediately closed, without saving.</span></span> <span data-ttu-id="36400-217">未保存の変更は失われます。</span><span class="sxs-lookup"><span data-stu-id="36400-217">Any unsaved changes will be lost.</span></span>

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a><span data-ttu-id="36400-218">関連項目</span><span class="sxs-lookup"><span data-stu-id="36400-218">See also</span></span>

- [<span data-ttu-id="36400-219">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="36400-219">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="36400-220">Excel JavaScript API を使用してワークシートを操作する</span><span class="sxs-lookup"><span data-stu-id="36400-220">Work with worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
- [<span data-ttu-id="36400-221">Excel JavaScript API を使用して範囲を操作する</span><span class="sxs-lookup"><span data-stu-id="36400-221">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
