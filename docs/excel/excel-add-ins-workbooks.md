---
title: Excel JavaScript API を使用してブックを操作する
description: Excel JavaScript API を使用して、ブックまたはアプリケーションレベルの機能を使用して一般的なタスクを実行する方法を示すコードサンプルです。
ms.date: 08/24/2020
localization_priority: Normal
ms.openlocfilehash: a7a35e2627863c648f8c3e31ab05b2714ca0aebe
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294130"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してブックを操作する

この記事では、Excel JavaScript API を使用して、ブックでタスクを実行する方法のコード サンプルを示しています。 オブジェクトがサポートするプロパティとメソッドの完全な一覧につい `Workbook` ては、「 [Workbook オブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.workbook)」を参照してください。 この記事では、[Application](/javascript/api/excel/excel.application) オブジェクトを使用して実行するブック レベルのアクションについても説明します。

Workbook オブジェクトは、Excel を操作するアドインのエントリ ポイントです。 このオブジェクトは、Excel データのアクセスや変更に使用するワークシート、テーブル、ピボットテーブル、その他のコレクションを保持します。 [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) オブジェクトは、個々のワークシートを使用して、ブックのすべてのデータにアドインからアクセスできるようにします。 具体的には、アドインからワークシートの追加、ワークシート間の移動、ワークシート イベントへのハンドラーの割り当てができます。 ワークシートへのアクセスと編集の方法については、「[Excel JavaScript API を使用してワークシートを操作する](excel-add-ins-worksheets.md)」を参照してください。

## <a name="get-the-active-cell-or-selected-range"></a>アクティブ セルまたは選択した範囲を取得する

Workbook オブジェクトには、ユーザーまたはアドインが選択したセルの範囲を取得する 2 つのメソッド `getActiveCell()` と `getSelectedRange()` があります。 `getActiveCell()` はブックからアクティブ セルを [Range オブジェクト](/javascript/api/excel/excel.range)として取得します。 次の例では、`getActiveCell()` を呼び出し、コンソールにセルのアドレスを表示します。

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

`getSelectedRange()` メソッドは現在選択されている単一の範囲を返します。 複数の範囲が選択されている場合は、InvalidSelection エラーがスローされます。 次の例では、`getSelectedRange()` を呼び出し、範囲の塗りつぶし色を黄色に設定します。

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a>ブックを作成する

アドインでは、アドインが現在実行されている Excel のインスタンスとは異なる新しいブックを作成できます。 Excel オブジェクトには、この目的の `createWorkbook` メソッドがあります。 このメソッドが呼び出されると、新しいブックが Excel の新しいインスタンスですぐに開いて表示されます。 アドインは前のブックで開いて実行されたままになります。

```js
Excel.createWorkbook();
```

`createWorkbook` メソッドは既存のブックのコピーの作成もできます。 このメソッドは、オプションのパラメーターとして .xlsx ファイルの base64 エンコード文字列表現を受け取ります。 文字列の引数は有効な .xlsx ファイルと見なされ、作成されるブックはそのファイルのコピーになります。

[ファイルスライシング](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)を使用して、アドインの現在のブックを base64 でエンコードされた文字列として取得できます。 次の例に示すように、[FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) クラスを使用して、ファイルを必要な base64 エンコード文字列に変換できます。

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

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one-preview"></a>既存のブックのコピーを現在のブックに挿入する (プレビュー)

> [!NOTE]
> `WorksheetCollection.addFromBase64` メソッドは現在、パブリック プレビューでのみ使用でき、Windows および Mac 上の Office でのみ使用できます。 [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

前の例は、既存のブックから作成された新しいブックを示しています。 既存のブックの一部またはすべてを、アドインに関連付けられているブックにコピーすることもできます。 ブックの [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) にある `addFromBase64` メソッドは、対象のブックのワークシートのコピーを現在のブックに挿入します。 他のブックのファイルは、`Excel.createWorkbook` 呼び出しの場合と同様に、base64 エンコード文字列として渡されます。

```TypeScript
addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet | string): OfficeExtension.ClientResult<string[]>;
```

次の例では、ブックのワークシートが現在のブックのアクティブ ワークシートの直後に挿入されています。 `null` が `sheetNamesToInsert?: string[]` パラメーターに渡されている点に注意してください。 つまり、すべてのワークシートが挿入されます。

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

## <a name="protect-the-workbooks-structure"></a>ブックのシート構成を保護する

アドインでは、ブックのシート構成を編集するユーザーの機能を制御できます。 Workbook オブジェクトの `protection` プロパティは [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) オブジェクトであり、`protect()` メソッドを備えています。 次の例では、ブックのシート構成の保護を切り替える基本的なシナリオを示します。

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

`protect` メソッドは、オプションの文字列パラメーターを受け取ります。 この文字列は、ユーザーが保護をバイパスしてブックのシート構成を変更するために必要なパスワードを表します。

保護は、不必要なデータ編集をできないようにするため、ワークシート レベルで設定することもできます。 詳細については、「[Excel JavaScript API を使用してワークシートを操作する](excel-add-ins-worksheets.md#data-protection)」の**データの保護**のセクションを参照してください。

> [!NOTE]
> Excel のブックの保護の詳細については、「[ブックを保護する](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517)」を参照してください。

## <a name="access-document-properties"></a>ドキュメント プロパティへのアクセス

Workbook オブジェクトは、[ドキュメント プロパティ](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75)と呼ばれる Office ファイルのメタデータにアクセスできます。 Workbook オブジェクトの `properties` プロパティは、これらのメタデータ値を含む [DocumentProperties](/javascript/api/excel/excel.documentproperties) オブジェクトです。 次の例は、プロパティを設定する方法を示して `author` います。

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="custom-properties"></a>カスタム プロパティ

カスタム プロパティを定義することもできます。 DocumentProperties オブジェクトには `custom` プロパティが含まれていて、ユーザー定義プロパティのキー値のペアのコレクションを表します。 次の例では、"Hello" という値を持つ **Introduction** という名前のカスタム プロパティを作成し、それを取得する方法を示します。

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

#### <a name="worksheet-level-custom-properties-preview"></a>ワークシートレベルのカスタムプロパティ (プレビュー)

> [!NOTE]
> ワークシートレベルのカスタムプロパティは現在プレビュー段階です。 [!INCLUDE [Information about using preview Excel APIs](../includes/using-excel-preview-apis.md)]

カスタムプロパティは、ワークシートレベルで設定することもできます。 これらはドキュメントレベルのカスタムプロパティに似ていますが、異なるワークシート間で同じキーを繰り返すことができる点が異なります。 次の例は、現在のワークシートで "α" という値を持つ、"worksheet **group** " という名前のカスタムプロパティを作成し、それを取得する方法を示しています。

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

## <a name="access-document-settings"></a>ドキュメント設定へのアクセス

ブックの設定は、カスタム プロパティのコレクションに似ています。 設定は 1 つの Excel ファイルとアドインのペアリングに固有であるのに対して、プロパティはファイルに接続しているだけである点が異なります。 次の例は、設定を作成してアクセスする方法を示しています。

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

## <a name="access-application-culture-settings"></a>Access アプリケーションのカルチャ設定

ブックには、特定のデータの表示方法に影響する言語とカルチャの設定が含まれています。 これらの設定は、アドインのユーザーが異なる言語とカルチャでブックを共有している場合に、データをローカライズするのに役立ちます。 アドインでは、文字列の解析を使用して、各ユーザーが独自のカルチャの形式でデータを表示できるように、システムのカルチャ設定に基づいて数値、日付、時刻の形式をローカライズできます。

`Application.cultureInfo` システムのカルチャ設定を [CultureInfo](/javascript/api/excel/excel.cultureinfo) オブジェクトとして定義します。 これには、数値の小数点の記号や日付の形式などの設定が含まれます。

一部のカルチャ設定は [、EXCEL UI を使用して変更](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e)できます。 システム設定は、オブジェクトに保持され `CultureInfo` ます。 ローカルの変更は、など、 [アプリケーション](/javascript/api/excel/excel.application)レベルのプロパティとして保持され `Application.decimalSeparator` ます。

次の例では、"," から、システム設定で使用される文字への数値文字列の小数点の区切り文字を変更します。

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

## <a name="add-custom-xml-data-to-the-workbook"></a>カスタム XML データをブックに追加する

Excel の Open XML **.xlsx** ファイル形式を使用すると、アドインでカスタム XML データをブックに埋め込むことができます。 このデータは、アドインに関係なく、ブックで保持されます。

ブックには [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection) が含まれます。これは [CustomXmlParts](/javascript/api/excel/excel.customxmlpart) のリストです。 これにより、XML 文字列と対応する一意の ID へのアクセスが提供されます。 これらの ID を設定として保管することにより、アドインはセッション間で XML パーツへのキーを保持できます。

以下のサンプルは、カスタム XML パーツを使用する方法を示しています。 最初のコード ブロックは、XML データをドキュメントに埋め込む方法を示しています。 レビュー担当者のリストを保管してから、ブックの設定を使用して XML の `id` を保存して、後から取得できるようにします。 2 番目のブロックでは、後からその XML にアクセスする方法が示されています。 "ContosoReviewXmlPartId" 設定がロードされ、ブックの `customXmlParts` に渡されます。 それから、XML データがコンソールに出力されます。

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
> `CustomXMLPart.namespaceUri` にデータが入れられるのは、トップレベルのカスタム XML 要素に `xmlns` 属性が含まれている場合に限ります。

## <a name="control-calculation-behavior"></a>計算の動作を制御する

### <a name="set-calculation-mode"></a>計算モードを設定する

既定では、Excel は、参照されているセルが変更されたときに数式の結果を再計算します。 この計算の動作を調整すると、アドインのパフォーマンス向上に役立つ場合があります。 Application オブジェクトには、`CalculationMode` 型のプロパティ `calculationMode` があります。 次のいずれかの値を設定できます。

- `automatic`: 既定の再計算動作。関連するデータが変更されるたびに、Excel は新しい数式の結果を計算します。
- `automaticExceptTables`: `automatic` と同様ですが、テーブル内の値に加えた変更は無視されます。
- `manual`: ユーザーまたはアドインが要求した場合にのみ計算します。

### <a name="set-calculation-type"></a>計算タイプを設定する

[Application](/javascript/api/excel/excel.application) オブジェクトは、強制的に即時再計算する方法を提供します。 `Application.calculate(calculationType)` は、指定した `calculationType` に基づいて手動再計算を開始します。 次の値を指定できます。

- `full`: 最後に再計算されてから変更されたかどうかに関係なく、開いているすべてのブックのすべての数式を再計算します。
- `fullRebuild`: 最後に再計算されてから変更されたかどうかに関係なく、依存関係のある数式を確認してから、開いているすべてのブックのすべての数式を再計算します。
- `recalculate`: すべてのアクティブなブックで、最後に計算されてから変更された数式 (またはプログラムで再計算用にマークされている数式)、およびそれに依存する数式を再計算します。

> [!NOTE]
> 再計算の詳細については、「[数式の再計算、反復計算、または精度を変更する](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4)」を参照してください。

### <a name="temporarily-suspend-calculations"></a>計算を一時的に中断する

Excel API では、アドインから `RequestContext.sync()` を呼び出すまで計算をオフにすることもできます。 これは、`suspendApiCalculationUntilNextSync()` で実行できます。 このメソッドは、アドインから大きな範囲を編集し、複数の編集の間でデータにアクセスする必要がない場合に使用します。

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="save-the-workbook"></a>ブックを保存する

`Workbook.save` は、ブックを永続記憶装置に保存します。 `save` メソッドにはオプションの `saveBehavior` パラメーターを 1 つ指定できます。値は次のいずれかになります。

- `Excel.SaveBehavior.save` (既定値): ファイル名や保存場所を指定するようにユーザーに促すダイアログは表示されず、そのままファイルが保存されます。 ファイルが以前に保存されていない場合は、既定の場所に保存されます。 ファイルが以前に保存されている場合は、同じ場所に保存されます。
- `Excel.SaveBehavior.prompt`: ファイルが以前に保存されていない場合は、ファイル名や保存場所を指定するようにユーザーに促すダイアログが表示されます。 ファイルが以前に保存されている場合、ファイルは同じ場所に保存され、ダイアログは表示されません。

> [!CAUTION]
> 保存を促すダイアログが表示されたのにユーザーがその操作をキャンセルすると、`save` は例外をスローします。

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook"></a>ブックを閉じる

`Workbook.close` は、ブックとそのブックに関連付けられているアドインを終了します (Excel アプリケーションは開いたまま)。 `close` メソッドにはオプションの `closeBehavior` パラメーターを 1 つ指定できます。値は次のいずれかになります。

- `Excel.CloseBehavior.save` (既定値): ファイルは閉じる前に保存されます。 そのファイルが以前に保存されていない場合は、ファイル名や保存場所を指定するようにユーザーに促すダイアログが表示されます。
- `Excel.CloseBehavior.skipSave`: ファイルはそのまま閉じられ、保存されません。 未保存の変更は失われます。

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a>関連項目

- [Excel JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)
- [Excel JavaScript API を使用してワークシートを操作する](excel-add-ins-worksheets.md)
- [Excel JavaScript API を使用して範囲を操作する](excel-add-ins-ranges.md)
