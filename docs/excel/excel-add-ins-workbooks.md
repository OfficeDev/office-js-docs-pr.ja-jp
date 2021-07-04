---
title: Excel JavaScript API を使用してブックを操作する
description: JavaScript API を使用してブックまたはアプリケーション レベルの機能で一般的なタスクを実行するExcel説明します。
ms.date: 06/07/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6a32ad5e50e23868d0c079697411366f3f68af17
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290755"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してブックを操作する

この記事では、Excel JavaScript API を使用して、ブックでタスクを実行する方法のコード サンプルを示しています。 オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Workbook` [、「Workbook オブジェクト (JavaScript API for Excel)」を参照してください](/javascript/api/excel/excel.workbook)。 この記事では、[Application](/javascript/api/excel/excel.application) オブジェクトを使用して実行するブック レベルのアクションについても説明します。

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

ファイルスライスを使用すると、アドインの現在のブックを base64 エンコード文字列 [として取得できます](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)。 次の例に示すように、[FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) クラスを使用して、ファイルを必要な base64 エンコード文字列に変換できます。

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

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one"></a>既存のブックのコピーを現在のブックに挿入する

前の例は、既存のブックから作成された新しいブックを示しています。 既存のブックの一部またはすべてを、アドインに関連付けられているブックにコピーすることもできます。 ブック [には](/javascript/api/excel/excel.workbook) 、ターゲット ブックのワークシートのコピーを自体に `insertWorksheetsFromBase64` 挿入するメソッドがあります。 他のブックのファイルは、呼び出しと同様に、base64 エンコードされた文字列として渡 `Excel.createWorkbook` されます。 

```TypeScript
insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions): OfficeExtension.ClientResult<string[]>;
```

> [!IMPORTANT]
> この `insertWorksheetsFromBase64` メソッドは、Excel Mac、Windows Web 上でサポートされています。 iOS ではサポートされていません。 さらに、このExcel on the webピボットテーブル、グラフ、コメント、またはスライサー要素を持つソース ワークシートはサポートされていません。 これらのオブジェクトが存在する場合、 `insertWorksheetsFromBase64` メソッドはエラーを返 `UnsupportedFeature` Excel on the web。 

次のコード サンプルは、別のブックから現在のブックにワークシートを挿入する方法を示しています。 このコード サンプルでは、まずオブジェクトを使用してブック ファイルを処理し、base64 エンコードされた文字列を抽出し、次にこの base64 エンコードされた文字列を現在のブック [`FileReader`](https://developer.mozilla.org/docs/Web/API/FileReader) に挿入します。 新しいワークシートは、Sheet1 という名前のワークシートの後 **に挿入されます**。 `[]` [InsertWorksheetOptions.sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetNamesToInsert)プロパティのパラメーターとして渡されます。 つまり、ターゲット ブックのすべてのワークシートが現在のブックに挿入されます。

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

保護は、不必要なデータ編集をできないようにするため、ワークシート レベルで設定することもできます。 詳細については、「[Excel JavaScript API を使用してワークシートを操作する](excel-add-ins-worksheets.md#data-protection)」の **データの保護** のセクションを参照してください。

> [!NOTE]
> Excel のブックの保護の詳細については、「[ブックを保護する](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517)」を参照してください。

## <a name="access-document-properties"></a>ドキュメント プロパティへのアクセス

Workbook オブジェクトは、[ドキュメント プロパティ](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75)と呼ばれる Office ファイルのメタデータにアクセスできます。 Workbook オブジェクトの `properties` プロパティは、これらのメタデータ値を含む [DocumentProperties](/javascript/api/excel/excel.documentproperties) オブジェクトです。 次の例は、プロパティを設定する方法を示 `author` しています。

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

#### <a name="worksheet-level-custom-properties"></a>ワークシート レベルのカスタム プロパティ

カスタム プロパティは、ワークシート レベルで設定することもできます。 これらはドキュメント レベルのカスタム プロパティに似ていますが、異なるワークシートで同じキーを繰り返す場合があります。 次の例は、現在のワークシートに値 **"Alpha"** を指定して WorksheetGroup という名前のカスタム プロパティを作成し、それを取得する方法を示しています。

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

## <a name="access-application-culture-settings"></a>アプリケーション カルチャの設定にアクセスする

ブックには、特定のデータの表示方法に影響する言語とカルチャの設定があります。 これらの設定は、アドインのユーザーが異なる言語やカルチャ間でブックを共有している場合にデータをローカライズするのに役立ちます。 アドインでは、文字列解析を使用して、システム カルチャ設定に基づいて数値、日付、時刻の形式をローカライズし、各ユーザーが独自のカルチャの形式でデータを表示できます。

`Application.cultureInfo` システム カルチャ設定を [CultureInfo オブジェクトとして定義](/javascript/api/excel/excel.cultureinfo) します。 これには、数値の小数点記号や日付形式のような設定が含まれる。

一部のカルチャ設定は[、UI を使用Excelできます](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e)。 システム設定はオブジェクトに保持 `CultureInfo` されます。 ローカルの変更は、アプリケーション レベル [の](/javascript/api/excel/excel.application)プロパティ (など) として保持されます `Application.decimalSeparator` 。

次のサンプルでは、数値文字列の小数点記号を ',' からシステム設定で使用される文字に変更します。

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

## <a name="detect-workbook-activation"></a>ブックのアクティブ化を検出する

アドインは、ブックがアクティブ化された場合に検出できます。 ユーザーが別 *のブック、* 別のアプリケーション、または (Excel on the web) Web ブラウザーの別のタブにフォーカスを切り替え、ブックが非アクティブになります。 ブックは、 *ユーザーが* ブックにフォーカスを返すときにアクティブ化されます。 ブックのアクティブ化によって、ブック データの更新など、アドイン内のコールバック関数をトリガーできます。

ブックがアクティブ化された場合を検出[](excel-add-ins-events.md#register-an-event-handler)するには、ブックの[onActivated イベントのイベント ハンドラー](/javascript/api/excel/excel.workbook#onActivated)を登録します。 イベントのイベント ハンドラーは `onActivated` 、イベントが発生すると [WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs) オブジェクトを受け取る。

> [!IMPORTANT]
> ブック `onActivated` を開いた場合、イベントは検出されません。 このイベントは、ユーザーがフォーカスを既に開いているブックに戻す場合にのみ検出されます。

次のコード サンプルは、イベント ハンドラーを登録し、 `onActivated` コールバック関数を設定する方法を示しています。

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

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [Excel JavaScript API を使用してワークシートを操作する](excel-add-ins-worksheets.md)
