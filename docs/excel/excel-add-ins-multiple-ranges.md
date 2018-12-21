---
title: Excel アドインで複数の範囲を同時に操作する
description: ''
ms.date: 09/04/2018
ms.openlocfilehash: f1217fc76d14269882a73ec5eb7758e519563456
ms.sourcegitcommit: 6870f0d96ed3da2da5a08652006c077a72d811b6
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/21/2018
ms.locfileid: "27383226"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a>Excel アドインで複数の範囲を同時に操作する (プレビュー)

Excel JavaScript ライブラリを使用すると、同時に複数の範囲に対してアドインによる操作の実行とプロパティの設定が可能になります。 範囲は連続している必要はありません。 コードがよりシンプルになることに加え、この方法でプロパティを設定すれば、各範囲に同じプロパティを個別に設定する方法よりも処理速度が格段に速くなります。

> [!NOTE]
> この記事で説明する API には、**Office 2016 クイック実行バージョン 1809 Build 10820.20000** 以降が必要です  ([Office Insider プログラム](https://products.office.com/office-insider)に参加して、適切なビルドを取得することが必要な場合があります)。また、Office JavaScript ライブラリのベータ版を [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) からロードする必要があります。 最後に、これらの API セットに関する参照ページはまだありません。 ただし、定義の種類ファイル [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) に説明が含まれています。

## <a name="rangeareas"></a>RangeAreas

範囲のセット (連続している必要はなし) は、`Excel.RangeAreas` オブジェクトで表されます。 `Range` 型と同様のプロパティとメソッドを持ちますが (多くの場合は同じまたは類似した名前)、以下に対しては調整が行われています。

- プロパティのデータ型と、セッターとゲッターの動作。
- メソッド パラメーターのデータ型と、メソッドの動作。
- メソッドの戻り値のデータ型。

次にいくつか例を示します。

- `RangeAreas` には `address` プロパティがあり、`Range.address` プロパティのように 1 つのアドレスを返すのではなく、複数の範囲のアドレスをコンマで区切った文字列を返します。
- `RangeAreas` には、一貫性がある場合、`RangeAreas` に指定された全範囲のデータ検証を表す `DataValidation` オブジェクトを返す `dataValidation` プロパティがあります。 `RangeAreas` に指定された全範囲に同じ `DataValidation` オブジェクトが適用されていない場合、このプロパティは `null` となります。 これは、`RangeAreas` オブジェクトに関する、汎用的ではありませんが一般的な原則です: *`RangeAreas` に指定された全範囲のプロパティの値に一貫性がない場合、`null` となります*。 より詳しい情報といくつかの例外については、「[RangeAreas のプロパティの読み取り](#read-properties-of-rangeareas)」を参照してください。
- `RangeAreas.cellCount` は、`RangeAreas` に指定された全範囲の合計セル数を取得します。
- `RangeAreas.calculate` は、`RangeAreas` に指定された全範囲のセルを再計算します。
- `RangeAreas.getEntireColumn` と `RangeAreas.getEntireRow` は、`RangeAreas` に指定された全範囲のセルの列 (または行) すべてを表す、別の `RangeAreas` オブジェクトを返します。 たとえば、`RangeAreas` が "A1:C4" と "F14:L15" を表す場合、`RangeAreas.getEntireColumn` は "A:C" と "F:L" を表す `RangeAreas` オブジェクトを返します。
- `RangeAreas.copyFrom` は、コピー操作のコピー元範囲を表す `Range` または `RangeAreas` パラメーターのいずれかを取得できます。

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a>RangeAreas でも利用可能な Range メンバーの全リスト

##### <a name="properties"></a>プロパティ

リストにあるプロパティを読み取るコードを書く前に、「[RangeAreas のプロパティの読み取り](#read-properties-of-rangeareas)」の内容を理解しておいてください。 繰り返される内容について細かい注意点があります。

- address
- addressLocal
- cellCount
- conditionalFormats
- context
- dataValidation
- format
- isEntireColumn
- isEntireRow
- style
- worksheet

##### <a name="methods"></a>メソッド

プレビュー段階の Range メソッドについてはマークが付いています。

- calculate()
- clear()
- convertDataTypeToText() (プレビュー)
- convertToLinkedDataType() (プレビュー)
- copyFrom() (プレビュー)
- getEntireColumn()
- getEntireRow()
- getIntersection()
- getIntersectionOrNullObject()
- getOffsetRange() (RangeAreas オブジェクトでの名前は getOffsetRangeAreas)
- getSpecialCells() (プレビュー)
- getSpecialCellsOrNullObject() (プレビュー)
- getTables() (プレビュー)
- getUsedRange() (RangeAreas オブジェクトでの名前は getUsedRangeAreas)
- getUsedRangeOrNullObject() (RangeAreas オブジェクトでの名前は getUsedRangeAreasOrNullObject)
- load()
- set()
- setDirty() (プレビュー)
- toJSON()
- track()
- untrack()

### <a name="rangearea-specific-properties-and-methods"></a>RangeArea 固有のプロパティとメソッド

`RangeAreas` 型には、`Range` オブジェクトには存在しないプロパティとメソッドがいくつかあります。 次にいくつか選択したものを示します。

- `areas`: `RangeAreas` オブジェクトが表す全範囲を含む `RangeCollection` オブジェクト。 `RangeCollection` オブジェクトも新しいオブジェクトであり、他の Excel コレクション オブジェクトと類似しています。 これには、範囲を表す `Range` オブジェクトの配列である `items` プロパティがあります。
- `areaCount`: `RangeAreas` で指定された範囲の合計数。
- `getOffsetRangeAreas`: [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-) と同じように動作します。ただし、`RangeAreas` を返し、元の `RangeAreas` で指定された範囲の 1 つからの各オフセットである範囲を含みます。

## <a name="create-rangeareas-and-set-properties"></a>RangeAreas の作成とプロパティの設定

`RangeAreas` オブジェクトの作成には、2 つの基本的な方法があります。

- `Worksheet.getRanges()` を呼び出して、範囲のアドレスがコンマで区切られた文字列を渡します。 含める対象の範囲が既に [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem) に指定されている場合、文字列にはアドレスではなくその名前を指定することができます。
- `Workbook.getSelectedRanges()` を呼び出します。 このメソッドは、現在アクティブなワークシート上で選択されている全範囲を表す `RangeAreas` を返します。

一度 `RangeAreas` オブジェクトを作成すると、`getOffsetRangeAreas` や `getIntersection` など、`RangeAreas` を返すオブジェクト上のメソッドを使用して別のオブジェクトを作成できます。

> [!NOTE]
> `RangeAreas` オブジェクトに新たな範囲を直接追加することはできません。 たとえば、`RangeAreas.areas` 内のコレクションには `add` メソッドが存在しません。

> [!WARNING]
> `RangeAreas.areas.items` 配列のメンバーの追加または削除を直接試行してはいけません。 これにより、後でコード内で望ましくない動作が発生します。 たとえば、追加の `Range` オブジェクトを配列にプッシュすることは可能ですが、エラーが発生します。`RangeAreas` のプロパティとメソッドは、その新しいアイテムがその場所に存在していないかのように動作するためです。 たとえば、`areaCount` プロパティにはこの方法でプッシュされた範囲は含まれません。また、`RangeAreas.getItemAt(index)` は、`index` が `areasCount-1`より大きい場合、エラーをスローします。 同様に、`RangeAreas.areas.items` 配列内の `Range` オブジェクトを、参照を取得してその `Range.delete` メソッドを呼び出すという方法で削除すると、バグとなります。`Range` オブジェクトは*削除されます*が、親 `RangeAreas` オブジェクトのプロパティとメソッドは、そのオブジェクトがまだ存在するものとして動作するためです。 たとえば、コードで `RangeAreas.calculate` を呼び出すと、Office は範囲を計算しようとしますが、範囲オブジェクトが既に存在しないためにエラーとなります。

`RangeAreas` に対してプロパティを設定すると、`RangeAreas.areas` コレクション内の全範囲の対応するプロパティが設定されます。

次に、複数の範囲にプロパティを設定する例を示します。 この関数は、**F3:F5** と **H3:H5** の範囲を強調表示します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

この例は、`getRanges` に渡す範囲のアドレスをハード コーディングできる場合や実行時に簡単に計算できる場合に適用されます。 たとえば、これが適切なのは次のような場合です。 

- コードが、既知のテンプレートのコンテキスト内で実行される。
- コードが、データのスキーマが既知であるインポート済みデータのコンテキスト内で実行される。

コーディング時に操作対象の範囲がわからない場合は、実行時に特定する必要があります。 次のセクションでは、そのような場合について説明します。

### <a name="discover-range-areas-programmatically"></a>プログラムを使用して範囲を特定する

`Range.getSpecialCells()` と `Range.getSpecialCellsOrNullObject()` メソッドを使用すると、セルの特性とセル値の種類を基に、操作対象のセルを実行時に特定することができます。 次に示すのは、TypeScript データ型ファイルの、このメソッドのシグネチャです。

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

このうち最初のものを使用する例を次に示します。 このコードの注意点は次のとおりです。

- 検索が必要なシートの部分を制限するために、まず `Worksheet.getUsedRange` を呼び出し、その範囲に関してのみ `getSpecialCells` を呼び出します。
- `Excel.SpecialCellType` 列挙からの値の文字列バージョンをパラメーターとして `getSpecialCells` に渡します。 代わりに渡すことができる他の値には、空のセルの場合は "Blanks"、数式ではなくリテラル値を含むセルの場合は "Constants"、`usedRange` 内の最初のセルと同じ条件付き書式を持つセルの場合は "SameConditionalFormat" などがあります。 最初のセルとは、左上隅のセルです。 列挙内の値の完全なリストについては、[beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) を参照してください。
- `getSpecialCells` メソッドは `RangeAreas` オブジェクトを返すため、数式を含むセルはすべて、連続していないセルであっても、ピンク色になります。 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

範囲内に対象の特性を持つセルが*まったくない*場合もあります。 `getSpecialCells` で対象のセルが見つからないと、**ItemNotFound** エラーがスローされます。 この場合、制御のフローが `catch` ブロック/メソッドに移ります (存在する場合)。 存在しない場合は、このエラーにより関数が停止します。 対象の特性を持つセルがない場合はエラーをスローするという動作が求められるシナリオもあるかもしれません。 

ただし、一般的ではありませんが、一致するセルがないということが通常であるようなシナリオでは、コードでこのような可能性があるかどうかを確認し、あれば、エラーをスローせずに適切に処理するようにしておく必要があります。 このようなシナリオの場合、`getSpecialCellsOrNullObject` メソッドを使用し、`RangeAreas.isNullObject` プロパティをテストします。 次に例を示します。 このコードの注意点は次のとおりです。

- `getSpecialCellsOrNullObject` メソッドは常にプロキシ オブジェクトを返します。そのため、通常の JavaScript 使用環境では `null` となることはありません。 ただし一致するセルが見つからなかった場合、オブジェクトの `isNullObject` プロパティは `true` に設定されます。
- `isNullObject` プロパティをテストする*前*に、`context.sync` を呼び出します。 これは、すべての `*OrNullObject` メソッドとプロパティの必要条件です。プロパティを読み取るためには常に、そのプロパティをロードして同期する必要があるためです。 ただし、*明示的*に `isNullObject` プロパティをロードする必要はありません。 `load` がオブジェクトに対して呼び出されていない場合であっても、プロパティは `context.sync` によって自動的にロードされます。 詳細については、「[\*OrNullObject メソッド](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods)」を参照してください。
- このコードをテストするには、最初に数式を含まないセルの範囲を選択してからコードを実行します。 次に、少なくとも 1 つのセルが数式を含む範囲を選択してからコードを再実行します。

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    var formulaRanges = range.getSpecialCellsOrNullObject("Formulas");
    return context.sync()
        .then(function() {
            if (formulaRanges.isNullObject) {
                console.log("No cells have formulas");
            }
            else {
                formulaRanges.format.fill.color = "pink";
            }
        })
        .then(context.sync);
})
```

わかりやすくするため、この記事内のすべての他の例では、`getSpecialCells` メソッドを `getSpecialCellsOrNullObject` の代わりに使用しています。

#### <a name="narrow-the-target-cells-with-cell-value-types"></a>セルの値の型に応じて対象のセルを絞り込む

オプションの 2 つめのパラメーター (列挙型 `Excel.SpecialCellValueType`) を使用すると、対象のセルをさらに絞り込むことができます。 このパラメーターは、"Formulas" または "Constants" を`getSpecialCells` または `getSpecialCellsOrNullObject` に渡す場合にのみ使用できます。 このパラメーターにより、特定の型の値を持つセルのみ対象として指定することができます。 4 つの基本的な型があります: "Error"、"Logical" (ブール値を意味します)、"Numbers"、"Text" です  (列挙の場合はこの 4 つ以外の値もあります。詳細は後述します)。次に例を示します。 このコードの注意点は次のとおりです。

- リテラル数値を持つセルのみ強調表示されます。 数式 (結果が数字の場合であっても)、ブール値、テキストを持つセル、およびエラー状態にあるセルは強調表示されません。
- コードをテストするには、リテラル数値を持ついくつかのセル、他の型のリテラル値を持ついくつかのセル、そして数式を持ついくつかのセルをそれぞれワークシートに含めるようにしてください。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

テキスト値のセルすべてとブール値 ("Logical") のセルすべてなど、セル値の型を複数操作する必要がある場合もあります。 `Excel.SpecialCellValueType` 列挙に含まれる値を使用すると、型を組み合わせることができます。 たとえば、"LogicalText" を使用すると、すべてのブール値のセルとテキスト値のセルを対象とすることができます。 4 つの基本的な型のうち、任意の 2 つまたは 3 つの型を組み合わせることができます。 基本的な型を組み合わせるこれらの列挙値の名前は、常にアルファベット順で指定します。 したがって、エラー値、テキスト値、ブール値のセルを組み合わせる場合は "ErrorLogicalText" を使用します。"LogicalErrorText" や "TextErrorLogical" とはしてはいけません。 既定のパラメーターである "All" は、4 つの型すべてを組み合わせます。 次の例では、結果が数値またはブール値となる数式を含むすべてのセルが強調表示されます。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaLogicalNumberRanges = usedRange.getSpecialCells("Formulas", "LogicalNumbers");
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

> [!NOTE]
> `Excel.SpecialCellValueType` パラメーターは、`Excel.SpecialCellType` パラメーターが "Formulas" または "Constants" の場合にのみ使用できます。

### <a name="get-rangeareas-within-rangeareas"></a>RangeAreas 内の RangeAreas を取得する

`RangeAreas` 型自体には、同じ 2 つのパラメーターを使用する `getSpecialCells` および `getSpecialCellsOrNullObject` メソッドもあります。 これらのメソッドは、`RangeAreas.areas` コレクション内の全範囲から対象のセルをすべて返します。 `Range`オブジェクトではなく `RangeAreas` に対して呼び出された場合のメソッドの動作には、少し異なる点が 1 つあります。最初のパラメーターとして "SameConditionalFormat" を渡した場合、*`RangeAreas.areas` コレクション内の最初の範囲*の左上隅のセルと同じ条件付き書式を持つセルがすべて返されます。 同じ点が "SameDataValidation" にも適用されます。`Range.getSpecialCells` にこれを渡すと、*範囲内*の左上隅のセルと同じデータ検証ルールを持つセルがすべて返されます。 一方、`RangeAreas.getSpecialCells` に渡した場合は、*`RangeAreas.areas` コレクション内の最初の範囲*の左上隅のセルと同じデータ検証ルールを持つセルがすべて返されます。

## <a name="read-properties-of-rangeareas"></a>RangeAreas のプロパティの読み取り

`RangeAreas` のプロパティ値の読み取りには、注意が必要です。`RangeAreas`内の範囲それぞれで、プロパティの値が異なる可能性があるためです。 一貫性のある値を返すことが*できる*場合には返す、というのが一般的なルールです。 たとえば、次のコードでは、ピンクの RGB コード (`#FFC0CB`) と `true` がコンソールに記録されます。`RangeAreas`オブジェクト内の範囲のどちらも、塗りつぶし色がピンクであり、列全体であるためです。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    var rangeAreas = sheet.getRanges("F:F, H:H");  
    rangeAreas.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // #FFC0CB
            console.log(rangeAreas.isEntireColumn); // true
        })
        .then(context.sync);
})
```

一貫性を期待できない場合、事態は複雑となります。 `RangeAreas` プロパティの動作は、次の 3 つの原則に従います。

- `RangeAreas` オブジェクトのブール値プロパティは、すべてのメンバー範囲でプロパティが true でない限り、`false` を返します。
- ブール値以外のプロパティ (`address` プロパティを除く) は、すべてのメンバー範囲で対応するプロパティが同じ値ではない限り、`null` を返します。
- `address` プロパティは、メンバー範囲のアドレスをコンマで区切った文字列を返します。

たとえば、次のコードでは、1 つの範囲のみが列全体であり、1 つの範囲のみがピンクで塗りつぶされている `RangeAreas` を作成します。 コンソールには、塗りつぶし色の場合は `null`、`isEntireRow` プロパティの場合は `false`、`address` プロパティの場合は "Sheet1!F3:F5, Sheet1!H:H" ("Sheet1" はシート名) が表示されます。 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H:H");

    var pinkColumnRange = sheet.getRange("H:H");
    pinkColumnRange.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn, address");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // null
            console.log(rangeAreas.isEntireColumn); // false
            console.log(rangeAreas.address); // "Sheet1!F3:F5, Sheet1!H:H"
        })
        .then(context.sync);
})
```

## <a name="see-also"></a>関連項目

- [Excel の JavaScript API の概要](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [Excel.​Range class](https://docs.microsoft.com/javascript/api/excel/excel.range)
- [RangeAreas オブジェクト (Excel の JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (API がプレビュー段階の間は、このリンクは機能しない場合があります。 その場合は、代わりに [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) を参照してください)。