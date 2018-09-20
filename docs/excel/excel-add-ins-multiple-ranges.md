---
title: Excel のアドインで同時に複数の範囲の操作をします。
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: bcb14d1f4c015fe675c2d65cb5f1198d485dd4c5
ms.sourcegitcommit: 3da2038e827dc3f274d63a01dc1f34c98b04557e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/19/2018
ms.locfileid: "24016459"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a>Excel のアドイン (プレビュー) で同時に複数のセル範囲を操作します。

Excel の JavaScript ライブラリは、アドインに操作を実行し、同時に複数の範囲のプロパティを設定するようにします。 範囲が隣接している必要はありません。 コードを簡単にするには、このプロパティの方法が、範囲ごとに個別に同じプロパティを設定するよりもはるかに高速に実行されます。

> [!NOTE]
> この資料に記載されている APIは、 **Office 2016 クイック実行バージョン 1809 10820.20000 の構築** 以降を必要とします。 (適切なビルドを取得するため、 [Office 内部プログラム](https://products.office.com/office-insider) に参加する必要があります。)　また、 [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)からベータ版の Office の JavaScript ライブラリを読み込む必要があります。 最後に、これらの API は、参照ページはまだ必要はありません。 次の定義型ファイルは、それらについての説明です: [ベータ版 office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)です。

## <a name="rangeareas"></a>RangeAreas

(場合によっては連続していない) 範囲のセットは、 `Excel.RangeAreas` オブジェクトで表示されます。 型(多くの場合、同じまたは同様の名前を持つ)に似たプロパティとメソッドを持っていますが、以下の補正が加えられました。`Range`

- プロパティのデータ型とセッターとゲッターの動作です。
- メソッドパラメーターのデータ型と、メソッドの動作です。
- メソッドのデータ型は、値を返します。

例:

- `RangeAreas` プロパティとともに1 つのアドレスとしてではなく、範囲のアドレスのコンマ区切りの文字列を返す`address` プロパティを持ちます。`Range.address`
- `RangeAreas` 一貫性がある場合、`RangeAreas`  内のすべての範囲のデータの入力規則を表す`DataValidation` オブジェクトを返す`dataValidation` プロパティを持ちます。 同一な `DataValidation` オブジェクトが、 `RangeAreas`内のすべての範囲に適用されない場合、プロパティは `null` です。 全般的に、汎用的でない場合、 `RangeAreas` オブジェクトの原則は: *もしプロパティが、 `RangeAreas`内のすべての範囲の値に一貫性を持っていない場合、 `null`です*。 いくつかの例外の詳細については、 [RangeAreas のプロパティを読み取り中](#reading-properties-of-rangeareas) を参照してください。
- `RangeAreas.cellCount` 内のすべての範囲内のセルの合計数を取得します。`RangeAreas`
- `RangeAreas.calculate` 内のすべての範囲のセルを再計算します。`RangeAreas`
- `RangeAreas.getEntireColumn` また、 `RangeAreas.getEntireRow` は、 `RangeAreas`内のすべての範囲のすべての列 (または行) を表す他の `RangeAreas` オブジェクトを返します。 例えば、 `RangeAreas` が”A1: C4"と”F14:L15”を表し、`RangeAreas.getEntireColumn` が "A:C"と"F:L"を表すオブジェクト `RangeAreas`を返します。
- `RangeAreas.copyFrom` コピー操作のソース範囲を表すパラメーター`Range` または `RangeAreas` のいずれがをとることができます。

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a>RangeAreas でも利用できる範囲のメンバーの完全なリスト

##### <a name="properties"></a>プロパティ

リストされた任意のプロパティを読み取るコードを記述する前に、「[RangeAreas のプロパティを読み取る](#reading-properties-of-rangeareas)」をご確認ください。 返されるものは微妙に異なります。

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

プレビューの範囲メソッドを示します。

- calculate()
- clear()
- convertDataTypeToText() (プレビュー)
- convertToLinkedDataType() (プレビュー)
- copyFrom() (プレビュー)
- getEntireColumn()
- getEntireRow()
- getIntersection()
- getIntersectionOrNullObject()
- getOffsetRange() (RangeAreas オブジェクトの named getOffsetRangeAreas を名前付け)
- getSpecialCells() (プレビュー)
- getSpecialCellsOrNullObject() (プレビュー)
- getTables() (プレビュー)
- getUsedRange() (RangeAreas オブジェクトの getUsedRangeAreas を名前付け)
- getUsedRangeOrNullObject() (RangeAreas オブジェクトの named getUsedRangeAreasOrNullObject を名前付け)
- load()
- set()
- setDirty() (プレビュー)
- toJSON()
- track()
- untrack()

### <a name="rangearea-specific-properties-and-methods"></a>RangeArea に固有のプロパティおよびメソッド

`RangeAreas` 型には、`Range` オブジェクト上にないいくつかのプロパティとメソッドがあります。 次はその一部です。

- `areas``RangeAreas` オブジェクトで表される範囲のすべてを含む `RangeCollection` オブジェクトです。 オブジェクトは、新しく、他の Excel のコレクション オブジェクトに似ています。`RangeCollection` 範囲を表す `Range` オブジェクトの配列である`items` プロパティを持ちます。
- `areaCount`: 範囲内の合計数は、 `RangeAreas`です。
- `getOffsetRangeAreas`:[   が返され、元の](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-) `RangeAreas`   の範囲の一つからそれぞれからのオフセットの範囲を含む点を除き、 Range.getOffsetRange`RangeAreas` と同じように動作します。

## <a name="create-rangeareas-and-set-properties"></a>RangeAreas の作成と、プロパティの設定

オブジェクトを2 つの基本的な方法で作成することができます。`RangeAreas`

- を呼び出し、コンマで区切られた範囲のアドレスを使用して文字列を渡します。`Worksheet.getRanges()` 含めたい任意の範囲が、 [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem)された場合は、文字列内のアドレスではなく、名前を含めることができます。
- `Workbook.getSelectedRanges()` を呼び出します。 このメソッドは、現在アクティブなワークシートで選択されているすべての範囲を表す `RangeAreas` を返します。

オブジェクトを作成したら、  `getOffsetRangeAreas` と `getIntersection`のような`RangeAreas` を返すオブジェクト上のメソッドを使用して他のユーザーを作成することができます。`RangeAreas`

> [!NOTE]
> オブジェクトに、追加の範囲を直接追加することはできません。`RangeAreas` 例えば、`RangeAreas.areas`内のコレクションは、 `add` メソッドを持ちません。


> [!WARNING] 
> の配列のメンバーを、直接追加または削除しないようにしてください。`RangeAreas.areas.items` コード内で望ましくない動作をしてしまいます。 たとえば、さらに配列上に追加の `Range` オブジェクトをプッシュすることは可能ですが、 `RangeAreas` プロパティやメソッドは、新しいアイテムがない場合と同様に動作するため、これを行うとエラーが発生します。 例えば、 `areaCount` プロパティには、この方法によりプッシュされた範囲を含みません。 `index` よりも大きい `areasCount-1`場合、 `RangeAreas.getItemAt(index)` は、エラーをスローします。 同様に、参照を取得し、その`Range.delete` メソッドを呼び出して、  `RangeAreas.areas.items`内の`Range` オブジェクトを削除すると、バグが発生します:  `Range` オブジェクト *は、* 削除されましたが、親の `RangeAreas` オブジェクトのプロパティとメソッドは、それがまだ存在するかのように動作、またはしようとします。 例えば、コードが `RangeAreas.calculate`を呼び出した場合、Office は、範囲を計算しようとしますが、がエラーが発生し、range オブジェクトは失われます。

上のプロパティの設定は、 `RangeAreas.areas` コレクション上のすべての範囲に対応するプロパティを設定します。`RangeAreas`

次は、複数の範囲のプロパティの設定の例です。 関数には、 **F3:F5** と **H3:H5**の範囲が強調表示されます。

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

この例では、 `getRanges`に渡す範囲のアドレスを渡すハード コードできるか 、簡単に実行時に自動的に計算できるシナリオを適用します。 これが正しいであろうシナリオの一部は次のとおりです。 

- コードは、既知のテンプレートのコンテキストで実行されます。
- コードは、データのスキーマがわかっているインポートされたデータのコンテキストで実行されます。

コーディング時に動作する必要がある範囲を知ることはできません、実行時に検出する必要があります。 次のセクションでは、これらのシナリオについて説明します。

### <a name="discover-range-areas-programmatically"></a>範囲の領域をプログラムで検出します。

`Range.getSpecialCells()` と `Range.getSpecialCellsOrNullObject()` メソッドを使用すると、実行時に、セルの特性とセルの値の型を基に操作し実行したい範囲を検索できます。 TypeScriptデータ型のファイルからのメソッドのシグネチャを次に示します。

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

次は、最初のシグネチャを使用する場合の例です。 このコードの注意点は次のとおりです。

- 最初の呼び出しで検索する必要があるシートの一部を制限して、その範囲のみで `Worksheet.getUsedRange`  `getSpecialCells` を呼び出します。
- 列挙型からの値の文字列バージョンをパラメーターとして`getSpecialCells`  に渡します。`Excel.SpecialCellType` 代わりに渡される他の値のいくつかは、空白のセルには「空白」、数式のかわりにリテラル値を持つセルには「定数」、 `usedRange`の最初のセルと同じ条件付き書式が設定されるセルには"SameConditionalFormat"です。 最初のセルは、上の左端のセルです。 列挙型の値の完全な一覧は、 [ベータ版の office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)を参照してください。
- `getSpecialCells` メソッドは、 `RangeAreas` オブジェクトを返します。数式を入力した全てのセルは、すべて連続していない場合でも、ピンク色に色分け表示されます。 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

場合によっては、対象となる特性を持つ *任意* のセルが範囲にありません。 `getSpecialCells` が対象となるセルをどうしても見つけられない場合、 **ItemNotFound** エラーがスローされます。 これは、一つだけある場合、コントロールのフローを `catch` ブロックまたはメソッドにそらします。 そうでない場合、エラーは、機能を停止します。 エラーをスローすることが、対象となる特性を持つセルが存在しない場合に、正確に必要されることであるシナリオがある可能性があります。 

通常の動作では、シナリオでは、一致するセルがない場合もありますが、これはおそらく一般的ではありません。コードは、この可能性を確認し、エラーをスローすることがなく適切に処理すること必要があります。 これらのシナリオでは、 `getSpecialCellsOrNullObject` メソッドを使用し、 `RangeAreas.isNullObject` プロパティをテストします。 次に例を示します。 このコードの注意点は次のとおりです。

- メソッドは、常にプロキシ オブジェクトを返します。したがって、通常 JavaScript という意味では、決して `null` にはなりません。`getSpecialCellsOrNullObject` 一致するセルが見つからない場合は、 `isNullObject` オブジェクトのプロパティが、 `true`に設定されます。
- プロパティをテストする前に、 を呼び出します。`context.sync`**`isNullObject` 読みだすために常にプロパティを読み込み同期させる必要があるため、これはすべての `*OrNullObject` メソッドとプロパティの必要条件です。 ただし、 `isNullObject` プロパティを *明示的に* 読み込む必要はありません。 がオブジェクト上に呼び出されない場合でも、 `context.sync` により自動的に読み込まれます。`load` 詳細については、「[\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods)」をご覧ください。
- 最初に数式のセルを含まない範囲を選択して実行することで、このコードをテストできます。 次に、少なくとも 1 つのセルに数式がある範囲を選択し、それをもう一度実行します。

```js
Excel.run(function (context) {
    const range = context.workbook.getSelectedRange();
    const formulaRanges = range.getSpecialCellsOrNullObject("Formulas");
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

この資料の他のすべての例をわかりやすくするためには、 `getSpecialCells` メソッドを `getSpecialCellsOrNullObject`の代わりに使用します。

#### <a name="narrow-the-target-cells-with-cell-value-types"></a>セルの値の型で対象セルを絞り込む

列挙型の省略可能な 2 番目のパラメーターがあり、対象とするセルをさらに絞り込みます。`Excel.SpecialCellValueType` 「数式」または「定数」のいずれかを渡す場合にのみに `getSpecialCells` または `getSpecialCellsOrNullObject`を使用できます。 パラメータは、必要な特定の種類の値のセルのみを指定します。 4 つの基本的な種類があります:「エラー」、「論理」(つまり、ブール値)、「番号」、および「テキスト」です。 (列挙型は、これら以外の他の値について次に説明する 4 つの他の値です。)次に、例を示します。 このコードの注意点は次のとおりです。

- リテラルの数値のあるセルのみがハイライト表示されます。 数式(結果は数値の場合でも)またはブール値、文字列、またはエラーの状態のセルが強調表示されます。
- コードをテストするには、リテラルの数値、他の種類のリテラル値、一部の数式のセルがワークシートにあることを確認してください。

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

場合によって 、すべてのテキスト値を持ち、すべてのブール値 (「論理」) を持つセルのよう1 つ以上のセル値型を操作する必要があります。 列挙型には値の種類を組み合わせることができます。`Excel.SpecialCellValueType` たとえば、"LogicalText"は、すべてのブール値、テキスト値を持つセルを対象にします。 4 つの基本的なタイプの内 2 つまたは 3 つを組み合わせることができます。 基本的な種類を組み合わせているこれらの列挙値の名前は、常にアルファベット順にします。 セルのエラー値、テキスト値、およびブール値を結合するには、"LogicalErrorText"または"TextErrorLogical"ではなく"ErrorLogicalText"を使用します。 「すべて」の既定のパラメーターは、4 種類全てを結合します。 次の使用例では、数値またはブール値を生成する数式を持つすべてのセルを強調表示しています。

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaLogicalNumberRanges = usedRange.getSpecialCells("Formulas", "LogicalNumbers");
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

> [!NOTE]
> `Excel.SpecialCellValueType` パラメーターは、 `Excel.SpecialCellType` パラメーターが "数式" または "定数" である場合にのみ使用できます。

### <a name="get-rangeareas-within-rangeareas"></a>RangeAreas 内の RangeAreas を取得します。

型自体も、同じの 2 つのパラメーターを受け取る `getSpecialCells` と `getSpecialCellsOrNullObject` メソッドを持ちます。`RangeAreas` これらのメソッドは、 `RangeAreas.areas` コレクション内のすべての範囲のセル範囲からすべての対象となるセルを返します。 オブジェクト の代わりに オブジェクトを呼び出した場合のメソッドの動作のには1 つの小さな違いがあります :"SameConditionalFormat"を最初のパラメーターとして渡すと、メソッドは、  コレクション 内の最初の範囲の上の左端のセル と同じ条件付き書式を持つすべてのセルを返します。`RangeAreas``Range`*`RangeAreas.areas`* 同じポイントは、"SameDataValidation"に適用されます:  に渡される場合、 範囲内の  一番左上と同じデータの入力規則をもつすべてのセルを返します。`Range.getSpecialCells`** しかｈしに渡された売位、  コレクション内の 最初の範囲の左端のセルと同じデータの入力規則が持つすべてのセルが返されます。`RangeAreas.getSpecialCells`*`RangeAreas.areas`*

## <a name="read-properties-of-rangeareas"></a>RangeAreas のプロパティの読み取り

プロパティ値を読み取るときは、指定したプロパティが、 `RangeAreas`内の別の範囲内の異なる値を持つ可能性があるため、注意が必要です。`RangeAreas` 一般的な規則は、一貫性のある値 *が* 返される場合、それが返されることです。 例えば、次のコードでは、ピンク色の (`#FFC0CB`) と `true` 用のRGBコードは、 `RangeAreas` オブジェクトの両方の範囲がピンク色の塗りであり、両方が全体の列であるため、コンソールに格納されます。

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    const rangeAreas = sheet.getRanges("F:F, H:H");  
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

整合性が可能でない場合、より複雑になります。 プロパティの動作は、次の 3 つの原則に従います。`RangeAreas`

- オブジェクトのブール型のプロパティは、すべてのメンバーの範囲が true でない限り、 `false` を返します。`RangeAreas`
- 非ブール値プロパティは、 `address` プロパティ例外を除いて、全てのメンバーの範囲に対応するプロパティが同じ値を持っていない限り、 `null` を返します。
- プロパティは、メンバーの範囲のアドレスのコンマ区切りの文字列を返します。`address`

たとえば、次のコードは、1 つだけ列全体であり、 1 つだけがピンク色で塗りつぶされます `RangeAreas` を生成します。 コンソールが、塗りつぶしの色に `null` を、 `false` を `isEntireRow` プロパティに、および"Sheet1!F3:F5、Sheet1!H:H"("Sheet1"は、シート名と仮定した場合) を `address` プロパティに表示します。 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H:H");

    const pinkColumnRange = sheet.getRange("H:H");
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

- [Excel JavaScript API の中心概念](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview)
- [Range オブジェクト (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)
- [RangeAreas オブジェクト (EXCELL用JavaScript API )](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (API がプレビュー中の場合、このリンクが動作しない 可能性があります。 代わりに、 [ベータ版 office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)を参照してください。)