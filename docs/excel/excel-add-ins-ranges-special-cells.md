---
title: JavaScript API を使用して範囲内の特別なセルExcel検索する
description: JavaScript API の Excelを使用して、数式、エラー、数値を含むセルなどの特別なセルを検索する方法について説明します。
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 1252fe599f93a3408fb161e2b8204600fa483339
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340520"
---
# <a name="find-special-cells-within-a-range-using-the-excel-javascript-api"></a>JavaScript API を使用して範囲内の特別なセルExcel検索する

この記事では、JavaScript API を使用して範囲内の特殊なセルを検索するExcel示します。 オブジェクトがサポートするプロパティとメソッドの`Range`完全な一覧については、「Excel[。Range クラス](/javascript/api/excel/excel.range)。

## <a name="find-ranges-with-special-cells"></a>特殊なセルを含む範囲を検索する

[Range.getSpecialCells](/javascript/api/excel/excel.range#excel-excel-range-getspecialcells-member(1)) メソッドと [Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#excel-excel-range-getspecialcellsornullobject-member(1)) メソッドは、セルの特性とセルの値の種類に基づいて範囲を検索します。 これらのメソッドでは両方とも、`RangeAreas` オブジェクトが返されます。 次に示すのは、TypeScript データ型ファイルの、このメソッドのシグネチャです。

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

次のコード サンプルでは、メソッドを `getSpecialCells` 使用して数式を含むすべてのセルを検索します。 このコードの注意点は次のとおりです。

- 検索が必要なシートの部分を制限するために、まず `Worksheet.getUsedRange` を呼び出し、その範囲に関してのみ `getSpecialCells` を呼び出します。
- `getSpecialCells` メソッドは `RangeAreas` オブジェクトを返すため、数式を含むセルはすべて、連続していないセルであっても、ピンク色になります。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let usedRange = sheet.getUsedRange();
    let formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    await context.sync();
});
```

対象の特性を含むセルが範囲内に存在しない場合、`getSpecialCells` によって **ItemNotFound** エラーがスローされます。 この場合、制御のフローが `catch` ブロックに移ります (存在する場合)。 ブロックが見当たらない場合 `catch` は、メソッドが停止します。

対象の特性を含むセルが常に存在するはずである場合、そうしたセルが存在しないなら、コードを使ってエラーをスローする必要があるかもしれません。 一致するセルがないということが有効なシナリオでは、コードでこのような可能性があるかどうかを確認し、あれば、エラーをスローせずに適切に処理するようにしておく必要があります。 `getSpecialCellsOrNullObject` メソッドと、返された `isNullObject` プロパティを使用して、この動作を実現できます。 次のコード サンプルでは、このパターンを使用します。 このコードについては、以下の点に注意してください。

- メソッド `getSpecialCellsOrNullObject` は常にプロキシ オブジェクトを返すの `null` で、通常の JavaScript の意味では返す必要がありません。 ただし一致するセルが見つからなかった場合、オブジェクトの `isNullObject` プロパティは `true` に設定されます。
- `isNullObject` プロパティをテストする *前* に、`context.sync` を呼び出します。 これは、すべての `*OrNullObject` メソッドとプロパティの必要条件です。プロパティを読み取るためには常に、そのプロパティをロードして同期する必要があるためです。 ただし、プロパティを明示的に *読み込む* 必要 `isNullObject` はありません。 オブジェクトで呼び出されない場合 `context.sync` でも `load` 、自動的に読み込まれます。 詳細については、「 [\*OrNullObject メソッドとプロパティ」を参照してください](../develop/application-specific-api-model.md#ornullobject-methods-and-properties)。
- このコードをテストするには、最初に数式を含まないセルの範囲を選択してからコードを実行します。 次に、少なくとも 1 つのセルが数式を含む範囲を選択してからコードを再実行します。

```js
await Excel.run(async (context) => {
    let range = context.workbook.getSelectedRange();
    let formulaRanges = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
    await context.sync();
        
    if (formulaRanges.isNullObject) {
        console.log("No cells have formulas");
    }
    else {
        formulaRanges.format.fill.color = "pink";
    }
    
    await context.sync();
});
```

わかりやすくするために、この記事の他のすべてのコード サンプルでは `getSpecialCells` 、 の代わりにメソッドを使用します  `getSpecialCellsOrNullObject`。

## <a name="narrow-the-target-cells-with-cell-value-types"></a>セルの値の型に応じて対象のセルを絞り込む

`Range.getSpecialCells()` メソッドと `Range.getSpecialCellsOrNullObject()` メソッドでは、対象セルをさらに絞り込むためにオプションとして使用される 2 番目のパラメーターを承諾します。 この 2 番目のパラメーターは、特定の種類の値を含むセルのみを指定するために使用される `Excel.SpecialCellValueType` パラメーターです。

> [!NOTE]
> `Excel.SpecialCellValueType` パラメーターは、`Excel.SpecialCellType` が `Excel.SpecialCellType.formulas` または `Excel.SpecialCellType.constants` の場合にのみ使用できます。

### <a name="test-for-a-single-cell-value-type"></a>単一のセル値の型のテスト

`Excel.SpecialCellValueType` 列挙型には、次の 4 つの基本型があります (このセクションで後述する他の値の組み合わせに加えて)。

- `Excel.SpecialCellValueType.errors`
- `Excel.SpecialCellValueType.logical` (ブール型)
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

次のコード サンプルでは、数値定数である特殊なセルを検索し、それらのセルをピンク色で色付けします。 このコードについては、以下の点に注意してください。

- リテラル数値を持つセルのみを強調表示します。 数式 (結果が数値の場合でも) またはブール値、テキスト、またはエラー状態のセルを持つセルは強調表示されます。
- コードをテストするには、リテラル数値を持ついくつかのセル、他の型のリテラル値を持ついくつかのセル、そして数式を持ついくつかのセルをそれぞれワークシートに含めるようにしてください。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let usedRange = sheet.getUsedRange();
    let constantNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.constants,
        Excel.SpecialCellValueType.numbers);
    constantNumberRanges.format.fill.color = "pink";

    await context.sync();
});
```

### <a name="test-for-multiple-cell-value-types"></a>複数のセル値の型のテスト

テキスト値のセルすべてとブール値 (`Excel.SpecialCellValueType.logical`) のセルすべてなど、セル値の型を複数操作する必要がある場合もあります。 `Excel.SpecialCellValueType` 列挙型には、結合された型の値があります。 たとえば、`Excel.SpecialCellValueType.logicalText` は、すべてのブール値のセルとテキスト値のセルを対象としています。 `Excel.SpecialCellValueType.all` は既定値であり、返されるセル値の型は制限されません。 次のコード サンプルでは、数値またはブール値を生成する数式ですべてのセルを色付けします。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let usedRange = sheet.getUsedRange();
    let formulaLogicalNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.formulas,
        Excel.SpecialCellValueType.logicalNumbers);
    formulaLogicalNumberRanges.format.fill.color = "pink";

    await context.sync();
});
```

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [JavaScript API を使用して文字列をExcelする](excel-add-ins-ranges-string-match.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
