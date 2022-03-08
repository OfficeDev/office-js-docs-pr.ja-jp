---
title: Excel アドインで複数の範囲を同時に操作する
description: JavaScript ライブラリExcelを使用して、複数の範囲で操作を実行し、プロパティを設定する方法について説明します。
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: 75b1248a15c37c548b11fa8ac47a809b045571e4
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340912"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins"></a>Excel アドインで複数の範囲を同時に操作する

Excel JavaScript ライブラリを使用すると、同時に複数の範囲に対してアドインによる操作の実行とプロパティの設定が可能になります。 範囲は連続している必要はありません。 コードがよりシンプルになることに加え、この方法でプロパティを設定すれば、各範囲に同じプロパティを個別に設定する方法よりも処理速度が格段に速くなります。

## <a name="rangeareas"></a>RangeAreas

一連の (不一視の可能性がある) 範囲は [、RangeAreas オブジェクトによって表](/javascript/api/excel/excel.rangeareas) されます。 `Range` 型と同様のプロパティとメソッドを持ちますが (多くの場合は同じまたは類似した名前)、以下に対しては調整が行われています。

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

### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a>RangeAreas でも利用可能な Range メンバーの全リスト

#### <a name="properties"></a>プロパティ

リストにあるプロパティを読み取るコードを書く前に、「[RangeAreas のプロパティの読み取り](#read-properties-of-rangeareas)」の内容を理解しておいてください。 繰り返される内容について細かい注意点があります。

- `address`
- `addressLocal`
- `cellCount`
- `conditionalFormats`
- `context`
- `dataValidation`
- `format`
- `isEntireColumn`
- `isEntireRow`
- `style`
- `worksheet`

#### <a name="methods"></a>メソッド

- `calculate()`
- `clear()`
- `convertDataTypeToText()`
- `convertToLinkedDataType()`
- `copyFrom()`
- `getEntireColumn()`
- `getEntireRow()`
- `getIntersection()`
- `getIntersectionOrNullObject()`
- `getOffsetRange()` (オブジェクトの `getOffsetRangeAreas` 名前 `RangeAreas` )
- `getSpecialCells()`
- `getSpecialCellsOrNullObject()`
- `getTables()`
- `getUsedRange()` (オブジェクトの `getUsedRangeAreas` 名前 `RangeAreas` )
- `getUsedRangeOrNullObject()` (オブジェクトの `getUsedRangeAreasOrNullObject` 名前 `RangeAreas` )
- `load()`
- `set()`
- `setDirty()`
- `toJSON()`
- `track()`
- `untrack()`

### <a name="rangearea-specific-properties-and-methods"></a>RangeArea 固有のプロパティとメソッド

`RangeAreas` 型には、`Range` オブジェクトには存在しないプロパティとメソッドがいくつかあります。 次に、選択した内容を示します。

- `areas`: `RangeAreas` オブジェクトが表す全範囲を含む `RangeCollection` オブジェクト。 `RangeCollection` オブジェクトも新しいオブジェクトであり、他の Excel コレクション オブジェクトと類似しています。 これには、範囲を表す `Range` オブジェクトの配列である `items` プロパティがあります。
- `areaCount`: `RangeAreas` で指定された範囲の合計数。
- `getOffsetRangeAreas`: [Range.getOffsetRange](/javascript/api/excel/excel.range#excel-excel-range-getoffsetrange-member(1)) と同じように動作します。ただし、`RangeAreas` を返し、元の `RangeAreas` で指定された範囲の 1 つからの各オフセットである範囲を含みます。

## <a name="create-rangeareas"></a>RangeAreas の作成

`RangeAreas` オブジェクトの作成には、2 つの基本的な方法があります。

- `Worksheet.getRanges()` を呼び出して、範囲のアドレスがコンマで区切られた文字列を渡します。 含める対象の範囲が既に [NamedItem](/javascript/api/excel/excel.nameditem) に指定されている場合、文字列にはアドレスではなくその名前を指定することができます。
- `Workbook.getSelectedRanges()` を呼び出します。 このメソッドは、現在アクティブなワークシート上で選択されている全範囲を表す `RangeAreas` を返します。

一度 `RangeAreas` オブジェクトを作成すると、`getOffsetRangeAreas` や `getIntersection` など、`RangeAreas` を返すオブジェクト上のメソッドを使用して別のオブジェクトを作成できます。

> [!NOTE]
> `RangeAreas` オブジェクトに新たな範囲を直接追加することはできません。 たとえば、`RangeAreas.areas` 内のコレクションには `add` メソッドが存在しません。

> [!WARNING]
> `RangeAreas.areas.items` 配列のメンバーの追加または削除を直接試行してはいけません。 これにより、後でコード内で望ましくない動作が発生します。 たとえば、追加の `Range` オブジェクトを配列にプッシュすることは可能ですが、エラーが発生します。`RangeAreas` のプロパティとメソッドは、その新しいアイテムがその場所に存在していないかのように動作するためです。 たとえば、`areaCount` プロパティにはこの方法でプッシュされた範囲は含まれません。また、`RangeAreas.getItemAt(index)` は、`index` が `areasCount-1`より大きい場合、エラーをスローします。 同様に、`RangeAreas.areas.items` 配列内の `Range` オブジェクトを、参照を取得してその `Range.delete` メソッドを呼び出すという方法で削除すると、バグとなります。`Range` オブジェクトは *削除されます* が、親 `RangeAreas` オブジェクトのプロパティとメソッドは、そのオブジェクトがまだ存在するものとして動作するためです。 たとえば、コードで `RangeAreas.calculate` を呼び出すと、Office は範囲を計算しようとしますが、範囲オブジェクトが既に存在しないためにエラーとなります。

## <a name="set-properties-on-multiple-ranges"></a>複数の範囲でのプロパティの設定

`RangeAreas` オブジェクトでプロパティを設定すると、`RangeAreas.areas` コレクション内の全範囲の対応するプロパティが設定されます。

次に、複数の範囲にプロパティを設定する例を示します。 この関数は、**F3:F5** と **H3:H5** の範囲を強調表示します。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    await context.sync();
});
```

この例は、`getRanges` に渡す範囲のアドレスをハード コーディングできる場合や実行時に簡単に計算できる場合に適用されます。 たとえば、これが適切なのは次のような場合です。

- コードが、既知のテンプレートのコンテキスト内で実行される。
- コードが、データのスキーマが既知であるインポート済みデータのコンテキスト内で実行される。

## <a name="get-special-cells-from-multiple-ranges"></a>複数の範囲からの特定のセルの取得

`RangeAreas` オブジェクトの `getSpecialCells` メソッドと `getSpecialCellsOrNullObject` メソッドは、`Range` オブジェクトの同じ名前のメソッドと同じように機能します。 これらのメソッドでは、`RangeAreas.areas` コレクション内のすべての範囲から、指定された特性を持つセルが返されます。 特殊なセルの詳細については、「範囲内の特殊 [なセルを検索する」を参照してください](excel-add-ins-ranges-special-cells.md)。

`RangeAreas` オブジェクトで `getSpecialCells` メソッドまたは `getSpecialCellsOrNullObject` メソッドを呼び出す場合:

- 最初のパラメーターとして `Excel.SpecialCellType.sameConditionalFormat` を渡した場合、このメソッドでは、`RangeAreas.areas` コレクション内の最初の範囲の左上隅のセルと同じ条件付き書式を持つセルがすべて返されます。
- 最初のパラメーターとして `Excel.SpecialCellType.sameDataValidation` を渡した場合、このメソッドでは、`RangeAreas.areas` コレクション内の最初の範囲の左上隅のセルと同じデータ検証ルールを持つセルがすべて返されます。

## <a name="read-properties-of-rangeareas"></a>RangeAreas のプロパティの読み取り

`RangeAreas` のプロパティ値の読み取りには、注意が必要です。`RangeAreas`内の範囲それぞれで、プロパティの値が異なる可能性があるためです。 一貫性のある値を返すことが *できる* 場合には返す、というのが一般的なルールです。 たとえば、次のコードでは、ピンク (`#FFC0CB`) `true` `RangeAreas` の RGB コードで、オブジェクト内の両方の範囲がピンク色の塗りつぶしを持ち、両方とも列全体のため、コンソールにログに記録されます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    let rangeAreas = sheet.getRanges("F:F, H:H");  
    rangeAreas.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn");
    await context.sync();

    console.log(rangeAreas.format.fill.color); // #FFC0CB
    console.log(rangeAreas.isEntireColumn); // true
});
```

一貫性を期待できない場合、事態は複雑となります。 `RangeAreas` プロパティの動作は、次の 3 つの原則に従います。

- `RangeAreas` オブジェクトのブール値プロパティは、すべてのメンバー範囲でプロパティが true でない限り、`false` を返します。
- ブール値以外のプロパティ (`address` プロパティを除く) は、すべてのメンバー範囲で対応するプロパティが同じ値ではない限り、`null` を返します。
- `address` プロパティは、メンバー範囲のアドレスをコンマで区切った文字列を返します。

たとえば、次のコードでは、1 つの範囲のみが列全体であり、1 つの範囲のみがピンクで塗りつぶされている `RangeAreas` を作成します。 コンソールには、塗りつぶし色の場合は `null`、`isEntireRow` プロパティの場合は `false`、`address` プロパティの場合は "Sheet1!F3:F5, Sheet1!H:H" ("Sheet1" はシート名) が表示されます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let rangeAreas = sheet.getRanges("F3:F5, H:H");

    let pinkColumnRange = sheet.getRange("H:H");
    pinkColumnRange.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn, address");
    await context.sync();

    console.log(rangeAreas.format.fill.color); // null
    console.log(rangeAreas.isEntireColumn); // false
    console.log(rangeAreas.address); // "Sheet1!F3:F5, Sheet1!H:H"
});
```

## <a name="see-also"></a>関連項目

- [Excel JavaScript API を使用した基本的なプログラミングの概念](../reference/overview/excel-add-ins-reference-overview.md)
- [JavaScript API を使用した大きな範囲の読み取りExcel書き込み](excel-add-ins-ranges-large.md)
