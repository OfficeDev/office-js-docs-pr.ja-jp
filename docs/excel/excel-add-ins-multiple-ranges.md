---
title: Excel のアドインで同時に複数の範囲の操作をします。
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: a00bbf15b53649147fb2c2b1dfa590f15c5739be
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506295"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a>Excel のアドイン (プレビュー) で同時に複数のセル範囲を操作します。

Excel の JavaScript ライブラリは、操作を実行し、同時に複数の範囲のプロパティを設定するように追加できます。範囲が隣接している必要はありません。コードを簡単にするだけでなく、このプロパティを設定する方法は、範囲ごとに個別に同じプロパティを設定するよりもはるかに高速実行されます。

> [!NOTE]
> この資料に記載されている Apiは、 **Office 2016 クイック実行バージョン ビルト 1809 10820.20000 の** 以降を日梅雨とします。(適切なビルドを取得する [Office 内部からのプログラム](https://products.office.com/office-insider) に参加する必要があります)。また、 [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)からベータ版の Office の JavaScript ライブラリを読み込む必要があります。最後に、これらの Api の参照ページまだありません。次の種類の定義ファイルにはそれらについての説明: [ベータ版の office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)です。

## <a name="rangeareas"></a>RangeAreas

(場合によっては連続していない) の範囲のセットは、 `Excel.RangeAreas` オブジェクトで表されます。`Range`に似ているプロパティとメソッドを持ちますが、 (多くの場合、同じまたは同様の名前を持つ)、以下のような修正が加えられました。

- プロパティのデータ型とセッターとゲッターの動作です。
- メソッドパラメーターのデータ型と、メソッドの動作です。
- メソッドのデータ型は、値を返します。

例:

- `RangeAreas` `Range.address` プロパティとともに1 つのアドレスとしてではなく、範囲のアドレスのコンマ区切りの文字列を返す`address` プロパティを持ちます。
- `RangeAreas` 一貫性のある場合に`RangeAreas`内のすべての範囲のデータの有効性を表す`DataValidation`オブジェクトを返す `dataValidation` プロパティを持ちます。`RangeAreas`内のすべての範囲に適用されない  `DataValidation`と同一の場合、  プロパティは、オブジェクトは`null` です。これは、全般的でかつ汎用的でない `RangeAreas` オブジェクト: *プロパティが、`RangeAreas`すべての範囲の値に一貫性を持っていない場合、  `null`です* 。いくつかの例外の詳細については、 [RangeAreas のプロパティを読み取り中](#reading-properties-of-rangeareas) を参照してください。
- `RangeAreas.cellCount` `RangeAreas` 内のすべての範囲内のセルの合計数を取得します。
- `RangeAreas.calculate` `RangeAreas` 内のすべての範囲のセルを再計算します。
- `RangeAreas.getEntireColumn` また `RangeAreas.getEntireRow` は、 `RangeAreas` 内のすべての範囲のすべての列 (または行) を表す `RangeAreas`オブジェクトを返します。例えば、 `RangeAreas` が、「a1: c4」および「F14:L15」を表す場合、 `RangeAreas.getEntireColumn` は、 "A:C"と"F:L"を表すオブジェクト`RangeAreas` を返します。
- `RangeAreas.copyFrom` コピー操作のソース範囲を表すパラメーターである`Range` または `RangeAreas` のいずれをとることができます。

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a>RangeAreas でも利用できる範囲のメンバーの完全なリスト

##### <a name="properties"></a>プロパティ

任意のプロパティを読み取るコードの一覧を記述する前に、 [RangeAreas のプロパティを読み取り中](#reading-properties-of-rangeareas) に精通します。何が返されるかは微妙です。

- アドレス
- addressLocal
- cellCount
- conditionalFormats
- コンテキスト
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
- getUsedRangeOrNullObject() (RangeAreas オブジェクトでは getUsedRangeAreasOrNullObject という名前)
- load()
- set()
- setDirty() (プレビュー)
- toJSON()
- track()
- untrack()

### <a name="rangearea-specific-properties-and-methods"></a>RangeArea に固有のプロパティおよびメソッド

`RangeAreas` 型には、 `Range` オブジェクト上にないいくつかのプロパティとメソッドがあります。それらのいくつかを次に示します。

- `areas`  `RangeAreas` で表される範囲のすべてを含むオブジェクトの  `RangeCollection` オブジェクトです。 `RangeCollection` オブジェクトは新しく、他の Excel のコレクション オブジェクトに似ています。 プロパティの範囲を表す `Range` オブジェクトの配列である`items` プロパティを持ちます。
- `areaCount`: `RangeAreas` 内の合計数。
- `getOffsetRangeAreas`:[   が返され、元の](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-) `RangeAreas`   の範囲の一つからそれぞれからのオフセットの範囲を含む点を除き、 Range.getOffsetRange`RangeAreas` と同じように動作します。

## <a name="create-rangeareas-and-set-properties"></a>RangeAreas の作成と、プロパティの設定

`RangeAreas`オブジェクトを 2 つの基本的な方法で作成することができます。

- `Worksheet.getRanges()` を呼び出し、コンマで区切られた範囲アドレスを含む文字列を渡します。含めたい任意の範囲を [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem)  とした場合、文字列にアドレスではなく、名前を含めることができます。
- `Workbook.getSelectedRanges()`を呼出します。このメソッドは、現在アクティブなワークシートで選択されているすべての範囲を表す`RangeAreas`を返します。

`RangeAreas` オブジェクトを作成したら、`getOffsetRangeAreas` および `getIntersection` のような `RangeAreas` を返すオブジェクト上のメソッドを使用して他のユーザーを作成することができます。

> [!NOTE]
> `RangeAreas`オブジェクトに、追加の範囲を直接追加することはできません。例えば、`RangeAreas.areas`内のコレクションは、`add`メソッドを持ちません。


> [!WARNING] 
> `RangeAreas.areas.items`  の配列のメンバーを直接追加または削除しないようにしてください。コード内で望ましくない動作をしてしまいます。例えば、配列に追加の `Range`  オブジェクトをプッシュすることは可能ですが、これにより `RangeAreas`  プロパティやメソッドは、新しいアイテムがないかのように動作するためエラーが発生します。例えば、`areaCount` プロパティには、このような方法でプッシュされた範囲を含んでおらず、`RangeAreas.getItemAt(index)` が `index`より大きい場合に、`areasCount-1` がエラーをスローします。同様に、参照を取得し、その メソッドを呼び出して、`Range` 内の <オブ`RangeAreas.areas.items`  `Range.delete`ジェクトを削除すると、バグが発生します。 オブジェクト`Range`は* 、* 削除されますが、親の`RangeAreas`オブジェクトのプロパティとメソッドは、それがまだ存在するかのように動作するか、動作しようとします。例えば、`RangeAreas.calculate` を呼出した場合、Office は範囲を計算しようとしますが、範囲オブジェクトがないためエラーが生じます。

`RangeAreas`上のプロパティの設定は、`RangeAreas.areas`コレクション上のすべての範囲に対応するプロパティを設定します。

次は、複数の範囲のプロパティの設定の例です。関数には、**F3:F5** と **H3:H5** の範囲が強調表示されます。

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

この例では、`getRanges`に渡す範囲のアドレスを渡すハード コードできるか 、簡単に実行時に自動的に計算できるシナリオを適用します。これが正しいであろうシナリオの一部は次のとおりです。 

- コードは、既知のテンプレートのコンテキストで実行されます。
- コードは、データのスキーマがわかっているインポートされたデータのコンテキストで実行されます。

コーディング時にどのような範囲で実行すれはよいかわからない場合には、ランタイムで検出する必要があります。

### <a name="discover-range-areas-programmatically"></a>範囲の領域をプログラムで検出します。

`Range.getSpecialCells()` と `Range.getSpecialCellsOrNullObject()` メソッドを使用すると、実行時に、セルの特性とセルの値の型を基に操作し範囲を検索できます。TypeScript データタイプのファイルからのメソッドのシグネチャを次に示します。

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

次は、最初のシグネチャを使用する場合の例です。このコードに関して以下に注意してください。

- その範囲のみで最初に`Worksheet.getUsedRange`を呼び出し、`getSpecialCells`を呼出して検索する必要があるシートの一部を制限します。
- `Excel.SpecialCellType`列挙型からの値の文字列バージョンをパラメータとして`getSpecialCells`に渡します。代わりに渡される他の値のいくつかは、空白のセルには「空白」、数式のかわりにリテラル値を持つセルには「定数」、`usedRange`の最初のセルと同じ条件付き書式が設定されるセルには「SameConditionalFormat」です。最初のセルは、上の左端のセルです。列挙型の値の一覧は、[ベータ版の「office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)」を参照してください。
- `getSpecialCells`メソッドは、`RangeAreas`オブジェクトを返します。数式を入力した全てのセルは、すべて連続していない場合も、ピンク色に色分け表示されます。 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

対象となる特性を持つ*任意*のセル範囲はありません。`getSpecialCells`が何も検出しない場合は、**ItemNotFound** エラーがスローされます。これがある場合は、`catch`ブロック/メソッドへのコントロールのフローを逸します。ない場合は、エラーは、機能を停止します。エラーをスローすることが、対象となる特性を持つセルが存在しない場合に、正確に必要されることであるシナリオがある可能性があります。 

通常の動作では、シナリオでは、一致するセルがない場合もありますが、これはおそらく一般的ではありません。コードは、この可能性を確認し、エラーをスローすることがなく適切に処理すること必要があります。これらのシナリオに関しては、`getSpecialCellsOrNullObject`メソッドを使用して`RangeAreas.isNullObject`プロパティをテストします。次に、例を示します。このコードに関して以下に注意してください。

- `getSpecialCellsOrNullObject`メソッドは、常にプロキシ オブジェクトを返します。したがって、通常 JavaScript という意味では、決して`null`にはなりません。一致するセルが見つからない場合は、オブジェクトの`isNullObject`プロパティが、`true`に設定されます。
- `isNullObject`プロパティをテストする*前に* 、`context.sync`を呼び出す。読み込むために常にプロパティを読み込み同期させる必要があるため、これはすべての`*OrNullObject`メソッドとプロパティの必要条件です。ただし、`isNullObject`プロパティを*明示的に*ロードする必要はありません。`load`がオブジェクトで呼び出されない場合でも、`context.sync`が自動的にロードします。詳細情報に関しては、[\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods) を参照してください 。
- 最初に数式のセルを含まない範囲を選択して実行することで、このコードをテストできます。次に、少なくとも 1 つのセルに数式がある範囲を選択し、それをもう一度実行します。

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

わかりやすくするために、この記事の他のすべての例では、`getSpecialCellsOrNullObject` ではなく`getSpecialCells` メソッドを使用しています。

#### <a name="narrow-the-target-cells-with-cell-value-types"></a>セルの値の型で対象セルを絞り込む

列挙型`Excel.SpecialCellValueType`の省略可能な 2 番目のパラメータがあり、対象に対してさらにセルを狭めるます。「数式」または「定数」のいずれかを`getSpecialCells`または`getSpecialCellsOrNullObject`に渡す場合にのみ使用できます。パラメータは、特定の種類の値のあるセルが必要であることを指定します。「エラー」、「論理」(ブール値を指す)、「番号」、および「テキスト」の 4 つの基本的な種類があります。(列挙型は、これら以外の他の値について次に説明する 4 つです。)次に、例を示します。このコードに関して次に注意してください。

- リテラルの数値のあるセルのみが強調表示されます。数式(結果は数値の場合でも)またはブール値、文字列、またはエラーの状態のセルが強調表示されます。
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

場合によってすべてのテキスト値およびすべてのブール値 (「論理」) のセルのように 1 つ以上のセル値の種類を操作する必要があります。`Excel.SpecialCellValueType`列挙型はタイプを組み合わせることができる値を持っています。たとえば、「LogicalText」はブール値、テキスト値を持つすべてのセルをターゲットにします。4 つの基本的なタイプの 2 つまたは 3 つを組み合わせることができます。基本的な種類を組み合わせるこれらの列挙値の名前は、常にアルファベット順にします。したがって、エラー値、テキスト値、およびブール値を持つセルを組み合わせるには、「LogicalErrorText」または「TextErrorLogical」ではなく「ErrorLogicalText」を使用します。「すべて」の既定のパラメータが全 4 種類を組み合わせます。次の使用例では、数値またはブール値を生成する式を持つすべてのセルが強調表示されています。

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
> `Excel.SpecialCellValueType` パラメータは、`Excel.SpecialCellType` パラメータが 「式」または「定数」である場合にのみ使用できます。

### <a name="get-rangeareas-within-rangeareas"></a>RangeAreas 内の RangeAreas を取得します。

`RangeAreas`タイプ自体も `getSpecialCells` と `getSpecialCellsOrNullObject` を持ち、これらは 2 つの同じパラメータを受け取るメソッドです。これらのメソッドは、`RangeAreas.areas` コレクションのすべての範囲から対象となるセルを返します。`Range` オブジェクトではなく`RangeAreas` オブジェクトを呼び出した場合のメソッドの動作には 1 つの小さな違いがあります。「SameConditionalFormat」を最初のパラメータとして渡すと、メソッドは、*`RangeAreas.areas` コレクション内の最初の範囲の*上の左端のセルと同じ条件付き書式を持つすべてのセルを返します。同じポイントが、「SameDataValidation」にも当てはまります。`Range.getSpecialCells`に渡される場合、 *範囲内の*上の左端のセルと同じデータの入力規則をもつすべてのセルを返します。しかし、`RangeAreas.getSpecialCells` に渡される場合、*`RangeAreas.areas` コレクション内の最初の範囲の*上の左端のセルと同じデータの入力規則を持つすべてのセルを返します。

## <a name="read-properties-of-rangeareas"></a>RangeAreas のプロパティの読み取り

`RangeAreas`のプロパティ値の読み取りには、`RangeAreas`内の指定したプロパティの別の範囲内の値が異なる可能性があるため、注意が必要です。一般的な規則では、一貫性のある値を返すことが*できる*場合は、返されます。例えば、次のコードでは、ピンク色の (`#FFC0CB`) と`true`用のRGBコードは、`RangeAreas`オブジェクトの両方の範囲がピンク色で塗りつぶされ、両方が全列であるため、コンソールに格納されます。

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

整合性が可能でない場合、より複雑になります。`RangeAreas`プロパティのビヘイビアーは、次の 3 つの原則に従います。

- `RangeAreas`オブジェクトのブール値プロパティは、すべてのメンバーの範囲が真でない限り、`false`を返します。
- 非ブール値プロパティは、 `address` プロパティ例外を除いて、全てのメンバーの範囲に対応するプロパティが同じ値を持っていない限り、 `null` を返します。
- `address`プロパティは、メンバーの範囲のアドレスのコンマ区切りの文字列を返します。

たとえば、次のコードは、1 つだけ列全体であり、 1 つだけがピンク色で塗りつぶされる`RangeAreas`を生成します。コンソールが、塗りつぶしの色に`null`を、`isEntireRow`プロパティに`false`を、および「Sheet1!F3:F5, Sheet1!H:H」(シート名は「Sheet1」と仮定) を`address`プロパティに表示します。 

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

- [Excel の JavaScript API を使用した基本的なプログラミングの概念](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [Range Object (Excel 向け JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.range)
- [RangeAreas Object (Excel 向け JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (API がプレビュー中の場合、このリンクが動作しない 可能性があります。代わりに[ベータ版 office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) を参照してください。)