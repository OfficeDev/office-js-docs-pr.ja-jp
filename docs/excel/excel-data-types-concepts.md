---
title: Excel JavaScript API データ型の主要概念
description: Office アドインで Excel データ型を使用するための主要概念について説明します。
ms.date: 09/01/2022
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: a5f742e47d698b215a999b966c424819e182ea49
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616022"
---
# <a name="excel-data-types-core-concepts-preview"></a>Excel データ型の主要概念 (プレビュー)

[!include[Data types preview availability note](../includes/excel-data-types-preview.md)]

この記事では、[Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) を使用してデータ型を操作する方法について説明します。 ここでは、データ型の開発の基本となる主要な概念を紹介します。

## <a name="the-valuesasjson-property"></a>`valuesAsJson` プロパティ

`valuesAsJson`プロパティ ([NamedItem](/javascript/api/excel/excel.nameditem) の場合は単数形`valueAsJson`) は、Excel でのデータ型の作成に不可欠です。 このプロパティは、[Range.values](/javascript/api/excel/excel.range#excel-excel-range-values-member) などの `values` プロパティの拡張です。 `values` と `valuesAsJson` プロパティはどちらもセル内の値にアクセスするに使用しますが、`values` プロパティは、文字列、数値、ブール値、またはエラーの 4 つの基本型の 1 つだけを返します (文字列として)。 一方、`valuesAsJson` は、4 つの基本型に関する拡張情報を返します。このプロパティは、書式設定された数値、エンティティ、Web イメージなどのデータ型を返すことができます。

次のオブジェクトは、`valuesAsJson` プロパティを提供します。

- [NamedItem](/javascript/api/excel/excel.nameditem) (as `valueAsJson`)
- [NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)
- [Range](/javascript/api/excel/excel.range)
- [RangeView](/javascript/api/excel/excel.rangeview)
- [TableColumn](/javascript/api/excel/excel.tablecolumn)
- [TableRow](/javascript/api/excel/excel.tablerow)

> [!NOTE]
> 一部のセル値は、ユーザーのロケールに基づいて変化します。 `valuesAsJsonLocal` プロパティはローカライズのサポートを提供し、`valuesAsJson` などのオブジェクトで利用可能です。

## <a name="cell-values"></a>セルの値

この`valuesAsJson` プロパティは、[CellValue](/javascript/api/excel/excel.cellvalue)型エイリアスを返します。これは、次のデータ型の[共用体](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#union-types)です。

- [ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)
- [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)
- [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)
- [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)
- [EntityCellValue](/javascript/api/excel/excel.entitycellvalue)
- [ErrorCellValue](/javascript/api/excel/excel.errorcellvalue)
- [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)
- [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)
- [ReferenceCellValue](/javascript/api/excel/excel.referencecellvalue)
- [StringCellValue](/javascript/api/excel/excel.stringcellvalue)
- [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)
- [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)

`CellValue` 型エイリアスは、[CellValueExtraProperties](/javascript/api/excel/excel.cellvalueextraproperties) オブジェクトも返します。これは、他の `*CellValue` 型との [積集合](https://www.typescriptlang.org/docs/handbook/2/objects.html#intersection-types) 部分です。 データ型自体ではありません。 `CellValueExtraProperties` オブジェクトのプロパティは、セル値の上書きに関連する詳細を指定するために、すべてのデータ型で使用されます。

### <a name="json-schema"></a>JSON スキーマ

`valuesAsJson` から返されたセルの値の型は、その型用に設計された JSON メタデータ スキーマを使用します。 各データ型に固有の追加のプロパティと共に、これらの JSON メタデータ スキーマには、共通の `type`、`basicType`、`basicValue` プロパティがあります。

`type` はデータの [CellValueType](/javascript/api/excel/excel.cellvaluetype) を定義します。 これは `basicType` 常に読み取り専用であり、データ型がサポートされていない場合や正しく書式設定されていない場合にフォールバックとして使用されます。 `basicValue` は `values` プロパティで返される値と一致します。 `basicValue` は、データ型機能をサポートしていない以前のバージョンの Excel など、計算で互換性のないシナリオが発生した場合にフォールバックとして使用されます。 これは`basicValue`、、`EntityCellValue``LinkedEntityCellValue`および`WebImageCellValue`データ型の`ArrayCellValue`読み取り専用です。

すべてのデータ型が共有する 3 つのフィールドに加えて、それぞれの `*CellValue` JSON メタデータ スキーマには、その型に従って使用可能なプロパティがあります。 たとえば、[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) 型には `altText` と `attribution`プロパティが含まれますが、[EntityCellValue](/javascript/api/excel/excel.entitycellvalue) 型には `properties` と `text` フィールドが用意されています。

次のセクションでは、書式設定された数値、エンティティ値、および Web 画像データ型の JSON コード サンプルを示します。

## <a name="formatted-number-values"></a>書式設定された数値

[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) オブジェクトを使用すると、Excel アドインで値用の`numberFormat`プロパティを定義できます。 割り当てられると、この数値形式は値を使用して計算を通過し、関数から返すことができます。

次の JSON コード サンプルは、フォーマットされた数値の完全なスキーマを示しています。 コード サンプルの `myDate`書式設定された数値は、Excel UI で **1/16/1990** と表示されます。 データ型機能の最小互換性要件が満たされていない場合、計算では、フォーマットされた数値の代わりに `basicValue` が使用されます。

```TypeScript
// This is an example of the complete JSON of a formatted number value.
// In this case, the number is formatted as a date.
const myDate: Excel.FormattedNumberCellValue = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    basicType: Excel.RangeValueType.double, // A read-only property. Used as a fallback in incompatible scenarios.
    numberFormat: "m/d/yyyy"
};
```

## <a name="entity-values"></a>エンティティの値:

エンティティ値は、オブジェクト指向プログラミングのオブジェクトと同様に、データ型のコンテナーです。 エンティティは、エンティティ値のプロパティとして配列もサポートします。 [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) オブジェクトを使用すると、アドインは `type`、`text`、`properties`などのプロパティを定義できます。 `properties` プロパティを使用すると、エンティティ値で追加のデータ型を定義および格納できます。

`basicType` プロパティと `basicValue` プロパティは、データ型を使用するための最小互換性要件が満たされていない場合に、計算がこのエンティティ データ型を読み取る方法を定義します。 そのシナリオでは、このエンティティ データ型は **#VALUE!** として表示されます。 Excel UI のエラー。

次の JSON コード サンプルは、テキスト、画像、日付、および追加のテキスト値を含むエンティティ値の完全なスキーマを示しています。

```TypeScript
// This is an example of the complete JSON for an entity value.
// The entity contains text and properties which contain an image, a date, and another text value.
const myEntity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: "A llama",
    properties: {
        image: myImage,
        "start date": myDate,
        "quote": {
            type: Excel.CellValueType.string,
            basicValue: "I love llamas."
        }
    }, 
    basicType: Excel.RangeValueType.error, // A read-only property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A read-only property. Used as a fallback in incompatible scenarios.
};
```

エンティティ値には、エンティティのカードを作成する `layouts` プロパティも用意されています。 カードは、Excel UI のモーダル ウィンドウとして表示され、セルに表示される内容を超えて、エンティティ値に含まれる追加情報を表示できます。 詳細については、「[エンティティ値データ型でカードを使用する](excel-data-types-entity-card.md)」を参照してください。

### <a name="linked-entities"></a>リンクされたエンティティ

リンクされたエンティティ値 または [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue) オブジェクトは、エンティティ値の種類です。 これらのオブジェクトは外部サービスによって提供されるデータを統合し、このデータを通常のエンティティ値のように [エンティティ カード](excel-data-types-entity-card.md) として表示できます。 Excel UI で使用できる [株価と地理データ型](https://support.microsoft.com/office/excel-data-types-stocks-and-geography-61a33056-9935-484f-8ac8-f1a89e210877) はリンクされたエンティティ値です。

## <a name="web-image-values"></a>Web 画像の値

[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) オブジェクトは、[エンティティ](#entity-values)の一部として、または範囲内の独立した値として画像を格納する機能を作成します。このオブジェクトには、`address`、`altText`、`relatedImagesAddress` など、多くのプロパティが用意されています。

`basicType` および `basicValue` プロパティは、データ型機能を使用するための最小互換性要件が満たされていない場合に、計算が Web 画像データ型を読み取る方法を定義します。 そのシナリオでは、この Web 画像データ型は **#VALUE!** として表示されます。 Excel UI のエラー。

次の JSON コード サンプルは、Web イメージの完全なスキーマを示しています。

```TypeScript
// This is an example of the complete JSON for a web image.
const myImage: Excel.WebImageCellValue = {
    type: Excel.CellValueType.webImage,
    address: "https://bit.ly/2YGOwtw", 
    basicType: Excel.RangeValueType.error, // A read-only property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A read-only property. Used as a fallback in incompatible scenarios.
};
```

## <a name="improved-error-support"></a>エラー サポートの改善

データ型 API は、既存の Excel UI エラーをオブジェクトとして公開します。 これらのエラーにオブジェクトとしてアクセスできるようになったので、アドインは `type`、`errorType`、 `errorSubType` などのプロパティを定義または取得できます。

データ型を介してサポートが拡張されたすべてのエラー オブジェクトの一覧を次に示します。

- [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)
- [BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)
- [CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)
- [ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)
- [Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)
- [FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)
- [GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)
- [NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)
- [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)
- [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)
- [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)
- [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)
- [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)
- [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)

各エラー オブジェクトは、`errorSubType` プロパティを使用して列挙型にアクセスでき、この列挙型にはエラーに関する追加のデータが含まれています。 たとえば、`BlockedErrorCellValue` エラー オブジェクトは、[BlockedErrorCellValueSubType](/javascript/api/excel/excel.blockederrorcellvaluesubtype) 列挙型にアクセスできます。 `BlockedErrorCellValueSubType`enum は、エラーの原因に関する追加データを提供します。

## <a name="see-also"></a>関連項目

- [Excel アドインのデータ型の概要](excel-data-types-overview.md)
- [エンティティ値データ型でカードを使用する](excel-data-types-entity-card.md)
- [Excel JavaScript API リファレンス](../reference/overview/excel-add-ins-reference-overview.md)
- [カスタム関数とデータ型](custom-functions-data-types-concepts.md)
