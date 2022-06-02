---
title: Excel JavaScript API データ型の主要概念
description: Office アドインで Excel データ型を使用するための主要概念について説明します。
ms.date: 05/26/2022
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 2259d28bc87e6452e526786c0b32135e4bb27d45
ms.sourcegitcommit: 35e7646c5ad0d728b1b158c24654423d999e0775
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/02/2022
ms.locfileid: "65833907"
---
# <a name="excel-data-types-core-concepts-preview"></a>Excel データ型の主要概念 (プレビュー)

[!include[Data types preview availability note](../includes/excel-data-types-preview.md)]

この記事では、[Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) を使用してデータ型を操作する方法について説明します。 ここでは、データ型の開発の基本となる主要な概念を紹介します。

## <a name="core-concepts"></a>中心概念

データ型の値を操作するには、[`Range.valuesAsJson`](/javascript/api/excel/excel.range#excel-excel-range-valuesasjson-member) プロパティを使用します。 このプロパティは [Range.values](/javascript/api/excel/excel.range#excel-excel-range-values-member) に似ていますが、`Range.values`は文字列、数値、ブール値、エラー値の 4 つの基本型のみを返します。 `Range.valuesAsJson` は、4 つの基本型に関する拡張情報を返します。このプロパティは、書式設定された数値、エンティティ、Web イメージなどのデータ型を返すことができます。

この`valuesAsJson` プロパティは、[CellValue](/javascript/api/excel/excel.cellvalue)型エイリアスを返します。これは、次のデータ型の[共用体](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#union-types)です。

- [ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)
- [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)
- [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)
- [EntityCellValue](/javascript/api/excel/excel.entitycellvalue)
- [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)
- [ErrorCellValue](/javascript/api/excel/excel.errorcellvalue)
- [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)
- [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)
- [ReferenceCellValue](/javascript/api/excel/excel.referencecellvalue)
- [StringCellValue](/javascript/api/excel/excel.stringcellvalue)
- [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)
- [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)

[CellValueExtraProperties](/javascript/api/excel/excel.cellvalueextraproperties) オブジェクトは、他の`*CellValue`型との[交差部分](https://www.typescriptlang.org/docs/handbook/2/objects.html#intersection-types)です。 データ型自体ではありません。 `CellValueExtraProperties` オブジェクトのプロパティは、セル値の上書きに関連する詳細を指定するために、すべてのデータ型で使用されます。

### <a name="json-schema"></a>JSON スキーマ

各データ型は、その型用に設計された JSON メタデータ スキーマを使用します。 これは、データの [CellValueType](/javascript/api/excel/excel.cellvaluetype) と `basicValue`、`numberFormat`、`address` などのセルに関する追加情報を定義します。 各`CellValueType`は、その型によって使用可能なプロパティがあります。 たとえば、 `webImage` の種類には、[altText](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-alttext-member) と [属性](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-attribution-member) プロパティが含まれます。 次のセクションでは、書式設定された数値、エンティティ値、および Web 画像データ型の JSON コード サンプルを示します。

各データ型の JSON メタデータ スキーマには、データ型機能の最小ビルド数要件を満たしていないバージョンの Excel など、計算で互換性のないシナリオが発生した場合に使用される 1 つ以上の読み取り専用プロパティも含まれます。 プロパティ `basicType` は、すべてのデータ型の JSON メタデータの一部であり、常に読み取り専用プロパティです。 `basicType` プロパティは、データ型がサポートされていないか、正しくフォーマットされていない場合のフォールバックとして使用されます。

## <a name="formatted-number-values"></a>書式設定された数値

[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) オブジェクトを使用すると、Excel アドインで値用の`numberFormat`プロパティを定義できます。 割り当てられると、この数値形式は値を使用して計算を通過し、関数から返すことができます。

次の JSON コード サンプルは、フォーマットされた数値の完全なスキーマを示しています。 コード サンプルの `myDate`書式設定された数値は、Excel UI で **1/16/1990** と表示されます。 データ型機能の最小互換性要件が満たされていない場合、計算では、フォーマットされた数値の代わりに `basicValue` が使用されます。

```TypeScript
// This is an example of the complete JSON of a formatted number value.
// In this case, the number is formatted as a date.
const myDate: Excel.FormattedNumberCellValue = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    basicType: Excel.RangeValueType.double, // A readonly property. Used as a fallback in incompatible scenarios.
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
    basicType: Excel.RangeValueType.error, // A readonly property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A readonly property. Used as a fallback in incompatible scenarios.
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
    basicType: Excel.RangeValueType.error, // A readonly property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A readonly property. Used as a fallback in incompatible scenarios.
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
