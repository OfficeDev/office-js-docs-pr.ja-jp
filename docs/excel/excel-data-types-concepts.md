---
title: Excel JavaScript API データ型の主要概念
description: Office アドインで Excel データ型を使用するための主要概念について説明します。
ms.date: 11/08/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 6155805245b14d3c3365d759bcd647419266f499
ms.sourcegitcommit: 3d37c42f5e465dac52d231d31717bdbb3bfa0e30
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/10/2021
ms.locfileid: "60889980"
---
# <a name="excel-data-types-core-concepts-preview"></a>Excel データ型の主要概念 (プレビュー)

> [!NOTE]
> 現在、データ型 API はパブリック プレビューでのみ使用できます。 プレビュー API は変更されることがあります。運用環境での使用は意図されていません。 運用環境やビジネス上重要なドキュメントでプレビュー API を使用しないでください。

> [!IMPORTANT]
> `Range.valuesAsJSON`のような、この記事で説明するデータ型の概念の一部は、アクティブな開発中であり、パブリック プレビューではまだ利用できません。 この記事は、概念的な紹介を目的としています。 この記事で説明されている、パブリック プレビューにはまだ含まれていない概念は、間もなくプレビューにリリースされる予定です。

この記事では、[Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) を使用してデータ型を操作する方法について説明します。 ここでは、データ型の開発の基本となる主要な概念を紹介します。

## <a name="core-concepts"></a>中心概念

データ型の値を操作するには、`Range.valuesAsJSON` プロパティを使用します。 このプロパティは [Range.values](/javascript/api/excel/excel.range#values) に似ていますが、`Range.values`は文字列、数値、ブール値、エラー値の 4 つの基本型のみを返します。 `Range.valuesAsJSON`4 つの基本型に関する拡張情報を返すことができます。このプロパティは、書式設定された数値、エンティティ、Web イメージなどのデータ型を返すことができます。

### <a name="json-schema"></a>JSON スキーマ

データ型は、データの [CellValueType](/javascript/api/excel/excel.cellvaluetype) と、 `basicValue`、 `numberFormat`、 `address`などの追加情報を定義する一貫性のある JSON スキーマを使用します。 各`CellValueType`は、その型によって使用可能なプロパティがあります。 たとえば、 `webImage` の種類には、[altText](/javascript/api/excel/excel.webimagecellvalue#altText) と [属性](/javascript/api/excel/excel.webimagecellvalue#attribution) プロパティが含まれます。 次のセクションでは、書式設定された数値、エンティティ値、および Web 画像データ型の JSON コード サンプルを示します。

## <a name="formatted-number-values"></a>書式設定された数値

[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) オブジェクトを使用すると、Excel アドインで値用の`numberFormat`プロパティを定義できます。 割り当てられると、この数値形式は値を使用して計算を通過し、関数から返すことができます。

次の JSON コード サンプルは、書式設定された数値を示しています。 コード サンプルの `myDate`書式設定された数値は、Excel UI で **1/16/1990** と表示されます。

```json
// This is an example of the JSON of a formatted number value.
// In this case, the number is formatted as a date.
const myDate = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    numberFormat: "m/d/yyyy"
};
```

## <a name="entity-values"></a>エンティティの値:

エンティティ値は、オブジェクト指向プログラミングのオブジェクトと同様に、データ型のコンテナーです。 エンティティは、エンティティ値のプロパティとして配列もサポートします。 [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) オブジェクトを使用すると、アドインは `type`、`text`、`properties`などのプロパティを定義できます。 `properties` プロパティを使用すると、エンティティ値で追加のデータ型を定義および格納できます。

次の JSON コード サンプルは、テキスト、画像、日付、および追加のテキスト値を含むエンティティ値を示しています。

```json
// This is an example of the JSON for an entity value.
// The entity contains text and properties which contain an image, a date, and another text value.
const myEntity = {
    type: Excel.CellValueType.entity,
    text: "A llama",
    properties: {
        image: myImage,
        "start date": myDate,
        "quote": {
            type: Excel.CellValueType.string,
            basicValue: "I love llamas."
        }
    }
};
```

## <a name="web-image-values"></a>Web 画像の値

[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) オブジェクトは、[エンティティ](#entity-values)の一部として、または範囲内の独立した値として画像を格納する機能を作成します。 このオブジェクトには、`address`、`altText`、 `relatedImagesAddress` など、多くのプロパティが用意されています。

次の JSON コード サンプルは、Web イメージを表す方法を示しています。

```json
// This is an example of the JSON for a web image.
const myImage = {
    type: Excel.CellValueType.webImage,
    address: "https://bit.ly/2YGOwtw"
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
- [Excel JavaScript API リファレンス](../reference/overview/excel-add-ins-reference-overview.md)
- [カスタム関数とデータ型の概要](custom-functions-data-types-overview.md)