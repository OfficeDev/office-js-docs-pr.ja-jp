---
title: Excel JavaScript API を使用して、非バウンド範囲に対する読み取りまたは書き込み
description: Excel JavaScript API を使用して、非バウンド範囲への読み取りまたは書き込みを行う方法について説明します。
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f7be2efc3e069ea3451088608ca5255a632ef863
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652886"
---
# <a name="read-or-write-to-an-unbounded-range-using-the-excel-javascript-api"></a>Excel JavaScript API を使用して、非バウンド範囲に対する読み取りまたは書き込み

この記事では、Excel JavaScript API を使用して、非バウンド範囲の読み取りおよび書き込みを行う方法について説明します。 オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel.Range クラス」を参照してください](/javascript/api/excel/excel.range)。

非バウンド範囲アドレスは、列全体または行全体を指定する範囲アドレスです。 次に例を示します。

- 列全体で構成される範囲アドレス:<ul><li>`C:C`</li><li>`A:F`</li></ul>
- 行全体で構成される範囲アドレス:<ul><li>`2:2`</li><li>`1:4`</li></ul>

## <a name="read-an-unbounded-range"></a>無制限の範囲の読み取り

API が無制限の範囲を取得する要求を行う場合 (`getRange('C:C')` など)、返される応答では、`null`、`values`、`text`、または `numberFormat` などのセル レベルのプロパティに `formula` 値が含まれます。 `address` または `cellCount` など、範囲のその他のプロパティには、無制限の範囲に有効な値が含まれます。

## <a name="write-to-an-unbounded-range"></a>無制限の範囲への書き込み

入力要求が大きすぎるため、セル レベルのプロパティ (、 など) を非バウンド `values` `numberFormat` `formula` 範囲に設定することはできません。 たとえば、次のコード例は、非バウンド範囲を指定しようとして `values` 無効です。 非バウンド範囲のセル レベルのプロパティを設定しようとすると、API はエラーを返します。

```js
// Note: This code sample attempts to specify `values` for an unbounded range, which is not a valid request. The sample will return an error. 
var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [Excel JavaScript API を使用してセルを使用する](excel-add-ins-cells.md)
- [Excel JavaScript API を使用して大きな範囲に対する読み取りまたは書き込み](excel-add-ins-ranges-large.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
