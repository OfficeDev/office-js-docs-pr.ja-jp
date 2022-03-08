---
title: JavaScript API を使用して、非バウンド範囲に対する読み取りExcel書き込み
description: JavaScript API を使用して、Excel範囲を読み取りまたは書き込みする方法について説明します。
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 6e9b0c56dfd04cd53e01c41fea23fbf826a6fa14
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340954"
---
# <a name="read-or-write-to-an-unbounded-range-using-the-excel-javascript-api"></a>JavaScript API を使用して、非バウンド範囲に対する読み取りExcel書き込み

この記事では、JavaScript API を使用して、非バウンド範囲に対する読み取りExcel説明します。 オブジェクトがサポートするプロパティとメソッドの`Range`完全な一覧については、「Excel[。Range クラス](/javascript/api/excel/excel.range)。

非バウンド範囲アドレスは、列全体または行全体を指定する範囲アドレスです。 例:

- 列全体で構成される範囲アドレス。
  - `C:C`
  - `A:F`
- 行全体で構成される範囲アドレス。
  - `2:2`
  - `1:4`

## <a name="read-an-unbounded-range"></a>無制限の範囲の読み取り

API が無制限の範囲を取得する要求を行う場合 (`getRange('C:C')` など)、返される応答では、`null`、`values`、`text`、または `numberFormat` などのセル レベルのプロパティに `formula` 値が含まれます。 `address` または `cellCount` など、範囲のその他のプロパティには、無制限の範囲に有効な値が含まれます。

## <a name="write-to-an-unbounded-range"></a>無制限の範囲への書き込み

入力要求が大きすぎる`values`ため、セル レベルのプロパティ (`numberFormat``formula`、 など) を非バウンド範囲に設定することはできません。 たとえば、次のコード例は `values` 、非バウンド範囲を指定しようとして無効です。 非バウンド範囲のセル レベルのプロパティを設定しようとすると、API はエラーを返します。

```js
// Note: This code sample attempts to specify `values` for an unbounded range, which is not a valid request. The sample will return an error. 
let range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [JavaScript API を使用した大きな範囲の読み取りExcel書き込み](excel-add-ins-ranges-large.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
