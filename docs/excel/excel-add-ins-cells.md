---
title: Excel JavaScript API を使用してセルを使用します。
description: セルの Excel JavaScript API 定義について説明し、セルを使用する方法について説明します。
ms.date: 04/07/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 5fcfeeef52f17c22d13ed3c1a10851f1d8e69204
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652977"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してセルを使用する

Excel JavaScript API には、"Cell" オブジェクトまたはクラスが含か指定されています。 代わりに、すべての Excel セルはオブジェクト `Range` です。 Excel UI の個々のセルは、Excel JavaScript API で 1 つのセルを持つオブジェクト `Range` に変換されます。

オブジェクト `Range` には、複数の連続するセルを含め、複数のセルを含めできます。 連続するセルは、(単一の行または列を含む) 未結合の四角形を形成します。 連続していないセルの操作については、「RangeAreas オブジェクトを使用して不連続セルを操作する [」を参照してください](#work-with-discontiguous-cells-using-the-rangeareas-object)。

オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel.Range クラス」を参照してください](/javascript/api/excel/excel.range)。

## <a name="excel-javascript-apis-that-mention-cells"></a>セルを言及する Excel JavaScript API

Excel JavaScript API に "Cell" オブジェクトまたはクラスがない場合でも、多数の API 名がセルを表します。 これらの API は、色、テキストの書式設定、フォントなど、セルのプロパティを制御します。

Excel JavaScript API の次のリストは、セルを参照します。

- [CellBorder](/javascript/api/excel/excel.cellborder)
- [CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)
- [CellProperties](/javascript/api/excel/excel.cellproperties)
- [CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)
- [CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)
- [CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)
- [CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)
- [CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)
- [ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)
- [SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a>RangeAreas オブジェクトを使用して不一視セルを使用する

[RangeAreas オブジェクトを](/javascript/api/excel/excel.rangeareas)使用すると、アドインは複数の範囲に対して一度に操作を実行できます。 これらの範囲は連続している可能性がありますが、必要はありません。 `RangeAreas` については、「[Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)」にさらに詳しい説明があります。

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [Excel JavaScript API を使用して範囲を取得する](excel-add-ins-ranges-get.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
