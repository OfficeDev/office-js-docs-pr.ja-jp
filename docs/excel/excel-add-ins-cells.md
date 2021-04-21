---
title: Excel JavaScript API を使用してセルを使用します。
description: セルの Excel JavaScript API 定義について説明し、セルを使用する方法について説明します。
ms.date: 04/16/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad8ca985b6bbdcf19920c36c371e690f61639f16
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917101"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してセルを使用する

Excel JavaScript API には、"Cell" オブジェクトまたはクラスがありません。 代わりに、すべての Excel セルはオブジェクト `Range` です。 Excel UI の個々のセルは、Excel JavaScript API の 1 つのセルを持つ `Range` オブジェクトに変換されます。

オブジェクト `Range` には、複数の連続するセルを含め、複数のセルを含めできます。 連続するセルは、(単一の行または列を含む) 未結合の四角形を形成します。 連続していないセルの操作については、「RangeAreas オブジェクトを使用して不連続セルを操作する [」を参照してください](#work-with-discontiguous-cells-using-the-rangeareas-object)。

オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Range Object (JavaScript API for Excel)」を参照してください](/javascript/api/excel/excel.range)。

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a>RangeAreas オブジェクトを使用して不一視セルを使用する

[RangeAreas オブジェクトを](/javascript/api/excel/excel.rangeareas)使用すると、アドインは複数の範囲に対して一度に操作を実行できます。 これらの範囲は連続している可能性がありますが、必要はありません。 `RangeAreas` については、「[Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)」にさらに詳しい説明があります。

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [Excel JavaScript API を使用して範囲を取得する](excel-add-ins-ranges-get.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
