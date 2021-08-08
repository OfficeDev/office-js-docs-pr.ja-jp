---
title: JavaScript API を使用してセルExcel作業します。
description: セルのExcel JavaScript API 定義について説明し、セルを使用する方法について説明します。
ms.date: 04/16/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 444feecd4aafb0e884de05b2ff198a3ca1423a16644c537865bcfb6905684a40
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57079347"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a>JavaScript API を使用してセルExcelする

Excel JavaScript API には、"Cell" オブジェクトまたはクラスがありません。 代わりに、すべてのセルExcelオブジェクト `Range` です。 Excel UI の個々のセルは、Excel JavaScript API の 1 つのセルを持つ `Range` オブジェクトに変換されます。

オブジェクト `Range` には、複数の連続するセルを含め、複数のセルを含めできます。 連続するセルは、(単一の行または列を含む) 未結合の四角形を形成します。 連続していないセルの操作については、「RangeAreas オブジェクトを使用して不連続セルを操作する [」を参照してください](#work-with-discontiguous-cells-using-the-rangeareas-object)。

オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Range Object (JavaScript API for Excel)」を参照してください](/javascript/api/excel/excel.range)。

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a>RangeAreas オブジェクトを使用して不一視セルを使用する

[RangeAreas オブジェクトを](/javascript/api/excel/excel.rangeareas)使用すると、アドインは複数の範囲に対して一度に操作を実行できます。 これらの範囲は連続している可能性がありますが、必要はありません。 `RangeAreas` については、「[Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)」にさらに詳しい説明があります。

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用して範囲Excelする](excel-add-ins-ranges-get.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
