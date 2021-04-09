---
title: Excel JavaScript API を使用して数式の前例を使用する
description: Excel JavaScript API を使用して数式の前例を取得する方法について説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0d21ae411615a22873a0f4dda185984f6191ac8e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652916"
---
# <a name="get-formula-precedents-using-the-excel-javascript-api"></a>Excel JavaScript API を使用して数式の前例を取得する

この記事では、Excel JavaScript API を使用して数式の前例を取得するコード サンプルを提供します。 オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel.Range クラス」を参照してください](/javascript/api/excel/excel.range)。

## <a name="get-formula-precedents"></a>数式の前例を取得する

Excel の数式は、多くの場合、他のセルを参照します。 セルが数式にデータを提供する場合、セルは数式 "前例" と呼ばれる。 セル間のリレーションシップに関連する Excel 機能の詳細については、「数式とセル間のリレーションシップを表示する [」を参照してください](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507)。 

[Range.getDirectPrecedents を](/javascript/api/excel/excel.range#getdirectprecedents--)使用すると、アドインは数式の直接の先行セルを検索できます。 `Range.getDirectPrecedents` オブジェクトを返 `WorkbookRangeAreas` します。 このオブジェクトには、ブック内のすべての前例のアドレスが含まれます。 このオブジェクトには、少なくとも 1 つの数式の前例を含 `RangeAreas` むワークシートごとに個別のオブジェクトがあります。 オブジェクトの操作の詳細については、「Excel アドインで複数の範囲を同時に操作する `RangeAreas` [」を参照してください](excel-add-ins-multiple-ranges.md)。

Excel UI では、[前 **例のトレース** ] ボタンは、前のセルから選択した数式に矢印を描画します。 Excel UI ボタンとは異なり、 `getDirectPrecedents` メソッドは矢印を描画しない。 

> [!IMPORTANT]
> メソッド `getDirectPrecedents` は、ブック間で先行セルを取得できない。 

次のコード サンプルでは、アクティブな範囲の直接の前例を取得し、それらの前のセルの背景色を黄色に変更します。 

> [!NOTE]
> 強調表示が適切に機能するには、アクティブな範囲に同じブック内の他のセルを参照する数式が含まれている必要があります。 

```js
Excel.run(function (context) {
    // Precedents are cells that provide data to the selected formula.
    var range = context.workbook.getActiveCell();
    var directPrecedents = range.getDirectPrecedents();
    range.load("address");
    directPrecedents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct precedent cells of ${range.address}:`);

            // Use the direct precedents API to loop through precedents of the active cell.
            for (var i = 0; i < directPrecedents.areas.items.length; i++) {
              // Highlight and print out the address of each precedent cell.
              directPrecedents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directPrecedents.areas.items[i].address}`);
            }
        })
        .then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [Excel JavaScript API を使用してセルを使用する](excel-add-ins-cells.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
