---
title: Excel JavaScript API を使用して範囲をグループ化する
description: Excel JavaScript API を使用して範囲の行または列をグループ化してアウトラインを作成する方法について説明します。
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 32f65cf88c23bd6368b37318d3ba20fde95b8436
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652923"
---
# <a name="group-ranges-for-an-outline-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してアウトラインの範囲をグループ化する

この記事では、Excel JavaScript API を使用してアウトラインの範囲をグループ化する方法を示すコード サンプルを提供します。 オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel.Range クラス」を参照してください](/javascript/api/excel/excel.range)。

## <a name="group-rows-or-columns-of-a-range-for-an-outline"></a>アウトラインの範囲の行または列をグループ化する

範囲の行または列をグループ化してアウトラインを作成 [できます](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF)。 これらのグループを折りたたみ、展開して、対応するセルを非表示にし、表示できます。 これにより、トップライン データの迅速な分析が容易になります。 [Range.group を使用して](/javascript/api/excel/excel.range#group-groupoption-)、これらのアウトライン グループを作成します。

アウトラインには階層を含め、小さなグループは大きなグループの下に入れ子にできます。 これにより、アウトラインをさまざまなレベルで表示できます。 表示されるアウトライン レベルを変更するには [、Worksheet.showOutlineLevels](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-) メソッドを使用してプログラムを使用します。 Excel は 8 つのレベルのアウトライン グループのみをサポートしています。

次のコード サンプルでは、行と列の両方に 2 つのレベルのグループを含むアウトラインを作成します。 次の図は、そのアウトラインのグループ化を示しています。 コード サンプルでは、グループ化されている範囲にアウトライン コントロールの行または列 (この例の "Totals") は含めされません。 グループは、コントロールの行または列ではなく、折りたたむものを定義します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // will be on row 10, meaning 4-9 will collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on rows 6 and 9, meaning 4-5 and 7-8 will collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // will be on column R, meaning C-Q will collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on columns G, L, and R, meaning C-F, H-K, and M-P will collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    return context.sync();
}).catch(errorHandlerFunction);
```

![2 つのレベルの 2 次元アウトラインを持つ範囲](../images/excel-outline.png)

## <a name="remove-grouping-from-rows-or-columns-of-a-range"></a>範囲の行または列からグループ化を削除する

行または列グループのグループ化を解除するには [、Range.ungroup メソッドを使用](/javascript/api/excel/excel.range#ungroup-groupoption-) します。 これにより、アウトラインから最も外側のレベルが削除されます。 同じ行または列の種類の複数のグループが指定した範囲内で同じレベルにある場合、それらのグループはすべてグループ化解除されます。

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [Excel JavaScript API を使用してセルを使用する](excel-add-ins-cells.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
