---
title: JavaScript API を使用して文字列をExcelする
description: JavaScript API を使用して範囲内の文字列を検索するExcel説明します。
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 042465e01af55bbb3f4325ea44edc27174d558f2
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340989"
---
# <a name="find-a-string-within-a-range-using-the-excel-javascript-api"></a>JavaScript API を使用して範囲内の文字列をExcelする

この記事では、JavaScript API を使用して範囲内の文字列を検索するコード サンプルExcel示します。 オブジェクトがサポートするプロパティとメソッドの`Range`完全な一覧については、「Excel[。Range クラス](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="match-a-string-within-a-range"></a>範囲内の文字列と一致する

`Range` オブジェクトには、範囲内で指定された文字列を検索するための `find` メソッドがあります。 このメソッドは、一致するテキストがある最初のセルの範囲を返します。

次のコード サンプルは、文字列 **Food** と等しい値を持つ最初のセルを検索して、そのアドレスをコンソールに記録します。 指定した文字列が範囲に存在しない場合、`ItemNotFound` エラーが `find` によってスローされます。 指定した文字列が範囲に存在しない可能性がある場合は、自分のコードで適切にシナリオを処理できるように、[findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) メソッドを使用するようにしてください。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let table = sheet.tables.getItem("ExpensesTable");
    let searchRange = table.getRange();
    let foundRange = searchRange.find("Food", {
        completeMatch: true, // Match the whole cell value.
        matchCase: false, // Don't match case.
        searchDirection: Excel.SearchDirection.forward // Start search at the beginning of the range.
    });

    foundRange.load("address");
    await context.sync();

    console.log(foundRange.address);
});
```

単一のセルを表す範囲に対して `find` メソッドが呼び出されると、ワークシート全体が検索されます。 検索はその単一のセルから始まり、`SearchCriteria.searchDirection` によって指定された方向へ行われ、場合によってはワークシートの最終部分で折り返されます。

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [JavaScript API を使用して範囲内の特別なセルExcel検索する](excel-add-ins-ranges-special-cells.md)
