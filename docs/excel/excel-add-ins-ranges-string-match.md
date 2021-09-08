---
title: JavaScript API を使用して文字列をExcelする
description: JavaScript API を使用して範囲内の文字列を検索するExcel説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9b649bb249cd24d7578bc4f8285e5d0a23d0e4cd
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937972"
---
# <a name="find-a-string-within-a-range-using-the-excel-javascript-api"></a>JavaScript API を使用して範囲内の文字列をExcelする

この記事では、JavaScript API を使用して範囲内の文字列を検索するコード サンプルExcel示します。 オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel。Range クラス](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="match-a-string-within-a-range"></a>範囲内の文字列と一致する

`Range` オブジェクトには、範囲内で指定された文字列を検索するための `find` メソッドがあります。 このメソッドは、一致するテキストがある最初のセルの範囲を返します。

次のコード サンプルは、文字列 **Food** と等しい値を持つ最初のセルを検索して、そのアドレスをコンソールに記録します。 指定した文字列が範囲に存在しない場合、`ItemNotFound` エラーが `find` によってスローされます。 指定した文字列が範囲に存在しない可能性がある場合は、自分のコードで適切にシナリオを処理できるように、[findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) メソッドを使用するようにしてください。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var table = sheet.tables.getItem("ExpensesTable");
    var searchRange = table.getRange();
    var foundRange = searchRange.find("Food", {
        completeMatch: true, // find will match the whole cell value
        matchCase: false, // find will not match case
        searchDirection: Excel.SearchDirection.forward // find will start searching at the beginning of the range
    });

    foundRange.load("address");
    return context.sync()
        .then(function() {
            console.log(foundRange.address);
    });
}).catch(errorHandlerFunction);
```

単一のセルを表す範囲に対して `find` メソッドが呼び出されると、ワークシート全体が検索されます。 検索はその単一のセルから始まり、`SearchCriteria.searchDirection` によって指定された方向へ行われ、場合によってはワークシートの最終部分で折り返されます。

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [JavaScript API を使用して範囲内の特別なセルExcel検索する](excel-add-ins-ranges-special-cells.md)
