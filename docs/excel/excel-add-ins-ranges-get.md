---
title: Excel JavaScript API を使用して範囲を取得する
description: Excel JavaScript API を使用して範囲を取得する方法について説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6aa9bb00bc9d24aeee5f1fef9e8d1531525e9d1f
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652928"
---
# <a name="get-a-range-using-the-excel-javascript-api"></a>Excel JavaScript API を使用して範囲を取得する

この記事では、Excel JavaScript API を使用してワークシート内の範囲を取得するさまざまな方法を示す例を示します。 オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel.Range クラス」を参照してください](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="get-range-by-address"></a>アドレスによって範囲を取得する

次のコード サンプルは **、Sample** という名前のワークシートから **アドレス B2:C5** の範囲を取得し、そのプロパティを読み込み、コンソールに `address` メッセージを書き込みます。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:C5");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range B2:C5 is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-range-by-name"></a>名前によって範囲を取得する

次のコード サンプルは、Sample という名前のワークシートからという名前の範囲を取得し、そのプロパティを読み込み、コンソール `MyRange`  `address` にメッセージを書き込みます。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("MyRange");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range "MyRange" is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-used-range"></a>使用範囲を取得する

次のコード サンプルは **、Sample** という名前のワークシートから使用範囲を取得し、そのプロパティを読み込み、コンソール `address` にメッセージを書き込みます。 使用範囲とは、値または書式設定が割り当てられているワークシート内のセルを含む、最小の範囲です。 ワークシート全体が空白の場合、メソッドは左上のセルだけで構成される `getUsedRange()` 範囲を返します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getUsedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the used range in the worksheet is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-entire-range"></a>範囲全体を取得する

次のコード サンプルは **、Sample** という名前のワークシートからワークシートの範囲全体を取得し、そのプロパティを読み込み、コンソール `address` にメッセージを書き込みます。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the entire worksheet range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [Excel JavaScript API を使用してセルを使用する](excel-add-ins-cells.md)
- [Excel JavaScript API を使用して範囲を挿入する](excel-add-ins-ranges-insert.md)
