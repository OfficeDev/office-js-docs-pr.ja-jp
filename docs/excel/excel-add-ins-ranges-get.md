---
title: JavaScript API を使用して範囲Excelする
description: JavaScript API を使用して範囲を取得するExcel説明します。
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 16c42ccf8f3496316fbf7b52e4d8139f819c6da1
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340940"
---
# <a name="get-a-range-using-the-excel-javascript-api"></a>JavaScript API を使用して範囲Excelする

この記事では、JavaScript API を使用してワークシート内の範囲を取得するさまざまな方法Excel示します。 オブジェクトがサポートするプロパティとメソッドの`Range`完全な一覧については、「Excel[。Range クラス](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="get-range-by-address"></a>アドレスによって範囲を取得する

次のコード サンプルは、**Sample**`address` という名前のワークシートから **アドレス B2:C5** の範囲を取得し、そのプロパティを読み込み、コンソールにメッセージを書き込みます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    
    let range = sheet.getRange("B2:C5");
    range.load("address");
    await context.sync();
    
    console.log(`The address of the range B2:C5 is "${range.address}"`);
});
```

## <a name="get-range-by-name"></a>名前によって範囲を取得する

次のコード サンプルは、**Sample**`address` という名前`MyRange`のワークシートからという名前の範囲を取得し、そのプロパティを読み込み、コンソールにメッセージを書き込みます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("MyRange");
    range.load("address");
    await context.sync();

    console.log(`The address of the range "MyRange" is "${range.address}"`);
});
```

## <a name="get-used-range"></a>使用範囲を取得する

次のコード サンプルは、**Sample**`address` という名前のワークシートから使用範囲を取得し、そのプロパティを読み込み、コンソールにメッセージを書き込みます。 使用範囲とは、値または書式設定が割り当てられているワークシート内のセルを含む、最小の範囲です。 ワークシート全体が空白の場合、 `getUsedRange()` メソッドは左上のセルだけで構成される範囲を返します。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getUsedRange();
    range.load("address");
    await context.sync();
    
    console.log(`The address of the used range in the worksheet is "${range.address}"`);
});
```

## <a name="get-entire-range"></a>範囲全体を取得する

次のコード サンプルは、**Sample**`address` という名前のワークシートからワークシートの範囲全体を取得し、そのプロパティを読み込み、コンソールにメッセージを書き込みます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange();
    range.load("address");
    await context.sync();
    
    console.log(`The address of the entire worksheet range is "${range.address}"`);
});
```

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [JavaScript API を使用して範囲Excel挿入する](excel-add-ins-ranges-insert.md)
