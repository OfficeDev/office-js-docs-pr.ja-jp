---
title: カスタム関数から Excel JavaScript API を呼び出す
description: カスタム関数から呼び出す Excel JavaScript API について説明します。
ms.date: 03/05/2021
localization_priority: Normal
ms.openlocfilehash: 4be1b1ee8ea4ae8b2f5d1d27195be18f7aa841da
ms.sourcegitcommit: d153f6d4c3e01d63ed24aa1349be16fa8ad51218
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/10/2021
ms.locfileid: "50613907"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a>カスタム関数から Excel JavaScript API を呼び出す

カスタム関数から Excel JavaScript API を呼び出して、範囲データを取得し、計算のコンテキストを追加します。 カスタム関数を使用して Excel JavaScript API を呼び出すのは、次の場合に役立ちます。

- カスタム関数は、計算前に Excel から情報を取得する必要があります。 この情報には、ドキュメントのプロパティ、範囲の形式、カスタム XML パーツ、ブック名、その他の Excel 固有の情報が含まれる場合があります。
- カスタム関数は、計算後の戻り値のセルの数値形式を設定します。

> [!IMPORTANT]
> カスタム関数から Excel JavaScript API を呼び出す場合は、共有 JavaScript ランタイムを使用する必要があります。 詳細については、「[Office アドインを構成して共有 JavaScript ランタイムを使用する ](../develop/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。

## <a name="code-sample"></a>コード サンプル

カスタム関数から Excel JavaScript API を呼び出す場合は、まずコンテキストが必要です。 コンテキストを [取得するには、Excel.RequestContext](/javascript/api/excel/excel.requestcontext) オブジェクトを使用します。 次に、ブックで必要な API を呼び出すコンテキストを使用します。

次のコード サンプルは、ブック内のセルから値を取得する `Excel.RequestContext` 方法を示しています。 このサンプルでは、パラメーター `address` は Excel JavaScript API [Worksheet.getRange](/javascript/api/excel/excel.worksheet#getRange_address_) メソッドに渡され、文字列として入力する必要があります。 たとえば、Excel UI に入力されたカスタム関数は、値を取得するセルのアドレスであるパターンに従 `=CONTOSO.GETRANGEVALUE("A1")` `"A1"` う必要があります。

```JavaScript
/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
async function getRangeValue(address) {
 // Retrieve the context object. 
 var context = new Excel.RequestContext();
 
 // Use the context object to access the cell at the input address. 
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a>カスタム関数を使用して Excel JavaScript API を呼び出す場合の制限事項

Excel の環境を変更するカスタム関数から Excel JavaScript API を呼び出さない。 つまり、カスタム関数は次の操作を行う必要があります。

- スプレッドシートにセルを挿入、削除、または書式設定します。
- 別のセルの値を変更します。
- ブックにシートを移動、名前変更、削除、または追加します。
- 計算モードや画面ビューなど、環境オプションを変更します。
- ブックに名前を追加します。
- プロパティを設定するか、ほとんどのメソッドを実行します。

Excel を変更すると、パフォーマンスが低下し、タイム アウトが発生し、無限ループが発生する可能性があります。 Excel の再計算が行なっている間は、予期しない結果になるので、カスタム関数の計算は実行できません。

代わりに、リボン ボタンまたは作業ウィンドウのコンテキストから Excel に変更を加えます。

## <a name="next-steps"></a>次の手順

- [Excel JavaScript API を使用した基本的なプログラミングの概念](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>関連項目

- [Excel カスタム関数と作業ウィンドウチュートリアルの間でデータとイベントを共有する](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Office アドインを構成して共有 JavaScript ランタイムを使用する](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
