---
title: カスタムExcel JavaScript API を呼び出す
description: カスタム関数Excel呼び出す JavaScript API について説明します。
ms.date: 08/30/2021
localization_priority: Normal
ms.openlocfilehash: 93b0c1a792c752102359b31b8baa808182c29c46
ms.sourcegitcommit: 3287eb4588d0af47f1ab8a59882bcc3f585169d8
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/02/2021
ms.locfileid: "58863528"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a>カスタムExcel JavaScript API を呼び出す

カスタムExcel JavaScript API を呼び出して、範囲データを取得し、計算のコンテキストを取得します。 カスタム関数Excel JavaScript API の呼び出しは、次の場合に役立ちます。

- カスタム関数は、計算前にデータからExcelする必要があります。 この情報には、ドキュメントのプロパティ、範囲の形式、カスタム XML パーツ、ブック名、その他のExcel情報が含まれます。
- カスタム関数は、計算後の戻り値のセルの数値形式を設定します。

> [!IMPORTANT]
> カスタム関数Excel JavaScript API を呼び出す場合は、共有 JavaScript ランタイムを使用する必要があります。 詳細については、「[Office アドインを構成して共有 JavaScript ランタイムを使用する ](../develop/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。

## <a name="code-sample"></a>コード サンプル

カスタム関数Excel JavaScript API を呼び出す場合は、まずコンテキストが必要です。 次の[Excel。コンテキストを取得する RequestContext](/javascript/api/excel/excel.requestcontext)オブジェクト。 次に、ブックで必要な API を呼び出すコンテキストを使用します。

次のコード サンプルは、ブック内のセルから値を取得する `Excel.RequestContext` 方法を示しています。 このサンプルでは、パラメーター `address` は JavaScript API [Worksheet.getRange](/javascript/api/excel/excel.worksheet#getRange_address_)メソッドExcel渡され、文字列として入力する必要があります。 たとえば、Excel UI に入力されたカスタム関数は、値を取得するセルのアドレスであるパターンに `=CONTOSO.GETRANGEVALUE("A1")` `"A1"` 従う必要があります。

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
 range.load("values");
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a>カスタム関数を使用Excel JavaScript API の呼び出しの制限事項

JavaScript API の環境Excel変更するカスタム関数から JavaScript API を呼び出Excel。 つまり、カスタム関数は次の操作を行う必要があります。

- スプレッドシートにセルを挿入、削除、または書式設定します。
- 別のセルの値を変更します。
- ブックにシートを移動、名前変更、削除、または追加します。
- 計算モードや画面ビューなど、環境オプションを変更します。
- ブックに名前を追加します。
- プロパティを設定するか、ほとんどのメソッドを実行します。

このExcelすると、パフォーマンスが低下し、タイム アウトが発生し、無限ループが発生する可能性があります。 カスタム関数の計算は、予測できない結果Excel再計算が行なっている間は実行できません。

代わりに、リボン ボタンExcel作業ウィンドウのコンテキストから変更を加えます。

## <a name="next-steps"></a>次の手順

- [Excel JavaScript API を使用した基本的なプログラミングの概念](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>関連項目

- [カスタム関数と作業ウィンドウのチュートリアルExcelデータとイベントを共有する](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Office アドインを構成して共有 JavaScript ランタイムを使用する](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
