---
title: カスタム関数から Excel JavaScript API を呼び出す
description: カスタム関数から呼び出すことができる Excel JavaScript API について説明します。
ms.date: 08/30/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8d1cbf6d07e4ede5b8309e899828f8f1d8ad1fa0
ms.sourcegitcommit: eef2064d7966db91f8401372dd255a32d76168c2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/31/2022
ms.locfileid: "67464833"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a>カスタム関数から Excel JavaScript API を呼び出す

カスタム関数から Excel JavaScript API を呼び出して、範囲データを取得し、計算のためのより多くのコンテキストを取得します。 カスタム関数を使用して Excel JavaScript API を呼び出すと、次の場合に役立ちます。

- カスタム関数は、計算前に Excel から情報を取得する必要があります。 この情報には、ドキュメントプロパティ、範囲形式、カスタム XML パーツ、ブック名、またはその他の Excel 固有の情報が含まれる場合があります。
- カスタム関数は、計算後に戻り値のセルの数値形式を設定します。

> [!IMPORTANT]
> カスタム関数から Excel JavaScript API を呼び出すには、 [共有ランタイム](../testing/runtimes.md#shared-runtime)を使用する必要があります。 詳細については、「 [共有ランタイムを使用するように Office アドインを構成](../develop/configure-your-add-in-to-use-a-shared-runtime.md) する」を参照してください。

## <a name="code-sample"></a>コード サンプル

カスタム関数から Excel JavaScript API を呼び出すには、まずコンテキストが必要です。 [Excel.RequestContext](/javascript/api/excel/excel.requestcontext) オブジェクトを使用してコンテキストを取得します。 次に、コンテキストを使用して、ブックに必要な API を呼び出します。

次のコード サンプルは、ブック内のセルから値を取得する方法 `Excel.RequestContext` を示しています。 このサンプルでは、 `address` パラメーターは Excel JavaScript API [Worksheet.getRange](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1)) メソッドに渡され、文字列として入力する必要があります。 たとえば、Excel UI に入力されたカスタム関数は、`"A1"`値を取得するセルのアドレスであるパターン`=CONTOSO.GETRANGEVALUE("A1")`に従う必要があります。

```JavaScript
/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
async function getRangeValue(address) {
 // Retrieve the context object. 
 const context = new Excel.RequestContext();
 
 // Use the context object to access the cell at the input address. 
 const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load("values");
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a>カスタム関数を使用して Excel JavaScript API を呼び出す場合の制限事項

カスタム関数アドインは Excel JavaScript API を呼び出すことができますが、呼び出す API については注意が必要です。 カスタム関数を実行しているセルの外側のセルを変更するカスタム関数から Excel JavaScript API を呼び出さないでください。 他のセルまたは Excel 環境を変更すると、Excel アプリケーションのパフォーマンス、タイムアウト、無限ループが低下する可能性があります。 つまり、カスタム関数では次の操作を行わないでください。

- スプレッドシートのセルを挿入、削除、書式設定します。
- 別のセルの値を変更します。
- ブックにシートを移動、名前変更、削除、または追加します。
- ブックに名前を追加します。
- プロパティを設定します。
- 計算モードや画面ビューなど、Excel 環境オプションを変更します。

カスタム関数アドインは、カスタム関数を実行しているセルの外部のセルから情報を読み取ることができますが、他のセルへの書き込み操作を実行しないでください。 代わりに、リボン ボタンまたは作業ウィンドウのコンテキストから、他のセルまたは Excel 環境に変更を加えます。 また、このシナリオでは予測できない結果が作成されるため、Excel 再計算が行われている間はカスタム関数の計算を実行しないでください。

## <a name="next-steps"></a>次の手順

- [Excel JavaScript API を使用した基本的なプログラミングの概念](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>関連項目

- [Excel カスタム関数と作業ウィンドウのチュートリアルの間でデータとイベントを共有する](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [共有ランタイムを使用するように Office アドインを構成する](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
