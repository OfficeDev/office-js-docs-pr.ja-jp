---
title: カスタム関数から Microsoft Excel Api を呼び出す
description: カスタム関数から呼び出すことができる Microsoft Excel Api について説明します。
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: a25d3f151f648560ee24a3da3f689cb9767bd52a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609805"
---
# <a name="call-microsoft-excel-apis-from-a-custom-function"></a>カスタム関数から Microsoft Excel Api を呼び出す

カスタム関数から Office .js Excel Api を呼び出して、範囲データを取得し、計算のためにより多くのコンテキストを取得します。

カスタム関数を使用した Office .js Api の呼び出しは、次のような場合に役立ちます。

- カスタム関数は、計算の前に Excel から情報を取得する必要があります。 この情報には、ドキュメントのプロパティ、範囲の書式、カスタム XML パーツ、ブック名、その他の Excel 固有の情報が含まれることがあります。
- ユーザー設定関数は、計算後の戻り値のセルの番号書式を設定します。

## <a name="code-sample"></a>コード サンプル

Office .js Api を呼び出すには、まずコンテキストが必要です。 オブジェクトを使用して `Excel.RequestContext` コンテキストを取得します。 その後、コンテキストを使用して、ブックで必要な Api を呼び出します。

次のコードサンプルは、ブックから値の範囲を取得する方法を示しています。

```JavaScript
/**
 * @customfunction
 * @param address range's address
 **/
async function getRangeValue (address) {
 var context = new Excel.RequestContext();
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-officejs-through-a-custom-function"></a>カスタム関数を使用して Office .js を呼び出す際の制限

Excel の環境を変更するカスタム関数から Office .js Api を呼び出さないでください。 これは、カスタム関数が以下の操作を行わないことを意味します。

- スプレッドシートのセルを挿入、削除、または書式設定します。
- 別のセルの値を変更します。
- ブックにシートを移動、名前変更、削除、または追加します。
- 計算モードや画面表示などの環境オプションを変更します。
- ブックに名前を追加します。
- プロパティを設定するか、ほとんどのメソッドを実行します。

Excel を変更すると、パフォーマンスが低下し、タイムアウトになり、無限ループが発生します。 Excel の再計算を実行しているときに、予期しない結果になる可能性があるため、カスタム関数の計算を実行することはできません。

代わりに、リボンボタンまたは作業ウィンドウのコンテキストから Excel に変更を加えます。

## <a name="next-steps"></a>次の手順

- [Excel JavaScript API を使用した基本的なプログラミングの概念](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>関連項目

- [Excel カスタム関数と作業ウィンドウチュートリアルの間でデータとイベントを共有する](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
