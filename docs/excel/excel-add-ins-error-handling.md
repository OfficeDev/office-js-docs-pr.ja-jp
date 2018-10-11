---
title: エラー処理
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: b07012516cbe15374d0707c157738117a9c8fe96
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459232"
---
# <a name="error-handling"></a>エラー処理

Excel の JavaScript API を使用してアドインをビルドする場合は、ランタイム エラーを考慮するためのエラー処理ロジックを含めるようにしてください。これは、API の非同期の性質のために重要です。

> [!NOTE]
>  **Sync()** メソッドと非同期であるため Excel の JavaScript API の詳細については、 [Excel の JavaScript API を使用して基本的なプログラミングの概念](excel-add-ins-core-concepts.md)を参照してください。

## <a name="best-practices"></a>ベスト プラクティス

このドキュメントのコード サンプル全体にわたり、`Excel.run`へのすべての呼び出しが`Excel.run`内で発生するエラーをキャッチする`catch`ステートメントに付属していることが分かります。Excel の JavaScript Api を使用してアドインをビルドするときは、同じパターンを使用することをお勧めします。

```js
Excel.run(function (context) { 
  
  // Excel JavaScript API calls here

  // Await the completion of context.sync() before continuing.
  return context.sync()
    .then(function () {
      console.log("Finished!");
    })
}).catch(errorHandlerFunction);     
```

## <a name="api-errors"></a>API エラー 

Excel JavaScript API 要求が正常に実行されない場合、API は次のプロパティを含むエラー オブジェクトを返します。 

- **コード**: `code` エラー メッセージのプロパティが含まれている文字列を含む、 `OfficeExtension.ErrorCodes` または `Excel.ErrorCodes` リストです。たとえば、エラー コード"InvalidReference"では、参照が指定された操作に対して有効ではないことを示します。エラー コードはローカライズされません。 

- **メッセージ**: `message` エラー メッセージのプロパティには、ローカライズされた文字列のエラーの概要が含まれています。エラー メッセージは、エンド ユーザーの消費対象ではありません。アドインがエンド ユーザーに示すエラー メッセージを確認するには、エラー コードと適切なビジネス ロジックを使用してください。

- **debugInfo**: 存在する場合、エラー メッセージの `debugInfo` プロパティは、エラーの根本原因を理解するために使用できる追加情報を提供します。 

> [!NOTE]
> `console.log()`を使用してコンソールにエラー メッセージを印刷する場合は、サーバー上にそれらのメッセージのみが表示されます。エンド ユーザーには、アドイン作業ウィンドウでまたはホスト アプリケーションの任意の場所にこれらのエラー メッセージは表示されません。

## <a name="see-also"></a>関連項目

- [Excel の JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)
- [OfficeExtension.Error オブジェクト (Excel JavaScript API)](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
