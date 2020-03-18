---
title: エラー処理
description: ランタイムエラーを考慮した Excel JavaScript API のエラー処理ロジックについて説明します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: bee5824d8854a55d5ac4041be1335ce239b31a9e
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717167"
---
# <a name="error-handling"></a>エラー処理

Excel JavaScript API を使用してアドインを作成する場合は、実行時エラーを考慮するために、エラー処理ロジックを含めます。 これは、API の非同期性のために重要になります。

> [!NOTE]
> `sync()`メソッドと EXCEL javascript api の非同期性の詳細については、「 [excel javascript api を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)」を参照してください。

## <a name="best-practices"></a>ベスト プラクティス

このドキュメントのコード サンプルでは、`Excel.run` へのすべての呼び出しに、`catch` 内で発生したエラーを検出するための `Excel.run` ステートメントが付いていることがわかります。 Excel JavaScript Api を使用してアドインを構築するときには、同じパターンを使用することをお勧めします。

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

- **code**:エラー メッセージの `code` プロパティには、`OfficeExtension.ErrorCodes` または `Excel.ErrorCodes` リストの一部である文字列が含まれます。 たとえば、エラー コード "InvalidReference" は、参照が指定された操作に対して有効でないことを示します。 エラー コードはローカライズされません。

- **message**: エラー メッセージの `message` プロパティには、ローカライズされた文字列のエラーの概要が含まれています。 このエラー メッセージは、エンド ユーザーが使用するためのものではありません。アドインによってエンド ユーザーに表示されるエラー メッセージは、エラー コードと適切なビジネス ロジックを使用して、判断する必要があります。

- **debugInfo**:存在する場合、エラー メッセージの `debugInfo` プロパティは、エラーの根本原因を理解するために使用できる追加情報を提供します。

> [!NOTE]
> `console.log()` を使用してエラー メッセージをコンソールに出力すると、それらのメッセージはサーバー上でのみ表示されます。 これらのエラー メッセージが、アドインの作業ウィンドウやホスト アプリケーション内のいずれかの場所で、エンド ユーザーに対して表示されることはありません。

## <a name="error-messages"></a>エラー メッセージ

次の表は、API から返される可能性のあるエラー一覧です。

|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |引数が無効であるか、存在しません。または形式が正しくありません。|
|InvalidRequest  |要求を処理できません。|
|InvalidReference|この参照は、現在の操作に対して無効です。|
|InvalidBinding  |このオブジェクトのバインドは、以前の更新プログラムが原因で無効になっています。|
|InvalidSelection|現在の選択内容は、この操作では無効です。|
|Unauthenticated |必要な認証情報が見つからないか、無効です。|
|AccessDenied |要求された操作を実行できません。|
|ItemNotFound |要求されたリソースは存在しません。|
|ActivityLimitReached|アクティビティの制限に達しました。|
|GeneralException|要求の処理中に内部エラーが発生しました。|
|NotImplemented  |要求された機能は実装されていません。|
|ServiceNotAvailable|サービスを利用できません。|
|Conflict|競合のため、要求を処理できませんでした。|
|ItemAlreadyExists|作成中のリソースはすでに存在しています。|
|UnsupportedOperation|試行中の操作はサポートされていません。|
|RequestAborted|実行時に要求が中止されました。|
|ApiNotAvailable|要求された API は使用できません。|
|InsertDeleteConflict|試行された挿入操作または削除操作で競合が発生しました。|
|InvalidOperation|試行された操作は、このオブジェクトでは無効です。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)
- [OfficeExtension.Error オブジェクト (JavaScript API for Excel)](/javascript/api/office/officeextension.error)
