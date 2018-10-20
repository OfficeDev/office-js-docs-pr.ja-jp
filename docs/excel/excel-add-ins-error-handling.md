---
title: エラー処理
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: caba29f7d6949cc6d9df1498ac0a3d4f5de6c4ee
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579815"
---
# <a name="error-handling"></a>エラー処理

Excel の JavaScript API を使用してアドインをビルドする場合は、ランタイム エラーを考慮するためのエラー処理ロジックを含めるようにしてください。これは、API の非同期の性質のために重要です。

> [!NOTE]
> **Sync()** メソッドと非同期であるため Excel の JavaScript API の詳細については、 [Excel の JavaScript API を使用して基本的なプログラミングの概念](excel-add-ins-core-concepts.md)を参照してください。

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
> `console.log()` を使用してエラー メッセージをコンソールに出力すると、それらのメッセージはサーバー上でのみ表示されます。これらのエラー メッセージが、アドインの作業ウィンドウやホスト アプリケーション内のいずれかの場所で、エンド ユーザーに対して表示されることはありません。

## <a name="error-messages"></a>エラー メッセージ

次の表は、API から返されるエラー一覧の定義を示します。

|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |引数が無効であるか、存在しません。または形式が正しくありません。|
|InvalidRequest  |要求を処理できません。|
|InvalidReference|この参照は、現在の操作に対して無効です。|
|InvalidBinding  |このオブジェクトのバインドは、以前の更新プログラムが原因で無効になっています。|
|InvalidSelection|現在の選択内容は、この操作では無効です。|
|認証されていません |必要な認証情報が見つからないか、無効です。|
|AccessDenied |要求された操作を実行できません。|
|ItemNotFound |要求されたリソースは存在しません。|
|ActivityLimitReached|アクティビティの制限に達しました。|
|GeneralException|リクエストの処理中に内部エラーが発生しました。|
|NotImplemented  |リクエストされた機能は実装されていません。|
|ServiceNotAvailable|サービスを利用できません。|
|一致しません|競合のため、要求を処理できませんでした。|
|ItemAlreadyExists|作成中のリソースはすでに存在しています。|
|UnsupportedOperation|試行中の操作はサポートされていません。|
|RequestAborted|実行時に要求が中止されました。|
|ApiNotAvailable|要求された API は使用できません。|
|InsertDeleteConflict|試行された挿入操作または削除操作で競合が発生しました。|
|InvalidOperation|試行された操作は、このオブジェクトでは無効です。|

## <a name="see-also"></a>関連項目

- [Excel の JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)
- [OfficeExtension.Error オブジェクト (Excel JavaScript API)](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
