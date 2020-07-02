---
title: エラー処理
description: ランタイムエラーを考慮した Excel JavaScript API のエラー処理ロジックについて説明します。
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: 8d410ae7eea315e14383b5aa08373ede3768cace
ms.sourcegitcommit: 065bf4f8e0d26194cee9689f7126702b391340cc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/01/2020
ms.locfileid: "45006445"
---
# <a name="error-handling"></a>エラー処理

When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.

> [!NOTE]
> `sync()`メソッドと Excel JAVASCRIPT api の非同期性の詳細については、「 [EXCEL javascript api を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)」を参照してください。

## <a name="best-practices"></a>ベスト プラクティス

Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.

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

- **code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.

- **message**: エラー メッセージの `message` プロパティには、ローカライズされた文字列のエラーの概要が含まれています。 このエラー メッセージは、エンド ユーザーが使用するためのものではありません。アドインによってエンド ユーザーに表示されるエラー メッセージは、エラー コードと適切なビジネス ロジックを使用して、判断する必要があります。

- **debugInfo**:存在する場合、エラー メッセージの `debugInfo` プロパティは、エラーの根本原因を理解するために使用できる追加情報を提供します。

> [!NOTE]
> `console.log()` を使用してエラー メッセージをコンソールに出力すると、それらのメッセージはサーバー上でのみ表示されます。 これらのエラー メッセージが、アドインの作業ウィンドウやホスト アプリケーション内のいずれかの場所で、エンド ユーザーに対して表示されることはありません。

## <a name="error-messages"></a>エラー メッセージ

次の表は、API から返される可能性のあるエラー一覧です。

|error.code | error.message |
|:----------|:--------------|
|`AccessDenied` |要求された操作を実行できません。|
|`ActivityLimitReached`|アクティビティの制限に達しました。|
|`ApiNotAvailable`|要求された API は使用できません。|
|`Conflict`|競合のため、要求を処理できませんでした。|
|`GeneralException`|要求の処理中に内部エラーが発生しました。|
|`InsertDeleteConflict`|試行された挿入操作または削除操作で競合が発生しました。|
|`InvalidArgument` |引数が無効であるか、存在しません。または形式が正しくありません。|
|`InvalidBinding`  |このオブジェクトのバインドは、以前の更新プログラムが原因で無効になっています。|
|`InvalidOperation`|試行された操作は、このオブジェクトでは無効です。|
|`InvalidReference`|この参照は、現在の操作に対して無効です。|
|`InvalidRequest`  |要求を処理できません。|
|`InvalidSelection`|現在の選択内容は、この操作では無効です。|
|`ItemAlreadyExists`|作成中のリソースはすでに存在しています。|
|`ItemNotFound` |要求されたリソースは存在しません。|
|`NotImplemented`  |要求された機能は実装されていません。|
|`RequestAborted`|実行時に要求が中止されました。|
|`ServiceNotAvailable`|サービスを利用できません。|
|`Unauthenticated` |必要な認証情報が見つからないか、無効です。|
|`UnsupportedOperation`|試行中の操作はサポートされていません。|
|`UnsupportedSheet`|このシートの種類は、マクロまたはグラフシートであるため、この操作をサポートしていません。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)
- [OfficeExtension.Error オブジェクト (JavaScript API for Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview)
