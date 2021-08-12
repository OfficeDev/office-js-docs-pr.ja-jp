---
title: JavaScript API のExcel処理
description: ランタイム エラーをExcel JavaScript API エラー処理ロジックについて説明します。
ms.date: 01/15/2021
localization_priority: Normal
ms.openlocfilehash: 020828a022605454b2b070bedb560ce7d53ff39f6902b042e5e9ac793f14b277
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084273"
---
# <a name="error-handling-with-the-excel-javascript-api"></a>JavaScript API のExcel処理

Excel JavaScript API を使用してアドインを作成する場合は、実行時エラーを考慮するために、エラー処理ロジックを含めます。 これは、API の非同期性のために重要になります。

> [!NOTE]
> JavaScript API のメソッドと非同期の性質の詳細については、「Excel アドイン」の `sync()` [「Excel JavaScript](excel-add-ins-core-concepts.md)オブジェクト モデルOffice参照してください。

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

JavaScript API Excelが正常に実行できない場合、API は次のプロパティを含むエラー オブジェクトを返します。

- **code**:エラー メッセージの `code` プロパティには、`OfficeExtension.ErrorCodes` または `Excel.ErrorCodes` リストの一部である文字列が含まれます。 たとえば、エラー コード "InvalidReference" は、参照が指定された操作に対して有効でないことを示します。 エラー コードはローカライズされません。

- **message**: エラー メッセージの `message` プロパティには、ローカライズされた文字列のエラーの概要が含まれています。 このエラー メッセージは、エンド ユーザーが使用するためのものではありません。アドインによってエンド ユーザーに表示されるエラー メッセージは、エラー コードと適切なビジネス ロジックを使用して、判断する必要があります。

- **debugInfo**:存在する場合、エラー メッセージの `debugInfo` プロパティは、エラーの根本原因を理解するために使用できる追加情報を提供します。

> [!NOTE]
> `console.log()` を使用してエラー メッセージをコンソールに出力すると、それらのメッセージはサーバー上でのみ表示されます。 エンド ユーザーは、アドインの作業ウィンドウまたはアプリケーション内の任意の場所にこれらのエラー メッセージOfficeしません。

## <a name="error-messages"></a>エラー メッセージ

次の表は、API から返される可能性のあるエラー一覧です。

|エラー コード | エラー メッセージ |
|:----------|:--------------|
|`AccessDenied` |要求された操作を実行できません。|
|`ActivityLimitReached`|アクティビティの制限に達しました。|
|`ApiNotAvailable`|要求された API は使用できません。|
|`ApiNotFound`|使用しようとしている API が見つかりませんでした。 このバージョンは、新しいバージョンの Excel。 詳細については[Excel JavaScript API 要件セットの記事](../reference/requirement-sets/excel-api-requirement-sets.md)を参照してください。|
|`BadPassword`|指定したパスワードが正しくありません。|
|`Conflict`|競合のため、要求を処理できませんでした。|
|`ContentLengthRequired`|`Content-length`HTTP ヘッダーが見つからない。|
|`GeneralException`|要求の処理中に内部エラーが発生しました。|
|`InactiveWorkbook`|複数のブックが開き、この API によって呼び出されるブックがフォーカスを失ったため、操作に失敗しました。|
|`InsertDeleteConflict`|試行された挿入操作または削除操作で競合が発生しました。|
|`InvalidArgument` |引数が無効であるか、存在しません。または形式が正しくありません。|
|`InvalidBinding`  |このオブジェクトのバインドは、以前の更新プログラムが原因で無効になっています。|
|`InvalidOperation`|試行された操作は、このオブジェクトでは無効です。|
|`InvalidReference`|この参照は、現在の操作に対して無効です。|
|`InvalidRequest`  |要求を処理できません。|
|`InvalidSelection`|現在の選択内容は、この操作では無効です。|
|`ItemAlreadyExists`|作成中のリソースはすでに存在しています。|
|`ItemNotFound` |要求されたリソースは存在しません。|
|`NonBlankCellOffSheet`|Microsoft Excelセルを挿入できないのは、空でないセルをワークシートの最後から押し出すためです。 空でないセルは空に表示されますが、空白の値、書式設定、または数式が含まれます。 挿入する行または列を十分に削除してから、もう一度やり直してください。|
|`NotImplemented`|要求された機能は実装されていません。|
|`RangeExceedsLimit`|範囲内のセル数がサポートされている最大数を超えました。 詳細については[、「リソースの制限](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)とパフォーマンスの最適化」Officeを参照してください。|
|`RequestAborted`|実行時に要求が中止されました。|
|`RequestPayloadSizeLimitExceeded`|要求ペイロードのサイズが制限を超えています。 詳細については[、「リソースの制限](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)とパフォーマンスの最適化」Officeを参照してください。 <br><br>このエラーは、このエラーが発生Excel on the web。|
|`ResponsePayloadSizeLimitExceeded`|応答ペイロードのサイズが制限を超えています。 詳細については[、「リソースの制限](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)とパフォーマンスの最適化」Officeを参照してください。  <br><br>このエラーは、このエラーが発生Excel on the web。|
|`ServiceNotAvailable`|サービスを利用できません。|
|`Unauthenticated` |必要な認証情報が見つからないか、無効です。|
|`UnsupportedOperation`|試行中の操作はサポートされていません。|
|`UnsupportedSheet`|このシートの種類はマクロ シートまたはグラフ シートで、この操作はサポートされていません。|

> [!NOTE]
> 前の表に、JavaScript API の使用中に発生する可能性Excel示します。 アプリケーション固有の Excel JavaScript API の代わりに共通 API を使用している場合は、「Office 一般的な[API](../reference/javascript-api-for-office-error-codes.md)エラー コード」を参照して、関連するエラー メッセージについて説明します。

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [OfficeExtension.Error オブジェクト (JavaScript API for Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [Office の一般的な API エラー コード](../reference/javascript-api-for-office-error-codes.md)
