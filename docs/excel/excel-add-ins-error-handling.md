---
title: JavaScript API のExcel処理
description: ランタイム エラー Excel説明する JavaScript API エラー処理ロジックについて説明します。
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: fa03cd9a3ccee9fce1cbb7025baf6c2463ff938d
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340541"
---
# <a name="error-handling-with-the-excel-javascript-api"></a>JavaScript API のExcel処理

Excel JavaScript API を使用してアドインを作成する場合は、実行時エラーを考慮するために、エラー処理ロジックを含めます。 これは、API の非同期性のために重要になります。

> [!NOTE]
> JavaScript API の`sync()`メソッドと非同期の性質の詳細については、「Excel アドイン」の「[Excel JavaScript](excel-add-ins-core-concepts.md) オブジェクト モデルOffice参照してください。

## <a name="best-practices"></a>ベスト プラクティス

コード サンプル[と](https://github.com/OfficeDev/Office-Add-in-samples) [Script Lab スニペット](../overview/explore-with-script-lab.md) `Excel.run` `catch` `Excel.run`では、すべての呼び出しにステートメントが伴い、 内で発生するエラーをキャッチします。 Excel JavaScript Api を使用してアドインを構築するときには、同じパターンを使用することをお勧めします。

```js
$("#run").click(() => tryCatch(run));

async function run() {
  await Excel.run(async (context) => {
      // Add your Excel JavaScript API calls here.

      // Await the completion of context.sync() before continuing.
    await context.sync();
    console.log("Finished!");
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

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

|エラー コード | エラー メッセージ | メモ |
|:----------|:--------------|:------|
|`AccessDenied` |要求された操作を実行できません。| |
|`ActivityLimitReached`|アクティビティの制限に達しました。| |
|`ApiNotAvailable`|要求された API は使用できません。| |
|`ApiNotFound`|使用しようとしている API が見つかりませんでした。 このバージョンは、新しいバージョンの Excel。 詳細についてはExcel [JavaScript API 要件セットの記事](../reference/requirement-sets/excel-api-requirement-sets.md)を参照してください。| |
|`BadPassword`|指定したパスワードが正しくありません。| |
|`Conflict`|競合のため、要求を処理できませんでした。| |
|`ContentLengthRequired`|HTTP `Content-length` ヘッダーが見つからない。| |
|`EmptyChartSeries`|グラフの系列が空のため、試行された操作は失敗しました。| |
|`FilteredRangeConflict`|試行された操作によって、フィルター処理された範囲との競合が発生します。| |
|`FormulaLengthExceedsLimit`|適用される数式のバイトコードが最大長制限を超えています。 32 Officeの場合、バイトコードの長さの制限は 16384 文字です。 64 ビット コンピューターでは、バイトコードの長さの制限は 32768 文字です。| このエラーは、デスクトップとExcel on the web両方で発生します。|
|`GeneralException`|要求の処理中に内部エラーが発生しました。| |
|`InactiveWorkbook`|複数のブックが開き、この API によって呼び出されるブックがフォーカスを失ったため、操作に失敗しました。| |
|`InsertDeleteConflict`|試行された挿入操作または削除操作で競合が発生しました。| |
|`InvalidArgument` |引数が無効であるか、存在しません。または形式が正しくありません。| |
|`InvalidBinding` |このオブジェクトのバインドは、以前の更新プログラムが原因で無効になっています。| |
|`InvalidOperation`|試行された操作は、このオブジェクトでは無効です。| |
|`InvalidOperationInCellEditMode`|[編集] セル モードの場合Excel操作は使用できません。 Enter キーまたは Tab **キーを使用** するか、別のセルを選択して編集モードを終了し、もう一度やり直します。| |
|`InvalidReference`|この参照は、現在の操作に対して無効です。| |
|`InvalidRequest`  |要求を処理できません。| |
|`InvalidSelection`|現在の選択内容は、この操作では無効です。| |
|`ItemAlreadyExists`|作成中のリソースはすでに存在しています。| |
|`ItemNotFound` |要求されたリソースは存在しません。| |
|`MemoryLimitReached`|メモリ制限に達しました。 アクションを完了する必要があります。| |
|`MergedRangeConflict`|操作を完了できません。 テーブルを別のテーブル、ピボットテーブル レポート、クエリ結果、結合セル、または XML マップと重ね合えすることはできません。|
|`NonBlankCellOffSheet`|Microsoft Excelセルを挿入できないのは、空でないセルをワークシートの最後から押し出すためです。 空でないセルは空に表示されますが、空白の値、書式設定、または数式が含まれます。 挿入する行または列を十分に削除してから、もう一度やり直してください。| |
|`NotImplemented`|要求された機能は実装されていません。| |
|`OperationCellsExceedLimit`|試行された操作は、33554000 セルの制限を超える値に影響します。| このエラー `TableColumnCollection.add API` が発生した場合は、ワークシート内に意図しないデータが含まれるのではなく、テーブルの外側にデータが含まれるか確認します。 特に、ワークシートの最も右の列にあるデータを確認します。 意図しないデータを削除して、このエラーを解決します。 操作が処理するセルの数を確認する方法の 1 つは、次の計算を実行する方法です。 `(number of table rows) x (16383 - (number of table columns))` 数値 16383 は、サポートされている列Excelです。 <br><br>このエラーは、このエラーが発生Excel on the web。 |
|`PivotTableRangeConflict`|操作が試行された場合、ピボットテーブル範囲との競合が発生します。| |
|`RangeExceedsLimit`|範囲内のセル数がサポートされている最大数を超えました。 詳細については[、「リソースの制限とパフォーマンスの最適化」のOfficeを](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)参照してください。| |
|`RefreshWorkbookLinksBlocked`|ユーザーが外部ブック リンクを更新するアクセス許可を付与しないので、操作に失敗しました。| |
|`RequestAborted`|実行時に要求が中止されました。| |
|`RequestPayloadSizeLimitExceeded`|要求ペイロードのサイズが制限を超えています。 詳細については[、「リソースの制限とパフォーマンスの最適化」のOfficeを](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)参照してください。| このエラーは、このエラーが発生Excel on the web。|
|`ResponsePayloadSizeLimitExceeded`|応答ペイロードのサイズが制限を超えています。 詳細については[、「リソースの制限とパフォーマンスの最適化」のOfficeを](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)参照してください。|  このエラーは、このエラーが発生Excel on the web。|
|`ServiceNotAvailable`|サービスを利用できません。| |
|`Unauthenticated` |必要な認証情報が見つからないか、無効です。| |
|`UnsupportedFeature`|ソース ワークシートにサポートされていない機能が 1 つ以上含まれているため、操作に失敗しました。| |
|`UnsupportedOperation`|試行中の操作はサポートされていません。| |
|`UnsupportedSheet`|このシートの種類はマクロ シートまたはグラフ シートで、この操作はサポートされていません。| |

> [!NOTE]
> 前の表に、JavaScript API の使用中に発生する可能性Excel一覧を示します。 アプリケーション固有の Excel JavaScript API の代わりに共通 API を使用している場合は、「Office 共通 [API](../reference/javascript-api-for-office-error-codes.md) エラー コード」を参照して、関連するエラー メッセージについて説明します。

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [OfficeExtension.Error オブジェクト (JavaScript API for Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [Office の一般的な API エラー コード](../reference/javascript-api-for-office-error-codes.md)
