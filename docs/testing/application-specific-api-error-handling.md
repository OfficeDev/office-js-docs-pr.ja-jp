---
title: アプリケーション固有の JavaScript API でのエラー処理
description: Excel、Word、PowerPoint、およびその他のアプリケーション固有の JavaScript API エラー処理ロジックについて説明し、ランタイム エラーを考慮します。
ms.date: 07/05/2022
ms.localizationpriority: medium
ms.openlocfilehash: b6f25f5740892df4729b72ee5ad87403853f45fb
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092995"
---
# <a name="error-handling-with-the-application-specific-javascript-apis"></a>アプリケーション固有の JavaScript API でのエラー処理

[アプリケーション固有の Office JavaScript API を](../develop/application-specific-api-model.md)使用してアドインをビルドする場合は、ランタイム エラーを考慮するエラー処理ロジックを必ず含めておきます。 API の非同期性のために、これを行うことは非常に重要です。

## <a name="best-practices"></a>ベスト プラクティス

[コード サンプル](https://github.com/OfficeDev/Office-Add-in-samples)と[Script Lab](../overview/explore-with-script-lab.md) スニペットでは、すべての呼び出しが 、`PowerPoint.run`エラー`Word.run`をキャッチするための`Excel.run`ステートメントを伴`catch`っていることがわかります。 アプリケーション固有の API を使用してアドインをビルドする場合は、同じパターンを使用することをお勧めします。

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

Office JavaScript API 要求が正常に実行されない場合、API は次のプロパティを含むエラー オブジェクトを返します。

- **code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.

- **message**: エラー メッセージの `message` プロパティには、ローカライズされた文字列のエラーの概要が含まれています。 このエラー メッセージは、エンド ユーザーが使用するためのものではありません。アドインによってエンド ユーザーに表示されるエラー メッセージは、エラー コードと適切なビジネス ロジックを使用して、判断する必要があります。

- **debugInfo**:存在する場合、エラー メッセージの `debugInfo` プロパティは、エラーの根本原因を理解するために使用できる追加情報を提供します。

> [!NOTE]
> コンソールにエラー メッセージを出力するために使用 `console.log()` する場合、これらのメッセージはサーバーでのみ表示されます。 エンド ユーザーは、アドイン作業ウィンドウまたは Office アプリケーション内の任意の場所にこれらのエラー メッセージを表示しません。 ユーザーにエラーを報告するには、「 [エラー通知](#error-notifications)」を参照してください。

## <a name="error-codes-and-messages"></a>エラー コードとメッセージ

次の表に、アプリケーション固有の API が返す可能性があるエラーを示します。

> [!NOTE]
> 上の表に、アプリケーション固有の API の使用中に発生する可能性があるエラー メッセージを示します。 Common API を使用している場合は、関連するエラー メッセージの詳細については、 [Office Common API のエラー コード](../reference/javascript-api-for-office-error-codes.md) を参照してください。

|エラー コード | エラー メッセージ | 備考 |
|:----------|:--------------|:------|
|`AccessDenied` |要求された操作を実行できません。|*なし。* |
|`ActivityLimitReached`|アクティビティの制限に達しました。|*なし。* |
|`ApiNotAvailable`|要求された API は使用できません。|*なし。* |
|`ApiNotFound`|使用しようとしている API が見つかりませんでした。 新しいバージョンの Excel で使用できる場合があります。 詳細については、 [Excel JavaScript API 要件セット](/javascript/api/requirement-sets/excel/excel-api-requirement-sets) に関する記事を参照してください。|*なし。* |
|`BadPassword`|指定したパスワードが正しくありません。|*なし。* |
|`Conflict`|競合のため、要求を処理できませんでした。|*なし。* |
|`ContentLengthRequired`|`Content-length` HTTP ヘッダーがありません。|*なし。* |
|`GeneralException`|要求の処理中に内部エラーが発生しました。|*なし。* |
|`InsertDeleteConflict`|試行された挿入操作または削除操作で競合が発生しました。|*なし。* |
|`InvalidArgument` |引数が無効であるか、存在しません。または形式が正しくありません。|*なし。* |
|`InvalidBinding` |このオブジェクトのバインドは、以前の更新プログラムが原因で無効になっています。|*なし。* |
|`InvalidOperation`|試行された操作は、このオブジェクトでは無効です。|*なし。* |
|`InvalidOperationInCellEditMode`|Excel がセルの編集モードになっている間は、操作を使用できません。 **Enter** キーまたは **Tab** キーを使用するか、別のセルを選択して編集モードを終了してから、もう一度やり直してください。|*なし。* |
|`InvalidReference`|この参照は、現在の操作に対して無効です。|*なし。* |
|`InvalidRequest`  |要求を処理できません。|*なし。* |
|`InvalidSelection`|現在の選択内容は、この操作では無効です。|*なし。* |
|`ItemAlreadyExists`|作成中のリソースはすでに存在しています。|*なし。* |
|`ItemNotFound` |要求されたリソースは存在しません。|*なし。* |
|`MemoryLimitReached`|メモリの制限に達しました。 アクションを完了できませんでした。|*なし。* |
|`NotImplemented`|要求された機能は実装されていません。| これは、API がプレビュー段階であるか、特定のプラットフォームでのみサポートされていることを意味する可能性があります (オンラインのみなど)。 詳細については、「 [Office アドインの Office クライアント アプリケーションとプラットフォームの可用性](/javascript/api/requirement-sets) 」を参照してください。|
|`RequestAborted`|実行時に要求が中止されました。|*なし。* |
|`RequestPayloadSizeLimitExceeded`|要求ペイロードのサイズが制限を超えました。 詳細については、 [Office アドインのリソースの制限とパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) に関する記事を参照してください。| このエラーは、Office on the webでのみ発生します。|
|`ResponsePayloadSizeLimitExceeded`|応答ペイロード のサイズが制限を超えました。 詳細については、 [Office アドインのリソースの制限とパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) に関する記事を参照してください。|  このエラーは、Office on the webでのみ発生します。|
|`ServiceNotAvailable`|サービスを利用できません。|*なし。* |
|`Unauthenticated` |必要な認証情報が見つからないか、無効です。|*なし。* |
|`UnsupportedFeature`|ソース ワークシートにサポートされていない 1 つ以上の機能が含まれているため、操作は失敗しました。|*なし。* |
|`UnsupportedOperation`|試行中の操作はサポートされていません。|*なし。* |

### <a name="excel-specific-error-codes-and-messages"></a>Excel 固有のエラー コードとメッセージ

|エラー コード | エラー メッセージ | 備考 |
|:----------|:--------------|:------|
|`EmptyChartSeries`|グラフ系列が空であるため、試行された操作は失敗しました。|*なし。* |
|`FilteredRangeConflict`|試行された操作により、フィルター処理された範囲との競合が発生します。|*なし。* |
|`FormulaLengthExceedsLimit`|適用された数式のバイトコードが最大長制限を超えています。 32 ビット コンピューターの Office の場合、バイトコードの長さの制限は 16384 文字です。 64 ビット コンピューターでは、バイトコードの長さの制限は 32768 文字です。| このエラーは、Excel on the webとデスクトップの両方で発生します。|
|`InactiveWorkbook`|複数のブックが開かれているため、この API によって呼び出されるブックにフォーカスが失われるため、操作は失敗しました。|*なし。* |
|`MergedRangeConflict`|操作を完了できません。 テーブルは、別のテーブル、ピボットテーブル レポート、クエリ結果、マージされたセル、または XML マップと重複することはできません。|*なし。* |
|`NonBlankCellOffSheet`|ワークシートの末尾に空でないセルをプッシュするため、Microsoft Excel では新しいセルを挿入できません。 これらの空でないセルは空に見えますが、空白の値、一部の書式設定、または数式があります。 挿入する行または列を十分に削除して、挿入する内容に余裕を持たせた後、もう一度やり直してください。|*なし。* |
|`OperationCellsExceedLimit`|試行された操作は、33554000 セルの制限を超える影響を与えます。| このエラーがトリガーされた `TableColumnCollection.add API` 場合は、ワークシート内でテーブルの外部に意図しないデータがないことを確認します。 特に、ワークシートの右端の列でデータを確認します。 このエラーを解決するには、意図しないデータを削除します。 操作が処理するセルの数を確認する 1 つの方法は、次の計算 `(number of table rows) x (16383 - (number of table columns))`を実行することです。 数値 16383 は、Excel がサポートする列の最大数です。 <br><br>このエラーは、Excel on the webでのみ発生します。 |
|`PivotTableRangeConflict`|試行された操作により、ピボットテーブル範囲との競合が発生します。|*なし。* |
|`RangeExceedsLimit`|範囲内のセル数が、サポートされている最大数を超えています。 詳細については、 [Office アドインのリソースの制限とパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) に関する記事を参照してください。|*なし。* |
|`RefreshWorkbookLinksBlocked`|ユーザーが外部ブックリンクを更新するアクセス許可を付与していないため、操作は失敗しました。|*なし。* |
|`UnsupportedSheet`|このシートの種類では、マクロ シートまたはグラフ シートであるため、この操作はサポートされていません。|*なし。* |

## <a name="error-notifications"></a>エラー通知

ユーザーにエラーを報告する方法は、使用している UI システムによって異なります。 REACTを UI システムとして使用している場合は、Fluent UI コンポーネントとデザイン要素を使用します。 この [Fluent UI ページ](https://developer.microsoft.com/fluentui#/controls/web)から適切なコントロールを選択します。 エラー メッセージは、メッセージ バー、ダイアログ、またはモーダルで伝達することをお勧めします。 エラーがユーザーの入力にある場合は、入力コントロールの近くに太字の赤でエラーを表示します。 [サンプル Office-Add-in-Microsoft-Graph-React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/0706cc67645675a48747f1fec1b1e5722b575b11/Samples/auth/Office-Add-in-Microsoft-Graph-React)では、MessageBar 要素を使用し、アドイン作業ウィンドウの [パーソナリティ] メニューを考慮するように変更します。

UI にReactを使用していない場合は、HTML と JavaScript で直接実装された古い Fabric UI コンポーネントの使用を検討してください。 一部のテンプレートの例は [、Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates) リポジトリに含まれています。 特にダイアログサブフォルダーとナビゲーション サブフォルダーを見てみましょう。 [Excel-Add-in-SalesLeads のサンプルでは](https://github.com/OfficeDev/Excel-Add-in-SalesLeads)、メッセージ バナーを使用します。

## <a name="see-also"></a>関連項目

- [OfficeExtension.Error オブジェクト (JavaScript API for Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [Office の一般的な API エラー コード](../reference/javascript-api-for-office-error-codes.md)
