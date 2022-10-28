---
title: アプリケーション固有の JavaScript API でのエラー処理
description: ランタイム エラーを考慮した Excel、Word、PowerPoint、およびその他のアプリケーション固有の JavaScript API エラー処理ロジックについて説明します。
ms.date: 10/27/2022
ms.localizationpriority: medium
ms.openlocfilehash: 21d8d3eef36f919f95459fd8e0b3037c1d5ae1b1
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/28/2022
ms.locfileid: "68767155"
---
# <a name="error-handling-with-the-application-specific-javascript-apis"></a>アプリケーション固有の JavaScript API でのエラー処理

[アプリケーション固有の Office JavaScript API を](../develop/application-specific-api-model.md)使用してアドインをビルドする場合は、実行時エラーを考慮するエラー処理ロジックを必ず含めます。 API の非同期的な性質上、これを行うことは重要です。

## <a name="best-practices"></a>ベスト プラクティス

[コード サンプル](https://github.com/OfficeDev/Office-Add-in-samples)と[Script Lab](../overview/explore-with-script-lab.md) スニペットでは、または `Word.run` のすべての`Excel.run``PowerPoint.run`呼び出しにエラーをキャッチするためのステートメントが付属`catch`していることがわかります。 アプリケーション固有の API を使用してアドインをビルドするときは、同じパターンを使用することをお勧めします。

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

- **code**: エラー メッセージのプロパティには`code`、*{application}* が `OfficeExtension.ErrorCodes` Excel、PowerPoint、または `{application}.ErrorCodes` Word を表す文字列が含まれています。 たとえば、エラー コード "InvalidReference" は、参照が指定された操作に対して有効でないことを示します。 エラー コードはローカライズされません。

- **message**: エラー メッセージの `message` プロパティには、ローカライズされた文字列のエラーの概要が含まれています。 エラー メッセージは、エンド ユーザーによる使用を目的としていません。エラー コードと適切なビジネス ロジックを使用して、アドインがエンド ユーザーに表示するエラー メッセージを特定する必要があります。

- **debugInfo**:存在する場合、エラー メッセージの `debugInfo` プロパティは、エラーの根本原因を理解するために使用できる追加情報を提供します。

> [!NOTE]
> を使用 `console.log()` してコンソールにエラー メッセージを出力する場合、それらのメッセージはサーバーでのみ表示されます。 エンド ユーザーは、アドイン作業ウィンドウや Office アプリケーション内のどこにもこれらのエラー メッセージを表示しません。 ユーザーにエラーを報告するには、「 [エラー通知](#error-notifications)」を参照してください。

## <a name="error-codes-and-messages"></a>エラー コードとメッセージ

次の表に、アプリケーション固有の API から返される可能性があるエラーを示します。

> [!NOTE]
> 次の表に、アプリケーション固有の API の使用中に発生する可能性があるエラー メッセージを示します。 Common API を使用している場合は、「 [Office Common API エラー コード](../reference/javascript-api-for-office-error-codes.md) 」を参照して、関連するエラー メッセージについて説明します。

|エラー コード | エラー メッセージ | メモ |
|:----------|:--------------|:------|
|`AccessDenied` |要求された操作を実行できません。|*なし。* |
|`ActivityLimitReached`|アクティビティの制限に達しました。|*なし。* |
|`ApiNotAvailable`|要求された API は使用できません。|*なし。* |
|`ApiNotFound`|使用しようとしている API が見つかりませんでした。 新しいバージョンの Office アプリケーションで使用できる場合があります。 詳細については、「 [Office アドインの Office クライアント アプリケーションとプラットフォームの可用性](/javascript/api/requirement-sets) 」を参照してください。|*なし。* |
|`BadPassword`|指定したパスワードが正しくありません。|*なし。* |
|`Conflict`|競合のため、要求を処理できませんでした。|*なし。* |
|`ContentLengthRequired`|`Content-length` HTTP ヘッダーがありません。|*なし。* |
|`GeneralException`|要求の処理中に内部エラーが発生しました。|*なし。* |
|`InsertDeleteConflict`|試行された挿入操作または削除操作で競合が発生しました。|*なし。* |
|`InvalidArgument` |引数が無効であるか、存在しません。または形式が正しくありません。|*なし。* |
|`InvalidBinding` |このオブジェクトのバインドは、以前の更新プログラムが原因で無効になっています。|*なし。* |
|`InvalidOperation`|試行された操作は、このオブジェクトでは無効です。|*なし。* |
|`InvalidReference`|この参照は、現在の操作に対して無効です。|*なし。* |
|`InvalidRequest`  |要求を処理できません。|*なし。* |
|`InvalidSelection`|現在の選択内容は、この操作では無効です。|*なし。* |
|`ItemAlreadyExists`|作成中のリソースはすでに存在しています。|*なし。* |
|`ItemNotFound` |要求されたリソースは存在しません。|*なし。* |
|`MemoryLimitReached`|メモリ制限に達しました。 アクションを完了できませんでした。|*なし。* |
|`NotImplemented`|要求された機能は実装されていません。| これは、API がプレビュー中であるか、特定のプラットフォーム (オンラインのみなど) でのみサポートされていることを意味する可能性があります。 詳細については、「 [Office アドインの Office クライアント アプリケーションとプラットフォームの可用性](/javascript/api/requirement-sets) 」を参照してください。|
|`RequestAborted`|実行時に要求が中止されました。|*なし。* |
|`RequestPayloadSizeLimitExceeded`|要求ペイロード のサイズが制限を超えています。 詳細については、 [Office アドインのリソース制限とパフォーマンスの最適化に関する](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) 記事を参照してください。| このエラーは、Office on the webでのみ発生します。|
|`ResponsePayloadSizeLimitExceeded`|応答ペイロードのサイズが制限を超えています。 詳細については、 [Office アドインのリソース制限とパフォーマンスの最適化に関する](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) 記事を参照してください。|  このエラーは、Office on the webでのみ発生します。|
|`ServiceNotAvailable`|サービスを利用できません。|*なし。* |
|`Unauthenticated` |必要な認証情報が見つからないか、無効です。|*なし。* |
|`UnsupportedFeature`|ソース ワークシートにサポートされていない機能が 1 つ以上含まれているため、操作が失敗しました。|*なし。* |
|`UnsupportedOperation`|試行中の操作はサポートされていません。|*なし。* |

### <a name="excel-specific-error-codes-and-messages"></a>Excel 固有のエラー コードとメッセージ

|エラー コード | エラー メッセージ | メモ |
|:----------|:--------------|:------|
|`EmptyChartSeries`|グラフ系列が空であるため、試行された操作は失敗しました。|*なし。* |
|`FilteredRangeConflict`|操作が試行されると、フィルター処理された範囲との競合が発生します。|*なし。* |
|`FormulaLengthExceedsLimit`|適用された数式のバイトコードが最大長制限を超えています。 32 ビット コンピューターの Office の場合、バイトコードの長さの制限は 16384 文字です。 64 ビット コンピューターでは、バイトコードの長さの制限は 32768 文字です。| このエラーは、Excel on the webとデスクトップの両方で発生します。|
|`GeneralException`|*各種。*|データ型 API は、動的エラー メッセージを含むエラーを返 `GeneralException` します。 これらのメッセージは、エラーの原因であるセルと、エラーの原因となっている問題を参照します。"セル A1 に必要なプロパティ `type`がありません。|
|`InactiveWorkbook`|複数のブックが開き、この API によって呼び出されているブックのフォーカスが失われたため、操作が失敗しました。|*なし。* |
|`InvalidOperationInCellEditMode`|Excel が [編集] セル モードの場合、操作は使用できません。 **Enter** キーまたは **Tab** キーを使用するか、別のセルを選択して編集モードを終了してから、もう一度やり直してください。|*なし。* |
|`MergedRangeConflict`|操作を完了できません。 テーブルは、別のテーブル、ピボットテーブル レポート、クエリ結果、マージされたセル、または XML マップと重複できません。|*なし。* |
|`NonBlankCellOffSheet`|Microsoft Excel では、ワークシートの末尾から空でないセルをプッシュするため、新しいセルを挿入できません。 これらの空でないセルは空で表示されますが、空白の値、書式、または数式が含まれる場合があります。 挿入する対象の領域を作るために十分な行または列を削除してから、もう一度やり直してください。|*なし。* |
|`OperationCellsExceedLimit`|試行された操作は、33554000 セルの制限を超える値に影響します。| このエラーがトリガーされた `TableColumnCollection.add API` 場合は、ワークシート内に意図しないデータがなく、テーブルの外部にないことを確認します。 特に、ワークシートの右端の列のデータを確認します。 意図しないデータを削除して、このエラーを解決します。 操作で処理されるセルの数を確認する 1 つの方法は、次の計算を実行することです。 `(number of table rows) x (16383 - (number of table columns))` 数値 16383 は、Excel でサポートされる列の最大数です。 <br><br>このエラーは、Excel on the webでのみ発生します。 |
|`PivotTableRangeConflict`|操作が試行されると、ピボットテーブル範囲との競合が発生します。|*なし。* |
|`RangeExceedsLimit`|範囲内のセル数が、サポートされている最大数を超えています。 詳細については、 [Office アドインのリソース制限とパフォーマンスの最適化に関する](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) 記事を参照してください。|*なし。* |
|`RefreshWorkbookLinksBlocked`|ユーザーに外部ブック リンクを更新するアクセス許可が付与されていないため、操作が失敗しました。|*なし。* |
|`UnsupportedSheet`|このシートの種類は、マクロ シートまたはグラフ シートであるため、この操作をサポートしていません。|*なし。* |

### <a name="word-specific-error-codes-and-messages"></a>Word 固有のエラー コードとメッセージ

|エラー コード | エラー メッセージ | メモ |
|:----------|:--------------|:------|
|`SearchDialogIsOpen`|検索ダイアログが開いています。|*なし。* |
|`SearchStringInvalidOrTooLong`|検索文字列が無効であるか、長すぎます。| 検索文字列の最大値は 255 文字です。 |

## <a name="error-notifications"></a>エラー通知

ユーザーにエラーを報告する方法は、使用している UI システムによって異なります。 ui システムとしてReactを使用している場合は、Fluent UI コンポーネントとデザイン要素を使用します。 この [Fluent UI ページ](https://developer.microsoft.com/fluentui#/controls/web)から適切なコントロールを選択します。 エラー メッセージは、メッセージ バー、ダイアログ、またはモーダルで伝達することをお勧めします。 エラーがユーザーの入力にある場合は、入力コントロールの近くでエラーを太字の赤で表示します。 サンプル[の Office-Add-in-Microsoft-Graph-React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/0706cc67645675a48747f1fec1b1e5722b575b11/Samples/auth/Office-Add-in-Microsoft-Graph-React) では、MessageBar 要素を使用し、アドイン作業ウィンドウの [パーソナリティ] メニューを考慮するように変更します。

UI にReactを使用していない場合は、HTML と JavaScript で直接実装されている古い Fabric UI コンポーネントの使用を検討してください。 テンプレートの例としては、 [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates) リポジトリがあります。 ダイアログとナビゲーション のサブフォルダーを特に見てみましょう。 [Excel-Add-in-SalesLeads のサンプルでは、](https://github.com/OfficeDev/Excel-Add-in-SalesLeads)メッセージ バナーを使用します。

## <a name="see-also"></a>関連項目

- [OfficeExtension.Error オブジェクト](/javascript/api/office/officeextension.error)
- [Office の一般的な API エラー コード](../reference/javascript-api-for-office-error-codes.md)
