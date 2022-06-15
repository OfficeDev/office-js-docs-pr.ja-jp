---
title: Excel JavaScript API でのエラー処理
description: ランタイム エラーを考慮する JavaScript API エラー処理ロジックExcelについて説明します。
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6fa5ca0c7ebf9400fcdd83c7bf4eb4b906f2e5b5
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/15/2022
ms.locfileid: "66090832"
---
# <a name="error-handling-with-the-excel-javascript-api"></a>Excel JavaScript API でのエラー処理

Excel JavaScript API を使用してアドインを作成する場合は、実行時エラーを考慮するために、エラー処理ロジックを含めます。 これは、API の非同期性のために重要になります。

> [!NOTE]
> Excel JavaScript API の`sync()`メソッドと非同期の性質の詳細については、「[Office アドインでの JavaScript オブジェクト モデルのExcel](excel-add-ins-core-concepts.md)」を参照してください。

## <a name="best-practices"></a>ベスト プラクティス

[コード サンプル](https://github.com/OfficeDev/Office-Add-in-samples)と[Script Lab](../overview/explore-with-script-lab.md) スニペットでは、すべての呼び出しに`Excel.run` `catch` `Excel.run`、. Excel JavaScript Api を使用してアドインを構築するときには、同じパターンを使用することをお勧めします。

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

Excel JavaScript API 要求が正常に実行されない場合、API は次のプロパティを含むエラー オブジェクトを返します。

- **code**:エラー メッセージの `code` プロパティには、`OfficeExtension.ErrorCodes` または `Excel.ErrorCodes` リストの一部である文字列が含まれます。 たとえば、エラー コード "InvalidReference" は、参照が指定された操作に対して有効でないことを示します。 エラー コードはローカライズされません。

- **message**: エラー メッセージの `message` プロパティには、ローカライズされた文字列のエラーの概要が含まれています。 このエラー メッセージは、エンド ユーザーが使用するためのものではありません。アドインによってエンド ユーザーに表示されるエラー メッセージは、エラー コードと適切なビジネス ロジックを使用して、判断する必要があります。

- **debugInfo**:存在する場合、エラー メッセージの `debugInfo` プロパティは、エラーの根本原因を理解するために使用できる追加情報を提供します。

> [!NOTE]
> コンソールにエラー メッセージを出力するために使用 `console.log()` する場合、これらのメッセージはサーバーでのみ表示されます。 エンド ユーザーは、アドイン作業ウィンドウまたはOffice アプリケーション内の任意の場所にこれらのエラー メッセージを表示しません。 ユーザーにエラーを報告するには、「 [エラー通知](#error-notifications)」を参照してください。

## <a name="error-messages"></a>エラー メッセージ

次の表は、API から返される可能性のあるエラー一覧です。

|エラー コード | エラー メッセージ | メモ |
|:----------|:--------------|:------|
|`AccessDenied` |要求された操作を実行できません。| |
|`ActivityLimitReached`|アクティビティの制限に達しました。| |
|`ApiNotAvailable`|要求された API は使用できません。| |
|`ApiNotFound`|使用しようとしている API が見つかりませんでした。 新しいバージョンのExcelで使用できる場合があります。 詳細については、[Excel JavaScript API 要件セット](/javascript/api/requirement-sets/excel/excel-api-requirement-sets)に関する記事を参照してください。| |
|`BadPassword`|指定したパスワードが正しくありません。| |
|`Conflict`|競合のため、要求を処理できませんでした。| |
|`ContentLengthRequired`|`Content-length` HTTP ヘッダーがありません。| |
|`EmptyChartSeries`|グラフ系列が空であるため、試行された操作は失敗しました。| |
|`FilteredRangeConflict`|試行された操作により、フィルター処理された範囲との競合が発生します。| |
|`FormulaLengthExceedsLimit`|適用された数式のバイトコードが最大長制限を超えています。 32 ビット コンピューターのOfficeの場合、バイトコードの長さの制限は 16384 文字です。 64 ビット コンピューターでは、バイトコードの長さの制限は 32768 文字です。| このエラーは、Excel on the webとデスクトップの両方で発生します。|
|`GeneralException`|要求の処理中に内部エラーが発生しました。| |
|`InactiveWorkbook`|複数のブックが開かれているため、この API によって呼び出されるブックにフォーカスが失われるため、操作は失敗しました。| |
|`InsertDeleteConflict`|試行された挿入操作または削除操作で競合が発生しました。| |
|`InvalidArgument` |引数が無効であるか、存在しません。または形式が正しくありません。| |
|`InvalidBinding` |このオブジェクトのバインドは、以前の更新プログラムが原因で無効になっています。| |
|`InvalidOperation`|試行された操作は、このオブジェクトでは無効です。| |
|`InvalidOperationInCellEditMode`|この操作は、セルの編集モードExcel中は使用できません。 **Enter** キーまたは **Tab** キーを使用するか、別のセルを選択して編集モードを終了してから、もう一度やり直してください。| |
|`InvalidReference`|この参照は、現在の操作に対して無効です。| |
|`InvalidRequest`  |要求を処理できません。| |
|`InvalidSelection`|現在の選択内容は、この操作では無効です。| |
|`ItemAlreadyExists`|作成中のリソースはすでに存在しています。| |
|`ItemNotFound` |要求されたリソースは存在しません。| |
|`MemoryLimitReached`|メモリの制限に達しました。 アクションを完了できませんでした。| |
|`MergedRangeConflict`|操作を完了できません。 テーブルは、別のテーブル、ピボットテーブル レポート、クエリ結果、マージされたセル、または XML マップと重複することはできません。|
|`NonBlankCellOffSheet`|Microsoft Excel、ワークシートの末尾に空でないセルをプッシュするため、新しいセルを挿入できません。 これらの空でないセルは空に見えますが、空白の値、一部の書式設定、または数式があります。 挿入する行または列を十分に削除して、挿入する内容に余裕を持たせた後、もう一度やり直してください。| |
|`NotImplemented`|要求された機能は実装されていません。| |
|`OperationCellsExceedLimit`|試行された操作は、33554000 セルの制限を超える影響を与えます。| このエラーがトリガーされた `TableColumnCollection.add API` 場合は、ワークシート内でテーブルの外部に意図しないデータがないことを確認します。 特に、ワークシートの右端の列でデータを確認します。 このエラーを解決するには、意図しないデータを削除します。 操作が処理するセルの数を確認する 1 つの方法は、次の計算 `(number of table rows) x (16383 - (number of table columns))`を実行することです。 数値 16383 は、Excelがサポートする列の最大数です。 <br><br>このエラーは、Excel on the webでのみ発生します。 |
|`PivotTableRangeConflict`|試行された操作により、ピボットテーブル範囲との競合が発生します。| |
|`RangeExceedsLimit`|範囲内のセル数が、サポートされている最大数を超えています。 詳細については、[Office アドインのリソース制限とパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)に関する記事を参照してください。| |
|`RefreshWorkbookLinksBlocked`|ユーザーが外部ブックリンクを更新するアクセス許可を付与していないため、操作は失敗しました。| |
|`RequestAborted`|実行時に要求が中止されました。| |
|`RequestPayloadSizeLimitExceeded`|要求ペイロードのサイズが制限を超えました。 詳細については、[Office アドインのリソース制限とパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)に関する記事を参照してください。| このエラーは、Excel on the webでのみ発生します。|
|`ResponsePayloadSizeLimitExceeded`|応答ペイロード のサイズが制限を超えました。 詳細については、[Office アドインのリソース制限とパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)に関する記事を参照してください。|  このエラーは、Excel on the webでのみ発生します。|
|`ServiceNotAvailable`|サービスを利用できません。| |
|`Unauthenticated` |必要な認証情報が見つからないか、無効です。| |
|`UnsupportedFeature`|ソース ワークシートにサポートされていない 1 つ以上の機能が含まれているため、操作は失敗しました。| |
|`UnsupportedOperation`|試行中の操作はサポートされていません。| |
|`UnsupportedSheet`|このシートの種類では、マクロ シートまたはグラフ シートであるため、この操作はサポートされていません。| |

> [!NOTE]
> 上記の表は、Excel JavaScript API の使用中に発生する可能性があるエラー メッセージの一覧です。 アプリケーション固有の Excel JavaScript API の代わりに Common API を使用している場合は、[一般的な API エラー コードOffice](../reference/javascript-api-for-office-error-codes.md)参照して、関連するエラー メッセージについて確認してください。

## <a name="error-notifications"></a>エラー通知

ユーザーにエラーを報告する方法は、使用している UI システムによって異なります。 REACTを UI システムとして使用している場合は、Fluent UI コンポーネントとデザイン要素を使用します。 この[Fluent UI ページ](https://developer.microsoft.com/fluentui#/controls/web)から適切なコントロールを選択します。 エラー メッセージは、メッセージ バー、ダイアログ、またはモーダルで伝達することをお勧めします。 エラーがユーザーの入力にある場合は、入力コントロールの近くに太字の赤でエラーを表示します。 サンプル [Office-Add-in-Microsoft-Graph-React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/0706cc67645675a48747f1fec1b1e5722b575b11/Samples/auth/Office-Add-in-Microsoft-Graph-React)では、MessageBar 要素を使用し、アドイン作業ウィンドウのパーソナリティ メニューを考慮するように変更します。

UI にReactを使用していない場合は、HTML と JavaScript で直接実装された古い Fabric UI コンポーネントの使用を検討してください。 テンプレートの例には[、Office-アドイン UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates) リポジトリがあります。 特にダイアログサブフォルダーとナビゲーション サブフォルダーを見てみましょう。 サンプル [Excel-アドイン SalesLeads では](https://github.com/OfficeDev/Excel-Add-in-SalesLeads)、メッセージ バナーが使用されています。

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [OfficeExtension.Error オブジェクト (JavaScript API for Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [Office の一般的な API エラー コード](../reference/javascript-api-for-office-error-codes.md)
