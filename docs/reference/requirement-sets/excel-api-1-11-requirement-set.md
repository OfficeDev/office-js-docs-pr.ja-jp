---
title: Excel JavaScript API 要件セット 1.11
description: ExcelApi 1.11 要件セットの詳細。
ms.date: 04/01/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 7beabf94523164280d29c7f34c8b2c1003698bcc
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51650843"
---
# <a name="whats-new-in-excel-javascript-api-111"></a>Excel JavaScript API 1.11 の新機能

ExcelApi 1.11 では、コメントとブック レベルのコントロール (ブックの保存や閉じるなど) のサポートが強化されました。 また、ローカライズのアカウントに役立つカルチャ設定へのアクセスも追加されました。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| コメント [メンション](../../excel/excel-add-ins-comments.md#mentions) |コメントを使用して他のブック ユーザーにタグを付け、通知します。 | [Comment](/javascript/api/excel/excel.comment)、 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| コメント [の解決](../../excel/excel-add-ins-comments.md#resolve-comment-threads) | コメント スレッドを解決し、解決状態を取得します。 | [Comment](/javascript/api/excel/excel.comment) |
| [カルチャの設定](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | 数値の書式設定など、ブックの文化システム設定を取得します。 | [CultureInfo](/javascript/api/excel/excel.cultureinfo)、 [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [アプリケーション](/javascript/api/excel/excel.application) |
| [切り取りと貼り付け (moveTo)](../../excel/excel-add-ins-ranges-cut-copy-paste.md) | 範囲の Excel の切り取りおよび貼り付け機能をレプリケートします。 | [Range](/javascript/api/excel/excel.range) |
| ブックを[保存](../../excel/excel-add-ins-workbooks.md#save-the-workbook)して[閉じる](../../excel/excel-add-ins-workbooks.md#close-the-workbook) | ブックを保存して閉じます。 | [Workbook](/javascript/api/excel/excel.workbook) |
| ワークシート のイベント | ワークシートの計算と非表示の行に関するその他のイベントとイベント情報。 | [WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs), [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs) |

## <a name="api-list"></a>API リスト

次の表に、Excel JavaScript API 要件セット 1.11 の API を示します。 Excel JavaScript API 要件セット 1.11 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット [1.11](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)以前の Excel API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|現在のシステム カルチャ設定に基づく情報を提供します。|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|数値の小数点として使用される文字列を取得します。|
||[ThousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|数値の 10 進数の左側に数字のグループを区切る文字列を取得します。|
||[useSystemSeparators](/javascript/api/excel/excel.application#usesystemseparators)|Excel のシステム区切り記号が有効になっている場合に指定します。|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|コメントに記載されているエンティティ (人など) を取得します。|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|リッチ コメント コンテンツ (コメントのメンションなど) を取得します。|
||[解決済み](/javascript/api/excel/excel.comment#resolved)|コメント スレッドの状態。|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|特別に書式設定された文字列とメンションの一覧を使用してコメント コンテンツを更新します。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|指定したセルで、指定した内容の新しいコメントを作成します。|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|コメントに記載されているエンティティの電子メール アドレス。|
||[id](/javascript/api/excel/excel.commentmention#id)|エンティティの ID。|
||[name](/javascript/api/excel/excel.commentmention#name)|コメントに記載されているエンティティの名前。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|コメントに記載されているエンティティ (人など)。|
||[解決済み](/javascript/api/excel/excel.commentreply#resolved)|コメントの返信の状態。|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|豊富なコメント コンテンツ (コメント内のメンションなど)。|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|特別に書式設定された文字列とメンションの一覧を使用してコメント コンテンツを更新します。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: CommentRichContent \| 文字列, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|コメントのコメント返信を作成します。|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|コメント内で言及されているすべてのエンティティ (人など) を含む配列。|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)|コメントのリッチ コンテンツ (例: メンション付きコメント コンテンツ、最初に言及されたエンティティの id 属性が 0、2 番目に指定されたエンティティの id 属性が 1) を指定します。|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|languagecode2-country/regioncode2 形式のカルチャ名 ("zh-cn" や "ja-us" など) を取得します。|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|数値を表示する文化的に適切な形式を定義します。|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|数値の小数点として使用される文字列を取得します。|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|数値の 10 進数の左側に数字のグループを区切る文字列を取得します。|
|[Range](/javascript/api/excel/excel.range)|[moveTo(destinationRange: Range \| string)](/javascript/api/excel/excel.range#moveto-destinationrange-)|セルの値、書式設定、および数式を現在の範囲から移動先の範囲に移動し、それらのセルの古い情報を置き換える。|
|[範囲の形式](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|範囲の書式設定のインデントを調整します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|現在のブックを閉じます。|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|現在のブックを保存します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|特定のワークシートで 1 つ以上の行の非表示状態が変更された場合に発生します。|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|計算が完了した範囲のアドレス。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|特定のワークシートで 1 つ以上の行の非表示状態が変更された場合に発生します。|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|特定のワークシートで変更されたエリアを表す範囲のアドレスを取得します。|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|イベントがトリガーされた方法を表す変更の種類を取得します。|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|データが変更されたワークシートの ID を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
