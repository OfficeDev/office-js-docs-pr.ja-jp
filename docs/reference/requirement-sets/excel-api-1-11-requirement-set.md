---
title: Excel JavaScript API 要件セット1.11
description: ExcelApi 1.11 の要件セットの詳細
ms.date: 05/06/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ab9fde262640aa243aaf2b88767225505e08b3b7
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612095"
---
# <a name="whats-new-in-excel-javascript-api-111"></a>Excel JavaScript API 1.11 の新機能

ExcelApi 1.11 は、コメントおよびブックレベルのコントロール (ブックの保存やクローズなど) のサポートが強化されました。 また、カルチャ設定へのアクセスを追加して、ローカライズに役立てることができます。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| コメント[メンション](../../excel/excel-add-ins-comments.md#mentions) |コメントを使用して、他のブックユーザーにタグ付けして通知します。 | [Comment](/javascript/api/excel/excel.comment)、 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| コメント[解決](../../excel/excel-add-ins-comments.md#resolve-comment-threads) | コメントスレッドを解決し、解決状態を取得します。 | [Comment](/javascript/api/excel/excel.comment) |
| [カルチャ設定](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | ブックのカルチャシステム設定 (数値の書式設定など) を取得します。 | [CultureInfo](/javascript/api/excel/excel.cultureinfo)、 [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [Application](/javascript/api/excel/excel.application) |
| [切り取りと貼り付け (moveTo)](../../excel/excel-add-ins-ranges-advanced.md#cut-copy-and-paste) | Excel の範囲のカットアンドペースト機能をレプリケートします。 | [Range](/javascript/api/excel/excel.range) |
| ブックを[保存](../../excel/excel-add-ins-workbooks.md#save-the-workbook)して[閉じる](../../excel/excel-add-ins-workbooks.md#close-the-workbook) | ブックを保存して閉じます。 | [Workbook](/javascript/api/excel/excel.workbook) |
| ワークシートイベント | ワークシートの計算および非表示の行に関するその他のイベントおよびイベント情報。 | [WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)、 [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs) |

## <a name="api-list"></a>API リスト

次の表に、Excel JavaScript API 要件セット1.11 の Api を示します。 Excel JavaScript API 要件セット1.11 またはそれ以前でサポートされているすべての Api の API リファレンスドキュメントを表示するには、「[要件セット1.11 またはそれ以前の Excel api](/javascript/api/excel?view=excel-js-1.11)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|現在のシステムのカルチャ設定に基づく情報を提供します。 これには、カルチャ名、数値形式、およびその他のカルチャに依存する設定が含まれます。|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|数値の小数点の記号として使用される文字列を取得します。 これは、Excel のローカル設定に基づいています。|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|数値の小数点の左側にある数字のグループを区切るために使用される文字列を取得します。 これは、Excel のローカル設定に基づいています。|
||[useSystemSeparators](/javascript/api/excel/excel.application#usesystemseparators)|Excel のシステム区切り記号を有効にするかどうかを指定します。|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|コメントに記載されているエンティティ (ユーザーなど) を取得します。|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|リッチコメントの内容 (コメント内のメンションなど) を取得します。 この文字列は、エンドユーザーに表示されることを意図したものではありません。 アドインでは、リッチコメントコンテンツを解析するためにのみ使用する必要があります。|
||[解析](/javascript/api/excel/excel.comment#resolved)|コメントスレッドの状態。 値 "true" は、コメントスレッドが解決されることを意味します。|
||[updateMentions (contentWithMentions ション: CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|特別に書式設定された文字列とメンションの一覧を使用して、コメントの内容を更新します。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add (cellAddress: Range \| string, content: CommentRichContent \| String, contenttype?: Excel)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|指定したセルで、指定した内容の新しいコメントを作成します。 `InvalidArgument`指定した範囲が1つのセルより大きい場合は、エラーがスローされます。|
|[コメントについて](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|コメントに記載されているエンティティの電子メールアドレス。|
||[id](/javascript/api/excel/excel.commentmention#id)|エンティティの id。 Id は、のいずれかの id と一致し `CommentRichContent.richContent` ます。|
||[name](/javascript/api/excel/excel.commentmention#name)|Comment で言及されているエンティティの名前。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|コメントに記載されているエンティティ (ユーザーなど)。|
||[解析](/javascript/api/excel/excel.commentreply#resolved)|コメントの返信状態。 値 "true" は、応答が解決された状態であることを意味します。|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|リッチコメントの内容 (コメント内のメンションなど)。 この文字列は、エンドユーザーに表示されることを意図したものではありません。 アドインでは、リッチコメントコンテンツを解析するためにのみ使用する必要があります。|
||[updateMentions (contentWithMentions ション: CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|特別に書式設定された文字列とメンションの一覧を使用して、コメントの内容を更新します。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add (content: CommentRichContent \| string, contenttype?: Excel)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|コメントのコメント返信を作成します。|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|コメント内で言及されているすべてのエンティティ (人物など) を含む配列。|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)|コメントのリッチコンテンツを指定します (たとえば、メンションを含むコメントコンテンツ、最初に説明したエンティティの id 属性は0、2番目に指定したエンティティの id 属性は1です)。|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|カルチャ名を languagecode2-country/regioncode2 の形式で取得します (例: "zh-cn-cn" または "en-us")。 これは、現在のシステム設定に基づいています。|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|数字を表示するためのカルチャに適した形式を定義します。 これは、現在のシステムのカルチャ設定に基づいています。|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|数値の小数点の記号として使用される文字列を取得します。 これは、現在のシステム設定に基づいています。|
||[番号 Groupseparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|数値の小数点の左側にある数字のグループを区切るために使用される文字列を取得します。 これは、現在のシステム設定に基づいています。|
|[Range](/javascript/api/excel/excel.range)|[moveTo (destinationRange: Range \| string)](/javascript/api/excel/excel.range#moveto-destinationrange-)|セルの値、書式設定、および数式を現在の範囲から移動先の範囲に移動し、そのセルの古い情報を置き換えます。|
|[範囲の形式](/javascript/api/excel/excel.rangeformat)|[adjustIndent (金額: 数値)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|範囲の書式のインデントを調整します。 [インデント] の値の範囲は 0 ~ 250 で、文字単位です。|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|現在のブックを閉じます。|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|現在のブックを保存します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|特定のワークシートで、1つまたは複数の行の非表示の状態が変更されたときに発生します。|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|計算を完了した範囲のアドレス。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|特定のワークシートで、1つまたは複数の行の非表示の状態が変更されたときに発生します。|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|特定のワークシートで変更されたエリアを表す範囲のアドレスを取得します。|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|イベントがトリガーされた方法を表す変更の種類を取得します。 詳細は「`Excel.RowHiddenChangeType`」をご覧ください。|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|データが変更されたワークシートの ID を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.11)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
