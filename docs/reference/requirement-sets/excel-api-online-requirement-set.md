---
title: Excel JavaScript API のオンラインのみの要件セット
description: ExcelApiOnline の要件セットの詳細
ms.date: 12/05/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad2a3cd627552baeb449397fa917fe10e86ebbaf
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814153"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel JavaScript API のオンラインのみの要件セット

`ExcelApiOnline`要件セットは、web 上の Excel でのみ使用可能な機能を含む特別な要件セットです。 この要件セットの Api は、web ホスト上の Excel の運用 Api (未提出の行動または構造上の変更による影響を受けない) と見なされます。 `ExcelApiOnline`他のプラットフォーム (Windows、Mac、iOS) の場合は "preview" Api と見なされますが、これらのプラットフォームではサポートされていない場合があります。

`ExcelApiOnline`要件セットの api がすべてのプラットフォームでサポートされている場合は、次にリリースされる`ExcelApi 1.[NEXT]`要件セット () に追加されます。 新しい要件が公開されると、これらの Api はから`ExcelApiOnline`削除されます。 この点は、プレビューからリリースに移行する API と同様に、昇格プロセスと考えることができます。

> [!IMPORTANT]
> `ExcelApiOnline`は、最新の番号付き要件セットのスーパーセットです。

> [!IMPORTANT]
> `ExcelApiOnline 1.1`は、オンライン専用 Api の唯一のバージョンです。 これは、web 上の Excel では、最新バージョンのユーザーが常に1つのバージョンを使用できるためです。

## <a name="recommended-usage"></a>推奨される使用法

Api `ExcelApiOnline`は web 上の Excel でのみサポートされているため、アドインでは、これらの api を呼び出す前に要件セットがサポートされているかどうかを確認する必要があります。 これにより、別のプラットフォームでオンラインのみの API を呼び出すことを回避できます。

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

クロスプラットフォームの要件セットに含まれる API は、 `isSetSupported`チェックを削除または編集する必要があります。 これにより、他のプラットフォームでアドインの機能が有効になります。 この変更を行うときは、これらのプラットフォームで機能をテストしてください。

> [!IMPORTANT]
> マニフェストでライセンス認証`ExcelApiOnline 1.1`の要件として指定することはできません。 [Set 要素](../manifest/set.md)で使用する有効な値ではありません。

## <a name="api-list"></a>API リスト

次の Api は、現在、 `ExcelApiOnline 1.1`要件セットの一部として web 上の Excel で使用できます。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|コメントに記載されているエンティティ (人物など) を取得します。|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|リッチコメントの内容 (コメントに含まれるメンションなど) を取得します。 この文字列は、エンドユーザーに表示されることを意図したものではありません。 アドインでは、リッチコメントコンテンツを解析するためにのみ使用する必要があります。|
||[updateMentions (contentWithMentions ション: CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|特別に書式設定された文字列とメンションの一覧を使用して、コメントの内容を更新します。|
|[コメントについて](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|コメントで言及されているエンティティの電子メールアドレスを取得または設定します。|
||[id](/javascript/api/excel/excel.commentmention#id)|エンティティの id を取得または設定します。 これは、のいずれかの`CommentRichContent.richContent`id と一致します。|
||[name](/javascript/api/excel/excel.commentmention#name)|コメントで言及されているエンティティの名前を取得または設定します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|コメントに記載されているエンティティ (人物など) を取得します。|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|リッチコメントの内容 (コメントに含まれるメンションなど) を取得します。 この文字列は、エンドユーザーに表示されることを意図したものではありません。 アドインでは、リッチコメントコンテンツを解析するためにのみ使用する必要があります。|
||[updateMentions (contentWithMentions ション: CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|特別に書式設定された文字列とメンションの一覧を使用して、コメントの内容を更新します。|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|コメント内で記述されているすべてのエンティティ (人物など) を含む配列。|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)||
|[Range](/javascript/api/excel/excel.range)|[moveTo (destinationRange: Range \| string)](/javascript/api/excel/excel.range#moveto-destinationrange-)|セルの値、書式設定、および数式を現在の範囲から移動先の範囲に移動し、そのセルの古い情報を置き換えます。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent (金額: 数値)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|範囲の書式のインデントを調整します。 [インデント] の値の範囲は 0 ~ 250 で、文字単位です。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-online)
- [Excel JavaScript プレビュー API](./excel-preview-apis.md)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)