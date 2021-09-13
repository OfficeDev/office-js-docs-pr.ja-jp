---
title: Word JavaScript API 要件セット 1.2
description: WordApi 1.2 要件セットの詳細
ms.date: 11/09/2020
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: de293cf67bbb452fe3c2b8c5de4896adf5cf7a43
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154822"
---
# <a name="whats-new-in-word-javascript-api-12"></a>Word JavaScript API 1.2 の新機能

WordApi 1.2 では、インライン画像のサポートが追加されました。

## <a name="api-list"></a>API リスト

次の表に、Word JavaScript API 要件セット 1.2 の API の一覧を示します。 Word JavaScript API 要件セット 1.2 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、「要件セット [1.2](/javascript/api/word?view=word-js-1.2&preserve-view=true)以前の Word API」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|画像を本文の指定された位置に挿入します。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|コンテンツ コントロール内の指定された位置にインライン画像を挿入します。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#delete__)|ドキュメントからインライン画像を削除します。|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertBreak_breakType__insertLocation_)|メイン文書の指定した位置に、区切りを挿入します。|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertFileFromBase64_base64File__insertLocation_)|指定した位置に文書を挿入します。|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertHtml_html__insertLocation_)|指定した位置に HTML を挿入します。|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|指定された位置にインライン画像を挿入します。|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertOoxml_ooxml__insertLocation_)|指定した位置に OOXML を挿入します。|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertParagraph_paragraphText__insertLocation_)|指定した位置に、段落を挿入します。|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertText_text__insertLocation_)|指定した位置にテキストを挿入します。|
||[段落](/javascript/api/word/word.inlinepicture#paragraph)|インライン イメージを含む親段落を取得します。|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.inlinepicture#select_selectionMode_)|インライン画像を選択します。|
|[Range](/javascript/api/word/word.range)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|指定された位置に画像を挿入します。|
||[inlinePictures](/javascript/api/word/word.range#inlinePictures)|範囲に含まれるインライン画像オブジェクトのコレクションを取得します。|

## <a name="see-also"></a>関連項目

- [Word JavaScript API リファレンス ドキュメント](/javascript/api/word)
- [Word JavaScript API の要件セット](word-api-requirement-sets.md)
