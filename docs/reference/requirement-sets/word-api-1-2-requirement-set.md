---
title: Word JavaScript API 要件セット1.2
description: WordApi 1.2 要件セットの詳細
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: ee9bf60a3a944a3a01a2ca5aa10d01958e3d3475
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996425"
---
# <a name="whats-new-in-word-javascript-api-12"></a>Word JavaScript API 1.2 の新機能

WordApi 1.2 インライン画像のサポートが追加されました。

## <a name="api-list"></a>API リスト

次の表に、Word JavaScript API 要件セット1.2 の Api を示します。 Word JavaScript API 要件セット1.2 またはそれ以前のバージョンでサポートされているすべての Api の API リファレンスドキュメントを表示するには、「 [要件セット1.2 またはそれ以前の Word api](/javascript/api/word?view=word-js-1.2&preserve-view=true)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[insertInlinePictureFromBase64 (base64EncodedImage: string, insertLocation: Word InsertLocation)](/javascript/api/word/word.body#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|画像を本文の指定された位置に挿入します。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64 (base64EncodedImage: string, insertLocation: Word InsertLocation)](/javascript/api/word/word.contentcontrol#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|コンテンツ コントロール内の指定された位置にインライン画像を挿入します。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#delete--)|ドキュメントからインライン画像を削除します。|
||[insertBreak (breakType: BreakType, Insertbreak: Word Insertbreak)](/javascript/api/word/word.inlinepicture#insertbreak-breaktype--insertlocation-)|メイン文書の指定した位置に、区切りを挿入します。|
||[insertFileFromBase64 (base64File: string, insertLocation: Word InsertLocation)](/javascript/api/word/word.inlinepicture#insertfilefrombase64-base64file--insertlocation-)|指定した位置に文書を挿入します。|
||[insertHtml (html: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#inserthtml-html--insertlocation-)|指定した位置に HTML を挿入します。|
||[insertInlinePictureFromBase64 (base64EncodedImage: string, insertLocation: Word InsertLocation)](/javascript/api/word/word.inlinepicture#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|指定された位置にインライン画像を挿入します。|
||[insertOoxml (ooxml: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertooxml-ooxml--insertlocation-)|指定した位置に OOXML を挿入します。|
||[insertParagraph (paragraphText: string, Insertparagraph: Word. Insertparagraph)](/javascript/api/word/word.inlinepicture#insertparagraph-paragraphtext--insertlocation-)|指定した位置に、段落を挿入します。|
||[insertText (text: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#inserttext-text--insertlocation-)|指定した位置にテキストを挿入します。|
||[段落](/javascript/api/word/word.inlinepicture#paragraph)|インライン イメージを含む親段落を取得します。|
||[select (selectionMode?:. SelectionMode)](/javascript/api/word/word.inlinepicture#select-selectionmode-)|インライン画像を選択します。|
|[Range](/javascript/api/word/word.range)|[insertInlinePictureFromBase64 (base64EncodedImage: string, insertLocation: Word InsertLocation)](/javascript/api/word/word.range#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|指定された位置に画像を挿入します。|
||[inlinePictures](/javascript/api/word/word.range#inlinepictures)|範囲に含まれるインライン画像オブジェクトのコレクションを取得します。|

## <a name="see-also"></a>関連項目

- [Word JavaScript API リファレンス ドキュメント](/javascript/api/word)
- [Word JavaScript API の要件セット](word-api-requirement-sets.md)
