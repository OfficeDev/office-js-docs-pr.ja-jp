---
title: PowerPointJavaScript API 要件セット 1.2
description: PowerPointApi 1.2 要件セットの詳細。
ms.date: 01/27/2021
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: 971617bc2bd70525fc3d5adf34fc0ad092ae66f9892ed52f0d83053b142caa10
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098690"
---
# <a name="whats-new-in-powerpoint-javascript-api-12"></a>JavaScript API 1.2 PowerPoint新機能

PowerPointApi 1.2 では、別のプレゼンテーションから現在のプレゼンテーションにスライドを挿入し、スライドを削除するためのサポートが追加されました。

最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [スライドの挿入と削除](../../powerpoint/insert-slides-into-presentation.md) | 別のプレゼンテーションから現在のプレゼンテーションに既存のスライドを挿入し、スライドを削除できます。 | [Slide.delete](/javascript/api/powerpoint/powerpoint.slide#delete--)、 [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|

## <a name="api-list"></a>API リスト

次の表に、JavaScript API PowerPointセット 1.2 の一覧を示します。 すべての JavaScript API (プレビュー API PowerPoint以前にリリースされた API を含む) の完全な一覧については[、JavaScript](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)API PowerPoint参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[書式設定](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|スライド挿入時に使用する書式を指定します。|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceSlideIds)|現在のプレゼンテーションに挿入されるソース プレゼンテーションのスライドを指定します。|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetSlideId)|プレゼンテーション内で新しいスライドを挿入する場所を指定します。|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint.InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertSlidesFromBase64_base64File__options_)|指定したスライドをプレゼンテーションから現在のプレゼンテーションに挿入します。|
||[スライド](/javascript/api/powerpoint/powerpoint.presentation#slides)|プレゼンテーション内のスライドの順序付きコレクションを返します。|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete__)|プレゼンテーションからスライドを削除します。|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|スライドの一意の ID を取得します。|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getCount__)|コレクション内のスライドの数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getItem_key_)|一意の ID を使用してスライドを取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_)|コレクション内の 0 から始るインデックスを使用してスライドを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getItemOrNullObject_id_)|一意の ID を使用してスライドを取得します。|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [PowerPointJavaScript API リファレンス ドキュメント](/javascript/api/powerpoint?view=powerpoint-js-1.2&preserve-view=true)
- [PowerPoint JavaScript API の要件セット](powerpoint-api-requirement-sets.md)
