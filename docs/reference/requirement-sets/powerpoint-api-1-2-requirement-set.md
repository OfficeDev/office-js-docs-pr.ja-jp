---
title: PowerPoint JavaScript API 要件セット 1.2
description: PowerPointApi 1.2 要件セットの詳細。
ms.date: 01/08/2021
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: 0f6d1e766de81fef5d071152f6116ab56613ec9d
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2021
ms.locfileid: "49841541"
---
# <a name="whats-new-in-powerpoint-javascript-api-12"></a>PowerPoint JavaScript API 1.2 の新機能

PowerPointApi 1.2 では、別のプレゼンテーションから現在のプレゼンテーションにスライドを挿入し、スライドを削除するためのサポートが追加されました。

最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| スライドの挿入と削除 | 別のプレゼンテーションから現在のプレゼンテーションに既存のスライドを挿入し、スライドを削除できます。 | [](/javascript/api/powerpoint/powerpoint.slide#delete--) [Slide.delete、Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|

## <a name="api-list"></a>API リスト

次の表に、PowerPoint JavaScript API 要件セット 1.2 を示します。 すべての PowerPoint JavaScript API (プレビュー API と以前にリリースされた API を含む) の完全な一覧については、 [すべての PowerPoint JavaScript](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[formatting](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|スライドの挿入時に使用する書式を指定します。|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceslideids)|現在のプレゼンテーションに挿入する元のプレゼンテーションのスライドを指定します。|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetslideid)|プレゼンテーション内の新しいスライドを挿入する場所を指定します。|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint.InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|プレゼンテーションから指定したスライドを現在のプレゼンテーションに挿入します。|
||[slides](/javascript/api/powerpoint/powerpoint.presentation#slides)|プレゼンテーション内のスライドの順序付きコレクションを返します。|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete--)|プレゼンテーションからスライドを削除します。|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|スライドの一意の ID を取得します。|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getcount--)|コレクション内のスライドの数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitem-key-)|一意の ID を使用してスライドを取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemat-index-)|コレクション内の 0 から始るインデックスを使用してスライドを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemornullobject-id-)|一意の ID を使用してスライドを取得します。|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [PowerPoint JavaScript API リファレンス ドキュメント](/javascript/api/powerpoint?view=powerpoint-js-1.2&preserve-view=true)
- [PowerPoint JavaScript API の要件セット](powerpoint-api-requirement-sets.md)
