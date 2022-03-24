---
title: PowerPoint JavaScript API 要件セット 1.2
description: PowerPointApi 1.2 要件セットの詳細。
ms.date: 01/27/2021
ms.prod: powerpoint
ms.localizationpriority: medium
ms.openlocfilehash: 0e8ae36a7a137db1645051628aa90a451caf4d56
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744226"
---
# <a name="whats-new-in-powerpoint-javascript-api-12"></a>JavaScript API 1.2 PowerPoint新機能

PowerPointApi 1.2 では、別のプレゼンテーションから現在のプレゼンテーションにスライドを挿入し、スライドを削除するためのサポートが追加されました。

最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [スライドの挿入と削除](../../powerpoint/insert-slides-into-presentation.md) | 別のプレゼンテーションから現在のプレゼンテーションに既存のスライドを挿入し、スライドを削除できます。 | [Slide.delete](/javascript/api/powerpoint/powerpoint.slide#delete--), [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-insertslidesfrombase64-member(1))|

## <a name="api-list"></a>API リスト

次の表に、JavaScript API PowerPointセット 1.2 の一覧を示します。 すべての JavaScript API (プレビュー API PowerPoint以前にリリースされた API を含む) の完全な一覧については、[JavaScript API PowerPoint参照してください](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[書式設定](/javascript/api/powerpoint/powerpoint.insertslideoptions#powerpoint-powerpoint-insertslideoptions-formatting-member)|スライド挿入時に使用する書式を指定します。|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#powerpoint-powerpoint-insertslideoptions-sourceslideids-member)|現在のプレゼンテーションに挿入されるソース プレゼンテーションのスライドを指定します。|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#powerpoint-powerpoint-insertslideoptions-targetslideid-member)|プレゼンテーション内で新しいスライドを挿入する場所を指定します。|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint.InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-insertslidesfrombase64-member(1))|指定したスライドをプレゼンテーションから現在のプレゼンテーションに挿入します。|
||[スライド](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-slides-member)|プレゼンテーション内のスライドの順序付きコレクションを返します。|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-delete-member(1))|プレゼンテーションからスライドを削除します。|
||[id](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-id-member)|スライドの一意の ID を取得します。|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getcount-member(1))|コレクション内のスライドの数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitem-member(1))|一意の ID を使用してスライドを取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitemat-member(1))|コレクション内の 0 から始るインデックスを使用してスライドを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitemornullobject-member(1))|一意の ID を使用してスライドを取得します。|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [PowerPoint JavaScript API リファレンス ドキュメント](/javascript/api/powerpoint?view=powerpoint-js-1.2&preserve-view=true)
- [PowerPoint JavaScript API の要件セット](powerpoint-api-requirement-sets.md)
