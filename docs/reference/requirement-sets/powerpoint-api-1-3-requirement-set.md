---
title: PowerPoint JavaScript API 要件セット 1.3
description: PowerPointApi 1.3 要件セットの詳細。
ms.date: 12/14/2021
ms.prod: powerpoint
ms.localizationpriority: medium
ms.openlocfilehash: 74f17923f7bc8b26416c39bdbbeea9cc0a94029a
ms.sourcegitcommit: e44a8109d9323aea42ace643e11717fb49f40baa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/15/2021
ms.locfileid: "61514248"
---
# <a name="whats-new-in-powerpoint-javascript-api-13"></a>JavaScript API 1.3 PowerPoint新機能

PowerPointApi 1.3 では、スライド管理とカスタム タグ付けのサポートが追加されました。

最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [スライドの管理](../../powerpoint/add-slides.md) | スライドの追加とスライド レイアウトとスライド マスターの管理のサポートを追加します。 | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| [Tags](../../powerpoint/tagging-presentations-slides-shapes.md) | アドインがキーと値のペアの形式でカスタム メタデータを添付できます。 | [Tag](/javascript/api/powerpoint/powerpoint.tag) |

## <a name="api-list"></a>API リスト

次の表に、JavaScript API PowerPointセット 1.3 の一覧を示します。 すべての JavaScript API (プレビュー API PowerPoint以前にリリースされた API を含む) の完全な一覧については[、JavaScript](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)API PowerPoint参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutId)|新しいスライドに使用するスライド レイアウトの ID を指定します。|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slideMasterId)|新しいスライドに使用するスライド マスターの ID を指定します。|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slideMasters)|プレゼンテーション内のオブジェクト `SlideMaster` のコレクションを返します。|
||[tags](/javascript/api/powerpoint/powerpoint.presentation#tags)|プレゼンテーションに添付されているタグのコレクションを返します。|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[delete()](/javascript/api/powerpoint/powerpoint.shape#delete__)|図形コレクションから図形を削除します。|
||[id](/javascript/api/powerpoint/powerpoint.shape#id)|図形の一意の ID を取得します。|
||[tags](/javascript/api/powerpoint/powerpoint.shape#tags)|図形内のタグのコレクションを返します。|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getCount__)|コレクション内の図形の数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getItem_key_)|一意の ID を使用して図形を取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemAt_index_)|コレクション内の 0 から始るインデックスを使用して図形を取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemOrNullObject_id_)|一意の ID を使用して図形を取得します。|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[レイアウト](/javascript/api/powerpoint/powerpoint.slide#layout)|スライドのレイアウトを取得します。|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|スライド内の図形のコレクションを返します。|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slideMaster)|スライドの `SlideMaster` 既定のコンテンツを表すオブジェクトを取得します。|
||[tags](/javascript/api/powerpoint/powerpoint.slide#tags)|スライド内のタグのコレクションを返します。|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint.AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_)|コレクションの最後に新しいスライドを追加します。|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|スライド レイアウトの一意の ID を取得します。|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|スライド レイアウトの名前を取得します。|
||[shapes](/javascript/api/powerpoint/powerpoint.slidelayout#shapes)|スライド レイアウト内の図形のコレクションを返します。|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getCount__)|コレクション内のレイアウトの数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItem_key_)|一意の ID を使用してレイアウトを取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemAt_index_)|コレクション内の 0 から始るインデックスを使用してレイアウトを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemOrNullObject_id_)|一意の ID を使用してレイアウトを取得します。|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|スライド マスターの一意の ID を取得します。|
||[レイアウト](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|スライドのスライド マスターによって提供されるレイアウトのコレクションを取得します。|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|スライド マスターの一意の名前を取得します。|
||[shapes](/javascript/api/powerpoint/powerpoint.slidemaster#shapes)|スライド マスターの図形のコレクションを返します。|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getCount__)|コレクション内のスライド マスターの数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItem_key_)|一意の ID を使用してスライド マスターを取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemAt_index_)|コレクション内の 0 から始るインデックスを使用してスライド マスターを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemOrNullObject_id_)|一意の ID を使用してスライド マスターを取得します。|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#key)|タグの一意の ID を取得します。|
||[value](/javascript/api/powerpoint/powerpoint.tag#value)|タグの値を取得します。|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_)|コレクションの末尾に新しいタグを追加します。|
||[delete(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#delete_key_)|このコレクション内の指定されたタグ `key` を削除します。|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#getCount__)|コレクション内のタグの数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItem_key_)|一意の ID を使用してタグを取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemAt_index_)|コレクション内の 0 から始るインデックスを使用してタグを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemOrNullObject_key_)|一意の ID を使用してタグを取得します。|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [PowerPoint JavaScript API リファレンス ドキュメント](/javascript/api/powerpoint?view=powerpoint-js-1.3&preserve-view=true)
- [PowerPoint JavaScript API の要件セット](powerpoint-api-requirement-sets.md)
