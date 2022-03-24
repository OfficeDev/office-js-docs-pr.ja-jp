---
title: PowerPoint JavaScript API 要件セット 1.3
description: PowerPointApi 1.3 要件セットの詳細。
ms.date: 12/14/2021
ms.prod: powerpoint
ms.localizationpriority: medium
ms.openlocfilehash: 185ece64559d124d8af7c4051d54267da7b11542
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746487"
---
# <a name="whats-new-in-powerpoint-javascript-api-13"></a>JavaScript API 1.3 PowerPoint新機能

PowerPointApi 1.3 では、スライド管理とカスタム タグ付けのサポートが追加されました。

最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [スライドの管理](../../powerpoint/add-slides.md) | スライドの追加とスライド レイアウトとスライド マスターの管理のサポートを追加します。 | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| [Tags](../../powerpoint/tagging-presentations-slides-shapes.md) | アドインがキーと値のペアの形式でカスタム メタデータを添付できます。 | [Tag](/javascript/api/powerpoint/powerpoint.tag) |

## <a name="api-list"></a>API リスト

次の表に、JavaScript API PowerPointセット 1.3 の一覧を示します。 すべての JavaScript API (プレビュー API PowerPoint以前にリリースされた API を含む) の完全な一覧については、[JavaScript API PowerPoint参照してください](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#powerpoint-powerpoint-addslideoptions-layoutid-member)|新しいスライドに使用するスライド レイアウトの ID を指定します。|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#powerpoint-powerpoint-addslideoptions-slidemasterid-member)|新しいスライドに使用するスライド マスターの ID を指定します。|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-slidemasters-member)|プレゼンテーション内のオブジェクト `SlideMaster` のコレクションを返します。|
||[tags](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-tags-member)|プレゼンテーションに添付されているタグのコレクションを返します。|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[delete()](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-delete-member(1))|図形コレクションから図形を削除します。|
||[id](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-id-member)|図形の一意の ID を取得します。|
||[tags](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-tags-member)|図形内のタグのコレクションを返します。|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getcount-member(1))|コレクション内の図形の数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getitem-member(1))|一意の ID を使用して図形を取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getitemat-member(1))|コレクション内の 0 から始るインデックスを使用して図形を取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getitemornullobject-member(1))|一意の ID を使用して図形を取得します。|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[レイアウト](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-layout-member)|スライドのレイアウトを取得します。|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-shapes-member)|スライド内の図形のコレクションを返します。|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-slidemaster-member)|スライドの `SlideMaster` 既定のコンテンツを表すオブジェクトを取得します。|
||[tags](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-tags-member)|スライド内のタグのコレクションを返します。|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint.AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-add-member(1))|コレクションの最後に新しいスライドを追加します。|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-id-member)|スライド レイアウトの一意の ID を取得します。|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-name-member)|スライド レイアウトの名前を取得します。|
||[shapes](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-shapes-member)|スライド レイアウト内の図形のコレクションを返します。|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getcount-member(1))|コレクション内のレイアウトの数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getitem-member(1))|一意の ID を使用してレイアウトを取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getitemat-member(1))|コレクション内の 0 から始るインデックスを使用してレイアウトを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getitemornullobject-member(1))|一意の ID を使用してレイアウトを取得します。|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-id-member)|スライド マスターの一意の ID を取得します。|
||[レイアウト](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-layouts-member)|スライドのスライド マスターによって提供されるレイアウトのコレクションを取得します。|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-name-member)|スライド マスターの一意の名前を取得します。|
||[shapes](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-shapes-member)|スライド マスターの図形のコレクションを返します。|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getcount-member(1))|コレクション内のスライド マスターの数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getitem-member(1))|一意の ID を使用してスライド マスターを取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getitemat-member(1))|コレクション内の 0 から始るインデックスを使用してスライド マスターを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getitemornullobject-member(1))|一意の ID を使用してスライド マスターを取得します。|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#powerpoint-powerpoint-tag-key-member)|タグの一意の ID を取得します。|
||[value](/javascript/api/powerpoint/powerpoint.tag#powerpoint-powerpoint-tag-value-member)|タグの値を取得します。|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-add-member(1))|コレクションの末尾に新しいタグを追加します。|
||[delete(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-delete-member(1))|このコレクション内の指定されたタグ `key` を削除します。|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getcount-member(1))|コレクション内のタグの数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getitem-member(1))|一意の ID を使用してタグを取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getitemat-member(1))|コレクション内の 0 から始るインデックスを使用してタグを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getitemornullobject-member(1))|一意の ID を使用してタグを取得します。|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [PowerPoint JavaScript API リファレンス ドキュメント](/javascript/api/powerpoint?view=powerpoint-js-1.3&preserve-view=true)
- [PowerPoint JavaScript API の要件セット](powerpoint-api-requirement-sets.md)
