---
title: PowerPoint JavaScript プレビュー Api
description: 今後の PowerPoint JavaScript Api についての詳細。
ms.date: 10/26/2020
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: 27a51054f930b560d2d2f9a00fc172394b26830d
ms.sourcegitcommit: a4e09546fd59579439025aca9cc58474b5ae7676
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/27/2020
ms.locfileid: "48774812"
---
# <a name="powerpoint-javascript-preview-apis"></a>PowerPoint JavaScript プレビュー Api

新しい PowerPoint JavaScript Api は最初は "プレビュー" で導入されており、これ以降のテストが行われ、ユーザーのフィードバックが取得された後、特定の番号付き要件の一部となります。

最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| スライドの挿入と削除 | 別のプレゼンテーションから現在のプレゼンテーションに既存のスライドを挿入したり、削除したりすることができます。 | [InsertSlidesFromBase64、プレゼンテーション](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)を[削除](/javascript/api/powerpoint/powerpoint.slide#delete--)します。|

## <a name="api-list"></a>API リスト

次の表に、現在プレビュー中の PowerPoint JavaScript Api を示します。 すべての PowerPoint JavaScript Api (プレビュー Api および以前リリースされた Api を含む) の完全なリストについては、「 [すべての Powerpoint Javascript api](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[書式](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|スライドの挿入時に使用する書式を指定します。|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceslideids)|現在のプレゼンテーションに挿入される、元のプレゼンテーションのスライドを指定します。 これらのスライドは、オブジェクトから取得できる Id で表され `Slide` ます。|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetslideid)|プレゼンテーションのどこに新しいスライドを挿入するかを指定します。|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64 (base64File: string, options?: InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|プレゼンテーションの指定したスライドを現在のプレゼンテーションに挿入します。|
||[スライド](/javascript/api/powerpoint/powerpoint.presentation#slides)|プレゼンテーション内のスライドの順序付けられたコレクションを返します。|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete--)|スライドをプレゼンテーションから削除します。 スライドが存在しない場合は何も実行しません。|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|スライドの一意の ID を取得します。|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getcount--)|コレクション内のスライド数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitem-key-)|一意の ID を使用してスライドを取得します。 スライドが存在しない場合は、例外がスローされます。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemat-index-)|コレクション内の0から始まるインデックスを使用してスライドを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemornullobject-id-)|一意の ID を使用してスライドを取得します。 スライドが存在しない場合は、 `isNullObject` プロパティが設定されているオブジェクトを返し `true` ます。|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [PowerPoint JavaScript API リファレンスドキュメント](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [PowerPoint JavaScript API の要件セット](powerpoint-api-requirement-sets.md)
