---
title: Excel JavaScript API 要件セット1.5
description: ExcelApi 1.5 の要件セットの詳細
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 346b5192d6d68046b9365d3159df9c3964a59271
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430850"
---
# <a name="whats-new-in-excel-javascript-api-15"></a>Excel JavaScript API 1.5 の新機能

ExcelApi 1.5 カスタム XML パーツを追加します。 これらには、workbook オブジェクトの [カスタム XML パーツコレクション](/javascript/api/excel/excel.workbook#customxmlparts) を通じてアクセスできます。

## <a name="custom-xml-part"></a>カスタム XML パーツ

* ID を使用してカスタム XML パーツを取得します。
* 名前空間が指定した名前空間に一致する、カスタム XML パーツの新しい範囲のコレクションを取得します。
* パーツに関連付けられている XML 文字列を取得します。
* パーツの ID と名前空間を指定します。
* 新しいカスタム XML パーツをブックに追加します。
* XML パーツ全体を設定します。
* カスタム XML パーツを削除します。
* xpath で識別される要素から、指定された名前を持つ属性を削除します。
* xpath で XML の内容を照会します。
* 属性の挿入、更新、および削除を行います。

## <a name="api-list"></a>API リスト

次の表に、Excel JavaScript API 要件セット1.5 の Api を示します。 Excel JavaScript API 要件セット1.5 またはそれ以前でサポートされているすべての Api の API リファレンスドキュメントを表示するには、「 [要件セット1.5 またはそれ以前の Excel api](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#delete--)|カスタム XML パーツを削除します。|
||[getXml ()](/javascript/api/excel/excel.customxmlpart#getxml--)|カスタム XML パーツのすべての XML コンテンツを取得します。|
||[id](/javascript/api/excel/excel.customxmlpart#id)|カスタム XML パーツの ID。 読み取り専用です。|
||[namespaceUri](/javascript/api/excel/excel.customxmlpart#namespaceuri)|カスタム XML パーツの名前空間 URI。 読み取り専用です。|
||[setXml (xml: string)](/javascript/api/excel/excel.customxmlpart#setxml-xml-)|カスタム XML パーツのすべての XML コンテンツを設定します。|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[add (xml: string)](/javascript/api/excel/excel.customxmlpartcollection#add-xml-)|ブックに新しいカスタム XML パーツを追加します。|
||[getByNamespace (namespaceUri: string)](/javascript/api/excel/excel.customxmlpartcollection#getbynamespace-namespaceuri-)|名前空間が指定した名前空間に一致する、カスタム XML パーツの新しい範囲のコレクションを取得します。|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#getcount--)|コレクションに含まれる CustomXml パーツの数を取得します。|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitem-id-)|ID に基づいて、カスタム XML パーツを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitemornullobject-id-)|ID に基づいて、カスタム XML パーツを取得します。|
||[items](/javascript/api/excel/excel.customxmlpartcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#getcount--)|コレクションに含まれる CustomXML パーツの数を取得します。|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitem-id-)|ID に基づいて、カスタム XML パーツを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitemornullobject-id-)|ID に基づいて、カスタム XML パーツを取得します。|
||[getOnlyItem ()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitem--)|コレクションに含まれる項目が 1 つだけの場合、このメソッドはその項目を返します。|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitemornullobject--)|コレクションに含まれる項目が 1 つだけの場合、このメソッドはその項目を返します。|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#id)|ピボットテーブルの ID。 読み取り専用です。|
|[ランタイム](/javascript/api/excel/excel.runtime)||[ブック](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#customxmlparts)|このブックに格納されているカスタム XML パーツのコレクションを表します。 読み取り専用です。|
|[ワークシート](/javascript/api/excel/excel.worksheet)|[getNext (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getnext-visibleonly-)|これに続くワークシートを取得します。 このメソッドに続くワークシートがない場合、このメソッドはエラーをスローします。|
||[getNextOrNullObject (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getnextornullobject-visibleonly-)|これに続くワークシートを取得します。 このメソッドに続くワークシートがない場合、このメソッドは null オブジェクトを返します。|
||[getPrevious (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getprevious-visibleonly-)|これより前のワークシートを取得します。 前のワークシートがない場合、このメソッドはエラーをスローします。|
||[getPreviousOrNullObject (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getpreviousornullobject-visibleonly-)|これより前のワークシートを取得します。 前のワークシートがない場合、このメソッドは null の objet を返します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getFirst (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getfirst-visibleonly-)|コレクション内の最初のワークシートを取得します。|
||[getLast (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getlast-visibleonly-)|コレクション内の最後のワークシートを取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
