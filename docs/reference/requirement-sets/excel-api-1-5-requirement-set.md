---
title: ExcelJavaScript API 要件セット 1.5
description: ExcelApi 1.5 要件セットの詳細。
ms.date: 03/19/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 01a13a0f531eae9eea2c213ba0da764fbe51ee15
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936571"
---
# <a name="whats-new-in-excel-javascript-api-15"></a>Excel JavaScript API 1.5 の新機能

ExcelApi 1.5 では、カスタム XML パーツが追加されます。 これらは、ブック オブジェクトの [カスタム XML パーツ コレクション](/javascript/api/excel/excel.workbook#customxmlparts) からアクセスできます。

## <a name="custom-xml-part"></a>カスタム XML パーツ

* ID を使用してカスタム XML パーツを取得します。
* 名前空間が指定した名前空間に一致する、カスタム XML パーツの新しい範囲のコレクションを取得します。
* パーツに関連付けられた XML 文字列を取得します。
* パーツの ID と名前空間を指定します。
* 新しいカスタム XML パーツをブックに追加します。
* XML パーツ全体を設定します。
* カスタム XML パーツを削除します。
* xpath で識別される要素から、指定された名前を持つ属性を削除します。
* xpath で XML の内容を照会します。
* 属性の挿入、更新、および削除。

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.5 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.5 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット[1.5](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)以前の Excel API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#delete__)|カスタム XML パーツを削除します。|
||[getXml()](/javascript/api/excel/excel.customxmlpart#getXml__)|カスタム XML パーツのすべての XML コンテンツを取得します。|
||[id](/javascript/api/excel/excel.customxmlpart#id)|カスタム XML パーツの ID。|
||[namespaceUri](/javascript/api/excel/excel.customxmlpart#namespaceUri)|カスタム XML パーツの名前空間 URI。|
||[setXml(xml: string)](/javascript/api/excel/excel.customxmlpart#setXml_xml_)|カスタム XML パーツのすべての XML コンテンツを設定します。|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[add(xml: string)](/javascript/api/excel/excel.customxmlpartcollection#add_xml_)|ブックに新しいカスタム XML パーツを追加します。|
||[getByNamespace(namespaceUri: string)](/javascript/api/excel/excel.customxmlpartcollection#getByNamespace_namespaceUri_)|名前空間が指定した名前空間に一致する、カスタム XML パーツの新しい範囲のコレクションを取得します。|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#getCount__)|コレクション内のカスタム XML パーツの数を取得します。|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getItem_id_)|ID に基づいて、カスタム XML パーツを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getItemOrNullObject_id_)|ID に基づいて、カスタム XML パーツを取得します。|
||[items](/javascript/api/excel/excel.customxmlpartcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#getCount__)|コレクションに含まれる CustomXML パーツの数を取得します。|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getItem_id_)|ID に基づいて、カスタム XML パーツを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getItemOrNullObject_id_)|ID に基づいて、カスタム XML パーツを取得します。|
||[getOnlyItem()](/javascript/api/excel/excel.customxmlpartscopedcollection#getOnlyItem__)|コレクションに含まれる項目が 1 つだけの場合、このメソッドはその項目を返します。|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#getOnlyItemOrNullObject__)|コレクションに含まれる項目が 1 つだけの場合、このメソッドはその項目を返します。|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#id)|ピボットテーブルの ID。|
|[RequestContext](/javascript/api/excel/excel.requestcontext)|[runtime](/javascript/api/excel/excel.requestcontext#runtime)||
|[ランタイム](/javascript/api/excel/excel.runtime)|||
|[Workbook](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#customXmlParts)|このブックに含まれるカスタム XML パーツのコレクションを表します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getNext(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getNext_visibleOnly_)|このワークシートに続くワークシートを取得します。|
||[getNextOrNullObject(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getNextOrNullObject_visibleOnly_)|このワークシートに続くワークシートを取得します。|
||[getPrevious(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getPrevious_visibleOnly_)|このワークシートの前のワークシートを取得します。|
||[getPreviousOrNullObject(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getPreviousOrNullObject_visibleOnly_)|このワークシートの前のワークシートを取得します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getFirst(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getFirst_visibleOnly_)|コレクション内の最初のワークシートを取得します。|
||[getLast(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getLast_visibleOnly_)|コレクション内の最後のワークシートを取得します。|

## <a name="see-also"></a>関連項目

* [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)
* [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
