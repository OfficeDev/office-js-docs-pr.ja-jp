---
title: Outlook で新規作成フォームのアイテム データを取得および設定する
description: 新規作成のシナリオで、受信者、件名、本文、予定の場所と時刻を含む Outlook アドインのアイテムのさまざまなプロパティを取得または設定します。
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: 7d290b698d5fcfc625baa1860fad297fa79b499bb7fcbcae8c7fa646ebd18f12
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57091442"
---
# <a name="get-and-set-item-data-in-a-compose-form-in-outlook"></a>Outlook で新規作成フォームのアイテム データを取得および設定する

新規作成のシナリオで、受信者、件名、本文、予定の場所と時刻を含む Outlook アドインのアイテムのさまざまなプロパティを取得または設定する方法について説明します。

## <a name="getting-and-setting-item-properties-for-a-compose-add-in"></a>新規作成アドインの item プロパティの取得と設定

新規作成フォームでは、閲覧フォームで公開されているのと同じ種類のプロパティのほとんど (出席者、受信者、件名、本文など) を取得でき、さらに、閲覧フォームではなく新規作成フォームのみに関連する少数の追加プロパティ (本文、bcc) を取得できます。

これらのプロパティのほとんどで、Outlook アドインとユーザーはユーザー インターフェイスの同じプロパティを同時に変更できるため、プロパティの取得と設定のメソッドは非同期になっています。表 1 に、アイテムレベルのプロパティ、および新規作成フォームでそれらのプロパティの取得と設定を行う関連する非同期メソッドを示します。[item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティと [item.conversationId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティは、ユーザーが変更できないため例外です。閲覧フォームの場合と同様に、新規作成フォームでも、直接親オブジェクトからプログラムを使用してプロパティを取得できます。

JavaScript API のアイテム プロパティにアクセスするOffice、Web Services (EWS) を使用してアイテム レベルのプロパティExchangeアクセスできます。 **ReadWriteMailbox** アクセス許可があれば、[mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドを使用して EWS 操作の [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) と [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation) アクセスし、ユーザーのメールボックス内のアイテムのより多くのプロパティを取得、設定することができます。

`makeEwsRequestAsync` 関数は、新規作成および読み取りの両フォームで利用可能です。 **ReadWriteMailbox** アクセス許可、および Office アドインのプラットフォームを介した EWS へのアクセスの詳細については、「[Outlook アドインのアクセス許可について](understanding-outlook-add-in-permissions.md)」と「[Outlook アドインから Web サービスを呼び出す](web-services.md)」を参照してください。

**表 1 新規作成フォームにおいてアイテム プロパティを取得または設定するための非同期メソッド**

<br/>

| プロパティ | プロパティの種類 | 取得する非同期メソッド | 設定する非同期メソッド |
|:-----|:-----|:-----|:-----|
|[bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[受信者](/javascript/api/outlook/office.Recipients)|[Recipients.getAsync](/javascript/api/outlook/office.Recipients#getAsync_options__callback_)|[Recipients.addAsync](/javascript/api/outlook/office.Recipients#addAsync_recipients__options__callback_), [Recipients.setAsync](/javascript/api/outlook/office.Recipients#setAsync_recipients__options__callback_)|
|[body](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[本文](/javascript/api/outlook/office.Body)|[Body.getAsync](/javascript/api/outlook/office.Body#getAsync_coercionType__options__callback_)|[Body.prependAsync](/javascript/api/outlook/office.Body#prependAsync_data__options__callback_), [Body.setAsync](/javascript/api/outlook/office.Body#setAsync_data__options__callback_), [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setSelectedDataAsync_data__options__callback_)|
|[cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|受信者|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Time](/javascript/api/outlook/office.Time)|[Time.getAsync](/javascript/api/outlook/office.Time#getAsync_options__callback_)|[Time.setAsync](/javascript/api/outlook/office.Time#setAsync_dateTime__options__callback_)|
|[location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[場所](/javascript/api/outlook/office.Location)|[Location.getAsync](/javascript/api/outlook/office.Location#getAsync_options__callback_)|[Location.setAsync](/javascript/api/outlook/office.Location#setAsync_location__options__callback_)|
|[optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|受信者|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|受信者|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|時刻|Time.getAsync|Time.setAsync|
|[subject](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[件名](/javascript/api/outlook/office.Subject)|[Subject.getAsync](/javascript/api/outlook/office.Subject#getAsync_options__callback_)|[Subject.setAsync](/javascript/api/outlook/office.Subject#setAsync_subject__options__callback_)|
|[to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|受信者|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|

## <a name="see-also"></a>関連項目

- [新規作成フォーム用の Outlook アドインを作成する](compose-scenario.md)
- [Outlook アドインのアクセス許可を理解する](understanding-outlook-add-in-permissions.md)
- [Outlook アドインから Web サービスを呼び出す](web-services.md)
- [閲覧または新規作成フォームの Outlook アイテム データを取得および設定する](item-data.md)
