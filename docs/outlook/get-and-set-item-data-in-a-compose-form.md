---
title: Outlook で新規作成フォームのアイテム データを取得および設定する
description: 新規作成のシナリオで、受信者、件名、本文、予定の場所と時刻を含む Outlook アドインのアイテムのさまざまなプロパティを取得または設定します。
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2ae4b6a30d08199207faf89079c57fbff46d6a0e
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467239"
---
# <a name="get-and-set-item-data-in-a-compose-form-in-outlook"></a>Outlook で新規作成フォームのアイテム データを取得および設定する

新規作成のシナリオで、受信者、件名、本文、予定の場所と時刻を含む Outlook アドインのアイテムのさまざまなプロパティを取得または設定する方法について説明します。

## <a name="getting-and-setting-item-properties-for-a-compose-add-in"></a>新規作成アドインの item プロパティの取得と設定

新規作成フォームでは、閲覧フォームで公開されているのと同じ種類のプロパティのほとんど (出席者、受信者、件名、本文など) を取得でき、さらに、閲覧フォームではなく新規作成フォームのみに関連する少数の追加プロパティ (本文、bcc) を取得できます。

For most of these properties, because it's possible that an Outlook add-in and the user can be modifying the same property in the user interface at the same time, the methods to get and set them are asynchronous. Table 1 lists the item-level properties and corresponding asynchronous methods to get and set them in a compose form. The  [item.itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) and [item.conversationId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) properties are exceptions because users cannot modify them. You can programmatically get them the same way in a compose form as in a read form, directly from the parent object.

Office JavaScript API でアイテム プロパティにアクセスする以外に、Exchange Web Services (EWS) を使用してアイテム レベルのプロパティにアクセスできます。 メールボックスの **読み取り/書き込み** アクセス許可を使用すると、 [mailbox.makeEwsRequestAsync メソッドを](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 使用して EWS 操作 ( [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) と [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation)) にアクセスし、ユーザーのメールボックス内のアイテムまたはアイテムのプロパティをさらに取得および設定できます。

このメソッドは `makeEwsRequestAsync` 、作成フォームと読み取りフォームの両方で使用できます。 **読み取り/書き込みメールボックス** のアクセス許可と Office アドイン プラットフォームを使用した EWS へのアクセスの詳細については、「[Outlook アドインのアクセス許可について」と「Outlook](understanding-outlook-add-in-permissions.md) アドイン [から Web サービスを呼び出す](web-services.md)」を参照してください。

**表 1 新規作成フォームにおいてアイテム プロパティを取得または設定するための非同期メソッド**

| プロパティ | プロパティの種類 | 取得する非同期メソッド | 設定する非同期メソッド |
|:-----|:-----|:-----|:-----|
|[bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[受信者](/javascript/api/outlook/office.recipients)|[Recipients.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1))|[Recipients.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1)), [Recipients.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))|
|[body](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[Body](/javascript/api/outlook/office.body)|[Body.getAsync](/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1))|[Body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1)), [Body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1)), [Body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1))|
|[cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|受信者|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[end](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[Time](/javascript/api/outlook/office.time)|[Time.getAsync](/javascript/api/outlook/office.time#outlook-office-time-getasync-member(1))|[Time.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))|
|[location](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[場所](/javascript/api/outlook/office.location)|[Location.getAsync](/javascript/api/outlook/office.location#outlook-office-location-getasync-member(1))|[Location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1))|
|[optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|受信者|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|受信者|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[start](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|時刻|Time.getAsync|Time.setAsync|
|[subject](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[件名](/javascript/api/outlook/office.subject)|[Subject.getAsync](/javascript/api/outlook/office.subject#outlook-office-subject-getasync-member(1))|[Subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1))|
|[to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|受信者|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|

## <a name="see-also"></a>関連項目

- [新規作成フォーム用の Outlook アドインを作成する](compose-scenario.md)
- [Outlook アドインのアクセス許可を理解する](understanding-outlook-add-in-permissions.md)
- [Outlook アドインから Web サービスを呼び出す](web-services.md)
- [閲覧または新規作成フォームの Outlook アイテム データを取得および設定する](item-data.md)
