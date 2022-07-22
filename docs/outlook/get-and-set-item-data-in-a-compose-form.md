---
title: Outlook で新規作成フォームのアイテム データを取得および設定する
description: 新規作成のシナリオで、受信者、件名、本文、予定の場所と時刻を含む Outlook アドインのアイテムのさまざまなプロパティを取得または設定します。
ms.date: 12/10/2019
ms.localizationpriority: medium
ms.openlocfilehash: ddc6cd0011060bc49d1fd5cd8e6c9ceebb2a8c08
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958980"
---
# <a name="get-and-set-item-data-in-a-compose-form-in-outlook"></a>Outlook で新規作成フォームのアイテム データを取得および設定する

新規作成のシナリオで、受信者、件名、本文、予定の場所と時刻を含む Outlook アドインのアイテムのさまざまなプロパティを取得または設定する方法について説明します。

## <a name="getting-and-setting-item-properties-for-a-compose-add-in"></a>新規作成アドインの item プロパティの取得と設定

新規作成フォームでは、閲覧フォームで公開されているのと同じ種類のプロパティのほとんど (出席者、受信者、件名、本文など) を取得でき、さらに、閲覧フォームではなく新規作成フォームのみに関連する少数の追加プロパティ (本文、bcc) を取得できます。

これらのプロパティのほとんどで、Outlook アドインとユーザーはユーザー インターフェイスの同じプロパティを同時に変更できるため、プロパティの取得と設定のメソッドは非同期になっています。表 1 に、アイテムレベルのプロパティ、および新規作成フォームでそれらのプロパティの取得と設定を行う関連する非同期メソッドを示します。[item.itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) プロパティと [item.conversationId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) プロパティは、ユーザーが変更できないため例外です。閲覧フォームの場合と同様に、新規作成フォームでも、直接親オブジェクトからプログラムを使用してプロパティを取得できます。

Office JavaScript API でアイテム プロパティにアクセスする以外に、Exchange Web Services (EWS) を使用してアイテム レベルのプロパティにアクセスできます。 **ReadWriteMailbox** アクセス許可があれば、[mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) メソッドを使用して EWS 操作の [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) と [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation) アクセスし、ユーザーのメールボックス内のアイテムのより多くのプロパティを取得、設定することができます。

このメソッドは `makeEwsRequestAsync` 、作成フォームと読み取りフォームの両方で使用できます。 **ReadWriteMailbox** アクセス許可、Office アドイン プラットフォームを経由した EWS へのアクセスの詳細については、「 [ユーザーのメールボックスにアクセスする Outlook アドインのためのアクセス許可を指定する](understanding-outlook-add-in-permissions.md)」および「 [Outlook アドインから Web サービスを呼び出す](web-services.md)」を参照してください。

**表 1 新規作成フォームにおいてアイテム プロパティを取得または設定するための非同期メソッド**

| プロパティ | プロパティの種類 | 取得する非同期メソッド | 設定する非同期メソッド |
|:-----|:-----|:-----|:-----|
|[bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[受信者](/javascript/api/outlook/office.recipients)|[Recipients.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1))|[Recipients.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1)), [Recipients.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))|
|[body](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[本文](/javascript/api/outlook/office.body)|[Body.getAsync](/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1))|[Body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1)), [Body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1)), [Body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1))|
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
