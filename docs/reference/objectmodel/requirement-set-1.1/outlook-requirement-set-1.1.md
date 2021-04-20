---
title: Outlook アドイン API 要件セット 1.1
description: Outlook アドインおよび Office JavaScript Api for the Mailbox API 1.1 の一部として導入された機能と Api。
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: f93b6d582043641903b362121c6e5eaf89c2ad1c
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431375"
---
# <a name="outlook-add-in-api-requirement-set-11"></a>Outlook アドイン API 要件セット 1.1

Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。 Outlook JavaScript API 1.1 (メールボックス 1.1) は、API の最初のバージョンです。

> [!NOTE]
> このドキュメントは、最新の要件セット以外の[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)のためのものです。

## <a name="whats-new-in-11"></a>1.1 の新機能

要件セット1.1 には、Outlook でサポートされているすべての [共通 API 要件セット](../../requirement-sets/office-add-in-requirement-sets.md) が含まれています。 アドインでメッセージと予定の本文にアクセスする機能、および現在のアイテムを変更する機能が追加されました。

### <a name="change-log"></a>変更ログ

- [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1&preserve-view=true) オブジェクトが追加されました。Outlook アドインでアイテムのコンテンツを追加および更新するメソッドを提供します。
- [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1&preserve-view=true) オブジェクトが追加されました。Outlook アドインで会議の場所を取得し設定する方法を提供します。
- [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1&preserve-view=true) オブジェクトが追加されました。Outlook のアドインで、予定またはメッセージの受信者を取得および設定するメソッドを提供します。
- [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1&preserve-view=true) オブジェクトが追加されました。Outlook のアドインで、予定またはメッセージの件名を取得および設定するメソッドを提供します。
- [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1&preserve-view=true) オブジェクトが追加されました。Outlook アドインで会議の開始時刻と終了時刻を取得および設定する方法を提供します。
- [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods) が追加されました。ファイルを添付ファイルとしてメッセージまたは予定に追加します。
- [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#methods) が追加されました。メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。
- [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#methods) が追加されました。メッセージまたは予定から添付ファイルを削除します。
- [Office.context.mailbox.item.body](office.context.mailbox.item.md#properties) が追加されました。アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。
- メッセージの [bcc](office.context.mailbox.item.md#properties) 行を追加しました。
- [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1&preserve-view=true) が追加されました。予定の受信者の種類を指定します。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](../../../outlook/outlook-add-ins-overview.md)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](../../../quickstarts/outlook-quickstart.md)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
