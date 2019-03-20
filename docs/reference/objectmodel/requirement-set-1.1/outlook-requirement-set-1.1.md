---
title: Outlook アドイン API 要件セット 1.1
description: ''
ms.date: 10/11/2018
localization_priority: Normal
ms.openlocfilehash: a074d0e38f8d872f0d75a68851aef947989c625e
ms.sourcegitcommit: c4d6ecdc41ea67291b6d155c3b246e31ec2e38b7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/19/2019
ms.locfileid: "30600257"
---
# <a name="outlook-add-in-api-requirement-set-11"></a>Outlook アドイン API 要件セット 1.1

JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!NOTE]
> このドキュメントは、最新の要件セット以外の[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)のためのものです。 

## <a name="whats-new-in-11"></a>1.1 の新機能

要件セット 1.1 には、要件セット 1.0 のすべての機能が含まれています。アドインでメッセージと予定の本文にアクセスする機能、および現在のアイテムを変更する機能が追加されました。

### <a name="change-log"></a>変更ログ

- [Body](/javascript/api/outlook_1_1/office.body) オブジェクトが追加されました。Outlook アドインでアイテムのコンテンツを追加および更新するメソッドを提供します。
- [Location](/javascript/api/outlook_1_1/office.location) オブジェクトが追加されました。Outlook アドインで会議の場所を取得し設定する方法を提供します。
- [Recipients](/javascript/api/outlook_1_1/office.recipients) オブジェクトが追加されました。Outlook のアドインで、予定またはメッセージの受信者を取得および設定するメソッドを提供します。
- [Subject](/javascript/api/outlook_1_1/office.subject) オブジェクトが追加されました。Outlook のアドインで、予定またはメッセージの件名を取得および設定するメソッドを提供します。
- [Time](/javascript/api/outlook_1_1/office.time) オブジェクトが追加されました。Outlook アドインで会議の開始時刻と終了時刻を取得および設定する方法を提供します。
- [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback) が追加されました。ファイルを添付ファイルとしてメッセージまたは予定に追加します。
- [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback) が追加されました。メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。
- [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) が追加されました。メッセージまたは予定から添付ファイルを削除します。
- [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-body) が追加されました。アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。
- メッセージの[bcc](office.context.mailbox.item.md#bcc-recipients)行を追加しました。
- [Office.MailboxEnums.RecipientType](/javascript/api/outlook_1_1/office.mailboxenums.recipienttype) が追加されました。予定の受信者の種類を指定します。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [作業の開始](https://docs.microsoft.com/outlook/add-ins/quick-start)
