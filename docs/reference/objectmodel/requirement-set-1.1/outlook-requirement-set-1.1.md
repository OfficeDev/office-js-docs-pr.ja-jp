---
title: Outlook アドイン API 要件セット 1.1
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: 13334029cd30742e6d7dd77cb569a1028a35106a
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433034"
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
- [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-bodyjavascriptapioutlook11officebody) が追加されました。アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。
- [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#bcc-recipientsjavascriptapioutlook11officerecipients) が追加されました。メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または設定します。
- [Office.MailboxEnums.RecipientType](/javascript/api/outlook_1_1/office.mailboxenums.recipienttype) が追加されました。予定の受信者の種類を指定します。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [作業の開始](https://docs.microsoft.com/outlook/add-ins/quick-start)