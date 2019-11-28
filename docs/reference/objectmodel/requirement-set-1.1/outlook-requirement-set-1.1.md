---
title: Outlook アドイン API 要件セット 1.1
description: ''
ms.date: 11/26/2019
localization_priority: Normal
ms.openlocfilehash: 1c2e8ea26cac7ff630961b176391ef1adf2249fd
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629238"
---
# <a name="outlook-add-in-api-requirement-set-11"></a>Outlook アドイン API 要件セット 1.1

JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。 Outlook JavaScript API 1.1 (メールボックス 1.1) は、API の最初のバージョンです。

> [!NOTE]
> このドキュメントは、最新の要件セット以外の[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)のためのものです。

## <a name="whats-new-in-11"></a>1.1 の新機能

要件セット1.1 には、Outlook でサポートされているすべての[共通 API 要件セット](../../requirement-sets/office-add-in-requirement-sets.md)が含まれています。 アドインでメッセージと予定の本文にアクセスする機能、および現在のアイテムを変更する機能が追加されました。

### <a name="change-log"></a>変更ログ

- [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1) オブジェクトが追加されました。Outlook アドインでアイテムのコンテンツを追加および更新するメソッドを提供します。
- [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1) オブジェクトが追加されました。Outlook アドインで会議の場所を取得し設定する方法を提供します。
- [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) オブジェクトが追加されました。Outlook のアドインで、予定またはメッセージの受信者を取得および設定するメソッドを提供します。
- [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1) オブジェクトが追加されました。Outlook のアドインで、予定またはメッセージの件名を取得および設定するメソッドを提供します。
- [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1) オブジェクトが追加されました。Outlook アドインで会議の開始時刻と終了時刻を取得および設定する方法を提供します。
- [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback) が追加されました。ファイルを添付ファイルとしてメッセージまたは予定に追加します。
- [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback) が追加されました。メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。
- [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) が追加されました。メッセージまたは予定から添付ファイルを削除します。
- [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-body) が追加されました。アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。
- メッセージの[bcc](office.context.mailbox.item.md#bcc-recipients)行を追加しました。
- [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1) が追加されました。予定の受信者の種類を指定します。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](/outlook/add-ins/quick-start)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
