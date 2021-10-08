---
title: Outlook アドイン API 要件セット 1.3
description: メールボックス API 1.3 の一部Outlook JavaScript API および Office JavaScript API 用に導入された機能と API。
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: a8688d5d63cd658084bd0ba4601ed85a631bf8d8
ms.sourcegitcommit: efd0966f6400c8e685017ce0c8c016a2cbab0d5c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/08/2021
ms.locfileid: "60237770"
---
# <a name="outlook-add-in-api-requirement-set-13"></a>Outlook アドイン API 要件セット 1.3

Office Outlook JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!NOTE]
> このドキュメントは、最新の要件セット以外の[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)のためのものです。

## <a name="whats-new-in-13"></a>1.3 の新機能

要件セット 1.3 には、要件セット [1.2 のすべての機能が含まれています](../requirement-set-1.2/outlook-requirement-set-1.2.md)。 次の機能が追加されました。

- [アドイン コマンド](../../../outlook/add-in-commands-for-outlook.md)のサポートが追加されました。
- 作成中のアイテムを保存または閉じる機能が追加されました。
- アドイン [が](/javascript/api/outlook/office.body?view=outlook-js-1.3&preserve-view=true) 本文全体を取得または設定できる拡張 Body オブジェクト。
- EWS 形式と REST 形式間で ID を変換する変換メソッドが追加されました。
- アイテム上にある情報バーに通知メッセージを追加する機能が追加されました。

### <a name="change-log"></a>変更ログ

- [Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3&preserve-view=true#getAsync_coercionType__options__callback_) が追加されました。現在の本文を指定された形式で返します。
- [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3&preserve-view=true#setAsync_data__options__callback_) が追加されました。本文全体を指定されたテキストに置換します。
- [Event](/javascript/api/office/office.addincommands.event?view=outlook-js-1.3&preserve-view=true) オブジェクトが追加されました。パラメーターとして、Outlook アドインの UI を使用しないコマンド関数に渡されます。処理の完了を通知するために使用されます。
- [Office.context.mailbox.item.close](office.context.mailbox.item.md#methods) が追加されました。作成中の現在のアイテムを閉じます。
- [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#methods) が追加されました。アイテムを非同期的に保存します。
- [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#properties) が追加されました。アイテムの通知メッセージを取得します。
- [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#methods) が追加されました。REST 形式のアイテム ID を EWS 形式に変換します。
- [Office.context.mailbox.convertToRestId](office.context.mailbox.md#methods) が追加されました。EWS 形式のアイテム ID を REST 形式に変換します。
- [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-1.3&preserve-view=true) が追加されました。予定またはメッセージの通知メッセージの種類を指定します。
- [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3&preserve-view=true) が追加されました。REST 形式のアイテム ID に対応する REST API のバージョンを指定します。
- [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3&preserve-view=true) オブジェクトが追加されました。Outlook アドインの通知メッセージにアクセスするメソッドを提供します。
- [NotificationMessageDetails](/javascript/api/outlook/office.notificationmessagedetails?view=outlook-js-1.3&preserve-view=true) 型を追加しました。`NotificationMessages.getAllAsync` メソッドによって返されます。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](../../../outlook/outlook-add-ins-overview.md)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](../../../quickstarts/outlook-quickstart.md)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
