---
title: Outlook アドイン API 要件セット 1.3
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: c34be8f30a2c674035e5ab0ca223f630d9bb5e5a
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432628"
---
# <a name="outlook-add-in-api-requirement-set-13"></a>Outlook アドイン API 要件セット 1.3

JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!NOTE]
> このドキュメントは、最新の要件セット以外の[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)のためのものです。 

## <a name="whats-new-in-13"></a>1.3 の新機能

要件セット 1.3 には、[要件セット 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) のすべての機能が含まれています。次の機能が追加されました。

- [アドイン コマンド](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook)のサポートが追加されました。
- 作成中のアイテムを保存または閉じる機能が追加されました。
- アドインで本文全体を取得または設定できるようにする [Body](/javascript/api/outlook_1_3/office.body) オブジェクトが強化されました。
- EWS 形式と REST 形式間で ID を変換する変換メソッドが追加されました。
- アイテム上にある情報バーに通知メッセージを追加する機能が追加されました。

### <a name="change-log"></a>変更ログ

- [Body.getAsync](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) が追加されました。現在の本文を指定された形式で返します。
- [Body.setAsync](/javascript/api/outlook_1_3/office.body#setasync-data--options--callback-) が追加されました。本文全体を指定されたテキストに置換します。
- [Office.context.officeTheme](office.context.md#officetheme-object) が追加されました。Office テーマの色にアクセスできるようにします。
- [Event](/javascript/api/office/office.addincommands.event) オブジェクトが追加されました。パラメーターとして、Outlook アドインの UI を使用しないコマンド関数に渡されます。処理の完了を通知するために使用されます。
- [Office.context.mailbox.item.close](office.context.mailbox.item.md#close) が追加されました。作成中の現在のアイテムを閉じます。
- [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#saveasyncoptions-callback) が追加されました。アイテムを非同期的に保存します。
- [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages) が追加されました。アイテムの通知メッセージを取得します。
- [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#converttoewsiditemid-restversion--string) が追加されました。REST 形式のアイテム ID を EWS 形式に変換します。
- [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) が追加されました。EWS 形式のアイテム ID を REST 形式に変換します。
- [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook_1_3/office.mailboxenums.itemnotificationmessagetype) が追加されました。予定またはメッセージの通知メッセージの種類を指定します。
- [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_3/office.mailboxenums.restversion) が追加されました。REST 形式のアイテム ID に対応する REST API のバージョンを指定します。
- [NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages) オブジェクトが追加されました。Outlook アドインの通知メッセージにアクセスするメソッドを提供します。
- [NotificationMessageDetails](/javascript/api/outlook_1_3/office.notificationmessagedetails) 型を追加しました。`NotificationMessages.getAllAsync` メソッドによって返されます。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [作業の開始](https://docs.microsoft.com/outlook/add-ins/quick-start)