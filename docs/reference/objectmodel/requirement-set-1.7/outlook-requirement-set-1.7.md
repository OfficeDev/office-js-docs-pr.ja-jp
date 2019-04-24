---
title: Outlook アドイン API 要件セット 1.7
description: ''
ms.date: 03/20/2019
localization_priority: Priority
ms.openlocfilehash: 8daf10239a704206d53a544185e030afa6b6a27a
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450283"
---
# <a name="outlook-add-in-api-requirement-set-17"></a>Outlook アドイン API 要件セット 1.7

JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

## <a name="whats-new-in-17"></a>1.7 の新機能

要件セット 1.7 には、[要件セット 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) のすべての機能が含まれています。 次の機能が追加されました。

- 予定とメッセージの定期的なパターン (会議出席依頼) に関する新しい API が追加されました。
- 新規作成モードでも使用できるように、item.from プロパティを変更しました。
- RecurrenceChanged イベント、RecipientsChanged イベント、AppointmentTimeChanged イベントのサポートが追加されました。

### <a name="change-log"></a>変更ログ

- [From](/javascript/api/outlook_1_7/office.from) が追加されました。From 値を取得するためのメソッドを提供する、新しいオブジェクトが追加されました。
- [Organizer](/javascript/api/outlook_1_7/office.organizer) が追加されました。Organizer 値を取得するためのメソッドを提供する、新しいオブジェクトが追加されました。
- [Recurrence](/javascript/api/outlook_1_7/office.recurrence) が追加されました。予定の定期的なパターンを取得および設定するためのメソッドを提供し、会議出席依頼であるメッセージの定期的なパターンのみを取得する新しいオブジェクトが追加されました。
- [RecurrenceTimeZone](/javascript/api/outlook_1_7/office.recurrencetimezone) が追加されました。定期的なパターンのタイムゾーン構成を表す新しいオブジェクトが追加されました。
- [SeriesTime](/javascript/api/outlook_1_7/office.seriestime) が追加されました。定期的な一連の予定の日付と時刻を取得および設定したり、定期的な一連の会議出席依頼の日時を取得したりするためのメソッドを提供する新しいオブジェクトが追加されました。
- [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback) が追加されました。サポートされているイベントのイベント ハンドラーを追加するための、新しいメソッドが追加されました。
- [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsfrom)が変更されました: 作成モードでfrom値を取得する機能を追加しました。
- [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsorganizer)が変更されました: 作成モードでオーガナイザー値を取得する機能を追加しました。
- [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrence) が追加されました。予定アイテムの定期的なパターンを管理するメソッドを提供するオブジェクトを取得または設定する新しいプロパティが追加されました。 このプロパティを使用して、会議出席依頼アイテムの定期的なパターンを取得することもできます。
- [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-options-callback) が追加されました。サポートされているイベントの種類のイベント ハンドラーを削除する新しいメソッドが追加されました。
- [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string) が追加されました。オカレンスが属する系列の ID を取得する新しいプロパティが追加されました。
- [Office.MailboxEnums.Days](/javascript/api/outlook_1_7/office.mailboxenums.days) が追加されました。曜日または日付の種類を指定する新しい列挙型が追加されました。
- [Office.MailboxEnums.Month](/javascript/api/outlook_1_7/office.mailboxenums.month) が追加されました。月を指定する新しい列挙型が追加されました。
- [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetimezone) が追加されました。定期的なアイテムに適用されるタイムゾーンを指定する新しい列挙型が追加されました。
- [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetype) が追加されました。定期的なアイテムの種類を指定する新しい列挙型が追加されました。
- [Office.MailboxEnums.WeekNumber](/javascript/api/outlook_1_7/office.mailboxenums.weeknumber) が追加されました。月の週を指定する新しい列挙型が追加されました。
- [Office.EventType](/javascript/api/office/office.eventtype)が変更されました: `RecurrenceChanged`、 `RecipientsChanged`と `AppointmentTimeChanged`のイベントにサポートを追加しました。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [作業の開始](/outlook/add-ins/quick-start)
