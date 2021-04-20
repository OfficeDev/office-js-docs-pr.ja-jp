---
title: Outlook アドイン API 要件セット 1.7
description: Outlook アドイン API の概要 (要件セット 1.7)
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: f8dc9fe097b56d3e940a5d7d945c5ccc50e07077
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431389"
---
# <a name="outlook-add-in-api-requirement-set-17"></a>Outlook アドイン API 要件セット 1.7

Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。

> [!NOTE]
> このドキュメントは、最新の要件セット以外の[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)のためのものです。

## <a name="whats-new-in-17"></a>1.7 の新機能

要件セット 1.7 には、[要件セット 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) のすべての機能が含まれています。 次の機能が追加されました。

- 予定とメッセージの定期的なパターン (会議出席依頼) に関する新しい API が追加されました。
- 新規作成モードでも使用できるように、item.from プロパティを変更しました。
- RecurrenceChanged イベント、RecipientsChanged イベント、AppointmentTimeChanged イベントのサポートが追加されました。

### <a name="change-log"></a>変更ログ

- [From](/javascript/api/outlook/office.from?view=outlook-js-1.7&preserve-view=true) が追加されました。From 値を取得するためのメソッドを提供する、新しいオブジェクトが追加されました。
- [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7&preserve-view=true) が追加されました。Organizer 値を取得するためのメソッドを提供する、新しいオブジェクトが追加されました。
- [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7&preserve-view=true) が追加されました。予定の定期的なパターンを取得および設定するためのメソッドを提供し、会議出席依頼であるメッセージの定期的なパターンのみを取得する新しいオブジェクトが追加されました。
- [RecurrenceTimeZone](/javascript/api/outlook/office.recurrencetimezone?view=outlook-js-1.7&preserve-view=true) が追加されました。定期的なパターンのタイムゾーン構成を表す新しいオブジェクトが追加されました。
- [SeriesTime](/javascript/api/outlook/office.seriestime?view=outlook-js-1.7&preserve-view=true) が追加されました。定期的な一連の予定の日付と時刻を取得および設定したり、定期的な一連の会議出席依頼の日時を取得したりするためのメソッドを提供する新しいオブジェクトが追加されました。
- [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#methods) が追加されました。サポートされているイベントのイベント ハンドラーを追加するための、新しいメソッドが追加されました。
- [Office.context.mailbox.item.from](office.context.mailbox.item.md#properties)が変更されました: 作成モードでfrom値を取得する機能を追加しました。
- [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#properties)が変更されました: 作成モードでオーガナイザー値を取得する機能を追加しました。
- [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#properties) が追加されました。予定アイテムの定期的なパターンを管理するメソッドを提供するオブジェクトを取得または設定する新しいプロパティが追加されました。 このプロパティを使用して、会議出席依頼アイテムの定期的なパターンを取得することもできます。
- [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#methods) が追加されました。サポートされているイベントの種類のイベント ハンドラーを削除する新しいメソッドが追加されました。
- [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#properties) が追加されました。オカレンスが属する系列の ID を取得する新しいプロパティが追加されました。
- [Office.MailboxEnums.Days](/javascript/api/outlook/office.mailboxenums.days?view=outlook-js-1.7&preserve-view=true) が追加されました。曜日または日付の種類を指定する新しい列挙型が追加されました。
- [Office.MailboxEnums.Month](/javascript/api/outlook/office.mailboxenums.month?view=outlook-js-1.7&preserve-view=true) が追加されました。月を指定する新しい列挙型が追加されました。
- [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook/office.mailboxenums.recurrencetimezone?view=outlook-js-1.7&preserve-view=true) が追加されました。定期的なアイテムに適用されるタイムゾーンを指定する新しい列挙型が追加されました。
- [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook/office.mailboxenums.recurrencetype?view=outlook-js-1.7&preserve-view=true) が追加されました。定期的なアイテムの種類を指定する新しい列挙型が追加されました。
- [Office.MailboxEnums.WeekNumber](/javascript/api/outlook/office.mailboxenums.weeknumber?view=outlook-js-1.7&preserve-view=true) が追加されました。月の週を指定する新しい列挙型が追加されました。
- [Office.EventType](/javascript/api/office/office.eventtype)が変更されました: `RecurrenceChanged`、 `RecipientsChanged`と `AppointmentTimeChanged`のイベントにサポートを追加しました。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](../../../outlook/outlook-add-ins-overview.md)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](../../../quickstarts/outlook-quickstart.md)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
