---
title: Outlook API 要件セット 1.11
description: アドイン API の要件セット 1.11 Outlook 1.11。
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 56066d7b3a6debaeed365a9ca05a3e894762dea3
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681832"
---
# <a name="outlook-add-in-api-requirement-set-111"></a>Outlook API 要件セット 1.11

Office Outlook JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

## <a name="whats-new-in-111"></a>1.11 の新機能

要件セット 1.11 には、要件セット [1.10 のすべての機能が含まれています](../requirement-set-1.10/outlook-requirement-set-1.10.md)。 次の機能が追加されました。

- イベント ベースのアクティブ化 [の新しいイベントを追加しました](../../../outlook/autolaunch.md#supported-events)。
- SessionData API を追加しました。

### <a name="change-log"></a>変更ログ

- [Office.context.mailbox.item.sessionData](office.context.mailbox.item.md#properties)の追加 : 新規作成モードでアイテムのセッション データを管理するための新しいプロパティを追加します。
- 追加された[Office。SessionData](/javascript/api/outlook/office.sessiondata?view=outlook-js-1.11&preserve-view=true): 新規作成アイテムのセッション データを表す新しいオブジェクトを追加します。
- イベント ベースのアクティブ化の [新しいイベントを追加](../../../outlook/autolaunch.md#supported-events)しました: 次のイベントのサポートを追加します。

  - `OnAppointmentAttachmentsChanged`
  - `OnAppointmentAttendeesChanged`
  - `OnAppointmentRecurrenceChanged`
  - `OnAppointmentTimeChanged`
  - `OnInfoBarDismissClicked`
  - `OnMessageAttachmentsChanged`
  - `OnMessageRecipientsChanged`

- 追加された[Office。AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true): イベントをサポートするオブジェクトを追加 `OnAppointmentTimeChanged` します。
- 追加された[Office。AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true): イベントとイベントをサポートするオブジェクト `OnAppointmentAttachmentsChanged` を追加 `OnMessageAttachmentsChanged` します。
- 追加された[Office。InfobarClickedEventArgs](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true): イベントをサポートするオブジェクトを追加 `OnInfoBarDismissClicked` します。
- 追加された[Office。RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true): イベントとイベントをサポートするオブジェクト `OnAppointmentAttendeesChanged` を追加 `OnMessageRecipientsChanged` します。
- 追加された[Office。RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true): イベントをサポートするオブジェクトを追加 `OnAppointmentRecurrenceChanged` します。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](../../../outlook/outlook-add-ins-overview.md)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](../../../quickstarts/outlook-quickstart.md)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)