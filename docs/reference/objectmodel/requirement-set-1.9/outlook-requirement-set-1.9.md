---
title: Outlook アドイン API 要件セット1.9
description: Outlook アドイン API の要件セット1.9。
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: b2174052a60580a895ef82a4b5d8f00ed6899feb
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/20/2020
ms.locfileid: "48628079"
---
# <a name="outlook-add-in-api-requirement-set-19"></a>Outlook アドイン API 要件セット1.9

Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。

## <a name="whats-new-in-19"></a>1.9 の新機能

要件セット1.9 には、 [要件セット 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md)のすべての機能が含まれています。 次の機能が追加されました。

- 追加、送信、カスタムプロパティ、および表示フォーム機能用の新しい Api が追加されました。
- のサポートが追加されました `Dialog.messageChild` 。

### <a name="change-log"></a>変更ログ

- CustomProperties を追加しました。 [getAll](/javascript/api/outlook/office.customproperties?view=outlook-js-1.9&preserve-view=true#getall--): すべての `CustomProperties` カスタムプロパティを取得する新しい関数をオブジェクトに追加します。
- [MessageChild](../../../develop/dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)が追加されました。この新しいメソッドを追加します。これにより、作業ウィンドウや、UI にない関数ファイルなどのホストページから、ページから開いたダイアログにメッセージを配信します。
- [Extendedpermissions マニフェスト要素](../../manifest/extendedpermissions.md)を追加しました。 [versionoverrides](../../manifest/versionoverrides.md)マニフェスト要素に子要素を追加します。 アドインが [メール追加機能](../../../outlook/append-on-send.md)をサポートするためには、拡張されたアクセス許可の `AppendOnSend` コレクションに拡張アクセス許可が含まれている必要があります。
- [DisplayAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayappointmentformasync-itemid--options--callback-)が追加されました。 `Mailbox` 既存の予定を表示するオブジェクトに新しい関数を追加します。 これは、メソッドの非同期バージョンです `displayAppointmentForm` 。
- 追加さ[れました。](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaymessageformasync-itemid--options--callback-)既存のメッセージを表示する新しい関数をオブジェクトに追加します。 `Mailbox` これは、メソッドの非同期バージョンです `displayMessageForm` 。
- [DisplayNewAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewappointmentformasync-parameters--options--callback-)が追加されまし `Mailbox` た。新しい予定のフォームを表示する新しい関数をオブジェクトに追加します。 これは、メソッドの非同期バージョンです `displayNewAppointmentForm` 。
- 追加さ [れました。](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewmessageformasync-parameters--options--callback-) `Mailbox` 新しいメッセージフォームを表示する新しい関数をオブジェクトに追加しています。 これは、メソッドの非同期バージョンです `displayNewMessageForm` 。
- 追加 [さ](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-)れました。このオブジェクトには、新規 `Body` 作成モードのアイテムの本文の最後にデータを追加する新しい関数が追加されています。
- [DisplayReplyAllFormAsync](office.context.mailbox.item.md#methods)が追加されまし `Item` た。読み取りモードで "全員に返信" フォームを表示する新しい関数をオブジェクトに追加します。 これは、メソッドの非同期バージョンです `displayReplyAllForm` 。
- [DisplayReplyFormAsync](office.context.mailbox.item.md#methods)が追加されました。 `Item` 読み取りモードで "Reply" フォームを表示するオブジェクトに新しい関数を追加します。 これは、メソッドの非同期バージョンです `displayReplyForm` 。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](../../../outlook/outlook-add-ins-overview.md)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](../../../quickstarts/outlook-quickstart.md)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
