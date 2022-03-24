---
title: Outlook API 要件セット 1.9
description: アドイン API の要件セット 1.9 Outlook 1.9。
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: c98db3909400a01ffa12d75acf4ee3c4a7752bf1
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745331"
---
# <a name="outlook-add-in-api-requirement-set-19"></a>Outlook API 要件セット 1.9

Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!NOTE]
> このドキュメントは、最新の要件セット以外の[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)のためのものです。

## <a name="whats-new-in-19"></a>1.9 の新機能

要件セット 1.9 には、要件セット [1.8 のすべての機能が含まれています](../requirement-set-1.8/outlook-requirement-set-1.8.md)。 次の機能が追加されました。

- 追加送信時、カスタム プロパティ、および表示フォーム機能用の新しい API が追加されました。
- のサポートが追加されました `Dialog.messageChild`。

### <a name="change-log"></a>変更ログ

- [CustomProperties.getAll](/javascript/api/outlook/office.customproperties?view=outlook-js-1.9&preserve-view=true#outlook-office-customproperties-getall-member(1)) の追加: すべての`CustomProperties`カスタム プロパティを取得するオブジェクトに新しい関数を追加します。
- [Dialog.messageChild](../../../develop/dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) の追加: 作業ウィンドウや UI レス関数ファイルなど、ホスト ページからメッセージをページから開いたダイアログに配信する新しいメソッドを追加します。
- [ExtendedPermissions マニフェスト要素の追加](../../manifest/extendedpermissions.md): [VersionOverrides](../../manifest/versionoverrides.md) マニフェスト要素に子要素を追加します。 アドインが追加送信[](../../../outlook/append-on-send.md)`AppendOnSend`機能をサポートするには、拡張アクセス許可を拡張アクセス許可のコレクションに含める必要があります。
- [Office.context.mailbox.displayAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#outlook-office-mailbox-displayappointmentformasync-member(1)): `Mailbox` 既存の予定を表示するオブジェクトに新しい関数を追加しました。 これは、メソッドの非同期バージョン `displayAppointmentForm` です。
- [Office.context.mailbox.displayMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#outlook-office-mailbox-displaymessageformasync-member(1)): `Mailbox` 既存のメッセージを表示するオブジェクトに新しい関数を追加しました。 これは、メソッドの非同期バージョン `displayMessageForm` です。
- [Office.context.mailbox.displayNewAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#outlook-office-mailbox-displaynewappointmentformasync-member(1)): `Mailbox` 新しい予定フォームを表示するオブジェクトに新しい関数を追加しました。 これは、メソッドの非同期バージョン `displayNewAppointmentForm` です。
- [Office.context.mailbox.displayNewMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#outlook-office-mailbox-displaynewmessageformasync-member(1)): `Mailbox` 新しいメッセージ フォームを表示するオブジェクトに新しい関数を追加しました。 これは、メソッドの非同期バージョン `displayNewMessageForm` です。
- [Office.context.mailbox.item.body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#outlook-office-body-appendonsendasync-member(1)): `Body` 新規作成モードでアイテム本文の末尾にデータを追加する新しい関数をオブジェクトに追加しました。
- [Office.context.mailbox.item.displayReplyAllFormAsync](office.context.mailbox.item.md#methods) を追加しました: `Item` 読み取りモードで "Reply all" フォームを表示する新しい関数をオブジェクトに追加します。 これは、メソッドの非同期バージョン `displayReplyAllForm` です。
- [Office.context.mailbox.item.displayReplyFormAsync](office.context.mailbox.item.md#methods) の追加: `Item` 読み取りモードで "Reply" フォームを表示するオブジェクトに新しい関数を追加します。 これは、メソッドの非同期バージョン `displayReplyForm` です。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](../../../outlook/outlook-add-ins-overview.md)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](../../../quickstarts/outlook-quickstart.md)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
