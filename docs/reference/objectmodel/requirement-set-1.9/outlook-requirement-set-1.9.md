---
title: Outlook API 要件セット 1.9
description: アドイン API の要件セット 1.9 Outlook 1.9。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 6ac4e119fea41a6f4bd1b3ab0bfe79f289278f3badeb5842fd895c8635d7f7b4
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57087665"
---
# <a name="outlook-add-in-api-requirement-set-19"></a>Outlook API 要件セット 1.9

Office Outlook JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!NOTE]
> このドキュメントは、最新の要件セット以外の[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)のためのものです。

## <a name="whats-new-in-19"></a>1.9 の新機能

要件セット 1.9 には、要件セット [1.8 のすべての機能が含まれています](../requirement-set-1.8/outlook-requirement-set-1.8.md)。 次の機能が追加されました。

- 追加送信時、カスタム プロパティ、および表示フォーム機能用の新しい API が追加されました。
- のサポートが追加されました `Dialog.messageChild` 。

### <a name="change-log"></a>変更ログ

- [CustomProperties.getAll](/javascript/api/outlook/office.customproperties?view=outlook-js-1.9&preserve-view=true#getAll__)を追加しました: すべてのカスタム プロパティを取得する新 `CustomProperties` しい関数をオブジェクトに追加します。
- [Dialog.messageChild](../../../develop/dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)を追加しました: 作業ウィンドウや UI レス関数ファイルなど、ホスト ページからメッセージをページから開いたダイアログに配信する新しいメソッドを追加します。
- [ExtendedPermissions マニフェスト要素を追加](../../manifest/extendedpermissions.md)しました: [VersionOverrides](../../manifest/versionoverrides.md)マニフェスト要素に子要素を追加します。 アドインが追加送信機能をサポートするには[](../../../outlook/append-on-send.md)、拡張アクセス許可を拡張アクセス許可のコレクション `AppendOnSend` に含める必要があります。
- [Office.context.mailbox.displayAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayAppointmentFormAsync_itemId__options__callback_): 既存の予定を表示するオブジェクトに新しい関数 `Mailbox` を追加しました。 これは、メソッドの非同期バージョン `displayAppointmentForm` です。
- [Office.context.mailbox.displayMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayMessageFormAsync_itemId__options__callback_): 既存のメッセージを表示するオブジェクトに新しい関数 `Mailbox` を追加しました。 これは、メソッドの非同期バージョン `displayMessageForm` です。
- [Office.context.mailbox.displayNewAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayNewAppointmentFormAsync_parameters__options__callback_): 新しい予定フォームを表示するオブジェクトに新しい関数 `Mailbox` を追加しました。 これは、メソッドの非同期バージョン `displayNewAppointmentForm` です。
- [Office.context.mailbox.displayNewMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayNewMessageFormAsync_parameters__options__callback_): 新しいメッセージ フォームを表示するオブジェクトに新しい関数 `Mailbox` を追加しました。 これは、メソッドの非同期バージョン `displayNewMessageForm` です。
- [Office.context.mailbox.item.body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendOnSendAsync_data__options__callback_): 新規作成モードでアイテム本文の末尾にデータを追加する新しい関数をオブジェクトに追加しました。 `Body`
- [Office.context.mailbox.item.displayReplyAllFormAsync](office.context.mailbox.item.md#methods): 読み取りモードで "Reply all" フォームを表示するオブジェクトに新しい関数を `Item` 追加しました。 これは、メソッドの非同期バージョン `displayReplyAllForm` です。
- [Office.context.mailbox.item.displayReplyFormAsync](office.context.mailbox.item.md#methods): 読み取りモードで "Reply" フォームを表示するオブジェクトに新しい関数 `Item` を追加しました。 これは、メソッドの非同期バージョン `displayReplyForm` です。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](../../../outlook/outlook-add-ins-overview.md)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](../../../quickstarts/outlook-quickstart.md)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
