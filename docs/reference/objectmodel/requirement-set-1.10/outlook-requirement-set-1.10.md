---
title: Outlook API 要件セット 1.10
description: アドイン API の要件セット 1.10 Outlook 1.10。
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: b54d327d37acd7b2c7fcff100cc7dbe7a39187c0
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154777"
---
# <a name="outlook-add-in-api-requirement-set-110"></a>Outlook API 要件セット 1.10

Office Outlook JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

## <a name="whats-new-in-110"></a>1.10 の新機能

要件セット 1.10 には、要件セット [1.9 のすべての機能が含まれています](../requirement-set-1.9/outlook-requirement-set-1.9.md)。 次の機能が追加されました。

- イベント ベースのアクティブ化機能と [メール署名機能用の新](../../../outlook/autolaunch.md) しい API が追加されました。
- 通知メッセージにカスタム アクションを含める機能が追加されました。

### <a name="change-log"></a>変更ログ

- [LaunchEvent 拡張ポイントを追加しました](../../manifest/extensionpoint.md#launchevent): サポートされている新しい種類の ExtensionPoint を追加します。 イベント ベースのアクティブ化機能を構成します。
- [LaunchEvents manifest 要素を追加](../../manifest/launchevents.md)しました: イベント ベースのアクティブ化機能の構成をサポートするマニフェスト要素を追加します。
- 変更[されたランタイム マニフェスト要素](../../manifest/runtimes.md): サポートOutlook追加します。 イベント ベースのアクティブ化機能に必要な HTML ファイルと JavaScript ファイルを参照します。
- [Office.context.mailbox.item.body.setSignatureAsync](/javascript/api/outlook/office.body?view=outlook-js-1.10&preserve-view=true#setSignatureAsync_data__options__callback_): オブジェクトに新しい関数を追加 `Body` しました。 作成モードでアイテム本文の署名を追加または置き換えます。
- [Office.context.mailbox.item.disableClientSignatureAsync](office.context.mailbox.item.md#methods): 新規作成モードで送信側メールボックスのクライアント署名を無効にする新しい関数を追加しました。
- [Office.context.mailbox.item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getComposeTypeAsync_options__callback_): 新規作成モードでメッセージの作成の種類を取得する新しい関数を追加しました。
- [Office.context.mailbox.item.isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods)を追加しました: 新規作成モードでアイテムでクライアント署名が有効になっているか確認する新しい関数を追加します。
- 追加された[Office。MailboxEnums.ActionType](/javascript/api/outlook/office.mailboxenums.actiontype): 新しい列挙型を追加します。 通知メッセージのカスタム アクションの種類を表します。
- [Office.MailboxEnums.ComposeType](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-1.10&preserve-view=true): 新規作成モードで使用できる新しい列挙型を追加しました。
- 追加された[Office。MailboxEnums.ItemNotificationMessageType.InsightMessage](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype): 列挙型に新しい型を追加 `ItemNotificationMessageType` します。 カスタム アクションを含む通知メッセージを表します。
- 追加された[Office。NotificationMessageAction](/javascript/api/outlook/office.notificationmessageaction): 通知のカスタム アクションを定義できる新しいオブジェクトを追加 `InsightMessage` します。
- 追加された[Office。NotificationMessageDetails.actions](/javascript/api/outlook/office.notificationmessagedetails#actions): カスタム アクションで通知を追加できる新しい `InsightMessage` プロパティを追加します。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](../../../outlook/outlook-add-ins-overview.md)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](../../../quickstarts/outlook-quickstart.md)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
