---
title: Outlook アドイン API 要件セットのプレビュー
description: Outlook アドイン用に現在プレビューされている機能と Api。
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 2f83f81dcf7aa7ab0e3a48fff4279c1e08ba6286
ms.sourcegitcommit: cba180ae712d88d8d9ec417b4d1c7112cd8fdd17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/09/2020
ms.locfileid: "49612751"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook アドイン API 要件セットのプレビュー

Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。

> [!IMPORTANT]
> このドキュメントは、[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)の **プレビュー** 用です。 この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。 アドイン マニフェストでこの要件を指定しないでください。

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> [Microsoft 365 テナントで対象指定リリースを構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)することで、web 上の Outlook の機能をプレビューできる場合があります。 該当する機能については、このページにある「プレビューアクセスを構成する」を確認してください。
>
> その他の機能については、Microsoft 365 アカウントを使用して web 上の Outlook のプレビュービットへのアクセス権を要求することができます。 [このフォーム](https://aka.ms/OWAPreview)を完成して送信します。 これらの機能については、「要求プレビューアクセス」を確認してください。

要件セットのプレビューには、 [要件セット 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md)のすべての機能が含まれています。

## <a name="features-in-preview"></a>プレビューの機能

次の機能はプレビュー段階です。

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a>Information Rights Management (IRM) で保護されたアイテムでのアドインのアクティブ化

これで、IRM で保護されたアイテムでアドインをアクティブ化できるようになります。 この機能を有効にするには、テナント管理者が `OBJMODEL` Office の [プログラムに **よるアクセスを許可** する] オプションを設定して使用権限を有効にする必要があります。 詳細については [、「使用権限と説明](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) 」を参照してください。

**利用可能**: Windows on Windows、build 13229.10000 (Microsoft 365 サブスクリプションに接続)

<br>

---

---

### <a name="additional-calendar-properties"></a>その他の予定表プロパティ

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

新規作成モードで予定の終日イベントプロパティを表す新しいオブジェクトを追加しました。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

新規作成モードで予定の秘密度を表す新しいオブジェクトを追加しました。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office...-Alldayevent](office.context.mailbox.item.md#properties)

予定が終日イベントであるかどうかを表す新しいプロパティを追加しました。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office. メールボックスの秘密度](office.context.mailbox.item.md#properties)

予定の秘密度を表す新しいプロパティを追加しました。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[MailboxEnums AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

`AppointmentSensitivityType`予定で利用可能な秘密度オプションを表す新しい列挙を追加しました。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

<br>

---

---

### <a name="event-based-activation"></a>イベントベースのライセンス認証

Outlook アドインでのイベントベースのアクティブ化機能のサポートが追加されました。詳細については [、「イベントベースのライセンス認証用の Outlook アドインを構成](../../../outlook/autolaunch.md) する」を参照してください。

#### <a name="launchevent-extension-point"></a>[LaunchEvent 拡張点](../../manifest/extensionpoint.md#launchevent-preview)

`LaunchEvent`マニフェストに拡張点サポートが追加されました。 イベントベースのライセンス認証機能を構成します。

**利用可能な** 機能: web 上の Outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="launchevents-manifest-element"></a>[LaunchEvents マニフェスト要素](../../manifest/launchevents.md)

`LaunchEvents`マニフェストに要素を追加しました。 イベントベースのアクティブ化機能の構成をサポートしています。

**利用可能な** 機能: web 上の Outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="runtimes-manifest-element"></a>[ランタイムマニフェスト要素](../../manifest/runtimes.md)

マニフェスト要素に Outlook サポートが追加されました `Runtimes` 。 イベントベースのアクティブ化機能に必要な HTML ファイルと JavaScript ファイルを参照します。

**利用可能な** 機能: web 上の Outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### <a name="integration-with-actionable-messages"></a>操作可能なメッセージとの統合

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)

<br>

---

---

### <a name="mail-signature"></a>メールの署名

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[SetSignatureAsync を示しています。](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

新規 `Body` 作成モードで、アイテムの本文の署名を追加または置換する新しい関数をオブジェクトに追加しました。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[DisableClientSignatureAsync を示します。](office.context.mailbox.item.md#methods)

新規作成モードの送信メールボックスのクライアント署名を無効にする新しい関数を追加しました。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[GetComposeTypeAsync を示します。](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

新規作成モードで、メッセージの作成の種類を取得する新しい関数を追加しました。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[。アイテム. isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods)

新規作成モードのアイテムでクライアント署名が有効になっているかどうかを確認する新しい関数を追加しました。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="officemailboxenumscomposetype"></a>[MailboxEnums Setype](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

新規 `ComposeType` 作成モードで使用可能な新しい列挙を追加しました。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### <a name="notification-messages-with-actions"></a>アクションを含む通知メッセージ

この機能を使用すると、既定の **アラーム** 処理に加えて、カスタムアクションを含む通知メッセージをアドインに含めることができます。 モダン Outlook on the web では、この機能は新規作成モードでのみ利用できます。

#### <a name="officenotificationmessagedetailsactions"></a>[Office の NotificationMessageDetails。アクション](/javascript/api/outlook/office.notificationmessagedetails#actions)

`InsightMessage`カスタムアクションを使用して通知を追加できるようにする新しいプロパティを追加しました。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)

#### <a name="officenotificationmessageaction"></a>[Office NotificationMessageAction](/javascript/api/outlook/office.notificationmessageaction)

通知のカスタムアクションを定義する新しいオブジェクトを追加しました `InsightMessage` 。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)

#### <a name="officemailboxenumsactiontype"></a>[MailboxEnums](/javascript/api/outlook/office.mailboxenums.actiontype)

新しい列挙を追加 `ActionType` しました。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[InsightMessage を MailboxEnums します。](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

Enum に新しい型を追加しました `InsightMessage` `ItemNotificationMessageType` 。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)

<br>

---

---

### <a name="office-theme"></a>Office テーマ

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Office テーマを取得する機能が追加されました。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

`OfficeThemeChanged` イベントが `Mailbox` に追加されました。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

<br>

---

---

### <a name="session-data"></a>セッション データ

#### <a name="officesessiondata"></a>[Office セッションデータ](/javascript/api/outlook/office.sessiondata)

アイテムのセッションデータを表す新しいオブジェクトを追加しました。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

#### <a name="officecontextmailboxitemsessiondata"></a>[Office. メールボックス (セッション)](office.context.mailbox.item.md#properties)

新規作成モードのアイテムのセッションデータを管理するための新しいプロパティを追加しました。

**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

## <a name="see-also"></a>関連項目

- [Outlook アドイン](../../../outlook/outlook-add-ins-overview.md)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](../../../quickstarts/outlook-quickstart.md)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
