---
title: Outlook アドイン API 要件セットのプレビュー
description: Outlook アドインのプレビュー中の機能と API。
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 92ba3510af0c8b9ebdf9ca4368c889b821a9cb3b
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173956"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook アドイン API 要件セットのプレビュー

Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。

> [!IMPORTANT]
> このドキュメントは、[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)の **プレビュー** 用です。 この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。 アドイン マニフェストでこの要件を指定しないでください。

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> Microsoft 365 テナントで対象指定リリースを構成することで、Outlook on the web の機能 [をプレビューできる場合があります](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。 該当する機能については、このページに「プレビュー アクセスを構成する」と示されています。
>
> その他の機能については、このフォームを入力して送信することにより、Microsoft 365 アカウントを使用して Web 上の Outlook のプレビュー ビットへのアクセスを [要求できる場合があります](https://aka.ms/OWAPreview)。 "プレビュー アクセスの要求" は、これらの機能に示されています。

要件セットのプレビューには、要件セット [1.9 のすべての機能が含まれます](../requirement-set-1.9/outlook-requirement-set-1.9.md)。

## <a name="features-in-preview"></a>プレビューの機能

次の機能はプレビュー段階です。

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a>Information Rights Management (IRM) で保護されたアイテムに対するアドインのアクティブ化

アドインは、IRM で保護されたアイテムに対してアクティブ化できます。 この機能を有効にするには、テナント管理者は、テナント管理者に対して [プログラムによるアクセスを許可する] カスタム ポリシー オプションを設定して、使用権限を有効にする `OBJMODEL` Office。  詳細 [については、「使用権と説明](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) 」を参照してください。

**利用できる** 場所 : Windows 上の Outlook、ビルド 13229.10000 (Microsoft 365 サブスクリプションに接続)

<br>

---

---

### <a name="additional-calendar-properties"></a>その他の予定表のプロパティ

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

新規作成モードの予定の全日イベント プロパティを表す新しいオブジェクトが追加されました。

**利用できる** 場所 : Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

新規作成モードの予定の感度を表す新しいオブジェクトが追加されました。

**利用できる** 場所 : Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office.context.mailbox.item.isAllDayEvent](office.context.mailbox.item.md#properties)

予定が 1 日のイベントの場合を表す新しいプロパティが追加されました。

**利用できる** 場所 : Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office.context.mailbox.item.sensitivity](office.context.mailbox.item.md#properties)

予定の感度を表す新しいプロパティを追加しました。

**利用できる** 場所 : Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[Office.MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

予定で使用可能な `AppointmentSensitivityType` 感度オプションを表す新しい列挙型が追加されました。

**利用できる** 場所 : Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)

<br>

---

---

### <a name="event-based-activation"></a>イベントベースのライセンス認証

Outlook アドインのイベント ベースのアクティブ化機能のサポートが追加されました。詳細 [については、「イベント ベースのアクティブ化のために Outlook アドインを構成する](../../../outlook/autolaunch.md) 」を参照してください。

#### <a name="launchevent-extension-point"></a>[LaunchEvent 拡張点](../../manifest/extensionpoint.md#launchevent-preview)

マニフェストに `LaunchEvent` 拡張ポイントのサポートを追加しました。 イベント ベースのアクティブ化機能を構成します。

**利用できる** 場所: Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン、プレビュー アクセス [の構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="launchevents-manifest-element"></a>[LaunchEvents マニフェスト要素](../../manifest/launchevents.md)

マニフェストに `LaunchEvents` 要素を追加しました。 イベント ベースのアクティブ化機能の構成をサポートしています。

**利用できる** 場所: Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン、プレビュー アクセス [の構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="runtimes-manifest-element"></a>[ランタイム マニフェスト要素](../../manifest/runtimes.md)

マニフェスト要素に Outlook サポートを `Runtimes` 追加しました。 イベント ベースのアクティブ化機能に必要な HTML ファイルと JavaScript ファイルを参照します。

**利用できる** 場所: Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン、プレビュー アクセス [の構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### <a name="integration-with-actionable-messages"></a>操作可能なメッセージとの統合

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。

**利用できる** 場所 : Outlook on Windows (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン)

<br>

---

---

### <a name="mail-signature"></a>メール署名

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[Office.context.mailbox.item.body.setSignatureAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

新規作成モードでアイテム本文の署名を追加または置換する新しい関数 `Body` をオブジェクトに追加しました。

**利用できる** 場所: Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン、プレビュー アクセス [の構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[Office.context.mailbox.item.disableClientSignatureAsync](office.context.mailbox.item.md#methods)

新規作成モードで送信側メールボックスのクライアント署名を無効にする新しい関数を追加しました。

**利用できる** 場所: Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン、プレビュー アクセス [の構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[Office.context.mailbox.item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

新規作成モードでメッセージの作成の種類を取得する新しい関数が追加されました。

**利用できる** 場所: Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン、プレビュー アクセス [の構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[Office.context.mailbox.item.isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods)

新規作成モードのアイテムでクライアント署名が有効になっているか確認する新しい関数が追加されました。

**利用できる** 場所: Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン、プレビュー アクセス [の構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="officemailboxenumscomposetype"></a>[Office.MailboxEnums.ComposeType](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

新規作成モードで使用可能な `ComposeType` 新しい列挙型が追加されました。

**利用できる** 場所: Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン、プレビュー アクセス [の構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### <a name="notification-messages-with-actions"></a>アクションを含む通知メッセージ

この機能を使用すると、既定の [閉じ] アクション以外のカスタム アクションを含む通知メッセージをアドインに **含** めできます。 最新の Outlook on the web では、この機能は新規作成モードでのみ使用できます。

#### <a name="officenotificationmessagedetailsactions"></a>[Office.NotificationMessageDetails.actions](/javascript/api/outlook/office.notificationmessagedetails#actions)

カスタム アクションで通知を追加できる新しい `InsightMessage` プロパティが追加されました。

**利用できる** 場所 : Outlook on Windows (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン)

#### <a name="officenotificationmessageaction"></a>[Office.NotificationMessageAction](/javascript/api/outlook/office.notificationmessageaction)

通知のカスタム アクションを定義する新しいオブジェクトが追加 `InsightMessage` されました。

**利用できる** 場所 : Outlook on Windows (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン)

#### <a name="officemailboxenumsactiontype"></a>[Office.MailboxEnums.ActionType](/javascript/api/outlook/office.mailboxenums.actiontype)

新しい列挙型を追加しました `ActionType` 。

**利用できる** 場所 : Outlook on Windows (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン)

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[Office.MailboxEnums.ItemNotificationMessageType.InsightMessage](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

列挙型に新しい `InsightMessage` 型を追加 `ItemNotificationMessageType` しました。

**利用できる** 場所 : Outlook on Windows (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン)

<br>

---

---

### <a name="office-theme"></a>Office テーマ

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Office テーマを取得する機能が追加されました。

**利用できる** 場所 : Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

`OfficeThemeChanged` イベントが `Mailbox` に追加されました。

**利用できる** 場所 : Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)

<br>

---

---

### <a name="session-data"></a>セッション データ

#### <a name="officesessiondata"></a>[Office.SessionData](/javascript/api/outlook/office.sessiondata)

アイテムのセッション データを表す新しいオブジェクトを追加しました。

**利用できる** 場所 : Outlook on Windows (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン)

#### <a name="officecontextmailboxitemsessiondata"></a>[Office.context.mailbox.item.sessionData](office.context.mailbox.item.md#properties)

新規作成モードでアイテムのセッション データを管理するための新しいプロパティが追加されました。

**利用できる** 場所 : Outlook on Windows (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン)

## <a name="see-also"></a>関連項目

- [Outlook アドイン](../../../outlook/outlook-add-ins-overview.md)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](../../../quickstarts/outlook-quickstart.md)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
