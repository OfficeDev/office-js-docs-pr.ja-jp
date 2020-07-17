---
title: Outlook アドイン API 要件セットのプレビュー
description: Outlook アドイン用に現在プレビューされている機能と Api。
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: 457195b7511d4dabca101242400d44154a57a781
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159221"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook アドイン API 要件セットのプレビュー

Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。

> [!IMPORTANT]
> このドキュメントは、[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)の**プレビュー**用です。 この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。 アドイン マニフェストでこの要件を指定しないでください。

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> [Microsoft 365 テナントで対象指定リリースを構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)することで、web 上の Outlook の機能をプレビューできる場合があります。 該当する機能については、このページにある「プレビューアクセスを構成する」を確認してください。
>
> その他の機能については、Microsoft 365 アカウントを使用して web 上の Outlook のプレビュービットへのアクセス権を要求することができます。[このフォーム](https://aka.ms/OWAPreview)を完成して送信します。 これらの機能については、「要求プレビューアクセス」を確認してください。

要件セットのプレビューには、[要件セット 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) のすべての機能が含まれています。

## <a name="features-in-preview"></a>プレビューの機能

次の機能はプレビュー段階です。

### <a name="additional-calendar-properties"></a>その他の予定表プロパティ

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

新規作成モードで予定の終日イベントプロパティを表す新しいオブジェクトを追加しました。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

新規作成モードで予定の秘密度を表す新しいオブジェクトを追加しました。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office...-Alldayevent](office.context.mailbox.item.md#properties)

予定が終日イベントであるかどうかを表す新しいプロパティを追加しました。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office. メールボックスの秘密度](office.context.mailbox.item.md#properties)

予定の秘密度を表す新しいプロパティを追加しました。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[MailboxEnums AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

`AppointmentSensitivityType`予定で利用可能な秘密度オプションを表す新しい列挙を追加しました。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

<br>

---

---

### <a name="append-on-send"></a>送信時に追加

追加-送信機能の使用方法については、「 [Outlook アドインで送信時に追加を実装](../../../outlook/append-on-send.md)する」を参照してください。

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[Office.......。](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

新規 `Body` 作成モードで、アイテムの本文の最後にデータを追加する新しい関数をオブジェクトに追加しました。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### <a name="extendedpermissions"></a>[ExtendedPermissions](../../manifest/extendedpermissions.md)

拡張された `AppendOnSend` アクセス許可のコレクションに拡張アクセス許可が含まれている必要があるマニフェストに、新しい要素を追加しました。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### <a name="async-versions-of-display-apis"></a>非同期バージョンの `display` api

#### <a name="officecontextmailboxdisplayappointmentformasync"></a>[DisplayAppointmentFormAsync の内容](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displayappointmentformasync-itemid--options--callback-)

既存の予定を表示するオブジェクトに新しい関数を追加 `Mailbox` しました。 これは、メソッドの非同期バージョンです `displayAppointmentForm` 。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

#### <a name="officecontextmailboxdisplaymessageformasync"></a>[Office. mailbox. displayMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaymessageformasync-itemid--options--callback-)

既存のメッセージを表示するオブジェクトに新しい関数を追加しまし `Mailbox` た。 これは、メソッドの非同期バージョンです `displayMessageForm` 。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

#### <a name="officecontextmailboxdisplaynewappointmentformasync"></a>[DisplayNewAppointmentFormAsync の内容](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewappointmentformasync-parameters--options--callback-)

`Mailbox`新しい予定のフォームを表示する新しい関数をオブジェクトに追加しました。 これは、メソッドの非同期バージョンです `displayNewAppointmentForm` 。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

#### <a name="officecontextmailboxdisplaynewmessageformasync"></a>[Office。 displayNewMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewmessageformasync-parameters--options--callback-)

`Mailbox`新しいメッセージフォームを表示する新しい関数をオブジェクトに追加しました。 これは、メソッドの非同期バージョンです `displayNewMessageForm` 。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

#### <a name="officecontextmailboxitemdisplayreplyallformasync"></a>[DisplayReplyAllFormAsync を示します。](office.context.mailbox.item.md#methods)

`Item`読み取りモードで "全員に返信" フォームを表示するオブジェクトに新しい関数を追加しました。 これは、メソッドの非同期バージョンです `displayReplyAllForm` 。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

#### <a name="officecontextmailboxitemdisplayreplyformasync"></a>[DisplayReplyFormAsync を示します。](office.context.mailbox.item.md#methods)

`Item`読み取りモードで "返信" フォームを表示するオブジェクトに新しい関数を追加しました。 これは、メソッドの非同期バージョンです `displayReplyForm` 。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

<br>

---

---

### <a name="event-based-activation"></a>イベントベースのライセンス認証

Outlook アドインでのイベントベースのアクティブ化機能のサポートが追加されました。詳細については[、「イベントベースのライセンス認証用の Outlook アドインを構成](../../../outlook/autolaunch.md)する」を参照してください。

#### <a name="launchevent-extension-point"></a>[LaunchEvent 拡張点](../../manifest/extensionpoint.md#launchevent-preview)

`LaunchEvent`マニフェストに拡張点サポートが追加されました。 イベントベースのライセンス認証機能を構成します。

**利用可能な**機能: web 上の Outlook (モダン、[要求のプレビューアクセス](https://aka.ms/OWAPreview))

#### <a name="launchevents-manifest-element"></a>[LaunchEvents マニフェスト要素](../../manifest/launchevents.md)

`LaunchEvents`マニフェストに要素を追加しました。 イベントベースのアクティブ化機能の構成をサポートしています。

**利用可能な**機能: web 上の Outlook (モダン、[要求のプレビューアクセス](https://aka.ms/OWAPreview))

#### <a name="runtimes-manifest-element"></a>[ランタイムマニフェスト要素](../../manifest/runtimes.md)

マニフェスト要素に Outlook サポートが追加されました `Runtimes` 。 イベントベースのアクティブ化機能に必要な HTML ファイルと JavaScript ファイルを参照します。

**利用可能な**機能: web 上の Outlook (モダン、[要求のプレビューアクセス](https://aka.ms/OWAPreview))

<br>

---

---

### <a name="get-all-custom-properties"></a>すべてのカスタムプロパティを取得する

#### <a name="custompropertiesgetall"></a>[CustomProperties getAll](/javascript/api/outlook/office.customproperties?view=outlook-js-preview#getall--)

すべてのカスタムプロパティを取得する新しい関数をオブジェクトに追加しまし `CustomProperties` た。

**利用可能な**対象: Outlook on Windows (microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)、Outlook on Mac (microsoft 365 サブスクリプションに接続)、outlook on the Outlook on iOS

<br>

---

---

### <a name="integration-with-actionable-messages"></a>操作可能なメッセージとの統合

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (クラシック)

<br>

---

---

### <a name="mail-signature"></a>メールの署名

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[SetSignatureAsync を示しています。](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

新規 `Body` 作成モードで、アイテムの本文の署名を追加または置換する新しい関数をオブジェクトに追加しました。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[DisableClientSignatureAsync を示します。](office.context.mailbox.item.md#methods)

新規作成モードの送信メールボックスのクライアント署名を無効にする新しい関数を追加しました。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[GetComposeTypeAsync を示します。](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

新規作成モードで、メッセージの作成の種類を取得する新しい関数を追加しました。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[。アイテム. isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods)

新規作成モードのアイテムでクライアント署名が有効になっているかどうかを確認する新しい関数を追加しました。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### <a name="officemailboxenumscomposetype"></a>[MailboxEnums Setype](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

新規 `ComposeType` 作成モードで使用可能な新しい列挙を追加しました。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### <a name="office-theme"></a>Office テーマ

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Office テーマを取得する機能が追加されました。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

`OfficeThemeChanged` イベントが `Mailbox` に追加されました。

**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)

<br>

---

---

### <a name="single-sign-on-sso"></a>シングル サインオン (SSO)

#### <a name="officeruntimeauthgetaccesstoken"></a>[OfficeRuntime.auth.getAccessToken](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

Microsoft Graph API の[アクセス トークンの取得](../../../outlook/authenticate-a-user-with-an-sso-token.md)をアドインに対して許可する、`getAccessToken` へのアクセスが追加されました。

**利用可能な**対象: Outlook on Windows (microsoft 365 サブスクリプションに接続)、Outlook on Mac (microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)、outlook on the web (クラシック)

## <a name="see-also"></a>関連項目

- [Outlook アドイン](../../../outlook/outlook-add-ins-overview.md)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](../../../quickstarts/outlook-quickstart.md)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
