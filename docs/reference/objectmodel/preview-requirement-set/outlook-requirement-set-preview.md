---
title: Outlook アドイン API 要件セットのプレビュー
description: Outlook アドインおよび Office JavaScript Api で現在プレビューされている機能と Api。
ms.date: 04/10/2020
localization_priority: Normal
ms.openlocfilehash: f8ef7b8c37dbd7539c30457c4922c1c16262381c
ms.sourcegitcommit: 76552b3e5725d9112c772595971b922c295e6b4c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/10/2020
ms.locfileid: "43225674"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook アドイン API 要件セットのプレビュー

Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。

> [!IMPORTANT]
> このドキュメントは、[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)の**プレビュー**用です。 この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。 アドイン マニフェストでこの要件を指定しないでください。

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

要件セットのプレビューには、[要件セット 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) のすべての機能が含まれています。

## <a name="features-in-preview"></a>プレビューの機能

次の機能はプレビュー段階です。

### <a name="additional-calendar-properties"></a>その他の予定表プロパティ

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

新規作成モードで予定の終日イベントプロパティを表す新しいオブジェクトを追加しました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

新規作成モードで予定の秘密度を表す新しいオブジェクトを追加しました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office...-Alldayevent](office.context.mailbox.item.md#properties)

予定が終日イベントであるかどうかを表す新しいプロパティを追加しました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office. メールボックスの秘密度](office.context.mailbox.item.md#properties)

予定の秘密度を表す新しいプロパティを追加しました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[MailboxEnums AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

予定で利用可能`AppointmentSensitivityType`な秘密度オプションを表す新しい列挙を追加しました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

<br>

---

---

### <a name="append-on-send"></a>送信時に追加

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[Office.......。](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

新規作成モードで、アイテム`Body`の本文の最後にデータを追加する新しい関数をオブジェクトに追加しました。

**利用可能な**対象: Outlook on Windows (Office 365 サブスクリプションに接続)、outlook on the web (モダン)

#### <a name="extendedpermissions"></a>[ExtendedPermissions](../../manifest/extendedpermissions.md)

拡張されたアクセス許可のコレクションに`AppendOnSend`拡張アクセス許可が含まれている必要があるマニフェストに、新しい要素を追加しました。

**利用可能な**対象: Outlook on Windows (Office 365 サブスクリプションに接続)、outlook on the web (モダン)

<br>

---

---

### <a name="integration-with-actionable-messages"></a>操作可能なメッセージとの統合

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (クラシック)

<br>

---

---

### <a name="mail-signature"></a>メールの署名

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[SetSignatureAsync を示しています。](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

新規作成モードで、アイテム`Body`の本文の署名を追加または置換する新しい関数をオブジェクトに追加しました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[DisableClientSignatureAsync を示します。](office.context.mailbox.item.md#methods)

新規作成モードの送信メールボックスのクライアント署名を無効にする新しい関数を追加しました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[GetComposeTypeAsync を示します。](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

新規作成モードで、メッセージの作成の種類を取得する新しい関数を追加しました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[。アイテム. isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods)

新規作成モードのアイテムでクライアント署名が有効になっているかどうかを確認する新しい関数を追加しました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

#### <a name="officemailboxenumscomposetype"></a>[MailboxEnums Setype](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

新規作成モードで`ComposeType`使用可能な新しい列挙を追加しました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

<br>

---

---

### <a name="office-theme"></a>Office テーマ

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Office テーマを取得する機能が追加されました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

`OfficeThemeChanged` イベントが `Mailbox` に追加されました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

<br>

---

---

### <a name="online-meeting-provider-integration"></a>オンライン会議プロバイダーの統合

Outlook mobile アドインでのオンライン会議統合のサポートが追加されました。詳細については、「[オンライン会議プロバイダー用の Outlook モバイルアドインを作成](../../../outlook/online-meeting.md)する」を参照してください。

#### <a name="mobileonlinemeetingcommandsurface-extension-point"></a>[MobileOnlineMeetingCommandSurface 拡張点](../../manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview)

マニフェスト`MobileOnlineMeetingCommandSurface`に拡張点を追加しました。 オンライン会議の統合を定義します。

**利用可能な**対象: Android on Outlook (Office 365 サブスクリプションに接続)

<br>

---

---

### <a name="sso"></a>SSO

#### <a name="officeruntimeauthgetaccesstoken"></a>[OfficeRuntime.auth.getAccessToken](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

Microsoft Graph API の[アクセス トークンの取得](../../../outlook/authenticate-a-user-with-an-sso-token.md)をアドインに対して許可する、`getAccessToken` へのアクセスが追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac、Outlook on the web (モダン)、Outlook on the web (クラシック)

## <a name="see-also"></a>関連項目

- [Outlook アドイン](../../../outlook/outlook-add-ins-overview.md)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](../../../quickstarts/outlook-quickstart.md)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
