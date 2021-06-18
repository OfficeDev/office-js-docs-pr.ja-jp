---
title: Outlook API プレビュー要件セット
description: 現在、アドインのプレビュー中Outlook API。
ms.date: 06/15/2021
localization_priority: Normal
ms.openlocfilehash: f9d8afc2b4347a8fb13f8ab98a163fb63968123f
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007763"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook API プレビュー要件セット

Office Outlook JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!IMPORTANT]
> このドキュメントは、[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)の **プレビュー** 用です。 この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。 アドイン マニフェストでこの要件を指定しないでください。

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> ターゲット リリースを構成することで、Outlook on the webの機能をプレビューできる場合[Microsoft 365があります](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。 該当する機能については、このページに「プレビュー アクセスを構成する」と表示されます。
>
> その他の機能については、このフォームを入力して送信することで、Outlook on the web アカウントを使用Microsoft 365プレビュー ビットへのアクセスを[要求できる場合があります](https://aka.ms/OWAPreview)。 "要求プレビュー アクセス" は、これらの機能に関して示されています。

プレビュー要件セットには、要件セット [1.10 のすべての機能が含まれています](../requirement-set-1.10/outlook-requirement-set-1.10.md)。

## <a name="features-in-preview"></a>プレビューの機能

次の機能はプレビュー段階です。

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a>Information Rights Management (IRM) によって保護されたアイテムに対するアドインのアクティブ化

アドインは IRM で保護されたアイテムでアクティブ化できます。 この機能を有効にするには、テナント管理者が[プログラムによるアクセスを許可する] カスタム ポリシー オプションを設定して、使用権を有効にする `OBJMODEL` 必要Office。  詳細 [については、「利用状況の権限と説明](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) 」を参照してください。

**で利用可能**: Outlook Windows ビルド 13229.10000 から (Microsoft 365 サブスクリプションに接続)

<br>

---

---

### <a name="additional-calendar-properties"></a>その他の予定表のプロパティ

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

作成モードで予定の全日イベント プロパティを表す新しいオブジェクトを追加しました。

**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

作成モードでの予定の感度を表す新しいオブジェクトを追加しました。

**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office.context.mailbox.item.isAllDayEvent](office.context.mailbox.item.md#properties)

予定が一日のイベントである場合を表す新しいプロパティを追加しました。

**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office.context.mailbox.item.sensitivity](office.context.mailbox.item.md#properties)

予定の感度を表す新しいプロパティを追加しました。

**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[Office。MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

予定で使用できる `AppointmentSensitivityType` 感度オプションを表す新しい列挙型を追加しました。

**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)

<br>

---

---

### <a name="event-based-activation"></a>イベントベースのライセンス認証

この機能は、要件セット [1.10 でリリースされました](../requirement-set-1.10/outlook-requirement-set-1.10.md)。 ただし、追加のイベントはプレビューで利用できます。 詳細については、「サポートされているイベント [」を参照してください](../../../outlook/autolaunch.md#supported-events)。

**で利用可能**: Outlook (Windowsサブスクリプションに接続) 、Microsoft 365 (Outlook on the web)

<br>

---

---

### <a name="integration-with-actionable-messages"></a>操作可能なメッセージとの統合

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。

**で利用可能**: Outlook (Windowsサブスクリプションに接続) 、Microsoft 365 (Outlook on the web)

<br>

---

---

### <a name="office-theme"></a>Office テーマ

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Office テーマを取得する機能が追加されました。

**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

`OfficeThemeChanged` イベントが `Mailbox` に追加されました。

**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)

<br>

---

---

### <a name="session-data"></a>セッション データ

#### <a name="officesessiondata"></a>[Office。SessionData](/javascript/api/outlook/office.sessiondata)

アイテムのセッション データを表す新しいオブジェクトを追加しました。

**で利用可能**: Outlook (Windowsサブスクリプションに接続) 、Microsoft 365 (Outlook on the web)

#### <a name="officecontextmailboxitemsessiondata"></a>[Office.context.mailbox.item.sessionData](office.context.mailbox.item.md#properties)

新規作成モードでアイテムのセッション データを管理するための新しいプロパティを追加しました。

**で利用可能**: Outlook (Windowsサブスクリプションに接続) 、Microsoft 365 (Outlook on the web)

<br>

---

---

### <a name="shared-mailboxes"></a>共有メールボックス

要件セット [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md)で、共有フォルダー (代理人アクセス) の機能サポートがリリースされました。 ただし、共有メールボックスのサポートはプレビューで利用できます。 詳細については、「共有フォルダーと共有 [メールボックスのシナリオを有効にする」を参照してください](../../../outlook/delegate-access.md)。

**で利用可能**: Outlook (Windowsサブスクリプションに接続) 、Microsoft 365 (Outlook on the web)

## <a name="see-also"></a>関連項目

- [Outlook アドイン](../../../outlook/outlook-add-ins-overview.md)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](../../../quickstarts/outlook-quickstart.md)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
