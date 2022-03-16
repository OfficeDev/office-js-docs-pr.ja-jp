---
title: Outlook API プレビュー要件セット
description: 現在、アドインのプレビュー中の機能Outlook API。
ms.date: 03/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 714be93351ff67ad49cd07154f145f19949efa68
ms.sourcegitcommit: 856f057a8c9b937bfb37e7d81a6b71dbed4b8ff4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/16/2022
ms.locfileid: "63511280"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook API プレビュー要件セット

Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!IMPORTANT]
> このドキュメントは、[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)の **プレビュー** 用です。 この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。 アドイン マニフェストでこの要件を指定しないでください。

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> ターゲット リリースを構成することで、Outlook on the webの機能をプレビュー Microsoft 365[があります](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。 該当する機能については、このページに「プレビュー アクセスを構成する」と表示されます。
>
> その他の機能については、このフォームを入力して送信することで、Outlook on the web アカウントをMicrosoft 365プレビュー ビットへのアクセスを[要求できる場合があります](https://aka.ms/OWAPreview)。 "要求プレビュー アクセス" は、これらの機能に関して示されています。

プレビュー要件セットには、要件セット [1.11 のすべての機能が含まれています](../requirement-set-1.11/outlook-requirement-set-1.11.md)。

## <a name="features-in-preview"></a>プレビューの機能

次の機能はプレビュー段階です。

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a>Information Rights Management (IRM) によって保護されたアイテムに対するアドインのアクティブ化

アドインは IRM で保護されたアイテムでアクティブ化できます。 この機能を有効にするには、`OBJMODEL`テナント管理者が[プログラムによるアクセスを許可する] カスタム  ポリシー オプションを設定して、使用権を有効にする必要Office。 詳細 [については、「利用状況の権限と説明](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) 」を参照してください。

**で利用可能**: Outlook Windows ビルド 13229.10000 から (Microsoft 365 サブスクリプションに接続)

<br>

---

---

### <a name="additional-calendar-properties"></a>その他の予定表のプロパティ

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

作成モードで予定の全日イベント プロパティを表す新しいオブジェクトを追加しました。

**使用可能な** 場所: Outlook (WindowsサブスクリプションにMicrosoft 365)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

作成モードでの予定の感度を表す新しいオブジェクトを追加しました。

**使用可能な** 場所: Outlook (WindowsサブスクリプションにMicrosoft 365)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office.context.mailbox.item.isAllDayEvent](office.context.mailbox.item.md#properties)

予定が一日のイベントである場合を表す新しいプロパティを追加しました。

**使用可能な** 場所: Outlook (WindowsサブスクリプションにMicrosoft 365)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office.context.mailbox.item.sensitivity](office.context.mailbox.item.md#properties)

予定の感度を表す新しいプロパティを追加しました。

**使用可能な** 場所: Outlook (WindowsサブスクリプションにMicrosoft 365)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[Office。MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

予定で使用できる感度 `AppointmentSensitivityType` オプションを表す新しい列挙型を追加しました。

**使用可能な** 場所: Outlook (WindowsサブスクリプションにMicrosoft 365)

<br>

---

---

### <a name="delay-delivery-time"></a>配信時間の遅延

#### <a name="officecontextmailboxitemdelaydeliverytime"></a>[Office.context.mailbox.item.delayDeliveryTime](office.context.mailbox.item.md#properties)

新規作成モードでメッセージの配信日時を管理できるオブジェクトを返す新しいプロパティを追加しました。

**使用可能な** 場所: Outlook (WindowsサブスクリプションにMicrosoft 365)

#### <a name="officedelaydeliverytime"></a>[Office。DelayDeliveryTime](/javascript/api/outlook/office.delaydeliverytime?view=outlook-js-preview&preserve-view=true)

新規作成モードでメッセージの配信日時を管理できる新しいオブジェクトを追加しました。

**使用可能な** 場所: Outlook (WindowsサブスクリプションにMicrosoft 365)

<br>

---

---

### <a name="event-based-activation"></a>イベントベースのライセンス認証

この機能は、要件セット [1.10 でリリースされました](../requirement-set-1.10/outlook-requirement-set-1.10.md)。 ただし、追加のイベントはプレビューで利用できます。 詳細については、「サポートされるイベント [」を参照してください](../../../outlook/autolaunch.md#supported-events)。

**使用可能な** 場所: Outlook (WindowsサブスクリプションにMicrosoft 365)

#### <a name="officeaddincommandseventcompletedoptionserrormessage"></a>[Office。AddinCommands.EventCompletedOptions.errorMessage](/javascript/api/office/office.addincommands.eventcompletedoptions?view=outlook-js-preview&preserve-view=true#office-office-addincommands-eventcompletedoptions-errormessage-member)

処理されたイベントを引き続き実行できない場合に、ユーザーにエラー メッセージを表示する新しいプロパティを追加しました。 例については、「Smart [Alerts」のチュートリアルを参照してください](../../../outlook/smart-alerts-onmessagesend-walkthrough.md)。

**使用可能な** 場所: Outlook (WindowsサブスクリプションにMicrosoft 365)

<br>

---

---

### <a name="integration-with-actionable-messages"></a>操作可能なメッセージとの統合

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。

**使用可能な** 場所: Outlook (Windowsサブスクリプションに接続されている)、Microsoft 365 (Outlook on the web)

<br>

---

---

### <a name="office-theme"></a>Office テーマ

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true#office-office-context-officetheme-member)

Office テーマを取得する機能が追加されました。

**使用可能な** 場所: Outlook (WindowsサブスクリプションにMicrosoft 365)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true)

`OfficeThemeChanged` イベントが `Mailbox` に追加されました。

**使用可能な** 場所: Outlook (WindowsサブスクリプションにMicrosoft 365)

<br>

---

---

### <a name="shared-mailboxes"></a>共有メールボックス

要件セット 1.8 では、共有フォルダー (つまり、代理人アクセス) の機能サポート [がリリースされました](../requirement-set-1.8/outlook-requirement-set-1.8.md)。 ただし、共有メールボックスのサポートはプレビューで利用できます。 詳細については、「[共有フォルダーと共有メールボックスのシナリオを有効にする](../../../outlook/delegate-access.md)」を参照してください。

**利用可能な** 機能: Outlook (Windows サブスクリプションにMicrosoft 365)、Outlook on the web (モダン)、Mac Outlookを使用できます。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](../../../outlook/outlook-add-ins-overview.md)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](../../../quickstarts/outlook-quickstart.md)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
