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
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="650b0-103">Outlook アドイン API 要件セットのプレビュー</span><span class="sxs-lookup"><span data-stu-id="650b0-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="650b0-104">Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="650b0-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="650b0-105">このドキュメントは、[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)の**プレビュー**用です。</span><span class="sxs-lookup"><span data-stu-id="650b0-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="650b0-106">この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。</span><span class="sxs-lookup"><span data-stu-id="650b0-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="650b0-107">アドイン マニフェストでこの要件を指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="650b0-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="650b0-108">要件セットのプレビューには、[要件セット 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) のすべての機能が含まれています。</span><span class="sxs-lookup"><span data-stu-id="650b0-108">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="650b0-109">プレビューの機能</span><span class="sxs-lookup"><span data-stu-id="650b0-109">Features in preview</span></span>

<span data-ttu-id="650b0-110">次の機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="650b0-110">The following features are in preview.</span></span>

### <a name="additional-calendar-properties"></a><span data-ttu-id="650b0-111">その他の予定表プロパティ</span><span class="sxs-lookup"><span data-stu-id="650b0-111">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="650b0-112">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="650b0-112">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

<span data-ttu-id="650b0-113">新規作成モードで予定の終日イベントプロパティを表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="650b0-113">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="650b0-114">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="650b0-114">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="650b0-115">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="650b0-115">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

<span data-ttu-id="650b0-116">新規作成モードで予定の秘密度を表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="650b0-116">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="650b0-117">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="650b0-117">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="650b0-118">Office...-Alldayevent</span><span class="sxs-lookup"><span data-stu-id="650b0-118">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="650b0-119">予定が終日イベントであるかどうかを表す新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="650b0-119">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="650b0-120">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="650b0-120">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="650b0-121">Office. メールボックスの秘密度</span><span class="sxs-lookup"><span data-stu-id="650b0-121">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="650b0-122">予定の秘密度を表す新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="650b0-122">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="650b0-123">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="650b0-123">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="650b0-124">MailboxEnums AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="650b0-124">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

<span data-ttu-id="650b0-125">予定で利用可能`AppointmentSensitivityType`な秘密度オプションを表す新しい列挙を追加しました。</span><span class="sxs-lookup"><span data-stu-id="650b0-125">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="650b0-126">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="650b0-126">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="650b0-127">送信時に追加</span><span class="sxs-lookup"><span data-stu-id="650b0-127">Append on send</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="650b0-128">Office.......。</span><span class="sxs-lookup"><span data-stu-id="650b0-128">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="650b0-129">新規作成モードで、アイテム`Body`の本文の最後にデータを追加する新しい関数をオブジェクトに追加しました。</span><span class="sxs-lookup"><span data-stu-id="650b0-129">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="650b0-130">**利用可能な**対象: Outlook on Windows (Office 365 サブスクリプションに接続)、outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="650b0-130">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="650b0-131">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="650b0-131">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="650b0-132">拡張されたアクセス許可のコレクションに`AppendOnSend`拡張アクセス許可が含まれている必要があるマニフェストに、新しい要素を追加しました。</span><span class="sxs-lookup"><span data-stu-id="650b0-132">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="650b0-133">**利用可能な**対象: Outlook on Windows (Office 365 サブスクリプションに接続)、outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="650b0-133">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="650b0-134">操作可能なメッセージとの統合</span><span class="sxs-lookup"><span data-stu-id="650b0-134">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="650b0-135">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="650b0-135">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="650b0-136">アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="650b0-136">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="650b0-137">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="650b0-137">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="650b0-138">メールの署名</span><span class="sxs-lookup"><span data-stu-id="650b0-138">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="650b0-139">SetSignatureAsync を示しています。</span><span class="sxs-lookup"><span data-stu-id="650b0-139">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

<span data-ttu-id="650b0-140">新規作成モードで、アイテム`Body`の本文の署名を追加または置換する新しい関数をオブジェクトに追加しました。</span><span class="sxs-lookup"><span data-stu-id="650b0-140">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="650b0-141">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="650b0-141">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="650b0-142">DisableClientSignatureAsync を示します。</span><span class="sxs-lookup"><span data-stu-id="650b0-142">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="650b0-143">新規作成モードの送信メールボックスのクライアント署名を無効にする新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="650b0-143">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="650b0-144">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="650b0-144">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="650b0-145">GetComposeTypeAsync を示します。</span><span class="sxs-lookup"><span data-stu-id="650b0-145">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

<span data-ttu-id="650b0-146">新規作成モードで、メッセージの作成の種類を取得する新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="650b0-146">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="650b0-147">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="650b0-147">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="650b0-148">。アイテム. isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="650b0-148">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="650b0-149">新規作成モードのアイテムでクライアント署名が有効になっているかどうかを確認する新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="650b0-149">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="650b0-150">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="650b0-150">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="650b0-151">MailboxEnums Setype</span><span class="sxs-lookup"><span data-stu-id="650b0-151">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

<span data-ttu-id="650b0-152">新規作成モードで`ComposeType`使用可能な新しい列挙を追加しました。</span><span class="sxs-lookup"><span data-stu-id="650b0-152">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="650b0-153">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="650b0-153">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="650b0-154">Office テーマ</span><span class="sxs-lookup"><span data-stu-id="650b0-154">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="650b0-155">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="650b0-155">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="650b0-156">Office テーマを取得する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="650b0-156">Added ability to get Office theme.</span></span>

<span data-ttu-id="650b0-157">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="650b0-157">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="650b0-158">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="650b0-158">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="650b0-159">`OfficeThemeChanged` イベントが `Mailbox` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="650b0-159">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="650b0-160">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="650b0-160">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="online-meeting-provider-integration"></a><span data-ttu-id="650b0-161">オンライン会議プロバイダーの統合</span><span class="sxs-lookup"><span data-stu-id="650b0-161">Online meeting provider integration</span></span>

<span data-ttu-id="650b0-162">Outlook mobile アドインでのオンライン会議統合のサポートが追加されました。詳細については、「[オンライン会議プロバイダー用の Outlook モバイルアドインを作成](../../../outlook/online-meeting.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="650b0-162">Added support for online-meeting integration in Outlook mobile add-ins. See [Create an Outlook mobile add-in for an online-meeting provider](../../../outlook/online-meeting.md) to learn more.</span></span>

#### <a name="mobileonlinemeetingcommandsurface-extension-point"></a>[<span data-ttu-id="650b0-163">MobileOnlineMeetingCommandSurface 拡張点</span><span class="sxs-lookup"><span data-stu-id="650b0-163">MobileOnlineMeetingCommandSurface extension point</span></span>](../../manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview)

<span data-ttu-id="650b0-164">マニフェスト`MobileOnlineMeetingCommandSurface`に拡張点を追加しました。</span><span class="sxs-lookup"><span data-stu-id="650b0-164">Added `MobileOnlineMeetingCommandSurface` extension point to manifest.</span></span> <span data-ttu-id="650b0-165">オンライン会議の統合を定義します。</span><span class="sxs-lookup"><span data-stu-id="650b0-165">It defines the online meeting integration.</span></span>

<span data-ttu-id="650b0-166">**利用可能な**対象: Android on Outlook (Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="650b0-166">**Available in**: Outlook on Android (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="sso"></a><span data-ttu-id="650b0-167">SSO</span><span class="sxs-lookup"><span data-stu-id="650b0-167">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="650b0-168">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="650b0-168">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="650b0-169">Microsoft Graph API の[アクセス トークンの取得](../../../outlook/authenticate-a-user-with-an-sso-token.md)をアドインに対して許可する、`getAccessToken` へのアクセスが追加されました。</span><span class="sxs-lookup"><span data-stu-id="650b0-169">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="650b0-170">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac、Outlook on the web (モダン)、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="650b0-170">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="650b0-171">関連項目</span><span class="sxs-lookup"><span data-stu-id="650b0-171">See also</span></span>

- [<span data-ttu-id="650b0-172">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="650b0-172">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="650b0-173">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="650b0-173">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="650b0-174">概要</span><span class="sxs-lookup"><span data-stu-id="650b0-174">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="650b0-175">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="650b0-175">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
