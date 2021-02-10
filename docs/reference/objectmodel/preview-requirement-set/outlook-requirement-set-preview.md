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
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="a3585-103">Outlook アドイン API 要件セットのプレビュー</span><span class="sxs-lookup"><span data-stu-id="a3585-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="a3585-104">Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="a3585-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a3585-105">このドキュメントは、[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)の **プレビュー** 用です。</span><span class="sxs-lookup"><span data-stu-id="a3585-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="a3585-106">この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。</span><span class="sxs-lookup"><span data-stu-id="a3585-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="a3585-107">アドイン マニフェストでこの要件を指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="a3585-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="a3585-108">Microsoft 365 テナントで対象指定リリースを構成することで、Outlook on the web の機能 [をプレビューできる場合があります](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。</span><span class="sxs-lookup"><span data-stu-id="a3585-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="a3585-109">該当する機能については、このページに「プレビュー アクセスを構成する」と示されています。</span><span class="sxs-lookup"><span data-stu-id="a3585-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="a3585-110">その他の機能については、このフォームを入力して送信することにより、Microsoft 365 アカウントを使用して Web 上の Outlook のプレビュー ビットへのアクセスを [要求できる場合があります](https://aka.ms/OWAPreview)。</span><span class="sxs-lookup"><span data-stu-id="a3585-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="a3585-111">"プレビュー アクセスの要求" は、これらの機能に示されています。</span><span class="sxs-lookup"><span data-stu-id="a3585-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="a3585-112">要件セットのプレビューには、要件セット [1.9 のすべての機能が含まれます](../requirement-set-1.9/outlook-requirement-set-1.9.md)。</span><span class="sxs-lookup"><span data-stu-id="a3585-112">The Preview Requirement set includes all of the features of [Requirement set 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="a3585-113">プレビューの機能</span><span class="sxs-lookup"><span data-stu-id="a3585-113">Features in preview</span></span>

<span data-ttu-id="a3585-114">次の機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="a3585-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="a3585-115">Information Rights Management (IRM) で保護されたアイテムに対するアドインのアクティブ化</span><span class="sxs-lookup"><span data-stu-id="a3585-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="a3585-116">アドインは、IRM で保護されたアイテムに対してアクティブ化できます。</span><span class="sxs-lookup"><span data-stu-id="a3585-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="a3585-117">この機能を有効にするには、テナント管理者は、テナント管理者に対して [プログラムによるアクセスを許可する] カスタム ポリシー オプションを設定して、使用権限を有効にする `OBJMODEL` Office。 </span><span class="sxs-lookup"><span data-stu-id="a3585-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="a3585-118">詳細 [については、「使用権と説明](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a3585-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="a3585-119">**利用できる** 場所 : Windows 上の Outlook、ビルド 13229.10000 (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a3585-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="a3585-120">その他の予定表のプロパティ</span><span class="sxs-lookup"><span data-stu-id="a3585-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="a3585-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="a3585-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="a3585-122">新規作成モードの予定の全日イベント プロパティを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="a3585-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="a3585-123">**利用できる** 場所 : Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a3585-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="a3585-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="a3585-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="a3585-125">新規作成モードの予定の感度を表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="a3585-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="a3585-126">**利用できる** 場所 : Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a3585-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="a3585-127">Office.context.mailbox.item.isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="a3585-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="a3585-128">予定が 1 日のイベントの場合を表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="a3585-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="a3585-129">**利用できる** 場所 : Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a3585-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="a3585-130">Office.context.mailbox.item.sensitivity</span><span class="sxs-lookup"><span data-stu-id="a3585-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="a3585-131">予定の感度を表す新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="a3585-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="a3585-132">**利用できる** 場所 : Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a3585-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="a3585-133">Office.MailboxEnums.AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="a3585-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="a3585-134">予定で使用可能な `AppointmentSensitivityType` 感度オプションを表す新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="a3585-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="a3585-135">**利用できる** 場所 : Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a3585-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="a3585-136">イベントベースのライセンス認証</span><span class="sxs-lookup"><span data-stu-id="a3585-136">Event-based activation</span></span>

<span data-ttu-id="a3585-137">Outlook アドインのイベント ベースのアクティブ化機能のサポートが追加されました。詳細 [については、「イベント ベースのアクティブ化のために Outlook アドインを構成する](../../../outlook/autolaunch.md) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a3585-137">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="a3585-138">LaunchEvent 拡張点</span><span class="sxs-lookup"><span data-stu-id="a3585-138">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="a3585-139">マニフェストに `LaunchEvent` 拡張ポイントのサポートを追加しました。</span><span class="sxs-lookup"><span data-stu-id="a3585-139">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="a3585-140">イベント ベースのアクティブ化機能を構成します。</span><span class="sxs-lookup"><span data-stu-id="a3585-140">It configures event-based activation functionality.</span></span>

<span data-ttu-id="a3585-141">**利用できる** 場所: Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン、プレビュー アクセス [の構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="a3585-141">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="a3585-142">LaunchEvents マニフェスト要素</span><span class="sxs-lookup"><span data-stu-id="a3585-142">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="a3585-143">マニフェストに `LaunchEvents` 要素を追加しました。</span><span class="sxs-lookup"><span data-stu-id="a3585-143">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="a3585-144">イベント ベースのアクティブ化機能の構成をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="a3585-144">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="a3585-145">**利用できる** 場所: Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン、プレビュー アクセス [の構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="a3585-145">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="a3585-146">ランタイム マニフェスト要素</span><span class="sxs-lookup"><span data-stu-id="a3585-146">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="a3585-147">マニフェスト要素に Outlook サポートを `Runtimes` 追加しました。</span><span class="sxs-lookup"><span data-stu-id="a3585-147">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="a3585-148">イベント ベースのアクティブ化機能に必要な HTML ファイルと JavaScript ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="a3585-148">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="a3585-149">**利用できる** 場所: Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン、プレビュー アクセス [の構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="a3585-149">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="a3585-150">操作可能なメッセージとの統合</span><span class="sxs-lookup"><span data-stu-id="a3585-150">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="a3585-151">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="a3585-151">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="a3585-152">アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="a3585-152">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="a3585-153">**利用できる** 場所 : Outlook on Windows (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="a3585-153">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="a3585-154">メール署名</span><span class="sxs-lookup"><span data-stu-id="a3585-154">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="a3585-155">Office.context.mailbox.item.body.setSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="a3585-155">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

<span data-ttu-id="a3585-156">新規作成モードでアイテム本文の署名を追加または置換する新しい関数 `Body` をオブジェクトに追加しました。</span><span class="sxs-lookup"><span data-stu-id="a3585-156">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="a3585-157">**利用できる** 場所: Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン、プレビュー アクセス [の構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="a3585-157">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="a3585-158">Office.context.mailbox.item.disableClientSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="a3585-158">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="a3585-159">新規作成モードで送信側メールボックスのクライアント署名を無効にする新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="a3585-159">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="a3585-160">**利用できる** 場所: Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン、プレビュー アクセス [の構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="a3585-160">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="a3585-161">Office.context.mailbox.item.getComposeTypeAsync</span><span class="sxs-lookup"><span data-stu-id="a3585-161">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

<span data-ttu-id="a3585-162">新規作成モードでメッセージの作成の種類を取得する新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="a3585-162">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="a3585-163">**利用できる** 場所: Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン、プレビュー アクセス [の構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="a3585-163">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="a3585-164">Office.context.mailbox.item.isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="a3585-164">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="a3585-165">新規作成モードのアイテムでクライアント署名が有効になっているか確認する新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="a3585-165">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="a3585-166">**利用できる** 場所: Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン、プレビュー アクセス [の構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="a3585-166">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="a3585-167">Office.MailboxEnums.ComposeType</span><span class="sxs-lookup"><span data-stu-id="a3585-167">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="a3585-168">新規作成モードで使用可能な `ComposeType` 新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="a3585-168">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="a3585-169">**利用できる** 場所: Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン、プレビュー アクセス [の構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="a3585-169">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="notification-messages-with-actions"></a><span data-ttu-id="a3585-170">アクションを含む通知メッセージ</span><span class="sxs-lookup"><span data-stu-id="a3585-170">Notification messages with actions</span></span>

<span data-ttu-id="a3585-171">この機能を使用すると、既定の [閉じ] アクション以外のカスタム アクションを含む通知メッセージをアドインに **含** めできます。</span><span class="sxs-lookup"><span data-stu-id="a3585-171">This feature allows your add-in to include a notification message with a custom action besides the default **Dismiss** action.</span></span> <span data-ttu-id="a3585-172">最新の Outlook on the web では、この機能は新規作成モードでのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="a3585-172">In modern Outlook on the web, this feature is available in Compose mode only.</span></span>

#### <a name="officenotificationmessagedetailsactions"></a>[<span data-ttu-id="a3585-173">Office.NotificationMessageDetails.actions</span><span class="sxs-lookup"><span data-stu-id="a3585-173">Office.NotificationMessageDetails.actions</span></span>](/javascript/api/outlook/office.notificationmessagedetails#actions)

<span data-ttu-id="a3585-174">カスタム アクションで通知を追加できる新しい `InsightMessage` プロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="a3585-174">Added a new property that enables you to add an `InsightMessage` notification with a custom action.</span></span>

<span data-ttu-id="a3585-175">**利用できる** 場所 : Outlook on Windows (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="a3585-175">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officenotificationmessageaction"></a>[<span data-ttu-id="a3585-176">Office.NotificationMessageAction</span><span class="sxs-lookup"><span data-stu-id="a3585-176">Office.NotificationMessageAction</span></span>](/javascript/api/outlook/office.notificationmessageaction)

<span data-ttu-id="a3585-177">通知のカスタム アクションを定義する新しいオブジェクトが追加 `InsightMessage` されました。</span><span class="sxs-lookup"><span data-stu-id="a3585-177">Added a new object where you define a custom action for your `InsightMessage` notification.</span></span>

<span data-ttu-id="a3585-178">**利用できる** 場所 : Outlook on Windows (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="a3585-178">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsactiontype"></a>[<span data-ttu-id="a3585-179">Office.MailboxEnums.ActionType</span><span class="sxs-lookup"><span data-stu-id="a3585-179">Office.MailboxEnums.ActionType</span></span>](/javascript/api/outlook/office.mailboxenums.actiontype)

<span data-ttu-id="a3585-180">新しい列挙型を追加しました `ActionType` 。</span><span class="sxs-lookup"><span data-stu-id="a3585-180">Added a new enum `ActionType`.</span></span>

<span data-ttu-id="a3585-181">**利用できる** 場所 : Outlook on Windows (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="a3585-181">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[<span data-ttu-id="a3585-182">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span><span class="sxs-lookup"><span data-stu-id="a3585-182">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span></span>](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

<span data-ttu-id="a3585-183">列挙型に新しい `InsightMessage` 型を追加 `ItemNotificationMessageType` しました。</span><span class="sxs-lookup"><span data-stu-id="a3585-183">Added a new type `InsightMessage` to the `ItemNotificationMessageType` enum.</span></span>

<span data-ttu-id="a3585-184">**利用できる** 場所 : Outlook on Windows (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="a3585-184">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="a3585-185">Office テーマ</span><span class="sxs-lookup"><span data-stu-id="a3585-185">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="a3585-186">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="a3585-186">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="a3585-187">Office テーマを取得する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="a3585-187">Added ability to get Office theme.</span></span>

<span data-ttu-id="a3585-188">**利用できる** 場所 : Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a3585-188">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="a3585-189">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="a3585-189">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="a3585-190">`OfficeThemeChanged` イベントが `Mailbox` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="a3585-190">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="a3585-191">**利用できる** 場所 : Windows 上の Outlook (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a3585-191">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="a3585-192">セッション データ</span><span class="sxs-lookup"><span data-stu-id="a3585-192">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="a3585-193">Office.SessionData</span><span class="sxs-lookup"><span data-stu-id="a3585-193">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="a3585-194">アイテムのセッション データを表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="a3585-194">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="a3585-195">**利用できる** 場所 : Outlook on Windows (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="a3585-195">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="a3585-196">Office.context.mailbox.item.sessionData</span><span class="sxs-lookup"><span data-stu-id="a3585-196">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="a3585-197">新規作成モードでアイテムのセッション データを管理するための新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="a3585-197">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="a3585-198">**利用できる** 場所 : Outlook on Windows (Microsoft 365 サブスクリプションに接続)、Outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="a3585-198">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

## <a name="see-also"></a><span data-ttu-id="a3585-199">関連項目</span><span class="sxs-lookup"><span data-stu-id="a3585-199">See also</span></span>

- [<span data-ttu-id="a3585-200">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="a3585-200">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="a3585-201">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="a3585-201">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="a3585-202">概要</span><span class="sxs-lookup"><span data-stu-id="a3585-202">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="a3585-203">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="a3585-203">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
