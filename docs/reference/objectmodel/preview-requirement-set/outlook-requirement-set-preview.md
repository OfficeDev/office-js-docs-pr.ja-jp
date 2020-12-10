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
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="a8f0f-103">Outlook アドイン API 要件セットのプレビュー</span><span class="sxs-lookup"><span data-stu-id="a8f0f-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="a8f0f-104">Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a8f0f-105">このドキュメントは、[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)の **プレビュー** 用です。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="a8f0f-106">この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="a8f0f-107">アドイン マニフェストでこの要件を指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="a8f0f-108">[Microsoft 365 テナントで対象指定リリースを構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)することで、web 上の Outlook の機能をプレビューできる場合があります。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="a8f0f-109">該当する機能については、このページにある「プレビューアクセスを構成する」を確認してください。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="a8f0f-110">その他の機能については、Microsoft 365 アカウントを使用して web 上の Outlook のプレビュービットへのアクセス権を要求することができます。 [このフォーム](https://aka.ms/OWAPreview)を完成して送信します。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="a8f0f-111">これらの機能については、「要求プレビューアクセス」を確認してください。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="a8f0f-112">要件セットのプレビューには、 [要件セット 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md)のすべての機能が含まれています。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-112">The Preview Requirement set includes all of the features of [Requirement set 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="a8f0f-113">プレビューの機能</span><span class="sxs-lookup"><span data-stu-id="a8f0f-113">Features in preview</span></span>

<span data-ttu-id="a8f0f-114">次の機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="a8f0f-115">Information Rights Management (IRM) で保護されたアイテムでのアドインのアクティブ化</span><span class="sxs-lookup"><span data-stu-id="a8f0f-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="a8f0f-116">これで、IRM で保護されたアイテムでアドインをアクティブ化できるようになります。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="a8f0f-117">この機能を有効にするには、テナント管理者が `OBJMODEL` Office の [プログラムに **よるアクセスを許可** する] オプションを設定して使用権限を有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="a8f0f-118">詳細については [、「使用権限と説明](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="a8f0f-119">**利用可能**: Windows on Windows、build 13229.10000 (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a8f0f-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="a8f0f-120">その他の予定表プロパティ</span><span class="sxs-lookup"><span data-stu-id="a8f0f-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="a8f0f-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="a8f0f-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="a8f0f-122">新規作成モードで予定の終日イベントプロパティを表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="a8f0f-123">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a8f0f-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="a8f0f-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="a8f0f-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="a8f0f-125">新規作成モードで予定の秘密度を表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="a8f0f-126">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a8f0f-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="a8f0f-127">Office...-Alldayevent</span><span class="sxs-lookup"><span data-stu-id="a8f0f-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="a8f0f-128">予定が終日イベントであるかどうかを表す新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="a8f0f-129">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a8f0f-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="a8f0f-130">Office. メールボックスの秘密度</span><span class="sxs-lookup"><span data-stu-id="a8f0f-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="a8f0f-131">予定の秘密度を表す新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="a8f0f-132">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a8f0f-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="a8f0f-133">MailboxEnums AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="a8f0f-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="a8f0f-134">`AppointmentSensitivityType`予定で利用可能な秘密度オプションを表す新しい列挙を追加しました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="a8f0f-135">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a8f0f-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="a8f0f-136">イベントベースのライセンス認証</span><span class="sxs-lookup"><span data-stu-id="a8f0f-136">Event-based activation</span></span>

<span data-ttu-id="a8f0f-137">Outlook アドインでのイベントベースのアクティブ化機能のサポートが追加されました。詳細については [、「イベントベースのライセンス認証用の Outlook アドインを構成](../../../outlook/autolaunch.md) する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-137">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="a8f0f-138">LaunchEvent 拡張点</span><span class="sxs-lookup"><span data-stu-id="a8f0f-138">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="a8f0f-139">`LaunchEvent`マニフェストに拡張点サポートが追加されました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-139">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="a8f0f-140">イベントベースのライセンス認証機能を構成します。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-140">It configures event-based activation functionality.</span></span>

<span data-ttu-id="a8f0f-141">**利用可能な** 機能: web 上の Outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="a8f0f-141">**Available in**: Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="a8f0f-142">LaunchEvents マニフェスト要素</span><span class="sxs-lookup"><span data-stu-id="a8f0f-142">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="a8f0f-143">`LaunchEvents`マニフェストに要素を追加しました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-143">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="a8f0f-144">イベントベースのアクティブ化機能の構成をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-144">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="a8f0f-145">**利用可能な** 機能: web 上の Outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="a8f0f-145">**Available in**: Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="a8f0f-146">ランタイムマニフェスト要素</span><span class="sxs-lookup"><span data-stu-id="a8f0f-146">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="a8f0f-147">マニフェスト要素に Outlook サポートが追加されました `Runtimes` 。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-147">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="a8f0f-148">イベントベースのアクティブ化機能に必要な HTML ファイルと JavaScript ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-148">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="a8f0f-149">**利用可能な** 機能: web 上の Outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="a8f0f-149">**Available in**: Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="a8f0f-150">操作可能なメッセージとの統合</span><span class="sxs-lookup"><span data-stu-id="a8f0f-150">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="a8f0f-151">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="a8f0f-151">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="a8f0f-152">アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-152">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="a8f0f-153">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="a8f0f-153">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="a8f0f-154">メールの署名</span><span class="sxs-lookup"><span data-stu-id="a8f0f-154">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="a8f0f-155">SetSignatureAsync を示しています。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-155">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

<span data-ttu-id="a8f0f-156">新規 `Body` 作成モードで、アイテムの本文の署名を追加または置換する新しい関数をオブジェクトに追加しました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-156">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="a8f0f-157">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="a8f0f-157">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="a8f0f-158">DisableClientSignatureAsync を示します。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-158">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="a8f0f-159">新規作成モードの送信メールボックスのクライアント署名を無効にする新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-159">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="a8f0f-160">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="a8f0f-160">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="a8f0f-161">GetComposeTypeAsync を示します。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-161">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

<span data-ttu-id="a8f0f-162">新規作成モードで、メッセージの作成の種類を取得する新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-162">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="a8f0f-163">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="a8f0f-163">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="a8f0f-164">。アイテム. isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="a8f0f-164">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="a8f0f-165">新規作成モードのアイテムでクライアント署名が有効になっているかどうかを確認する新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-165">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="a8f0f-166">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="a8f0f-166">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="a8f0f-167">MailboxEnums Setype</span><span class="sxs-lookup"><span data-stu-id="a8f0f-167">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="a8f0f-168">新規 `ComposeType` 作成モードで使用可能な新しい列挙を追加しました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-168">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="a8f0f-169">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="a8f0f-169">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="notification-messages-with-actions"></a><span data-ttu-id="a8f0f-170">アクションを含む通知メッセージ</span><span class="sxs-lookup"><span data-stu-id="a8f0f-170">Notification messages with actions</span></span>

<span data-ttu-id="a8f0f-171">この機能を使用すると、既定の **アラーム** 処理に加えて、カスタムアクションを含む通知メッセージをアドインに含めることができます。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-171">This feature allows your add-in to include a notification message with a custom action besides the default **Dismiss** action.</span></span> <span data-ttu-id="a8f0f-172">モダン Outlook on the web では、この機能は新規作成モードでのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-172">In modern Outlook on the web, this feature is available in Compose mode only.</span></span>

#### <a name="officenotificationmessagedetailsactions"></a>[<span data-ttu-id="a8f0f-173">Office の NotificationMessageDetails。アクション</span><span class="sxs-lookup"><span data-stu-id="a8f0f-173">Office.NotificationMessageDetails.actions</span></span>](/javascript/api/outlook/office.notificationmessagedetails#actions)

<span data-ttu-id="a8f0f-174">`InsightMessage`カスタムアクションを使用して通知を追加できるようにする新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-174">Added a new property that enables you to add an `InsightMessage` notification with a custom action.</span></span>

<span data-ttu-id="a8f0f-175">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="a8f0f-175">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officenotificationmessageaction"></a>[<span data-ttu-id="a8f0f-176">Office NotificationMessageAction</span><span class="sxs-lookup"><span data-stu-id="a8f0f-176">Office.NotificationMessageAction</span></span>](/javascript/api/outlook/office.notificationmessageaction)

<span data-ttu-id="a8f0f-177">通知のカスタムアクションを定義する新しいオブジェクトを追加しました `InsightMessage` 。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-177">Added a new object where you define a custom action for your `InsightMessage` notification.</span></span>

<span data-ttu-id="a8f0f-178">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="a8f0f-178">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsactiontype"></a>[<span data-ttu-id="a8f0f-179">MailboxEnums</span><span class="sxs-lookup"><span data-stu-id="a8f0f-179">Office.MailboxEnums.ActionType</span></span>](/javascript/api/outlook/office.mailboxenums.actiontype)

<span data-ttu-id="a8f0f-180">新しい列挙を追加 `ActionType` しました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-180">Added a new enum `ActionType`.</span></span>

<span data-ttu-id="a8f0f-181">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="a8f0f-181">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[<span data-ttu-id="a8f0f-182">InsightMessage を MailboxEnums します。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-182">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span></span>](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

<span data-ttu-id="a8f0f-183">Enum に新しい型を追加しました `InsightMessage` `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-183">Added a new type `InsightMessage` to the `ItemNotificationMessageType` enum.</span></span>

<span data-ttu-id="a8f0f-184">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="a8f0f-184">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="a8f0f-185">Office テーマ</span><span class="sxs-lookup"><span data-stu-id="a8f0f-185">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="a8f0f-186">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="a8f0f-186">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="a8f0f-187">Office テーマを取得する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-187">Added ability to get Office theme.</span></span>

<span data-ttu-id="a8f0f-188">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a8f0f-188">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="a8f0f-189">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="a8f0f-189">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="a8f0f-190">`OfficeThemeChanged` イベントが `Mailbox` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-190">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="a8f0f-191">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a8f0f-191">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="a8f0f-192">セッション データ</span><span class="sxs-lookup"><span data-stu-id="a8f0f-192">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="a8f0f-193">Office セッションデータ</span><span class="sxs-lookup"><span data-stu-id="a8f0f-193">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="a8f0f-194">アイテムのセッションデータを表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-194">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="a8f0f-195">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a8f0f-195">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="a8f0f-196">Office. メールボックス (セッション)</span><span class="sxs-lookup"><span data-stu-id="a8f0f-196">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="a8f0f-197">新規作成モードのアイテムのセッションデータを管理するための新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="a8f0f-197">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="a8f0f-198">**利用可能な** 対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a8f0f-198">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

## <a name="see-also"></a><span data-ttu-id="a8f0f-199">関連項目</span><span class="sxs-lookup"><span data-stu-id="a8f0f-199">See also</span></span>

- [<span data-ttu-id="a8f0f-200">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="a8f0f-200">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="a8f0f-201">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="a8f0f-201">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="a8f0f-202">概要</span><span class="sxs-lookup"><span data-stu-id="a8f0f-202">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="a8f0f-203">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="a8f0f-203">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
