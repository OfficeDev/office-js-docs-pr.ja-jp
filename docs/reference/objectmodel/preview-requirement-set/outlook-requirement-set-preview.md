---
title: Outlook アドイン API 要件セットのプレビュー
description: Outlook アドイン用に現在プレビューされている機能と Api。
ms.date: 05/19/2020
localization_priority: Normal
ms.openlocfilehash: 3183c81a9af99f480c2dbecc787695501380cea7
ms.sourcegitcommit: 8499a4247d1cb1e96e99c17cb520f4a8a41667e3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2020
ms.locfileid: "44292295"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="db6a5-103">Outlook アドイン API 要件セットのプレビュー</span><span class="sxs-lookup"><span data-stu-id="db6a5-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="db6a5-104">Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="db6a5-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="db6a5-105">このドキュメントは、[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)の**プレビュー**用です。</span><span class="sxs-lookup"><span data-stu-id="db6a5-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="db6a5-106">この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。</span><span class="sxs-lookup"><span data-stu-id="db6a5-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="db6a5-107">アドイン マニフェストでこの要件を指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="db6a5-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="db6a5-108">[Microsoft 365 テナントで対象指定リリースを構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)することで、web 上の Outlook の機能をプレビューできる場合があります。</span><span class="sxs-lookup"><span data-stu-id="db6a5-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="db6a5-109">該当する機能については、このページにある「プレビューアクセスを構成する」を確認してください。</span><span class="sxs-lookup"><span data-stu-id="db6a5-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="db6a5-110">その他の機能については、Microsoft 365 アカウントを使用して web 上の Outlook のプレビュービットへのアクセス権を要求することができます。[このフォーム](https://aka.ms/OWAPreview)を完成して送信します。</span><span class="sxs-lookup"><span data-stu-id="db6a5-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="db6a5-111">これらの機能については、「要求プレビューアクセス」を確認してください。</span><span class="sxs-lookup"><span data-stu-id="db6a5-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="db6a5-112">要件セットのプレビューには、[要件セット 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) のすべての機能が含まれています。</span><span class="sxs-lookup"><span data-stu-id="db6a5-112">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="db6a5-113">プレビューの機能</span><span class="sxs-lookup"><span data-stu-id="db6a5-113">Features in preview</span></span>

<span data-ttu-id="db6a5-114">次の機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="db6a5-114">The following features are in preview.</span></span>

### <a name="additional-calendar-properties"></a><span data-ttu-id="db6a5-115">その他の予定表プロパティ</span><span class="sxs-lookup"><span data-stu-id="db6a5-115">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="db6a5-116">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="db6a5-116">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

<span data-ttu-id="db6a5-117">新規作成モードで予定の終日イベントプロパティを表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="db6a5-117">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="db6a5-118">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="db6a5-118">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="db6a5-119">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="db6a5-119">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

<span data-ttu-id="db6a5-120">新規作成モードで予定の秘密度を表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="db6a5-120">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="db6a5-121">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="db6a5-121">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="db6a5-122">Office...-Alldayevent</span><span class="sxs-lookup"><span data-stu-id="db6a5-122">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="db6a5-123">予定が終日イベントであるかどうかを表す新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="db6a5-123">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="db6a5-124">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="db6a5-124">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="db6a5-125">Office. メールボックスの秘密度</span><span class="sxs-lookup"><span data-stu-id="db6a5-125">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="db6a5-126">予定の秘密度を表す新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="db6a5-126">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="db6a5-127">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="db6a5-127">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="db6a5-128">MailboxEnums AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="db6a5-128">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

<span data-ttu-id="db6a5-129">`AppointmentSensitivityType`予定で利用可能な秘密度オプションを表す新しい列挙を追加しました。</span><span class="sxs-lookup"><span data-stu-id="db6a5-129">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="db6a5-130">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="db6a5-130">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="db6a5-131">送信時に追加</span><span class="sxs-lookup"><span data-stu-id="db6a5-131">Append on send</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="db6a5-132">Office.......。</span><span class="sxs-lookup"><span data-stu-id="db6a5-132">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="db6a5-133">新規 `Body` 作成モードで、アイテムの本文の最後にデータを追加する新しい関数をオブジェクトに追加しました。</span><span class="sxs-lookup"><span data-stu-id="db6a5-133">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="db6a5-134">**利用可能な**対象: Outlook on Windows (Office 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="db6a5-134">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="db6a5-135">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="db6a5-135">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="db6a5-136">拡張された `AppendOnSend` アクセス許可のコレクションに拡張アクセス許可が含まれている必要があるマニフェストに、新しい要素を追加しました。</span><span class="sxs-lookup"><span data-stu-id="db6a5-136">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="db6a5-137">**利用可能な**対象: Outlook on Windows (Office 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="db6a5-137">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="db6a5-138">イベントベースのライセンス認証</span><span class="sxs-lookup"><span data-stu-id="db6a5-138">Event-based activation</span></span>

<span data-ttu-id="db6a5-139">Outlook アドインでのイベントベースのアクティブ化機能のサポートが追加されました。詳細については[、「イベントベースのライセンス認証用の Outlook アドインを構成](../../../outlook/autolaunch.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="db6a5-139">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="db6a5-140">LaunchEvent 拡張点</span><span class="sxs-lookup"><span data-stu-id="db6a5-140">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="db6a5-141">`LaunchEvent`マニフェストに拡張点サポートが追加されました。</span><span class="sxs-lookup"><span data-stu-id="db6a5-141">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="db6a5-142">イベントベースのライセンス認証機能を構成します。</span><span class="sxs-lookup"><span data-stu-id="db6a5-142">It configures event-based activation functionality.</span></span>

<span data-ttu-id="db6a5-143">**利用可能な**機能: web 上の Outlook (モダン、[要求のプレビューアクセス](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="db6a5-143">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="db6a5-144">LaunchEvents マニフェスト要素</span><span class="sxs-lookup"><span data-stu-id="db6a5-144">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="db6a5-145">`LaunchEvents`マニフェストに要素を追加しました。</span><span class="sxs-lookup"><span data-stu-id="db6a5-145">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="db6a5-146">イベントベースのアクティブ化機能の構成をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="db6a5-146">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="db6a5-147">**利用可能な**機能: web 上の Outlook (モダン、[要求のプレビューアクセス](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="db6a5-147">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="db6a5-148">ランタイムマニフェスト要素</span><span class="sxs-lookup"><span data-stu-id="db6a5-148">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="db6a5-149">マニフェスト要素に Outlook サポートが追加されました `Runtimes` 。</span><span class="sxs-lookup"><span data-stu-id="db6a5-149">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="db6a5-150">イベントベースのアクティブ化機能に必要な HTML ファイルと JavaScript ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="db6a5-150">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="db6a5-151">**利用可能な**機能: web 上の Outlook (モダン、[要求のプレビューアクセス](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="db6a5-151">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="db6a5-152">操作可能なメッセージとの統合</span><span class="sxs-lookup"><span data-stu-id="db6a5-152">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="db6a5-153">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="db6a5-153">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="db6a5-154">アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="db6a5-154">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="db6a5-155">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="db6a5-155">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="db6a5-156">メールの署名</span><span class="sxs-lookup"><span data-stu-id="db6a5-156">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="db6a5-157">SetSignatureAsync を示しています。</span><span class="sxs-lookup"><span data-stu-id="db6a5-157">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

<span data-ttu-id="db6a5-158">新規 `Body` 作成モードで、アイテムの本文の署名を追加または置換する新しい関数をオブジェクトに追加しました。</span><span class="sxs-lookup"><span data-stu-id="db6a5-158">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="db6a5-159">**利用可能な**対象: Outlook on Windows (Office 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="db6a5-159">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="db6a5-160">DisableClientSignatureAsync を示します。</span><span class="sxs-lookup"><span data-stu-id="db6a5-160">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="db6a5-161">新規作成モードの送信メールボックスのクライアント署名を無効にする新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="db6a5-161">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="db6a5-162">**利用可能な**対象: Outlook on Windows (Office 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="db6a5-162">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="db6a5-163">GetComposeTypeAsync を示します。</span><span class="sxs-lookup"><span data-stu-id="db6a5-163">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

<span data-ttu-id="db6a5-164">新規作成モードで、メッセージの作成の種類を取得する新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="db6a5-164">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="db6a5-165">**利用可能な**対象: Outlook on Windows (Office 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="db6a5-165">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="db6a5-166">。アイテム. isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="db6a5-166">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="db6a5-167">新規作成モードのアイテムでクライアント署名が有効になっているかどうかを確認する新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="db6a5-167">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="db6a5-168">**利用可能な**対象: Outlook on Windows (Office 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="db6a5-168">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="db6a5-169">MailboxEnums Setype</span><span class="sxs-lookup"><span data-stu-id="db6a5-169">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

<span data-ttu-id="db6a5-170">新規 `ComposeType` 作成モードで使用可能な新しい列挙を追加しました。</span><span class="sxs-lookup"><span data-stu-id="db6a5-170">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="db6a5-171">**利用可能な**対象: Outlook on Windows (Office 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="db6a5-171">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="db6a5-172">Office テーマ</span><span class="sxs-lookup"><span data-stu-id="db6a5-172">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="db6a5-173">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="db6a5-173">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="db6a5-174">Office テーマを取得する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="db6a5-174">Added ability to get Office theme.</span></span>

<span data-ttu-id="db6a5-175">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="db6a5-175">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="db6a5-176">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="db6a5-176">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="db6a5-177">`OfficeThemeChanged` イベントが `Mailbox` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="db6a5-177">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="db6a5-178">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="db6a5-178">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="single-sign-on-sso"></a><span data-ttu-id="db6a5-179">シングル サインオン (SSO)</span><span class="sxs-lookup"><span data-stu-id="db6a5-179">Single sign-on (SSO)</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="db6a5-180">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="db6a5-180">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="db6a5-181">Microsoft Graph API の[アクセス トークンの取得](../../../outlook/authenticate-a-user-with-an-sso-token.md)をアドインに対して許可する、`getAccessToken` へのアクセスが追加されました。</span><span class="sxs-lookup"><span data-stu-id="db6a5-181">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="db6a5-182">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac、Outlook on the web (モダン)、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="db6a5-182">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="db6a5-183">関連項目</span><span class="sxs-lookup"><span data-stu-id="db6a5-183">See also</span></span>

- [<span data-ttu-id="db6a5-184">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="db6a5-184">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="db6a5-185">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="db6a5-185">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="db6a5-186">概要</span><span class="sxs-lookup"><span data-stu-id="db6a5-186">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="db6a5-187">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="db6a5-187">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
