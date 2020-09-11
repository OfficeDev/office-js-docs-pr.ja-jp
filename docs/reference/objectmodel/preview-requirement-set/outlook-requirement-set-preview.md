---
title: Outlook アドイン API 要件セットのプレビュー
description: Outlook アドイン用に現在プレビューされている機能と Api。
ms.date: 09/02/2020
localization_priority: Normal
ms.openlocfilehash: 24cad394f0f3ffb95a05a81ccb38ee4aa72a3797
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431067"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="9bfbd-103">Outlook アドイン API 要件セットのプレビュー</span><span class="sxs-lookup"><span data-stu-id="9bfbd-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="9bfbd-104">Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9bfbd-105">このドキュメントは、[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)の**プレビュー**用です。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="9bfbd-106">この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="9bfbd-107">アドイン マニフェストでこの要件を指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="9bfbd-108">[Microsoft 365 テナントで対象指定リリースを構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)することで、web 上の Outlook の機能をプレビューできる場合があります。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="9bfbd-109">該当する機能については、このページにある「プレビューアクセスを構成する」を確認してください。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="9bfbd-110">その他の機能については、Microsoft 365 アカウントを使用して web 上の Outlook のプレビュービットへのアクセス権を要求することができます。 [このフォーム](https://aka.ms/OWAPreview)を完成して送信します。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="9bfbd-111">これらの機能については、「要求プレビューアクセス」を確認してください。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="9bfbd-112">要件セットのプレビューには、[要件セット 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) のすべての機能が含まれています。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-112">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="9bfbd-113">プレビューの機能</span><span class="sxs-lookup"><span data-stu-id="9bfbd-113">Features in preview</span></span>

<span data-ttu-id="9bfbd-114">次の機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="9bfbd-115">Information Rights Management (IRM) で保護されたアイテムでのアドインのアクティブ化</span><span class="sxs-lookup"><span data-stu-id="9bfbd-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="9bfbd-116">これで、IRM で保護されたアイテムでアドインをアクティブ化できるようになります。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="9bfbd-117">この機能を有効にするには、テナント管理者が `OBJMODEL` Office の [プログラムに **よるアクセスを許可** する] オプションを設定して使用権限を有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="9bfbd-118">詳細については [、「使用権限と説明](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="9bfbd-119">**利用可能**: Windows on Windows、build 13229.10000 (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="9bfbd-120">その他の予定表プロパティ</span><span class="sxs-lookup"><span data-stu-id="9bfbd-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="9bfbd-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="9bfbd-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="9bfbd-122">新規作成モードで予定の終日イベントプロパティを表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="9bfbd-123">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="9bfbd-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="9bfbd-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="9bfbd-125">新規作成モードで予定の秘密度を表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="9bfbd-126">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="9bfbd-127">Office...-Alldayevent</span><span class="sxs-lookup"><span data-stu-id="9bfbd-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="9bfbd-128">予定が終日イベントであるかどうかを表す新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="9bfbd-129">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="9bfbd-130">Office. メールボックスの秘密度</span><span class="sxs-lookup"><span data-stu-id="9bfbd-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="9bfbd-131">予定の秘密度を表す新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="9bfbd-132">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="9bfbd-133">MailboxEnums AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="9bfbd-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="9bfbd-134">`AppointmentSensitivityType`予定で利用可能な秘密度オプションを表す新しい列挙を追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="9bfbd-135">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="9bfbd-136">送信時に追加</span><span class="sxs-lookup"><span data-stu-id="9bfbd-136">Append on send</span></span>

<span data-ttu-id="9bfbd-137">追加-送信機能の使用方法については、「 [Outlook アドインで送信時に追加を実装](../../../outlook/append-on-send.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-137">To learn about using the append-on-send feature, see [Implement append on send in your Outlook add-in](../../../outlook/append-on-send.md).</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="9bfbd-138">Office.......。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-138">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-)

<span data-ttu-id="9bfbd-139">新規 `Body` 作成モードで、アイテムの本文の最後にデータを追加する新しい関数をオブジェクトに追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-139">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="9bfbd-140">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9bfbd-140">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="9bfbd-141">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="9bfbd-141">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="9bfbd-142">拡張された `AppendOnSend` アクセス許可のコレクションに拡張アクセス許可が含まれている必要があるマニフェストに、新しい要素を追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-142">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="9bfbd-143">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9bfbd-143">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="async-versions-of-display-apis"></a><span data-ttu-id="9bfbd-144">非同期バージョンの `display` api</span><span class="sxs-lookup"><span data-stu-id="9bfbd-144">Async versions of `display` APIs</span></span>

#### <a name="officecontextmailboxdisplayappointmentformasync"></a>[<span data-ttu-id="9bfbd-145">DisplayAppointmentFormAsync の内容</span><span class="sxs-lookup"><span data-stu-id="9bfbd-145">Office.context.mailbox.displayAppointmentFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displayappointmentformasync-itemid--options--callback-)

<span data-ttu-id="9bfbd-146">既存の予定を表示するオブジェクトに新しい関数を追加 `Mailbox` しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-146">Added a new function to the `Mailbox` object that displays an existing appointment.</span></span> <span data-ttu-id="9bfbd-147">これは、メソッドの非同期バージョンです `displayAppointmentForm` 。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-147">This is the async version of the `displayAppointmentForm` method.</span></span>

<span data-ttu-id="9bfbd-148">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-148">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxdisplaymessageformasync"></a>[<span data-ttu-id="9bfbd-149">Office. mailbox. displayMessageFormAsync</span><span class="sxs-lookup"><span data-stu-id="9bfbd-149">Office.context.mailbox.displayMessageFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaymessageformasync-itemid--options--callback-)

<span data-ttu-id="9bfbd-150">既存のメッセージを表示するオブジェクトに新しい関数を追加しまし `Mailbox` た。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-150">Added a new function to the `Mailbox` object that displays an existing message.</span></span> <span data-ttu-id="9bfbd-151">これは、メソッドの非同期バージョンです `displayMessageForm` 。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-151">This is the async version of the `displayMessageForm` method.</span></span>

<span data-ttu-id="9bfbd-152">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-152">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxdisplaynewappointmentformasync"></a>[<span data-ttu-id="9bfbd-153">DisplayNewAppointmentFormAsync の内容</span><span class="sxs-lookup"><span data-stu-id="9bfbd-153">Office.context.mailbox.displayNewAppointmentFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaynewappointmentformasync-parameters--options--callback-)

<span data-ttu-id="9bfbd-154">`Mailbox`新しい予定のフォームを表示する新しい関数をオブジェクトに追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-154">Added a new function to the `Mailbox` object that displays a new appointment form.</span></span> <span data-ttu-id="9bfbd-155">これは、メソッドの非同期バージョンです `displayNewAppointmentForm` 。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-155">This is the async version of the `displayNewAppointmentForm` method.</span></span>

<span data-ttu-id="9bfbd-156">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-156">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxdisplaynewmessageformasync"></a>[<span data-ttu-id="9bfbd-157">Office。 displayNewMessageFormAsync</span><span class="sxs-lookup"><span data-stu-id="9bfbd-157">Office.context.mailbox.displayNewMessageFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaynewmessageformasync-parameters--options--callback-)

<span data-ttu-id="9bfbd-158">`Mailbox`新しいメッセージフォームを表示する新しい関数をオブジェクトに追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-158">Added a new function to the `Mailbox` object that displays a new message form.</span></span> <span data-ttu-id="9bfbd-159">これは、メソッドの非同期バージョンです `displayNewMessageForm` 。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-159">This is the async version of the `displayNewMessageForm` method.</span></span>

<span data-ttu-id="9bfbd-160">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-160">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemdisplayreplyallformasync"></a>[<span data-ttu-id="9bfbd-161">DisplayReplyAllFormAsync を示します。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-161">Office.context.mailbox.item.displayReplyAllFormAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="9bfbd-162">`Item`読み取りモードで "全員に返信" フォームを表示するオブジェクトに新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-162">Added a new function to the `Item` object that displays the "Reply all" form in Read mode.</span></span> <span data-ttu-id="9bfbd-163">これは、メソッドの非同期バージョンです `displayReplyAllForm` 。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-163">This is the async version of the `displayReplyAllForm` method.</span></span>

<span data-ttu-id="9bfbd-164">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-164">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemdisplayreplyformasync"></a>[<span data-ttu-id="9bfbd-165">DisplayReplyFormAsync を示します。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-165">Office.context.mailbox.item.displayReplyFormAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="9bfbd-166">`Item`読み取りモードで "返信" フォームを表示するオブジェクトに新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-166">Added a new function to the `Item` object that displays the "Reply" form in Read mode.</span></span> <span data-ttu-id="9bfbd-167">これは、メソッドの非同期バージョンです `displayReplyForm` 。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-167">This is the async version of the `displayReplyForm` method.</span></span>

<span data-ttu-id="9bfbd-168">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-168">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="9bfbd-169">イベントベースのライセンス認証</span><span class="sxs-lookup"><span data-stu-id="9bfbd-169">Event-based activation</span></span>

<span data-ttu-id="9bfbd-170">Outlook アドインでのイベントベースのアクティブ化機能のサポートが追加されました。詳細については [、「イベントベースのライセンス認証用の Outlook アドインを構成](../../../outlook/autolaunch.md) する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-170">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="9bfbd-171">LaunchEvent 拡張点</span><span class="sxs-lookup"><span data-stu-id="9bfbd-171">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="9bfbd-172">`LaunchEvent`マニフェストに拡張点サポートが追加されました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-172">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="9bfbd-173">イベントベースのライセンス認証機能を構成します。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-173">It configures event-based activation functionality.</span></span>

<span data-ttu-id="9bfbd-174">**利用可能な**機能: web 上の Outlook (モダン、 [要求のプレビューアクセス](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="9bfbd-174">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="9bfbd-175">LaunchEvents マニフェスト要素</span><span class="sxs-lookup"><span data-stu-id="9bfbd-175">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="9bfbd-176">`LaunchEvents`マニフェストに要素を追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-176">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="9bfbd-177">イベントベースのアクティブ化機能の構成をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-177">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="9bfbd-178">**利用可能な**機能: web 上の Outlook (モダン、 [要求のプレビューアクセス](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="9bfbd-178">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="9bfbd-179">ランタイムマニフェスト要素</span><span class="sxs-lookup"><span data-stu-id="9bfbd-179">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="9bfbd-180">マニフェスト要素に Outlook サポートが追加されました `Runtimes` 。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-180">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="9bfbd-181">イベントベースのアクティブ化機能に必要な HTML ファイルと JavaScript ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-181">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="9bfbd-182">**利用可能な**機能: web 上の Outlook (モダン、 [要求のプレビューアクセス](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="9bfbd-182">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

<br>

---

---

### <a name="get-all-custom-properties"></a><span data-ttu-id="9bfbd-183">すべてのカスタムプロパティを取得する</span><span class="sxs-lookup"><span data-stu-id="9bfbd-183">Get all custom properties</span></span>

#### <a name="custompropertiesgetall"></a>[<span data-ttu-id="9bfbd-184">CustomProperties getAll</span><span class="sxs-lookup"><span data-stu-id="9bfbd-184">CustomProperties.getAll</span></span>](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true#getall--)

<span data-ttu-id="9bfbd-185">すべてのカスタムプロパティを取得する新しい関数をオブジェクトに追加しまし `CustomProperties` た。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-185">Added a new function to the `CustomProperties` object that gets all custom properties.</span></span>

<span data-ttu-id="9bfbd-186">**利用可能な**対象: Outlook on Windows (microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)、Outlook on Mac (microsoft 365 サブスクリプションに接続)、outlook on the Outlook on iOS</span><span class="sxs-lookup"><span data-stu-id="9bfbd-186">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to a Microsoft 365 subscription), Outlook on Android, Outlook on iOS</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="9bfbd-187">操作可能なメッセージとの統合</span><span class="sxs-lookup"><span data-stu-id="9bfbd-187">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="9bfbd-188">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="9bfbd-188">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="9bfbd-189">アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-189">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="9bfbd-190">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-190">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="9bfbd-191">メールの署名</span><span class="sxs-lookup"><span data-stu-id="9bfbd-191">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="9bfbd-192">SetSignatureAsync を示しています。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-192">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

<span data-ttu-id="9bfbd-193">新規 `Body` 作成モードで、アイテムの本文の署名を追加または置換する新しい関数をオブジェクトに追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-193">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="9bfbd-194">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9bfbd-194">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="9bfbd-195">DisableClientSignatureAsync を示します。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-195">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="9bfbd-196">新規作成モードの送信メールボックスのクライアント署名を無効にする新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-196">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="9bfbd-197">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9bfbd-197">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="9bfbd-198">GetComposeTypeAsync を示します。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-198">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

<span data-ttu-id="9bfbd-199">新規作成モードで、メッセージの作成の種類を取得する新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-199">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="9bfbd-200">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9bfbd-200">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="9bfbd-201">。アイテム. isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="9bfbd-201">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="9bfbd-202">新規作成モードのアイテムでクライアント署名が有効になっているかどうかを確認する新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-202">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="9bfbd-203">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9bfbd-203">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="9bfbd-204">MailboxEnums Setype</span><span class="sxs-lookup"><span data-stu-id="9bfbd-204">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="9bfbd-205">新規 `ComposeType` 作成モードで使用可能な新しい列挙を追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-205">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="9bfbd-206">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、 [構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9bfbd-206">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="notification-messages-with-actions"></a><span data-ttu-id="9bfbd-207">アクションを含む通知メッセージ</span><span class="sxs-lookup"><span data-stu-id="9bfbd-207">Notification messages with actions</span></span>

<span data-ttu-id="9bfbd-208">この機能を使用すると、既定の **アラーム** 処理に加えて、カスタムアクションを含む通知メッセージをアドインに含めることができます。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-208">This feature allows your add-in to include a notification message with a custom action besides the default **Dismiss** action.</span></span>

#### <a name="officenotificationmessagedetailsactions"></a>[<span data-ttu-id="9bfbd-209">Office の NotificationMessageDetails。アクション</span><span class="sxs-lookup"><span data-stu-id="9bfbd-209">Office.NotificationMessageDetails.actions</span></span>](/javascript/api/outlook/office.notificationmessagedetails#actions)

<span data-ttu-id="9bfbd-210">`InsightMessage`カスタムアクションを使用して通知を追加できるようにする新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-210">Added a new property that enables you to add an `InsightMessage` notification with a custom action.</span></span>

<span data-ttu-id="9bfbd-211">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-211">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officenotificationmessageaction"></a>[<span data-ttu-id="9bfbd-212">Office NotificationMessageAction</span><span class="sxs-lookup"><span data-stu-id="9bfbd-212">Office.NotificationMessageAction</span></span>](/javascript/api/outlook/office.notificationmessageaction)

<span data-ttu-id="9bfbd-213">通知のカスタムアクションを定義する新しいオブジェクトを追加しました `InsightMessage` 。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-213">Added a new object where you define a custom action for your `InsightMessage` notification.</span></span>

<span data-ttu-id="9bfbd-214">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-214">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsactiontype"></a>[<span data-ttu-id="9bfbd-215">MailboxEnums</span><span class="sxs-lookup"><span data-stu-id="9bfbd-215">Office.MailboxEnums.ActionType</span></span>](/javascript/api/outlook/office.mailboxenums.actiontype)

<span data-ttu-id="9bfbd-216">新しい列挙を追加 `ActionType` しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-216">Added a new enum `ActionType`.</span></span>

<span data-ttu-id="9bfbd-217">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-217">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[<span data-ttu-id="9bfbd-218">InsightMessage を MailboxEnums します。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-218">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span></span>](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

<span data-ttu-id="9bfbd-219">Enum に新しい型を追加しました `InsightMessage` `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-219">Added a new type `InsightMessage` to the `ItemNotificationMessageType` enum.</span></span>

<span data-ttu-id="9bfbd-220">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-220">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="9bfbd-221">Office テーマ</span><span class="sxs-lookup"><span data-stu-id="9bfbd-221">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="9bfbd-222">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="9bfbd-222">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="9bfbd-223">Office テーマを取得する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-223">Added ability to get Office theme.</span></span>

<span data-ttu-id="9bfbd-224">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-224">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="9bfbd-225">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="9bfbd-225">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="9bfbd-226">`OfficeThemeChanged` イベントが `Mailbox` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-226">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="9bfbd-227">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-227">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="9bfbd-228">セッション データ</span><span class="sxs-lookup"><span data-stu-id="9bfbd-228">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="9bfbd-229">Office セッションデータ</span><span class="sxs-lookup"><span data-stu-id="9bfbd-229">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="9bfbd-230">アイテムのセッションデータを表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-230">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="9bfbd-231">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-231">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="9bfbd-232">Office. メールボックス (セッション)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-232">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="9bfbd-233">新規作成モードのアイテムのセッションデータを管理するための新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-233">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="9bfbd-234">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-234">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="single-sign-on-sso"></a><span data-ttu-id="9bfbd-235">シングル サインオン (SSO)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-235">Single sign-on (SSO)</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="9bfbd-236">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="9bfbd-236">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="9bfbd-237">Microsoft Graph API の[アクセス トークンの取得](../../../outlook/authenticate-a-user-with-an-sso-token.md)をアドインに対して許可する、`getAccessToken` へのアクセスが追加されました。</span><span class="sxs-lookup"><span data-stu-id="9bfbd-237">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="9bfbd-238">**利用可能な**対象: Outlook on Windows (microsoft 365 サブスクリプションに接続)、Outlook on Mac (microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)、outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="9bfbd-238">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on Mac (connected to a Microsoft 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="9bfbd-239">関連項目</span><span class="sxs-lookup"><span data-stu-id="9bfbd-239">See also</span></span>

- [<span data-ttu-id="9bfbd-240">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="9bfbd-240">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="9bfbd-241">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="9bfbd-241">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="9bfbd-242">概要</span><span class="sxs-lookup"><span data-stu-id="9bfbd-242">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="9bfbd-243">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="9bfbd-243">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
