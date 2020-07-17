---
title: Outlook アドイン API 要件セットのプレビュー
description: Outlook アドイン用に現在プレビューされている機能と Api。
ms.date: 06/17/2020
localization_priority: Normal
ms.openlocfilehash: d91d1e16382a9ada71210657d6111f548c85ccfd
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094422"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="bab54-103">Outlook アドイン API 要件セットのプレビュー</span><span class="sxs-lookup"><span data-stu-id="bab54-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="bab54-104">Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="bab54-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bab54-105">このドキュメントは、[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)の**プレビュー**用です。</span><span class="sxs-lookup"><span data-stu-id="bab54-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="bab54-106">この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。</span><span class="sxs-lookup"><span data-stu-id="bab54-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="bab54-107">アドイン マニフェストでこの要件を指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="bab54-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="bab54-108">[Microsoft 365 テナントで対象指定リリースを構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)することで、web 上の Outlook の機能をプレビューできる場合があります。</span><span class="sxs-lookup"><span data-stu-id="bab54-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="bab54-109">該当する機能については、このページにある「プレビューアクセスを構成する」を確認してください。</span><span class="sxs-lookup"><span data-stu-id="bab54-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="bab54-110">その他の機能については、Microsoft 365 アカウントを使用して web 上の Outlook のプレビュービットへのアクセス権を要求することができます。[このフォーム](https://aka.ms/OWAPreview)を完成して送信します。</span><span class="sxs-lookup"><span data-stu-id="bab54-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="bab54-111">これらの機能については、「要求プレビューアクセス」を確認してください。</span><span class="sxs-lookup"><span data-stu-id="bab54-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="bab54-112">要件セットのプレビューには、[要件セット 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) のすべての機能が含まれています。</span><span class="sxs-lookup"><span data-stu-id="bab54-112">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="bab54-113">プレビューの機能</span><span class="sxs-lookup"><span data-stu-id="bab54-113">Features in preview</span></span>

<span data-ttu-id="bab54-114">次の機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="bab54-114">The following features are in preview.</span></span>

### <a name="additional-calendar-properties"></a><span data-ttu-id="bab54-115">その他の予定表プロパティ</span><span class="sxs-lookup"><span data-stu-id="bab54-115">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="bab54-116">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="bab54-116">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

<span data-ttu-id="bab54-117">新規作成モードで予定の終日イベントプロパティを表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="bab54-117">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="bab54-118">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="bab54-118">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="bab54-119">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="bab54-119">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

<span data-ttu-id="bab54-120">新規作成モードで予定の秘密度を表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="bab54-120">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="bab54-121">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="bab54-121">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="bab54-122">Office...-Alldayevent</span><span class="sxs-lookup"><span data-stu-id="bab54-122">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="bab54-123">予定が終日イベントであるかどうかを表す新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="bab54-123">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="bab54-124">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="bab54-124">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="bab54-125">Office. メールボックスの秘密度</span><span class="sxs-lookup"><span data-stu-id="bab54-125">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="bab54-126">予定の秘密度を表す新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="bab54-126">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="bab54-127">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="bab54-127">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="bab54-128">MailboxEnums AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="bab54-128">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

<span data-ttu-id="bab54-129">`AppointmentSensitivityType`予定で利用可能な秘密度オプションを表す新しい列挙を追加しました。</span><span class="sxs-lookup"><span data-stu-id="bab54-129">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="bab54-130">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="bab54-130">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="bab54-131">送信時に追加</span><span class="sxs-lookup"><span data-stu-id="bab54-131">Append on send</span></span>

<span data-ttu-id="bab54-132">追加-送信機能の使用方法については、「 [Outlook アドインで送信時に追加を実装](../../../outlook/append-on-send.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bab54-132">To learn about using the append-on-send feature, see [Implement append on send in your Outlook add-in](../../../outlook/append-on-send.md).</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="bab54-133">Office.......。</span><span class="sxs-lookup"><span data-stu-id="bab54-133">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="bab54-134">新規 `Body` 作成モードで、アイテムの本文の最後にデータを追加する新しい関数をオブジェクトに追加しました。</span><span class="sxs-lookup"><span data-stu-id="bab54-134">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="bab54-135">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bab54-135">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="bab54-136">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="bab54-136">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="bab54-137">拡張された `AppendOnSend` アクセス許可のコレクションに拡張アクセス許可が含まれている必要があるマニフェストに、新しい要素を追加しました。</span><span class="sxs-lookup"><span data-stu-id="bab54-137">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="bab54-138">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bab54-138">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="async-versions-of-display-apis"></a><span data-ttu-id="bab54-139">非同期バージョンの `display` api</span><span class="sxs-lookup"><span data-stu-id="bab54-139">Async versions of `display` APIs</span></span>

#### <a name="officecontextmailboxdisplayappointmentformasync"></a>[<span data-ttu-id="bab54-140">DisplayAppointmentFormAsync の内容</span><span class="sxs-lookup"><span data-stu-id="bab54-140">Office.context.mailbox.displayAppointmentFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displayappointmentformasync-itemid--options--callback-)

<span data-ttu-id="bab54-141">既存の予定を表示するオブジェクトに新しい関数を追加 `Mailbox` しました。</span><span class="sxs-lookup"><span data-stu-id="bab54-141">Added a new function to the `Mailbox` object that displays an existing appointment.</span></span> <span data-ttu-id="bab54-142">これは、メソッドの非同期バージョンです `displayAppointmentForm` 。</span><span class="sxs-lookup"><span data-stu-id="bab54-142">This is the async version of the `displayAppointmentForm` method.</span></span>

<span data-ttu-id="bab54-143">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="bab54-143">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxdisplaymessageformasync"></a>[<span data-ttu-id="bab54-144">Office. mailbox. displayMessageFormAsync</span><span class="sxs-lookup"><span data-stu-id="bab54-144">Office.context.mailbox.displayMessageFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaymessageformasync-itemid--options--callback-)

<span data-ttu-id="bab54-145">既存のメッセージを表示するオブジェクトに新しい関数を追加しまし `Mailbox` た。</span><span class="sxs-lookup"><span data-stu-id="bab54-145">Added a new function to the `Mailbox` object that displays an existing message.</span></span> <span data-ttu-id="bab54-146">これは、メソッドの非同期バージョンです `displayMessageForm` 。</span><span class="sxs-lookup"><span data-stu-id="bab54-146">This is the async version of the `displayMessageForm` method.</span></span>

<span data-ttu-id="bab54-147">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="bab54-147">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxdisplaynewappointmentformasync"></a>[<span data-ttu-id="bab54-148">DisplayNewAppointmentFormAsync の内容</span><span class="sxs-lookup"><span data-stu-id="bab54-148">Office.context.mailbox.displayNewAppointmentFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewappointmentformasync-parameters--options--callback-)

<span data-ttu-id="bab54-149">`Mailbox`新しい予定のフォームを表示する新しい関数をオブジェクトに追加しました。</span><span class="sxs-lookup"><span data-stu-id="bab54-149">Added a new function to the `Mailbox` object that displays a new appointment form.</span></span> <span data-ttu-id="bab54-150">これは、メソッドの非同期バージョンです `displayNewAppointmentForm` 。</span><span class="sxs-lookup"><span data-stu-id="bab54-150">This is the async version of the `displayNewAppointmentForm` method.</span></span>

<span data-ttu-id="bab54-151">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="bab54-151">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxdisplaynewmessageformasync"></a>[<span data-ttu-id="bab54-152">Office。 displayNewMessageFormAsync</span><span class="sxs-lookup"><span data-stu-id="bab54-152">Office.context.mailbox.displayNewMessageFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewmessageformasync-parameters--options--callback-)

<span data-ttu-id="bab54-153">`Mailbox`新しいメッセージフォームを表示する新しい関数をオブジェクトに追加しました。</span><span class="sxs-lookup"><span data-stu-id="bab54-153">Added a new function to the `Mailbox` object that displays a new message form.</span></span> <span data-ttu-id="bab54-154">これは、メソッドの非同期バージョンです `displayNewMessageForm` 。</span><span class="sxs-lookup"><span data-stu-id="bab54-154">This is the async version of the `displayNewMessageForm` method.</span></span>

<span data-ttu-id="bab54-155">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="bab54-155">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemdisplayreplyallformasync"></a>[<span data-ttu-id="bab54-156">DisplayReplyAllFormAsync を示します。</span><span class="sxs-lookup"><span data-stu-id="bab54-156">Office.context.mailbox.item.displayReplyAllFormAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="bab54-157">`Item`読み取りモードで "全員に返信" フォームを表示するオブジェクトに新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="bab54-157">Added a new function to the `Item` object that displays the "Reply all" form in Read mode.</span></span> <span data-ttu-id="bab54-158">これは、メソッドの非同期バージョンです `displayReplyAllForm` 。</span><span class="sxs-lookup"><span data-stu-id="bab54-158">This is the async version of the `displayReplyAllForm` method.</span></span>

<span data-ttu-id="bab54-159">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="bab54-159">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemdisplayreplyformasync"></a>[<span data-ttu-id="bab54-160">DisplayReplyFormAsync を示します。</span><span class="sxs-lookup"><span data-stu-id="bab54-160">Office.context.mailbox.item.displayReplyFormAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="bab54-161">`Item`読み取りモードで "返信" フォームを表示するオブジェクトに新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="bab54-161">Added a new function to the `Item` object that displays the "Reply" form in Read mode.</span></span> <span data-ttu-id="bab54-162">これは、メソッドの非同期バージョンです `displayReplyForm` 。</span><span class="sxs-lookup"><span data-stu-id="bab54-162">This is the async version of the `displayReplyForm` method.</span></span>

<span data-ttu-id="bab54-163">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="bab54-163">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="bab54-164">イベントベースのライセンス認証</span><span class="sxs-lookup"><span data-stu-id="bab54-164">Event-based activation</span></span>

<span data-ttu-id="bab54-165">Outlook アドインでのイベントベースのアクティブ化機能のサポートが追加されました。詳細については[、「イベントベースのライセンス認証用の Outlook アドインを構成](../../../outlook/autolaunch.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bab54-165">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="bab54-166">LaunchEvent 拡張点</span><span class="sxs-lookup"><span data-stu-id="bab54-166">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="bab54-167">`LaunchEvent`マニフェストに拡張点サポートが追加されました。</span><span class="sxs-lookup"><span data-stu-id="bab54-167">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="bab54-168">イベントベースのライセンス認証機能を構成します。</span><span class="sxs-lookup"><span data-stu-id="bab54-168">It configures event-based activation functionality.</span></span>

<span data-ttu-id="bab54-169">**利用可能な**機能: web 上の Outlook (モダン、[要求のプレビューアクセス](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="bab54-169">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="bab54-170">LaunchEvents マニフェスト要素</span><span class="sxs-lookup"><span data-stu-id="bab54-170">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="bab54-171">`LaunchEvents`マニフェストに要素を追加しました。</span><span class="sxs-lookup"><span data-stu-id="bab54-171">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="bab54-172">イベントベースのアクティブ化機能の構成をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="bab54-172">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="bab54-173">**利用可能な**機能: web 上の Outlook (モダン、[要求のプレビューアクセス](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="bab54-173">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="bab54-174">ランタイムマニフェスト要素</span><span class="sxs-lookup"><span data-stu-id="bab54-174">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="bab54-175">マニフェスト要素に Outlook サポートが追加されました `Runtimes` 。</span><span class="sxs-lookup"><span data-stu-id="bab54-175">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="bab54-176">イベントベースのアクティブ化機能に必要な HTML ファイルと JavaScript ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="bab54-176">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="bab54-177">**利用可能な**機能: web 上の Outlook (モダン、[要求のプレビューアクセス](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="bab54-177">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

<br>

---

---

### <a name="get-all-custom-properties"></a><span data-ttu-id="bab54-178">すべてのカスタムプロパティを取得する</span><span class="sxs-lookup"><span data-stu-id="bab54-178">Get all custom properties</span></span>

#### <a name="custompropertiesgetall"></a>[<span data-ttu-id="bab54-179">CustomProperties getAll</span><span class="sxs-lookup"><span data-stu-id="bab54-179">CustomProperties.getAll</span></span>](/javascript/api/outlook/office.customproperties?view=outlook-js-preview#getall--)

<span data-ttu-id="bab54-180">すべてのカスタムプロパティを取得する新しい関数をオブジェクトに追加しまし `CustomProperties` た。</span><span class="sxs-lookup"><span data-stu-id="bab54-180">Added a new function to the `CustomProperties` object that gets all custom properties.</span></span>

<span data-ttu-id="bab54-181">**利用可能な**対象: Outlook on Windows (microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)、Outlook on Mac (microsoft 365 サブスクリプションに接続)、outlook on the Outlook on IOS、Outlook on iOS</span><span class="sxs-lookup"><span data-stu-id="bab54-181">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Microsoft 365 subscription), Outlook on Android, Outlook on iOS</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="bab54-182">操作可能なメッセージとの統合</span><span class="sxs-lookup"><span data-stu-id="bab54-182">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="bab54-183">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="bab54-183">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="bab54-184">アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="bab54-184">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="bab54-185">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="bab54-185">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="bab54-186">メールの署名</span><span class="sxs-lookup"><span data-stu-id="bab54-186">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="bab54-187">SetSignatureAsync を示しています。</span><span class="sxs-lookup"><span data-stu-id="bab54-187">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

<span data-ttu-id="bab54-188">新規 `Body` 作成モードで、アイテムの本文の署名を追加または置換する新しい関数をオブジェクトに追加しました。</span><span class="sxs-lookup"><span data-stu-id="bab54-188">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="bab54-189">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bab54-189">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="bab54-190">DisableClientSignatureAsync を示します。</span><span class="sxs-lookup"><span data-stu-id="bab54-190">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="bab54-191">新規作成モードの送信メールボックスのクライアント署名を無効にする新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="bab54-191">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="bab54-192">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bab54-192">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="bab54-193">GetComposeTypeAsync を示します。</span><span class="sxs-lookup"><span data-stu-id="bab54-193">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

<span data-ttu-id="bab54-194">新規作成モードで、メッセージの作成の種類を取得する新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="bab54-194">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="bab54-195">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bab54-195">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="bab54-196">。アイテム. isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="bab54-196">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="bab54-197">新規作成モードのアイテムでクライアント署名が有効になっているかどうかを確認する新しい関数を追加しました。</span><span class="sxs-lookup"><span data-stu-id="bab54-197">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="bab54-198">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bab54-198">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="bab54-199">MailboxEnums Setype</span><span class="sxs-lookup"><span data-stu-id="bab54-199">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

<span data-ttu-id="bab54-200">新規 `ComposeType` 作成モードで使用可能な新しい列挙を追加しました。</span><span class="sxs-lookup"><span data-stu-id="bab54-200">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="bab54-201">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)、web 上の outlook (モダン、[構成プレビューアクセス](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bab54-201">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="bab54-202">Office テーマ</span><span class="sxs-lookup"><span data-stu-id="bab54-202">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="bab54-203">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="bab54-203">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="bab54-204">Office テーマを取得する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="bab54-204">Added ability to get Office theme.</span></span>

<span data-ttu-id="bab54-205">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="bab54-205">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="bab54-206">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="bab54-206">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="bab54-207">`OfficeThemeChanged` イベントが `Mailbox` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="bab54-207">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="bab54-208">**利用可能な**対象: Outlook on Windows (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="bab54-208">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="single-sign-on-sso"></a><span data-ttu-id="bab54-209">シングル サインオン (SSO)</span><span class="sxs-lookup"><span data-stu-id="bab54-209">Single sign-on (SSO)</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="bab54-210">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="bab54-210">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="bab54-211">Microsoft Graph API の[アクセス トークンの取得](../../../outlook/authenticate-a-user-with-an-sso-token.md)をアドインに対して許可する、`getAccessToken` へのアクセスが追加されました。</span><span class="sxs-lookup"><span data-stu-id="bab54-211">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="bab54-212">**利用可能な**対象: Outlook on Windows (microsoft 365 サブスクリプションに接続)、Outlook on Mac (microsoft 365 サブスクリプションに接続)、outlook on the web (モダン)、outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="bab54-212">**Available in**: Outlook on Windows (connected to Microsoft 365 subscription), Outlook on Mac (connected to Microsoft 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="bab54-213">関連項目</span><span class="sxs-lookup"><span data-stu-id="bab54-213">See also</span></span>

- [<span data-ttu-id="bab54-214">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="bab54-214">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="bab54-215">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="bab54-215">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="bab54-216">概要</span><span class="sxs-lookup"><span data-stu-id="bab54-216">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="bab54-217">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="bab54-217">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
