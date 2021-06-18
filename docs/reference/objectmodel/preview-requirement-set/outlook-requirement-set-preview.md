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
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="727cf-103">Outlook API プレビュー要件セット</span><span class="sxs-lookup"><span data-stu-id="727cf-103">Outlook add-in API preview requirement set</span></span>

<span data-ttu-id="727cf-104">Office Outlook JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="727cf-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="727cf-105">このドキュメントは、[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)の **プレビュー** 用です。</span><span class="sxs-lookup"><span data-stu-id="727cf-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="727cf-106">この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。</span><span class="sxs-lookup"><span data-stu-id="727cf-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="727cf-107">アドイン マニフェストでこの要件を指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="727cf-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="727cf-108">ターゲット リリースを構成することで、Outlook on the webの機能をプレビューできる場合[Microsoft 365があります](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。</span><span class="sxs-lookup"><span data-stu-id="727cf-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="727cf-109">該当する機能については、このページに「プレビュー アクセスを構成する」と表示されます。</span><span class="sxs-lookup"><span data-stu-id="727cf-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="727cf-110">その他の機能については、このフォームを入力して送信することで、Outlook on the web アカウントを使用Microsoft 365プレビュー ビットへのアクセスを[要求できる場合があります](https://aka.ms/OWAPreview)。</span><span class="sxs-lookup"><span data-stu-id="727cf-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="727cf-111">"要求プレビュー アクセス" は、これらの機能に関して示されています。</span><span class="sxs-lookup"><span data-stu-id="727cf-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="727cf-112">プレビュー要件セットには、要件セット [1.10 のすべての機能が含まれています](../requirement-set-1.10/outlook-requirement-set-1.10.md)。</span><span class="sxs-lookup"><span data-stu-id="727cf-112">The preview requirement set includes all of the features of [requirement set 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="727cf-113">プレビューの機能</span><span class="sxs-lookup"><span data-stu-id="727cf-113">Features in preview</span></span>

<span data-ttu-id="727cf-114">次の機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="727cf-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="727cf-115">Information Rights Management (IRM) によって保護されたアイテムに対するアドインのアクティブ化</span><span class="sxs-lookup"><span data-stu-id="727cf-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="727cf-116">アドインは IRM で保護されたアイテムでアクティブ化できます。</span><span class="sxs-lookup"><span data-stu-id="727cf-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="727cf-117">この機能を有効にするには、テナント管理者が[プログラムによるアクセスを許可する] カスタム ポリシー オプションを設定して、使用権を有効にする `OBJMODEL` 必要Office。 </span><span class="sxs-lookup"><span data-stu-id="727cf-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="727cf-118">詳細 [については、「利用状況の権限と説明](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="727cf-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="727cf-119">**で利用可能**: Outlook Windows ビルド 13229.10000 から (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="727cf-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="727cf-120">その他の予定表のプロパティ</span><span class="sxs-lookup"><span data-stu-id="727cf-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="727cf-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="727cf-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="727cf-122">作成モードで予定の全日イベント プロパティを表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="727cf-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="727cf-123">**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)</span><span class="sxs-lookup"><span data-stu-id="727cf-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="727cf-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="727cf-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="727cf-125">作成モードでの予定の感度を表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="727cf-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="727cf-126">**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)</span><span class="sxs-lookup"><span data-stu-id="727cf-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="727cf-127">Office.context.mailbox.item.isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="727cf-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="727cf-128">予定が一日のイベントである場合を表す新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="727cf-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="727cf-129">**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)</span><span class="sxs-lookup"><span data-stu-id="727cf-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="727cf-130">Office.context.mailbox.item.sensitivity</span><span class="sxs-lookup"><span data-stu-id="727cf-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="727cf-131">予定の感度を表す新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="727cf-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="727cf-132">**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)</span><span class="sxs-lookup"><span data-stu-id="727cf-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="727cf-133">Office。MailboxEnums.AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="727cf-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="727cf-134">予定で使用できる `AppointmentSensitivityType` 感度オプションを表す新しい列挙型を追加しました。</span><span class="sxs-lookup"><span data-stu-id="727cf-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="727cf-135">**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)</span><span class="sxs-lookup"><span data-stu-id="727cf-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="727cf-136">イベントベースのライセンス認証</span><span class="sxs-lookup"><span data-stu-id="727cf-136">Event-based activation</span></span>

<span data-ttu-id="727cf-137">この機能は、要件セット [1.10 でリリースされました](../requirement-set-1.10/outlook-requirement-set-1.10.md)。</span><span class="sxs-lookup"><span data-stu-id="727cf-137">This feature was released in [requirement set 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span> <span data-ttu-id="727cf-138">ただし、追加のイベントはプレビューで利用できます。</span><span class="sxs-lookup"><span data-stu-id="727cf-138">However, additional events are now available in preview.</span></span> <span data-ttu-id="727cf-139">詳細については、「サポートされているイベント [」を参照してください](../../../outlook/autolaunch.md#supported-events)。</span><span class="sxs-lookup"><span data-stu-id="727cf-139">To learn more, refer to [Supported events](../../../outlook/autolaunch.md#supported-events).</span></span>

<span data-ttu-id="727cf-140">**で利用可能**: Outlook (Windowsサブスクリプションに接続) 、Microsoft 365 (Outlook on the web)</span><span class="sxs-lookup"><span data-stu-id="727cf-140">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="727cf-141">操作可能なメッセージとの統合</span><span class="sxs-lookup"><span data-stu-id="727cf-141">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="727cf-142">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="727cf-142">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="727cf-143">アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="727cf-143">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="727cf-144">**で利用可能**: Outlook (Windowsサブスクリプションに接続) 、Microsoft 365 (Outlook on the web)</span><span class="sxs-lookup"><span data-stu-id="727cf-144">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="727cf-145">Office テーマ</span><span class="sxs-lookup"><span data-stu-id="727cf-145">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="727cf-146">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="727cf-146">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="727cf-147">Office テーマを取得する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="727cf-147">Added ability to get Office theme.</span></span>

<span data-ttu-id="727cf-148">**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)</span><span class="sxs-lookup"><span data-stu-id="727cf-148">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="727cf-149">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="727cf-149">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="727cf-150">`OfficeThemeChanged` イベントが `Mailbox` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="727cf-150">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="727cf-151">**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)</span><span class="sxs-lookup"><span data-stu-id="727cf-151">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="727cf-152">セッション データ</span><span class="sxs-lookup"><span data-stu-id="727cf-152">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="727cf-153">Office。SessionData</span><span class="sxs-lookup"><span data-stu-id="727cf-153">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="727cf-154">アイテムのセッション データを表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="727cf-154">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="727cf-155">**で利用可能**: Outlook (Windowsサブスクリプションに接続) 、Microsoft 365 (Outlook on the web)</span><span class="sxs-lookup"><span data-stu-id="727cf-155">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="727cf-156">Office.context.mailbox.item.sessionData</span><span class="sxs-lookup"><span data-stu-id="727cf-156">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="727cf-157">新規作成モードでアイテムのセッション データを管理するための新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="727cf-157">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="727cf-158">**で利用可能**: Outlook (Windowsサブスクリプションに接続) 、Microsoft 365 (Outlook on the web)</span><span class="sxs-lookup"><span data-stu-id="727cf-158">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="shared-mailboxes"></a><span data-ttu-id="727cf-159">共有メールボックス</span><span class="sxs-lookup"><span data-stu-id="727cf-159">Shared mailboxes</span></span>

<span data-ttu-id="727cf-160">要件セット [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md)で、共有フォルダー (代理人アクセス) の機能サポートがリリースされました。</span><span class="sxs-lookup"><span data-stu-id="727cf-160">Feature support for shared folders (that is, delegate access) was released in [requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span> <span data-ttu-id="727cf-161">ただし、共有メールボックスのサポートはプレビューで利用できます。</span><span class="sxs-lookup"><span data-stu-id="727cf-161">However, support for shared mailboxes is now available in preview.</span></span> <span data-ttu-id="727cf-162">詳細については、「共有フォルダーと共有 [メールボックスのシナリオを有効にする」を参照してください](../../../outlook/delegate-access.md)。</span><span class="sxs-lookup"><span data-stu-id="727cf-162">To learn more, refer to [Enable shared folders and shared mailbox scenarios](../../../outlook/delegate-access.md).</span></span>

<span data-ttu-id="727cf-163">**で利用可能**: Outlook (Windowsサブスクリプションに接続) 、Microsoft 365 (Outlook on the web)</span><span class="sxs-lookup"><span data-stu-id="727cf-163">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

## <a name="see-also"></a><span data-ttu-id="727cf-164">関連項目</span><span class="sxs-lookup"><span data-stu-id="727cf-164">See also</span></span>

- [<span data-ttu-id="727cf-165">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="727cf-165">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="727cf-166">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="727cf-166">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="727cf-167">概要</span><span class="sxs-lookup"><span data-stu-id="727cf-167">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="727cf-168">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="727cf-168">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
