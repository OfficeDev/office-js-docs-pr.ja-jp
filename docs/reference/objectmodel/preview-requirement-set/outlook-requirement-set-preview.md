---
title: Outlook API プレビュー要件セット
description: 現在、アドインのプレビュー中Outlook API。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 98bf56c169967ad7c994d1793afa8678d31f6892
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591059"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="a359c-103">Outlook API プレビュー要件セット</span><span class="sxs-lookup"><span data-stu-id="a359c-103">Outlook add-in API preview requirement set</span></span>

<span data-ttu-id="a359c-104">Office Outlook JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="a359c-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a359c-105">このドキュメントは、[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)の **プレビュー** 用です。</span><span class="sxs-lookup"><span data-stu-id="a359c-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="a359c-106">この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。</span><span class="sxs-lookup"><span data-stu-id="a359c-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="a359c-107">アドイン マニフェストでこの要件を指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="a359c-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="a359c-108">Web 上の機能をプレビューするには、Outlookテナントでターゲット リリース[を構成Microsoft 365があります](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。</span><span class="sxs-lookup"><span data-stu-id="a359c-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="a359c-109">該当する機能については、このページに「プレビュー アクセスを構成する」と表示されます。</span><span class="sxs-lookup"><span data-stu-id="a359c-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="a359c-110">その他の機能については、このフォームを入力して送信することで、Microsoft 365 アカウントを使用して web 上の Outlook のプレビュー ビットへのアクセスを[要求できる場合があります](https://aka.ms/OWAPreview)。</span><span class="sxs-lookup"><span data-stu-id="a359c-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="a359c-111">"要求プレビュー アクセス" は、これらの機能に関して示されています。</span><span class="sxs-lookup"><span data-stu-id="a359c-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="a359c-112">プレビュー要件セットには、要件セット [1.10 のすべての機能が含まれています](../requirement-set-1.10/outlook-requirement-set-1.10.md)。</span><span class="sxs-lookup"><span data-stu-id="a359c-112">The preview requirement set includes all of the features of [requirement set 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="a359c-113">プレビューの機能</span><span class="sxs-lookup"><span data-stu-id="a359c-113">Features in preview</span></span>

<span data-ttu-id="a359c-114">次の機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="a359c-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="a359c-115">Information Rights Management (IRM) によって保護されたアイテムに対するアドインのアクティブ化</span><span class="sxs-lookup"><span data-stu-id="a359c-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="a359c-116">アドインは IRM で保護されたアイテムでアクティブ化できます。</span><span class="sxs-lookup"><span data-stu-id="a359c-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="a359c-117">この機能を有効にするには、テナント管理者が[プログラムによるアクセスを許可する] カスタム ポリシー オプションを設定して、使用権を有効にする `OBJMODEL` 必要Office。 </span><span class="sxs-lookup"><span data-stu-id="a359c-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="a359c-118">詳細 [については、「利用状況の権限と説明](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a359c-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="a359c-119">**で利用可能**: Outlook Windows ビルド 13229.10000 から (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="a359c-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="a359c-120">その他の予定表のプロパティ</span><span class="sxs-lookup"><span data-stu-id="a359c-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="a359c-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="a359c-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="a359c-122">作成モードで予定の全日イベント プロパティを表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="a359c-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="a359c-123">**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)</span><span class="sxs-lookup"><span data-stu-id="a359c-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="a359c-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="a359c-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="a359c-125">作成モードでの予定の感度を表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="a359c-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="a359c-126">**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)</span><span class="sxs-lookup"><span data-stu-id="a359c-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="a359c-127">Office.context.mailbox.item.isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="a359c-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="a359c-128">予定が一日のイベントである場合を表す新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="a359c-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="a359c-129">**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)</span><span class="sxs-lookup"><span data-stu-id="a359c-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="a359c-130">Office.context.mailbox.item.sensitivity</span><span class="sxs-lookup"><span data-stu-id="a359c-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="a359c-131">予定の感度を表す新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="a359c-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="a359c-132">**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)</span><span class="sxs-lookup"><span data-stu-id="a359c-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="a359c-133">Office。MailboxEnums.AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="a359c-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="a359c-134">予定で使用できる `AppointmentSensitivityType` 感度オプションを表す新しい列挙型を追加しました。</span><span class="sxs-lookup"><span data-stu-id="a359c-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="a359c-135">**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)</span><span class="sxs-lookup"><span data-stu-id="a359c-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="a359c-136">操作可能なメッセージとの統合</span><span class="sxs-lookup"><span data-stu-id="a359c-136">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="a359c-137">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="a359c-137">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="a359c-138">アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="a359c-138">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="a359c-139">**で利用** できる: Outlook (Windows サブスクリプションに接続されている) Microsoft 365、Outlook (モダン)</span><span class="sxs-lookup"><span data-stu-id="a359c-139">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="a359c-140">Office テーマ</span><span class="sxs-lookup"><span data-stu-id="a359c-140">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="a359c-141">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="a359c-141">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="a359c-142">Office テーマを取得する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="a359c-142">Added ability to get Office theme.</span></span>

<span data-ttu-id="a359c-143">**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)</span><span class="sxs-lookup"><span data-stu-id="a359c-143">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="a359c-144">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="a359c-144">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="a359c-145">`OfficeThemeChanged` イベントが `Mailbox` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="a359c-145">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="a359c-146">**で利用可能**: Outlook (WindowsサブスクリプションにMicrosoft 365)</span><span class="sxs-lookup"><span data-stu-id="a359c-146">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="a359c-147">セッション データ</span><span class="sxs-lookup"><span data-stu-id="a359c-147">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="a359c-148">Office。SessionData</span><span class="sxs-lookup"><span data-stu-id="a359c-148">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="a359c-149">アイテムのセッション データを表す新しいオブジェクトを追加しました。</span><span class="sxs-lookup"><span data-stu-id="a359c-149">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="a359c-150">**で利用** できる: Outlook (Windows サブスクリプションに接続されている) Microsoft 365、Outlook (モダン)</span><span class="sxs-lookup"><span data-stu-id="a359c-150">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="a359c-151">Office.context.mailbox.item.sessionData</span><span class="sxs-lookup"><span data-stu-id="a359c-151">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="a359c-152">新規作成モードでアイテムのセッション データを管理するための新しいプロパティを追加しました。</span><span class="sxs-lookup"><span data-stu-id="a359c-152">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="a359c-153">**で利用** できる: Outlook (Windows サブスクリプションに接続されている) Microsoft 365、Outlook (モダン)</span><span class="sxs-lookup"><span data-stu-id="a359c-153">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

## <a name="see-also"></a><span data-ttu-id="a359c-154">関連項目</span><span class="sxs-lookup"><span data-stu-id="a359c-154">See also</span></span>

- [<span data-ttu-id="a359c-155">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="a359c-155">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="a359c-156">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="a359c-156">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="a359c-157">概要</span><span class="sxs-lookup"><span data-stu-id="a359c-157">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="a359c-158">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="a359c-158">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
