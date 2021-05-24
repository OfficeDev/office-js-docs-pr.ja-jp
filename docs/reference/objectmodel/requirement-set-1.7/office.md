---
title: Office名前空間 - 要件セット 1.7
description: Office API 要件セット 1.7 をOutlookアドインで使用できる名前空間メンバーを指定します。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 19c80c0c8c4aaf31c42aad16b3f474e92b7cdaec
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590975"
---
# <a name="office-mailbox-requirement-set-17"></a><span data-ttu-id="3ee08-103">Office (メールボックス要件セット 1.7)</span><span class="sxs-lookup"><span data-stu-id="3ee08-103">Office (Mailbox requirement set 1.7)</span></span>

<span data-ttu-id="3ee08-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3ee08-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee08-106">要件</span><span class="sxs-lookup"><span data-stu-id="3ee08-106">Requirements</span></span>

|<span data-ttu-id="3ee08-107">要件</span><span class="sxs-lookup"><span data-stu-id="3ee08-107">Requirement</span></span>| <span data-ttu-id="3ee08-108">値</span><span class="sxs-lookup"><span data-stu-id="3ee08-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee08-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3ee08-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3ee08-110">1.1</span><span class="sxs-lookup"><span data-stu-id="3ee08-110">1.1</span></span>|
|[<span data-ttu-id="3ee08-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3ee08-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3ee08-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3ee08-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="3ee08-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="3ee08-113">Properties</span></span>

| <span data-ttu-id="3ee08-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="3ee08-114">Property</span></span> | <span data-ttu-id="3ee08-115">モード</span><span class="sxs-lookup"><span data-stu-id="3ee08-115">Modes</span></span> | <span data-ttu-id="3ee08-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="3ee08-116">Return type</span></span> | <span data-ttu-id="3ee08-117">最小値</span><span class="sxs-lookup"><span data-stu-id="3ee08-117">Minimum</span></span><br><span data-ttu-id="3ee08-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="3ee08-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3ee08-119">context</span><span class="sxs-lookup"><span data-stu-id="3ee08-119">context</span></span>](office.context.md) | <span data-ttu-id="3ee08-120">作成</span><span class="sxs-lookup"><span data-stu-id="3ee08-120">Compose</span></span><br><span data-ttu-id="3ee08-121">Read</span><span class="sxs-lookup"><span data-stu-id="3ee08-121">Read</span></span> | [<span data-ttu-id="3ee08-122">Context</span><span class="sxs-lookup"><span data-stu-id="3ee08-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="3ee08-123">1.1</span><span class="sxs-lookup"><span data-stu-id="3ee08-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="3ee08-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="3ee08-124">Enumerations</span></span>

| <span data-ttu-id="3ee08-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="3ee08-125">Enumeration</span></span> | <span data-ttu-id="3ee08-126">モード</span><span class="sxs-lookup"><span data-stu-id="3ee08-126">Modes</span></span> | <span data-ttu-id="3ee08-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="3ee08-127">Return type</span></span> | <span data-ttu-id="3ee08-128">最小値</span><span class="sxs-lookup"><span data-stu-id="3ee08-128">Minimum</span></span><br><span data-ttu-id="3ee08-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="3ee08-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3ee08-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="3ee08-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="3ee08-131">作成</span><span class="sxs-lookup"><span data-stu-id="3ee08-131">Compose</span></span><br><span data-ttu-id="3ee08-132">Read</span><span class="sxs-lookup"><span data-stu-id="3ee08-132">Read</span></span> | <span data-ttu-id="3ee08-133">String</span><span class="sxs-lookup"><span data-stu-id="3ee08-133">String</span></span> | [<span data-ttu-id="3ee08-134">1.1</span><span class="sxs-lookup"><span data-stu-id="3ee08-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3ee08-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="3ee08-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="3ee08-136">作成</span><span class="sxs-lookup"><span data-stu-id="3ee08-136">Compose</span></span><br><span data-ttu-id="3ee08-137">Read</span><span class="sxs-lookup"><span data-stu-id="3ee08-137">Read</span></span> | <span data-ttu-id="3ee08-138">String</span><span class="sxs-lookup"><span data-stu-id="3ee08-138">String</span></span> | [<span data-ttu-id="3ee08-139">1.1</span><span class="sxs-lookup"><span data-stu-id="3ee08-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3ee08-140">EventType</span><span class="sxs-lookup"><span data-stu-id="3ee08-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="3ee08-141">作成</span><span class="sxs-lookup"><span data-stu-id="3ee08-141">Compose</span></span><br><span data-ttu-id="3ee08-142">Read</span><span class="sxs-lookup"><span data-stu-id="3ee08-142">Read</span></span> | <span data-ttu-id="3ee08-143">String</span><span class="sxs-lookup"><span data-stu-id="3ee08-143">String</span></span> | [<span data-ttu-id="3ee08-144">1.5</span><span class="sxs-lookup"><span data-stu-id="3ee08-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="3ee08-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="3ee08-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="3ee08-146">作成</span><span class="sxs-lookup"><span data-stu-id="3ee08-146">Compose</span></span><br><span data-ttu-id="3ee08-147">Read</span><span class="sxs-lookup"><span data-stu-id="3ee08-147">Read</span></span> | <span data-ttu-id="3ee08-148">String</span><span class="sxs-lookup"><span data-stu-id="3ee08-148">String</span></span> | [<span data-ttu-id="3ee08-149">1.1</span><span class="sxs-lookup"><span data-stu-id="3ee08-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="3ee08-150">名前空間</span><span class="sxs-lookup"><span data-stu-id="3ee08-150">Namespaces</span></span>

<span data-ttu-id="3ee08-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7&preserve-view=true): 、 など、Outlook固有の列挙の `ItemType` `EntityType` `AttachmentType` 数 `RecipientType` が `ResponseType` 含まれています `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="3ee08-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="3ee08-152">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="3ee08-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="3ee08-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="3ee08-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="3ee08-154">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="3ee08-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="3ee08-155">型</span><span class="sxs-lookup"><span data-stu-id="3ee08-155">Type</span></span>

*   <span data-ttu-id="3ee08-156">String</span><span class="sxs-lookup"><span data-stu-id="3ee08-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3ee08-157">プロパティ</span><span class="sxs-lookup"><span data-stu-id="3ee08-157">Properties</span></span>

|<span data-ttu-id="3ee08-158">名前</span><span class="sxs-lookup"><span data-stu-id="3ee08-158">Name</span></span>| <span data-ttu-id="3ee08-159">型</span><span class="sxs-lookup"><span data-stu-id="3ee08-159">Type</span></span>| <span data-ttu-id="3ee08-160">説明</span><span class="sxs-lookup"><span data-stu-id="3ee08-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="3ee08-161">String</span><span class="sxs-lookup"><span data-stu-id="3ee08-161">String</span></span>|<span data-ttu-id="3ee08-162">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="3ee08-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="3ee08-163">String</span><span class="sxs-lookup"><span data-stu-id="3ee08-163">String</span></span>|<span data-ttu-id="3ee08-164">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="3ee08-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3ee08-165">要件</span><span class="sxs-lookup"><span data-stu-id="3ee08-165">Requirements</span></span>

|<span data-ttu-id="3ee08-166">要件</span><span class="sxs-lookup"><span data-stu-id="3ee08-166">Requirement</span></span>| <span data-ttu-id="3ee08-167">値</span><span class="sxs-lookup"><span data-stu-id="3ee08-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee08-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3ee08-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3ee08-169">1.1</span><span class="sxs-lookup"><span data-stu-id="3ee08-169">1.1</span></span>|
|[<span data-ttu-id="3ee08-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3ee08-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3ee08-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3ee08-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="3ee08-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="3ee08-172">CoercionType: String</span></span>

<span data-ttu-id="3ee08-173">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="3ee08-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3ee08-174">型</span><span class="sxs-lookup"><span data-stu-id="3ee08-174">Type</span></span>

*   <span data-ttu-id="3ee08-175">String</span><span class="sxs-lookup"><span data-stu-id="3ee08-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3ee08-176">プロパティ</span><span class="sxs-lookup"><span data-stu-id="3ee08-176">Properties</span></span>

|<span data-ttu-id="3ee08-177">名前</span><span class="sxs-lookup"><span data-stu-id="3ee08-177">Name</span></span>| <span data-ttu-id="3ee08-178">型</span><span class="sxs-lookup"><span data-stu-id="3ee08-178">Type</span></span>| <span data-ttu-id="3ee08-179">説明</span><span class="sxs-lookup"><span data-stu-id="3ee08-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="3ee08-180">String</span><span class="sxs-lookup"><span data-stu-id="3ee08-180">String</span></span>|<span data-ttu-id="3ee08-181">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="3ee08-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="3ee08-182">String</span><span class="sxs-lookup"><span data-stu-id="3ee08-182">String</span></span>|<span data-ttu-id="3ee08-183">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="3ee08-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3ee08-184">要件</span><span class="sxs-lookup"><span data-stu-id="3ee08-184">Requirements</span></span>

|<span data-ttu-id="3ee08-185">要件</span><span class="sxs-lookup"><span data-stu-id="3ee08-185">Requirement</span></span>| <span data-ttu-id="3ee08-186">値</span><span class="sxs-lookup"><span data-stu-id="3ee08-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee08-187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3ee08-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3ee08-188">1.1</span><span class="sxs-lookup"><span data-stu-id="3ee08-188">1.1</span></span>|
|[<span data-ttu-id="3ee08-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3ee08-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3ee08-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3ee08-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="3ee08-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="3ee08-191">EventType: String</span></span>

<span data-ttu-id="3ee08-192">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="3ee08-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="3ee08-193">型</span><span class="sxs-lookup"><span data-stu-id="3ee08-193">Type</span></span>

*   <span data-ttu-id="3ee08-194">String</span><span class="sxs-lookup"><span data-stu-id="3ee08-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3ee08-195">プロパティ</span><span class="sxs-lookup"><span data-stu-id="3ee08-195">Properties</span></span>

| <span data-ttu-id="3ee08-196">名前</span><span class="sxs-lookup"><span data-stu-id="3ee08-196">Name</span></span> | <span data-ttu-id="3ee08-197">型</span><span class="sxs-lookup"><span data-stu-id="3ee08-197">Type</span></span> | <span data-ttu-id="3ee08-198">説明</span><span class="sxs-lookup"><span data-stu-id="3ee08-198">Description</span></span> | <span data-ttu-id="3ee08-199">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="3ee08-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="3ee08-200">String</span><span class="sxs-lookup"><span data-stu-id="3ee08-200">String</span></span> | <span data-ttu-id="3ee08-201">選択した予定または系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="3ee08-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="3ee08-202">1.7</span><span class="sxs-lookup"><span data-stu-id="3ee08-202">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="3ee08-203">String</span><span class="sxs-lookup"><span data-stu-id="3ee08-203">String</span></span> | <span data-ttu-id="3ee08-204">作業ウィンドウOutlook表示する場合は、別のアイテムが選択されています。</span><span class="sxs-lookup"><span data-stu-id="3ee08-204">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="3ee08-205">1.5</span><span class="sxs-lookup"><span data-stu-id="3ee08-205">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="3ee08-206">String</span><span class="sxs-lookup"><span data-stu-id="3ee08-206">String</span></span> | <span data-ttu-id="3ee08-207">選択したアイテムまたは予定の場所の受信者リストが変更されました。</span><span class="sxs-lookup"><span data-stu-id="3ee08-207">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="3ee08-208">1.7</span><span class="sxs-lookup"><span data-stu-id="3ee08-208">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="3ee08-209">String</span><span class="sxs-lookup"><span data-stu-id="3ee08-209">String</span></span> | <span data-ttu-id="3ee08-210">選択した系列の定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="3ee08-210">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="3ee08-211">1.7</span><span class="sxs-lookup"><span data-stu-id="3ee08-211">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3ee08-212">要件</span><span class="sxs-lookup"><span data-stu-id="3ee08-212">Requirements</span></span>

|<span data-ttu-id="3ee08-213">要件</span><span class="sxs-lookup"><span data-stu-id="3ee08-213">Requirement</span></span>| <span data-ttu-id="3ee08-214">値</span><span class="sxs-lookup"><span data-stu-id="3ee08-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee08-215">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3ee08-215">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3ee08-216">1.5</span><span class="sxs-lookup"><span data-stu-id="3ee08-216">1.5</span></span> |
|[<span data-ttu-id="3ee08-217">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3ee08-217">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3ee08-218">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3ee08-218">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="3ee08-219">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="3ee08-219">SourceProperty: String</span></span>

<span data-ttu-id="3ee08-220">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="3ee08-220">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3ee08-221">型</span><span class="sxs-lookup"><span data-stu-id="3ee08-221">Type</span></span>

*   <span data-ttu-id="3ee08-222">String</span><span class="sxs-lookup"><span data-stu-id="3ee08-222">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3ee08-223">プロパティ</span><span class="sxs-lookup"><span data-stu-id="3ee08-223">Properties</span></span>

|<span data-ttu-id="3ee08-224">名前</span><span class="sxs-lookup"><span data-stu-id="3ee08-224">Name</span></span>| <span data-ttu-id="3ee08-225">型</span><span class="sxs-lookup"><span data-stu-id="3ee08-225">Type</span></span>| <span data-ttu-id="3ee08-226">説明</span><span class="sxs-lookup"><span data-stu-id="3ee08-226">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="3ee08-227">String</span><span class="sxs-lookup"><span data-stu-id="3ee08-227">String</span></span>|<span data-ttu-id="3ee08-228">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="3ee08-228">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="3ee08-229">String</span><span class="sxs-lookup"><span data-stu-id="3ee08-229">String</span></span>|<span data-ttu-id="3ee08-230">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="3ee08-230">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3ee08-231">要件</span><span class="sxs-lookup"><span data-stu-id="3ee08-231">Requirements</span></span>

|<span data-ttu-id="3ee08-232">要件</span><span class="sxs-lookup"><span data-stu-id="3ee08-232">Requirement</span></span>| <span data-ttu-id="3ee08-233">値</span><span class="sxs-lookup"><span data-stu-id="3ee08-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee08-234">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3ee08-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3ee08-235">1.1</span><span class="sxs-lookup"><span data-stu-id="3ee08-235">1.1</span></span>|
|[<span data-ttu-id="3ee08-236">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3ee08-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3ee08-237">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3ee08-237">Compose or Read</span></span>|
