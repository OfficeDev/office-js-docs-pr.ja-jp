---
title: Office名前空間 - 要件セット 1.9
description: Office API 要件セット 1.9 を使用Outlookアドインで使用できる名前空間メンバーを指定します。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 203b901c619e19a8e5b9255e36274e2f6e1d1658
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590947"
---
# <a name="office-mailbox-requirement-set-19"></a><span data-ttu-id="92021-103">Office (メールボックス要件セット 1.9)</span><span class="sxs-lookup"><span data-stu-id="92021-103">Office (Mailbox requirement set 1.9)</span></span>

<span data-ttu-id="92021-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="92021-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="92021-106">要件</span><span class="sxs-lookup"><span data-stu-id="92021-106">Requirements</span></span>

|<span data-ttu-id="92021-107">要件</span><span class="sxs-lookup"><span data-stu-id="92021-107">Requirement</span></span>| <span data-ttu-id="92021-108">値</span><span class="sxs-lookup"><span data-stu-id="92021-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="92021-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92021-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="92021-110">1.1</span><span class="sxs-lookup"><span data-stu-id="92021-110">1.1</span></span>|
|[<span data-ttu-id="92021-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92021-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="92021-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="92021-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="92021-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="92021-113">Properties</span></span>

| <span data-ttu-id="92021-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="92021-114">Property</span></span> | <span data-ttu-id="92021-115">モード</span><span class="sxs-lookup"><span data-stu-id="92021-115">Modes</span></span> | <span data-ttu-id="92021-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="92021-116">Return type</span></span> | <span data-ttu-id="92021-117">最小値</span><span class="sxs-lookup"><span data-stu-id="92021-117">Minimum</span></span><br><span data-ttu-id="92021-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="92021-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="92021-119">context</span><span class="sxs-lookup"><span data-stu-id="92021-119">context</span></span>](office.context.md) | <span data-ttu-id="92021-120">作成</span><span class="sxs-lookup"><span data-stu-id="92021-120">Compose</span></span><br><span data-ttu-id="92021-121">Read</span><span class="sxs-lookup"><span data-stu-id="92021-121">Read</span></span> | [<span data-ttu-id="92021-122">Context</span><span class="sxs-lookup"><span data-stu-id="92021-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="92021-123">1.1</span><span class="sxs-lookup"><span data-stu-id="92021-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="92021-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="92021-124">Enumerations</span></span>

| <span data-ttu-id="92021-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="92021-125">Enumeration</span></span> | <span data-ttu-id="92021-126">モード</span><span class="sxs-lookup"><span data-stu-id="92021-126">Modes</span></span> | <span data-ttu-id="92021-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="92021-127">Return type</span></span> | <span data-ttu-id="92021-128">最小値</span><span class="sxs-lookup"><span data-stu-id="92021-128">Minimum</span></span><br><span data-ttu-id="92021-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="92021-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="92021-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="92021-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="92021-131">作成</span><span class="sxs-lookup"><span data-stu-id="92021-131">Compose</span></span><br><span data-ttu-id="92021-132">Read</span><span class="sxs-lookup"><span data-stu-id="92021-132">Read</span></span> | <span data-ttu-id="92021-133">String</span><span class="sxs-lookup"><span data-stu-id="92021-133">String</span></span> | [<span data-ttu-id="92021-134">1.1</span><span class="sxs-lookup"><span data-stu-id="92021-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="92021-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="92021-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="92021-136">作成</span><span class="sxs-lookup"><span data-stu-id="92021-136">Compose</span></span><br><span data-ttu-id="92021-137">Read</span><span class="sxs-lookup"><span data-stu-id="92021-137">Read</span></span> | <span data-ttu-id="92021-138">String</span><span class="sxs-lookup"><span data-stu-id="92021-138">String</span></span> | [<span data-ttu-id="92021-139">1.1</span><span class="sxs-lookup"><span data-stu-id="92021-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="92021-140">EventType</span><span class="sxs-lookup"><span data-stu-id="92021-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="92021-141">作成</span><span class="sxs-lookup"><span data-stu-id="92021-141">Compose</span></span><br><span data-ttu-id="92021-142">Read</span><span class="sxs-lookup"><span data-stu-id="92021-142">Read</span></span> | <span data-ttu-id="92021-143">String</span><span class="sxs-lookup"><span data-stu-id="92021-143">String</span></span> | [<span data-ttu-id="92021-144">1.5</span><span class="sxs-lookup"><span data-stu-id="92021-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="92021-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="92021-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="92021-146">作成</span><span class="sxs-lookup"><span data-stu-id="92021-146">Compose</span></span><br><span data-ttu-id="92021-147">Read</span><span class="sxs-lookup"><span data-stu-id="92021-147">Read</span></span> | <span data-ttu-id="92021-148">String</span><span class="sxs-lookup"><span data-stu-id="92021-148">String</span></span> | [<span data-ttu-id="92021-149">1.1</span><span class="sxs-lookup"><span data-stu-id="92021-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="92021-150">名前空間</span><span class="sxs-lookup"><span data-stu-id="92021-150">Namespaces</span></span>

<span data-ttu-id="92021-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.9&preserve-view=true): 、 など、Outlook固有の列挙の `ItemType` `EntityType` `AttachmentType` 数 `RecipientType` が `ResponseType` 含まれています `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="92021-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.9&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="92021-152">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="92021-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="92021-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="92021-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="92021-154">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="92021-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="92021-155">型</span><span class="sxs-lookup"><span data-stu-id="92021-155">Type</span></span>

*   <span data-ttu-id="92021-156">String</span><span class="sxs-lookup"><span data-stu-id="92021-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="92021-157">プロパティ</span><span class="sxs-lookup"><span data-stu-id="92021-157">Properties</span></span>

|<span data-ttu-id="92021-158">名前</span><span class="sxs-lookup"><span data-stu-id="92021-158">Name</span></span>| <span data-ttu-id="92021-159">型</span><span class="sxs-lookup"><span data-stu-id="92021-159">Type</span></span>| <span data-ttu-id="92021-160">説明</span><span class="sxs-lookup"><span data-stu-id="92021-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="92021-161">String</span><span class="sxs-lookup"><span data-stu-id="92021-161">String</span></span>|<span data-ttu-id="92021-162">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="92021-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="92021-163">String</span><span class="sxs-lookup"><span data-stu-id="92021-163">String</span></span>|<span data-ttu-id="92021-164">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="92021-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92021-165">要件</span><span class="sxs-lookup"><span data-stu-id="92021-165">Requirements</span></span>

|<span data-ttu-id="92021-166">要件</span><span class="sxs-lookup"><span data-stu-id="92021-166">Requirement</span></span>| <span data-ttu-id="92021-167">値</span><span class="sxs-lookup"><span data-stu-id="92021-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="92021-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92021-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="92021-169">1.1</span><span class="sxs-lookup"><span data-stu-id="92021-169">1.1</span></span>|
|[<span data-ttu-id="92021-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92021-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="92021-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="92021-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="92021-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="92021-172">CoercionType: String</span></span>

<span data-ttu-id="92021-173">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="92021-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="92021-174">型</span><span class="sxs-lookup"><span data-stu-id="92021-174">Type</span></span>

*   <span data-ttu-id="92021-175">String</span><span class="sxs-lookup"><span data-stu-id="92021-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="92021-176">プロパティ</span><span class="sxs-lookup"><span data-stu-id="92021-176">Properties</span></span>

|<span data-ttu-id="92021-177">名前</span><span class="sxs-lookup"><span data-stu-id="92021-177">Name</span></span>| <span data-ttu-id="92021-178">型</span><span class="sxs-lookup"><span data-stu-id="92021-178">Type</span></span>| <span data-ttu-id="92021-179">説明</span><span class="sxs-lookup"><span data-stu-id="92021-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="92021-180">String</span><span class="sxs-lookup"><span data-stu-id="92021-180">String</span></span>|<span data-ttu-id="92021-181">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="92021-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="92021-182">String</span><span class="sxs-lookup"><span data-stu-id="92021-182">String</span></span>|<span data-ttu-id="92021-183">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="92021-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92021-184">要件</span><span class="sxs-lookup"><span data-stu-id="92021-184">Requirements</span></span>

|<span data-ttu-id="92021-185">要件</span><span class="sxs-lookup"><span data-stu-id="92021-185">Requirement</span></span>| <span data-ttu-id="92021-186">値</span><span class="sxs-lookup"><span data-stu-id="92021-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="92021-187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92021-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="92021-188">1.1</span><span class="sxs-lookup"><span data-stu-id="92021-188">1.1</span></span>|
|[<span data-ttu-id="92021-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92021-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="92021-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="92021-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="92021-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="92021-191">EventType: String</span></span>

<span data-ttu-id="92021-192">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="92021-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="92021-193">型</span><span class="sxs-lookup"><span data-stu-id="92021-193">Type</span></span>

*   <span data-ttu-id="92021-194">String</span><span class="sxs-lookup"><span data-stu-id="92021-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="92021-195">プロパティ</span><span class="sxs-lookup"><span data-stu-id="92021-195">Properties</span></span>

| <span data-ttu-id="92021-196">名前</span><span class="sxs-lookup"><span data-stu-id="92021-196">Name</span></span> | <span data-ttu-id="92021-197">型</span><span class="sxs-lookup"><span data-stu-id="92021-197">Type</span></span> | <span data-ttu-id="92021-198">説明</span><span class="sxs-lookup"><span data-stu-id="92021-198">Description</span></span> | <span data-ttu-id="92021-199">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="92021-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="92021-200">String</span><span class="sxs-lookup"><span data-stu-id="92021-200">String</span></span> | <span data-ttu-id="92021-201">選択した予定または系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="92021-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="92021-202">1.7</span><span class="sxs-lookup"><span data-stu-id="92021-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="92021-203">String</span><span class="sxs-lookup"><span data-stu-id="92021-203">String</span></span> | <span data-ttu-id="92021-204">アイテムに添付ファイルが追加または削除されました。</span><span class="sxs-lookup"><span data-stu-id="92021-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="92021-205">1.8</span><span class="sxs-lookup"><span data-stu-id="92021-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="92021-206">String</span><span class="sxs-lookup"><span data-stu-id="92021-206">String</span></span> | <span data-ttu-id="92021-207">選択した予定の場所が変更されました。</span><span class="sxs-lookup"><span data-stu-id="92021-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="92021-208">1.8</span><span class="sxs-lookup"><span data-stu-id="92021-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="92021-209">String</span><span class="sxs-lookup"><span data-stu-id="92021-209">String</span></span> | <span data-ttu-id="92021-210">作業ウィンドウOutlook表示する場合は、別のアイテムが選択されています。</span><span class="sxs-lookup"><span data-stu-id="92021-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="92021-211">1.5</span><span class="sxs-lookup"><span data-stu-id="92021-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="92021-212">String</span><span class="sxs-lookup"><span data-stu-id="92021-212">String</span></span> | <span data-ttu-id="92021-213">選択したアイテムまたは予定の場所の受信者リストが変更されました。</span><span class="sxs-lookup"><span data-stu-id="92021-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="92021-214">1.7</span><span class="sxs-lookup"><span data-stu-id="92021-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="92021-215">String</span><span class="sxs-lookup"><span data-stu-id="92021-215">String</span></span> | <span data-ttu-id="92021-216">選択した系列の定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="92021-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="92021-217">1.7</span><span class="sxs-lookup"><span data-stu-id="92021-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="92021-218">要件</span><span class="sxs-lookup"><span data-stu-id="92021-218">Requirements</span></span>

|<span data-ttu-id="92021-219">要件</span><span class="sxs-lookup"><span data-stu-id="92021-219">Requirement</span></span>| <span data-ttu-id="92021-220">値</span><span class="sxs-lookup"><span data-stu-id="92021-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="92021-221">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92021-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="92021-222">1.5</span><span class="sxs-lookup"><span data-stu-id="92021-222">1.5</span></span> |
|[<span data-ttu-id="92021-223">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92021-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="92021-224">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="92021-224">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="92021-225">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="92021-225">SourceProperty: String</span></span>

<span data-ttu-id="92021-226">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="92021-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="92021-227">型</span><span class="sxs-lookup"><span data-stu-id="92021-227">Type</span></span>

*   <span data-ttu-id="92021-228">String</span><span class="sxs-lookup"><span data-stu-id="92021-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="92021-229">プロパティ</span><span class="sxs-lookup"><span data-stu-id="92021-229">Properties</span></span>

|<span data-ttu-id="92021-230">名前</span><span class="sxs-lookup"><span data-stu-id="92021-230">Name</span></span>| <span data-ttu-id="92021-231">型</span><span class="sxs-lookup"><span data-stu-id="92021-231">Type</span></span>| <span data-ttu-id="92021-232">説明</span><span class="sxs-lookup"><span data-stu-id="92021-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="92021-233">String</span><span class="sxs-lookup"><span data-stu-id="92021-233">String</span></span>|<span data-ttu-id="92021-234">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="92021-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="92021-235">String</span><span class="sxs-lookup"><span data-stu-id="92021-235">String</span></span>|<span data-ttu-id="92021-236">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="92021-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92021-237">要件</span><span class="sxs-lookup"><span data-stu-id="92021-237">Requirements</span></span>

|<span data-ttu-id="92021-238">要件</span><span class="sxs-lookup"><span data-stu-id="92021-238">Requirement</span></span>| <span data-ttu-id="92021-239">値</span><span class="sxs-lookup"><span data-stu-id="92021-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="92021-240">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92021-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="92021-241">1.1</span><span class="sxs-lookup"><span data-stu-id="92021-241">1.1</span></span>|
|[<span data-ttu-id="92021-242">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92021-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="92021-243">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="92021-243">Compose or Read</span></span>|
