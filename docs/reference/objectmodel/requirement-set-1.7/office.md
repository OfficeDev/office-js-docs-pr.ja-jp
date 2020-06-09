---
title: Office 名前空間-要件セット1.7
description: メールボックス API 要件セット1.7 を使用した Outlook アドインで使用可能な Office 名前空間メンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 718de46689fc2fcb52ad455763581ecab06a4c39
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612200"
---
# <a name="office-mailbox-requirement-set-17"></a><span data-ttu-id="2f552-103">Office (メールボックス要件セット 1.7)</span><span class="sxs-lookup"><span data-stu-id="2f552-103">Office (Mailbox requirement set 1.7)</span></span>

<span data-ttu-id="2f552-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2f552-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="2f552-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="2f552-106">Requirements</span></span>

|<span data-ttu-id="2f552-107">要件</span><span class="sxs-lookup"><span data-stu-id="2f552-107">Requirement</span></span>| <span data-ttu-id="2f552-108">値</span><span class="sxs-lookup"><span data-stu-id="2f552-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f552-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2f552-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f552-110">1.1</span><span class="sxs-lookup"><span data-stu-id="2f552-110">1.1</span></span>|
|[<span data-ttu-id="2f552-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2f552-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f552-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2f552-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="2f552-113">Properties</span><span class="sxs-lookup"><span data-stu-id="2f552-113">Properties</span></span>

| <span data-ttu-id="2f552-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="2f552-114">Property</span></span> | <span data-ttu-id="2f552-115">モード</span><span class="sxs-lookup"><span data-stu-id="2f552-115">Modes</span></span> | <span data-ttu-id="2f552-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="2f552-116">Return type</span></span> | <span data-ttu-id="2f552-117">最小値</span><span class="sxs-lookup"><span data-stu-id="2f552-117">Minimum</span></span><br><span data-ttu-id="2f552-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="2f552-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2f552-119">context</span><span class="sxs-lookup"><span data-stu-id="2f552-119">context</span></span>](office.context.md) | <span data-ttu-id="2f552-120">作成</span><span class="sxs-lookup"><span data-stu-id="2f552-120">Compose</span></span><br><span data-ttu-id="2f552-121">Read</span><span class="sxs-lookup"><span data-stu-id="2f552-121">Read</span></span> | [<span data-ttu-id="2f552-122">Context</span><span class="sxs-lookup"><span data-stu-id="2f552-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="2f552-123">1.1</span><span class="sxs-lookup"><span data-stu-id="2f552-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="2f552-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="2f552-124">Enumerations</span></span>

| <span data-ttu-id="2f552-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="2f552-125">Enumeration</span></span> | <span data-ttu-id="2f552-126">モード</span><span class="sxs-lookup"><span data-stu-id="2f552-126">Modes</span></span> | <span data-ttu-id="2f552-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="2f552-127">Return type</span></span> | <span data-ttu-id="2f552-128">最小値</span><span class="sxs-lookup"><span data-stu-id="2f552-128">Minimum</span></span><br><span data-ttu-id="2f552-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="2f552-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2f552-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="2f552-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="2f552-131">作成</span><span class="sxs-lookup"><span data-stu-id="2f552-131">Compose</span></span><br><span data-ttu-id="2f552-132">Read</span><span class="sxs-lookup"><span data-stu-id="2f552-132">Read</span></span> | <span data-ttu-id="2f552-133">String</span><span class="sxs-lookup"><span data-stu-id="2f552-133">String</span></span> | [<span data-ttu-id="2f552-134">1.1</span><span class="sxs-lookup"><span data-stu-id="2f552-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2f552-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="2f552-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="2f552-136">作成</span><span class="sxs-lookup"><span data-stu-id="2f552-136">Compose</span></span><br><span data-ttu-id="2f552-137">Read</span><span class="sxs-lookup"><span data-stu-id="2f552-137">Read</span></span> | <span data-ttu-id="2f552-138">String</span><span class="sxs-lookup"><span data-stu-id="2f552-138">String</span></span> | [<span data-ttu-id="2f552-139">1.1</span><span class="sxs-lookup"><span data-stu-id="2f552-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2f552-140">EventType</span><span class="sxs-lookup"><span data-stu-id="2f552-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="2f552-141">作成</span><span class="sxs-lookup"><span data-stu-id="2f552-141">Compose</span></span><br><span data-ttu-id="2f552-142">Read</span><span class="sxs-lookup"><span data-stu-id="2f552-142">Read</span></span> | <span data-ttu-id="2f552-143">String</span><span class="sxs-lookup"><span data-stu-id="2f552-143">String</span></span> | [<span data-ttu-id="2f552-144">1.5</span><span class="sxs-lookup"><span data-stu-id="2f552-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="2f552-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="2f552-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="2f552-146">作成</span><span class="sxs-lookup"><span data-stu-id="2f552-146">Compose</span></span><br><span data-ttu-id="2f552-147">Read</span><span class="sxs-lookup"><span data-stu-id="2f552-147">Read</span></span> | <span data-ttu-id="2f552-148">String</span><span class="sxs-lookup"><span data-stu-id="2f552-148">String</span></span> | [<span data-ttu-id="2f552-149">1.1</span><span class="sxs-lookup"><span data-stu-id="2f552-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="2f552-150">名前空間</span><span class="sxs-lookup"><span data-stu-id="2f552-150">Namespaces</span></span>

<span data-ttu-id="2f552-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7):、、、、、など、多数の Outlook 固有の列挙を含み `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` ます。</span><span class="sxs-lookup"><span data-stu-id="2f552-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="2f552-152">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="2f552-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="2f552-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="2f552-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="2f552-154">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="2f552-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="2f552-155">型</span><span class="sxs-lookup"><span data-stu-id="2f552-155">Type</span></span>

*   <span data-ttu-id="2f552-156">String</span><span class="sxs-lookup"><span data-stu-id="2f552-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2f552-157">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2f552-157">Properties:</span></span>

|<span data-ttu-id="2f552-158">名前</span><span class="sxs-lookup"><span data-stu-id="2f552-158">Name</span></span>| <span data-ttu-id="2f552-159">種類</span><span class="sxs-lookup"><span data-stu-id="2f552-159">Type</span></span>| <span data-ttu-id="2f552-160">説明</span><span class="sxs-lookup"><span data-stu-id="2f552-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="2f552-161">String</span><span class="sxs-lookup"><span data-stu-id="2f552-161">String</span></span>|<span data-ttu-id="2f552-162">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="2f552-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="2f552-163">String</span><span class="sxs-lookup"><span data-stu-id="2f552-163">String</span></span>|<span data-ttu-id="2f552-164">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="2f552-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2f552-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="2f552-165">Requirements</span></span>

|<span data-ttu-id="2f552-166">要件</span><span class="sxs-lookup"><span data-stu-id="2f552-166">Requirement</span></span>| <span data-ttu-id="2f552-167">値</span><span class="sxs-lookup"><span data-stu-id="2f552-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f552-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2f552-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f552-169">1.1</span><span class="sxs-lookup"><span data-stu-id="2f552-169">1.1</span></span>|
|[<span data-ttu-id="2f552-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2f552-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f552-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2f552-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="2f552-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="2f552-172">CoercionType: String</span></span>

<span data-ttu-id="2f552-173">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="2f552-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2f552-174">型</span><span class="sxs-lookup"><span data-stu-id="2f552-174">Type</span></span>

*   <span data-ttu-id="2f552-175">String</span><span class="sxs-lookup"><span data-stu-id="2f552-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2f552-176">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2f552-176">Properties:</span></span>

|<span data-ttu-id="2f552-177">名前</span><span class="sxs-lookup"><span data-stu-id="2f552-177">Name</span></span>| <span data-ttu-id="2f552-178">種類</span><span class="sxs-lookup"><span data-stu-id="2f552-178">Type</span></span>| <span data-ttu-id="2f552-179">説明</span><span class="sxs-lookup"><span data-stu-id="2f552-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="2f552-180">String</span><span class="sxs-lookup"><span data-stu-id="2f552-180">String</span></span>|<span data-ttu-id="2f552-181">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="2f552-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="2f552-182">String</span><span class="sxs-lookup"><span data-stu-id="2f552-182">String</span></span>|<span data-ttu-id="2f552-183">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="2f552-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2f552-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="2f552-184">Requirements</span></span>

|<span data-ttu-id="2f552-185">要件</span><span class="sxs-lookup"><span data-stu-id="2f552-185">Requirement</span></span>| <span data-ttu-id="2f552-186">値</span><span class="sxs-lookup"><span data-stu-id="2f552-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f552-187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2f552-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f552-188">1.1</span><span class="sxs-lookup"><span data-stu-id="2f552-188">1.1</span></span>|
|[<span data-ttu-id="2f552-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2f552-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f552-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2f552-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="2f552-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="2f552-191">EventType: String</span></span>

<span data-ttu-id="2f552-192">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="2f552-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="2f552-193">型</span><span class="sxs-lookup"><span data-stu-id="2f552-193">Type</span></span>

*   <span data-ttu-id="2f552-194">String</span><span class="sxs-lookup"><span data-stu-id="2f552-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2f552-195">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2f552-195">Properties:</span></span>

| <span data-ttu-id="2f552-196">名前</span><span class="sxs-lookup"><span data-stu-id="2f552-196">Name</span></span> | <span data-ttu-id="2f552-197">種類</span><span class="sxs-lookup"><span data-stu-id="2f552-197">Type</span></span> | <span data-ttu-id="2f552-198">説明</span><span class="sxs-lookup"><span data-stu-id="2f552-198">Description</span></span> | <span data-ttu-id="2f552-199">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="2f552-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="2f552-200">String</span><span class="sxs-lookup"><span data-stu-id="2f552-200">String</span></span> | <span data-ttu-id="2f552-201">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="2f552-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="2f552-202">1.7</span><span class="sxs-lookup"><span data-stu-id="2f552-202">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="2f552-203">String</span><span class="sxs-lookup"><span data-stu-id="2f552-203">String</span></span> | <span data-ttu-id="2f552-204">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="2f552-204">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="2f552-205">1.5</span><span class="sxs-lookup"><span data-stu-id="2f552-205">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="2f552-206">String</span><span class="sxs-lookup"><span data-stu-id="2f552-206">String</span></span> | <span data-ttu-id="2f552-207">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="2f552-207">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="2f552-208">1.7</span><span class="sxs-lookup"><span data-stu-id="2f552-208">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="2f552-209">String</span><span class="sxs-lookup"><span data-stu-id="2f552-209">String</span></span> | <span data-ttu-id="2f552-210">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="2f552-210">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="2f552-211">1.7</span><span class="sxs-lookup"><span data-stu-id="2f552-211">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2f552-212">Requirements</span><span class="sxs-lookup"><span data-stu-id="2f552-212">Requirements</span></span>

|<span data-ttu-id="2f552-213">要件</span><span class="sxs-lookup"><span data-stu-id="2f552-213">Requirement</span></span>| <span data-ttu-id="2f552-214">値</span><span class="sxs-lookup"><span data-stu-id="2f552-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f552-215">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2f552-215">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f552-216">1.5</span><span class="sxs-lookup"><span data-stu-id="2f552-216">1.5</span></span> |
|[<span data-ttu-id="2f552-217">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2f552-217">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f552-218">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2f552-218">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="2f552-219">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="2f552-219">SourceProperty: String</span></span>

<span data-ttu-id="2f552-220">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="2f552-220">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2f552-221">型</span><span class="sxs-lookup"><span data-stu-id="2f552-221">Type</span></span>

*   <span data-ttu-id="2f552-222">String</span><span class="sxs-lookup"><span data-stu-id="2f552-222">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2f552-223">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2f552-223">Properties:</span></span>

|<span data-ttu-id="2f552-224">名前</span><span class="sxs-lookup"><span data-stu-id="2f552-224">Name</span></span>| <span data-ttu-id="2f552-225">種類</span><span class="sxs-lookup"><span data-stu-id="2f552-225">Type</span></span>| <span data-ttu-id="2f552-226">説明</span><span class="sxs-lookup"><span data-stu-id="2f552-226">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="2f552-227">String</span><span class="sxs-lookup"><span data-stu-id="2f552-227">String</span></span>|<span data-ttu-id="2f552-228">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="2f552-228">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="2f552-229">String</span><span class="sxs-lookup"><span data-stu-id="2f552-229">String</span></span>|<span data-ttu-id="2f552-230">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="2f552-230">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2f552-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="2f552-231">Requirements</span></span>

|<span data-ttu-id="2f552-232">要件</span><span class="sxs-lookup"><span data-stu-id="2f552-232">Requirement</span></span>| <span data-ttu-id="2f552-233">値</span><span class="sxs-lookup"><span data-stu-id="2f552-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f552-234">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2f552-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f552-235">1.1</span><span class="sxs-lookup"><span data-stu-id="2f552-235">1.1</span></span>|
|[<span data-ttu-id="2f552-236">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2f552-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f552-237">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2f552-237">Compose or Read</span></span>|
