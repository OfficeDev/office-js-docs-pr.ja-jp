---
title: Office 名前空間-要件セット1.8
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: b23afd7b84dcd18e120f6aea4bd4fb0952791f1c
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814167"
---
# <a name="office"></a><span data-ttu-id="34399-102">Office</span><span class="sxs-lookup"><span data-stu-id="34399-102">Office</span></span>

<span data-ttu-id="34399-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="34399-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="34399-105">要件</span><span class="sxs-lookup"><span data-stu-id="34399-105">Requirements</span></span>

|<span data-ttu-id="34399-106">要件</span><span class="sxs-lookup"><span data-stu-id="34399-106">Requirement</span></span>| <span data-ttu-id="34399-107">値</span><span class="sxs-lookup"><span data-stu-id="34399-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="34399-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="34399-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="34399-109">1.1</span><span class="sxs-lookup"><span data-stu-id="34399-109">1.1</span></span>|
|[<span data-ttu-id="34399-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="34399-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34399-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="34399-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="34399-112">Properties</span><span class="sxs-lookup"><span data-stu-id="34399-112">Properties</span></span>

| <span data-ttu-id="34399-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="34399-113">Property</span></span> | <span data-ttu-id="34399-114">モード</span><span class="sxs-lookup"><span data-stu-id="34399-114">Modes</span></span> | <span data-ttu-id="34399-115">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="34399-115">Return type</span></span> | <span data-ttu-id="34399-116">最小値</span><span class="sxs-lookup"><span data-stu-id="34399-116">Minimum</span></span><br><span data-ttu-id="34399-117">要件セット</span><span class="sxs-lookup"><span data-stu-id="34399-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="34399-118">context</span><span class="sxs-lookup"><span data-stu-id="34399-118">context</span></span>](office.context.md) | <span data-ttu-id="34399-119">作成</span><span class="sxs-lookup"><span data-stu-id="34399-119">Compose</span></span><br><span data-ttu-id="34399-120">読み取り</span><span class="sxs-lookup"><span data-stu-id="34399-120">Read</span></span> | [<span data-ttu-id="34399-121">Context</span><span class="sxs-lookup"><span data-stu-id="34399-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8) | [<span data-ttu-id="34399-122">1.1</span><span class="sxs-lookup"><span data-stu-id="34399-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="34399-123">列挙型</span><span class="sxs-lookup"><span data-stu-id="34399-123">Enumerations</span></span>

| <span data-ttu-id="34399-124">列挙体</span><span class="sxs-lookup"><span data-stu-id="34399-124">Enumeration</span></span> | <span data-ttu-id="34399-125">モード</span><span class="sxs-lookup"><span data-stu-id="34399-125">Modes</span></span> | <span data-ttu-id="34399-126">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="34399-126">Return type</span></span> | <span data-ttu-id="34399-127">最小値</span><span class="sxs-lookup"><span data-stu-id="34399-127">Minimum</span></span><br><span data-ttu-id="34399-128">要件セット</span><span class="sxs-lookup"><span data-stu-id="34399-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="34399-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="34399-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="34399-130">作成</span><span class="sxs-lookup"><span data-stu-id="34399-130">Compose</span></span><br><span data-ttu-id="34399-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="34399-131">Read</span></span> | <span data-ttu-id="34399-132">String</span><span class="sxs-lookup"><span data-stu-id="34399-132">String</span></span> | [<span data-ttu-id="34399-133">1.1</span><span class="sxs-lookup"><span data-stu-id="34399-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="34399-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="34399-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="34399-135">作成</span><span class="sxs-lookup"><span data-stu-id="34399-135">Compose</span></span><br><span data-ttu-id="34399-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="34399-136">Read</span></span> | <span data-ttu-id="34399-137">String</span><span class="sxs-lookup"><span data-stu-id="34399-137">String</span></span> | [<span data-ttu-id="34399-138">1.1</span><span class="sxs-lookup"><span data-stu-id="34399-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="34399-139">EventType</span><span class="sxs-lookup"><span data-stu-id="34399-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="34399-140">作成</span><span class="sxs-lookup"><span data-stu-id="34399-140">Compose</span></span><br><span data-ttu-id="34399-141">読み取り</span><span class="sxs-lookup"><span data-stu-id="34399-141">Read</span></span> | <span data-ttu-id="34399-142">String</span><span class="sxs-lookup"><span data-stu-id="34399-142">String</span></span> | [<span data-ttu-id="34399-143">1.5</span><span class="sxs-lookup"><span data-stu-id="34399-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="34399-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="34399-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="34399-145">作成</span><span class="sxs-lookup"><span data-stu-id="34399-145">Compose</span></span><br><span data-ttu-id="34399-146">読み取り</span><span class="sxs-lookup"><span data-stu-id="34399-146">Read</span></span> | <span data-ttu-id="34399-147">String</span><span class="sxs-lookup"><span data-stu-id="34399-147">String</span></span> | [<span data-ttu-id="34399-148">1.1</span><span class="sxs-lookup"><span data-stu-id="34399-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="34399-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="34399-149">Namespaces</span></span>

<span data-ttu-id="34399-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="34399-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="34399-151">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="34399-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="34399-152">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="34399-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="34399-153">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="34399-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="34399-154">型</span><span class="sxs-lookup"><span data-stu-id="34399-154">Type</span></span>

*   <span data-ttu-id="34399-155">String</span><span class="sxs-lookup"><span data-stu-id="34399-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="34399-156">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="34399-156">Properties:</span></span>

|<span data-ttu-id="34399-157">名前</span><span class="sxs-lookup"><span data-stu-id="34399-157">Name</span></span>| <span data-ttu-id="34399-158">種類</span><span class="sxs-lookup"><span data-stu-id="34399-158">Type</span></span>| <span data-ttu-id="34399-159">説明</span><span class="sxs-lookup"><span data-stu-id="34399-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="34399-160">String</span><span class="sxs-lookup"><span data-stu-id="34399-160">String</span></span>|<span data-ttu-id="34399-161">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="34399-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="34399-162">String</span><span class="sxs-lookup"><span data-stu-id="34399-162">String</span></span>|<span data-ttu-id="34399-163">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="34399-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="34399-164">要件</span><span class="sxs-lookup"><span data-stu-id="34399-164">Requirements</span></span>

|<span data-ttu-id="34399-165">要件</span><span class="sxs-lookup"><span data-stu-id="34399-165">Requirement</span></span>| <span data-ttu-id="34399-166">値</span><span class="sxs-lookup"><span data-stu-id="34399-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="34399-167">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="34399-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="34399-168">1.1</span><span class="sxs-lookup"><span data-stu-id="34399-168">1.1</span></span>|
|[<span data-ttu-id="34399-169">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="34399-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34399-170">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="34399-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="34399-171">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="34399-171">CoercionType: String</span></span>

<span data-ttu-id="34399-172">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="34399-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="34399-173">型</span><span class="sxs-lookup"><span data-stu-id="34399-173">Type</span></span>

*   <span data-ttu-id="34399-174">String</span><span class="sxs-lookup"><span data-stu-id="34399-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="34399-175">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="34399-175">Properties:</span></span>

|<span data-ttu-id="34399-176">名前</span><span class="sxs-lookup"><span data-stu-id="34399-176">Name</span></span>| <span data-ttu-id="34399-177">種類</span><span class="sxs-lookup"><span data-stu-id="34399-177">Type</span></span>| <span data-ttu-id="34399-178">説明</span><span class="sxs-lookup"><span data-stu-id="34399-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="34399-179">String</span><span class="sxs-lookup"><span data-stu-id="34399-179">String</span></span>|<span data-ttu-id="34399-180">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="34399-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="34399-181">String</span><span class="sxs-lookup"><span data-stu-id="34399-181">String</span></span>|<span data-ttu-id="34399-182">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="34399-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="34399-183">要件</span><span class="sxs-lookup"><span data-stu-id="34399-183">Requirements</span></span>

|<span data-ttu-id="34399-184">要件</span><span class="sxs-lookup"><span data-stu-id="34399-184">Requirement</span></span>| <span data-ttu-id="34399-185">値</span><span class="sxs-lookup"><span data-stu-id="34399-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="34399-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="34399-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="34399-187">1.1</span><span class="sxs-lookup"><span data-stu-id="34399-187">1.1</span></span>|
|[<span data-ttu-id="34399-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="34399-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34399-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="34399-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="34399-190">EventType: String</span><span class="sxs-lookup"><span data-stu-id="34399-190">EventType: String</span></span>

<span data-ttu-id="34399-191">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="34399-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="34399-192">型</span><span class="sxs-lookup"><span data-stu-id="34399-192">Type</span></span>

*   <span data-ttu-id="34399-193">String</span><span class="sxs-lookup"><span data-stu-id="34399-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="34399-194">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="34399-194">Properties:</span></span>

| <span data-ttu-id="34399-195">名前</span><span class="sxs-lookup"><span data-stu-id="34399-195">Name</span></span> | <span data-ttu-id="34399-196">種類</span><span class="sxs-lookup"><span data-stu-id="34399-196">Type</span></span> | <span data-ttu-id="34399-197">説明</span><span class="sxs-lookup"><span data-stu-id="34399-197">Description</span></span> | <span data-ttu-id="34399-198">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="34399-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="34399-199">String</span><span class="sxs-lookup"><span data-stu-id="34399-199">String</span></span> | <span data-ttu-id="34399-200">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="34399-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="34399-201">1.7</span><span class="sxs-lookup"><span data-stu-id="34399-201">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="34399-202">String</span><span class="sxs-lookup"><span data-stu-id="34399-202">String</span></span> | <span data-ttu-id="34399-203">添付ファイルがアイテムに追加またはアイテムから削除されています。</span><span class="sxs-lookup"><span data-stu-id="34399-203">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="34399-204">1.8</span><span class="sxs-lookup"><span data-stu-id="34399-204">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="34399-205">String</span><span class="sxs-lookup"><span data-stu-id="34399-205">String</span></span> | <span data-ttu-id="34399-206">選択した予定の場所が変更されました。</span><span class="sxs-lookup"><span data-stu-id="34399-206">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="34399-207">1.8</span><span class="sxs-lookup"><span data-stu-id="34399-207">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="34399-208">String</span><span class="sxs-lookup"><span data-stu-id="34399-208">String</span></span> | <span data-ttu-id="34399-209">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="34399-209">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="34399-210">1.5</span><span class="sxs-lookup"><span data-stu-id="34399-210">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="34399-211">String</span><span class="sxs-lookup"><span data-stu-id="34399-211">String</span></span> | <span data-ttu-id="34399-212">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="34399-212">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="34399-213">1.7</span><span class="sxs-lookup"><span data-stu-id="34399-213">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="34399-214">String</span><span class="sxs-lookup"><span data-stu-id="34399-214">String</span></span> | <span data-ttu-id="34399-215">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="34399-215">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="34399-216">1.7</span><span class="sxs-lookup"><span data-stu-id="34399-216">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="34399-217">要件</span><span class="sxs-lookup"><span data-stu-id="34399-217">Requirements</span></span>

|<span data-ttu-id="34399-218">要件</span><span class="sxs-lookup"><span data-stu-id="34399-218">Requirement</span></span>| <span data-ttu-id="34399-219">値</span><span class="sxs-lookup"><span data-stu-id="34399-219">Value</span></span>|
|---|---|
|[<span data-ttu-id="34399-220">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="34399-220">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="34399-221">1.5</span><span class="sxs-lookup"><span data-stu-id="34399-221">1.5</span></span> |
|[<span data-ttu-id="34399-222">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="34399-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34399-223">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="34399-223">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="34399-224">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="34399-224">SourceProperty: String</span></span>

<span data-ttu-id="34399-225">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="34399-225">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="34399-226">型</span><span class="sxs-lookup"><span data-stu-id="34399-226">Type</span></span>

*   <span data-ttu-id="34399-227">String</span><span class="sxs-lookup"><span data-stu-id="34399-227">String</span></span>

##### <a name="properties"></a><span data-ttu-id="34399-228">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="34399-228">Properties:</span></span>

|<span data-ttu-id="34399-229">名前</span><span class="sxs-lookup"><span data-stu-id="34399-229">Name</span></span>| <span data-ttu-id="34399-230">種類</span><span class="sxs-lookup"><span data-stu-id="34399-230">Type</span></span>| <span data-ttu-id="34399-231">説明</span><span class="sxs-lookup"><span data-stu-id="34399-231">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="34399-232">String</span><span class="sxs-lookup"><span data-stu-id="34399-232">String</span></span>|<span data-ttu-id="34399-233">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="34399-233">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="34399-234">String</span><span class="sxs-lookup"><span data-stu-id="34399-234">String</span></span>|<span data-ttu-id="34399-235">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="34399-235">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="34399-236">要件</span><span class="sxs-lookup"><span data-stu-id="34399-236">Requirements</span></span>

|<span data-ttu-id="34399-237">要件</span><span class="sxs-lookup"><span data-stu-id="34399-237">Requirement</span></span>| <span data-ttu-id="34399-238">値</span><span class="sxs-lookup"><span data-stu-id="34399-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="34399-239">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="34399-239">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="34399-240">1.1</span><span class="sxs-lookup"><span data-stu-id="34399-240">1.1</span></span>|
|[<span data-ttu-id="34399-241">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="34399-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="34399-242">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="34399-242">Compose or Read</span></span>|
