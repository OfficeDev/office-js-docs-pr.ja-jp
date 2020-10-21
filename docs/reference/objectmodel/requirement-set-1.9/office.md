---
title: Office 名前空間-要件セット1.9
description: メールボックス API 要件セット1.9 を使用した Outlook アドインで使用可能な Office 名前空間メンバー。
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: e6a932c528dea692ff5fd7ea8d3e1454bb9a7e03
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/20/2020
ms.locfileid: "48628065"
---
# <a name="office-mailbox-requirement-set-19"></a><span data-ttu-id="aeb82-103">Office (メールボックス要件セット 1.9)</span><span class="sxs-lookup"><span data-stu-id="aeb82-103">Office (Mailbox requirement set 1.9)</span></span>

<span data-ttu-id="aeb82-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="aeb82-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="aeb82-106">要件</span><span class="sxs-lookup"><span data-stu-id="aeb82-106">Requirements</span></span>

|<span data-ttu-id="aeb82-107">要件</span><span class="sxs-lookup"><span data-stu-id="aeb82-107">Requirement</span></span>| <span data-ttu-id="aeb82-108">値</span><span class="sxs-lookup"><span data-stu-id="aeb82-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="aeb82-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aeb82-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="aeb82-110">1.1</span><span class="sxs-lookup"><span data-stu-id="aeb82-110">1.1</span></span>|
|[<span data-ttu-id="aeb82-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aeb82-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="aeb82-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aeb82-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="aeb82-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="aeb82-113">Properties</span></span>

| <span data-ttu-id="aeb82-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="aeb82-114">Property</span></span> | <span data-ttu-id="aeb82-115">モード</span><span class="sxs-lookup"><span data-stu-id="aeb82-115">Modes</span></span> | <span data-ttu-id="aeb82-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="aeb82-116">Return type</span></span> | <span data-ttu-id="aeb82-117">最小値</span><span class="sxs-lookup"><span data-stu-id="aeb82-117">Minimum</span></span><br><span data-ttu-id="aeb82-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="aeb82-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="aeb82-119">context</span><span class="sxs-lookup"><span data-stu-id="aeb82-119">context</span></span>](office.context.md) | <span data-ttu-id="aeb82-120">作成</span><span class="sxs-lookup"><span data-stu-id="aeb82-120">Compose</span></span><br><span data-ttu-id="aeb82-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="aeb82-121">Read</span></span> | [<span data-ttu-id="aeb82-122">Context</span><span class="sxs-lookup"><span data-stu-id="aeb82-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="aeb82-123">1.1</span><span class="sxs-lookup"><span data-stu-id="aeb82-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="aeb82-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="aeb82-124">Enumerations</span></span>

| <span data-ttu-id="aeb82-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="aeb82-125">Enumeration</span></span> | <span data-ttu-id="aeb82-126">モード</span><span class="sxs-lookup"><span data-stu-id="aeb82-126">Modes</span></span> | <span data-ttu-id="aeb82-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="aeb82-127">Return type</span></span> | <span data-ttu-id="aeb82-128">最小値</span><span class="sxs-lookup"><span data-stu-id="aeb82-128">Minimum</span></span><br><span data-ttu-id="aeb82-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="aeb82-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="aeb82-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="aeb82-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="aeb82-131">作成</span><span class="sxs-lookup"><span data-stu-id="aeb82-131">Compose</span></span><br><span data-ttu-id="aeb82-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="aeb82-132">Read</span></span> | <span data-ttu-id="aeb82-133">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-133">String</span></span> | [<span data-ttu-id="aeb82-134">1.1</span><span class="sxs-lookup"><span data-stu-id="aeb82-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="aeb82-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="aeb82-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="aeb82-136">作成</span><span class="sxs-lookup"><span data-stu-id="aeb82-136">Compose</span></span><br><span data-ttu-id="aeb82-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="aeb82-137">Read</span></span> | <span data-ttu-id="aeb82-138">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-138">String</span></span> | [<span data-ttu-id="aeb82-139">1.1</span><span class="sxs-lookup"><span data-stu-id="aeb82-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="aeb82-140">EventType</span><span class="sxs-lookup"><span data-stu-id="aeb82-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="aeb82-141">作成</span><span class="sxs-lookup"><span data-stu-id="aeb82-141">Compose</span></span><br><span data-ttu-id="aeb82-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="aeb82-142">Read</span></span> | <span data-ttu-id="aeb82-143">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-143">String</span></span> | [<span data-ttu-id="aeb82-144">1.5</span><span class="sxs-lookup"><span data-stu-id="aeb82-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="aeb82-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="aeb82-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="aeb82-146">作成</span><span class="sxs-lookup"><span data-stu-id="aeb82-146">Compose</span></span><br><span data-ttu-id="aeb82-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="aeb82-147">Read</span></span> | <span data-ttu-id="aeb82-148">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-148">String</span></span> | [<span data-ttu-id="aeb82-149">1.1</span><span class="sxs-lookup"><span data-stu-id="aeb82-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="aeb82-150">名前空間</span><span class="sxs-lookup"><span data-stu-id="aeb82-150">Namespaces</span></span>

<span data-ttu-id="aeb82-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.9&preserve-view=true):、、、、、など、多数の Outlook 固有の列挙を含み `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` ます。</span><span class="sxs-lookup"><span data-stu-id="aeb82-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.9&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="aeb82-152">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="aeb82-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="aeb82-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="aeb82-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="aeb82-154">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="aeb82-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="aeb82-155">型</span><span class="sxs-lookup"><span data-stu-id="aeb82-155">Type</span></span>

*   <span data-ttu-id="aeb82-156">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="aeb82-157">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="aeb82-157">Properties:</span></span>

|<span data-ttu-id="aeb82-158">名前</span><span class="sxs-lookup"><span data-stu-id="aeb82-158">Name</span></span>| <span data-ttu-id="aeb82-159">種類</span><span class="sxs-lookup"><span data-stu-id="aeb82-159">Type</span></span>| <span data-ttu-id="aeb82-160">説明</span><span class="sxs-lookup"><span data-stu-id="aeb82-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="aeb82-161">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-161">String</span></span>|<span data-ttu-id="aeb82-162">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="aeb82-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="aeb82-163">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-163">String</span></span>|<span data-ttu-id="aeb82-164">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="aeb82-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aeb82-165">要件</span><span class="sxs-lookup"><span data-stu-id="aeb82-165">Requirements</span></span>

|<span data-ttu-id="aeb82-166">要件</span><span class="sxs-lookup"><span data-stu-id="aeb82-166">Requirement</span></span>| <span data-ttu-id="aeb82-167">値</span><span class="sxs-lookup"><span data-stu-id="aeb82-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="aeb82-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aeb82-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="aeb82-169">1.1</span><span class="sxs-lookup"><span data-stu-id="aeb82-169">1.1</span></span>|
|[<span data-ttu-id="aeb82-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aeb82-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="aeb82-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aeb82-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="aeb82-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="aeb82-172">CoercionType: String</span></span>

<span data-ttu-id="aeb82-173">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="aeb82-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="aeb82-174">型</span><span class="sxs-lookup"><span data-stu-id="aeb82-174">Type</span></span>

*   <span data-ttu-id="aeb82-175">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="aeb82-176">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="aeb82-176">Properties:</span></span>

|<span data-ttu-id="aeb82-177">名前</span><span class="sxs-lookup"><span data-stu-id="aeb82-177">Name</span></span>| <span data-ttu-id="aeb82-178">種類</span><span class="sxs-lookup"><span data-stu-id="aeb82-178">Type</span></span>| <span data-ttu-id="aeb82-179">説明</span><span class="sxs-lookup"><span data-stu-id="aeb82-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="aeb82-180">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-180">String</span></span>|<span data-ttu-id="aeb82-181">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="aeb82-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="aeb82-182">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-182">String</span></span>|<span data-ttu-id="aeb82-183">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="aeb82-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aeb82-184">要件</span><span class="sxs-lookup"><span data-stu-id="aeb82-184">Requirements</span></span>

|<span data-ttu-id="aeb82-185">要件</span><span class="sxs-lookup"><span data-stu-id="aeb82-185">Requirement</span></span>| <span data-ttu-id="aeb82-186">値</span><span class="sxs-lookup"><span data-stu-id="aeb82-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="aeb82-187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aeb82-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="aeb82-188">1.1</span><span class="sxs-lookup"><span data-stu-id="aeb82-188">1.1</span></span>|
|[<span data-ttu-id="aeb82-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aeb82-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="aeb82-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aeb82-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="aeb82-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="aeb82-191">EventType: String</span></span>

<span data-ttu-id="aeb82-192">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="aeb82-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="aeb82-193">型</span><span class="sxs-lookup"><span data-stu-id="aeb82-193">Type</span></span>

*   <span data-ttu-id="aeb82-194">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="aeb82-195">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="aeb82-195">Properties:</span></span>

| <span data-ttu-id="aeb82-196">名前</span><span class="sxs-lookup"><span data-stu-id="aeb82-196">Name</span></span> | <span data-ttu-id="aeb82-197">種類</span><span class="sxs-lookup"><span data-stu-id="aeb82-197">Type</span></span> | <span data-ttu-id="aeb82-198">説明</span><span class="sxs-lookup"><span data-stu-id="aeb82-198">Description</span></span> | <span data-ttu-id="aeb82-199">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="aeb82-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="aeb82-200">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-200">String</span></span> | <span data-ttu-id="aeb82-201">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="aeb82-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="aeb82-202">1.7</span><span class="sxs-lookup"><span data-stu-id="aeb82-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="aeb82-203">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-203">String</span></span> | <span data-ttu-id="aeb82-204">添付ファイルがアイテムに追加またはアイテムから削除されています。</span><span class="sxs-lookup"><span data-stu-id="aeb82-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="aeb82-205">1.8</span><span class="sxs-lookup"><span data-stu-id="aeb82-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="aeb82-206">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-206">String</span></span> | <span data-ttu-id="aeb82-207">選択した予定の場所が変更されました。</span><span class="sxs-lookup"><span data-stu-id="aeb82-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="aeb82-208">1.8</span><span class="sxs-lookup"><span data-stu-id="aeb82-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="aeb82-209">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-209">String</span></span> | <span data-ttu-id="aeb82-210">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="aeb82-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="aeb82-211">1.5</span><span class="sxs-lookup"><span data-stu-id="aeb82-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="aeb82-212">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-212">String</span></span> | <span data-ttu-id="aeb82-213">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="aeb82-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="aeb82-214">1.7</span><span class="sxs-lookup"><span data-stu-id="aeb82-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="aeb82-215">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-215">String</span></span> | <span data-ttu-id="aeb82-216">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="aeb82-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="aeb82-217">1.7</span><span class="sxs-lookup"><span data-stu-id="aeb82-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="aeb82-218">要件</span><span class="sxs-lookup"><span data-stu-id="aeb82-218">Requirements</span></span>

|<span data-ttu-id="aeb82-219">要件</span><span class="sxs-lookup"><span data-stu-id="aeb82-219">Requirement</span></span>| <span data-ttu-id="aeb82-220">値</span><span class="sxs-lookup"><span data-stu-id="aeb82-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="aeb82-221">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aeb82-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="aeb82-222">1.5</span><span class="sxs-lookup"><span data-stu-id="aeb82-222">1.5</span></span> |
|[<span data-ttu-id="aeb82-223">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aeb82-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="aeb82-224">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aeb82-224">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="aeb82-225">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="aeb82-225">SourceProperty: String</span></span>

<span data-ttu-id="aeb82-226">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="aeb82-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="aeb82-227">型</span><span class="sxs-lookup"><span data-stu-id="aeb82-227">Type</span></span>

*   <span data-ttu-id="aeb82-228">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="aeb82-229">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="aeb82-229">Properties:</span></span>

|<span data-ttu-id="aeb82-230">名前</span><span class="sxs-lookup"><span data-stu-id="aeb82-230">Name</span></span>| <span data-ttu-id="aeb82-231">種類</span><span class="sxs-lookup"><span data-stu-id="aeb82-231">Type</span></span>| <span data-ttu-id="aeb82-232">説明</span><span class="sxs-lookup"><span data-stu-id="aeb82-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="aeb82-233">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-233">String</span></span>|<span data-ttu-id="aeb82-234">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="aeb82-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="aeb82-235">String</span><span class="sxs-lookup"><span data-stu-id="aeb82-235">String</span></span>|<span data-ttu-id="aeb82-236">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="aeb82-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aeb82-237">要件</span><span class="sxs-lookup"><span data-stu-id="aeb82-237">Requirements</span></span>

|<span data-ttu-id="aeb82-238">要件</span><span class="sxs-lookup"><span data-stu-id="aeb82-238">Requirement</span></span>| <span data-ttu-id="aeb82-239">値</span><span class="sxs-lookup"><span data-stu-id="aeb82-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="aeb82-240">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aeb82-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="aeb82-241">1.1</span><span class="sxs-lookup"><span data-stu-id="aeb82-241">1.1</span></span>|
|[<span data-ttu-id="aeb82-242">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aeb82-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="aeb82-243">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aeb82-243">Compose or Read</span></span>|
