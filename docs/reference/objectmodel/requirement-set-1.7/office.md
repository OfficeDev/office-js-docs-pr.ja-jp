---
title: Office 名前空間-要件セット1.7
description: メールボックス API 要件セット1.7 を使用した Outlook アドインで使用可能な Office 名前空間メンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 7991fd56097bbdebbfd4d4494a900626a1d3e02b
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891251"
---
# <a name="office-mailbox-requirement-set-17"></a><span data-ttu-id="fe183-103">Office (メールボックス要件セット 1.7)</span><span class="sxs-lookup"><span data-stu-id="fe183-103">Office (Mailbox requirement set 1.7)</span></span>

<span data-ttu-id="fe183-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fe183-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe183-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="fe183-106">Requirements</span></span>

|<span data-ttu-id="fe183-107">要件</span><span class="sxs-lookup"><span data-stu-id="fe183-107">Requirement</span></span>| <span data-ttu-id="fe183-108">値</span><span class="sxs-lookup"><span data-stu-id="fe183-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe183-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fe183-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="fe183-110">1.1</span><span class="sxs-lookup"><span data-stu-id="fe183-110">1.1</span></span>|
|[<span data-ttu-id="fe183-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fe183-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="fe183-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="fe183-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="fe183-113">Properties</span><span class="sxs-lookup"><span data-stu-id="fe183-113">Properties</span></span>

| <span data-ttu-id="fe183-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="fe183-114">Property</span></span> | <span data-ttu-id="fe183-115">モード</span><span class="sxs-lookup"><span data-stu-id="fe183-115">Modes</span></span> | <span data-ttu-id="fe183-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="fe183-116">Return type</span></span> | <span data-ttu-id="fe183-117">最小値</span><span class="sxs-lookup"><span data-stu-id="fe183-117">Minimum</span></span><br><span data-ttu-id="fe183-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="fe183-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="fe183-119">context</span><span class="sxs-lookup"><span data-stu-id="fe183-119">context</span></span>](office.context.md) | <span data-ttu-id="fe183-120">作成</span><span class="sxs-lookup"><span data-stu-id="fe183-120">Compose</span></span><br><span data-ttu-id="fe183-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="fe183-121">Read</span></span> | [<span data-ttu-id="fe183-122">Context</span><span class="sxs-lookup"><span data-stu-id="fe183-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="fe183-123">1.1</span><span class="sxs-lookup"><span data-stu-id="fe183-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="fe183-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="fe183-124">Enumerations</span></span>

| <span data-ttu-id="fe183-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="fe183-125">Enumeration</span></span> | <span data-ttu-id="fe183-126">モード</span><span class="sxs-lookup"><span data-stu-id="fe183-126">Modes</span></span> | <span data-ttu-id="fe183-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="fe183-127">Return type</span></span> | <span data-ttu-id="fe183-128">最小値</span><span class="sxs-lookup"><span data-stu-id="fe183-128">Minimum</span></span><br><span data-ttu-id="fe183-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="fe183-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="fe183-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="fe183-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="fe183-131">作成</span><span class="sxs-lookup"><span data-stu-id="fe183-131">Compose</span></span><br><span data-ttu-id="fe183-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="fe183-132">Read</span></span> | <span data-ttu-id="fe183-133">String</span><span class="sxs-lookup"><span data-stu-id="fe183-133">String</span></span> | [<span data-ttu-id="fe183-134">1.1</span><span class="sxs-lookup"><span data-stu-id="fe183-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="fe183-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="fe183-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="fe183-136">作成</span><span class="sxs-lookup"><span data-stu-id="fe183-136">Compose</span></span><br><span data-ttu-id="fe183-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="fe183-137">Read</span></span> | <span data-ttu-id="fe183-138">String</span><span class="sxs-lookup"><span data-stu-id="fe183-138">String</span></span> | [<span data-ttu-id="fe183-139">1.1</span><span class="sxs-lookup"><span data-stu-id="fe183-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="fe183-140">EventType</span><span class="sxs-lookup"><span data-stu-id="fe183-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="fe183-141">作成</span><span class="sxs-lookup"><span data-stu-id="fe183-141">Compose</span></span><br><span data-ttu-id="fe183-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="fe183-142">Read</span></span> | <span data-ttu-id="fe183-143">String</span><span class="sxs-lookup"><span data-stu-id="fe183-143">String</span></span> | [<span data-ttu-id="fe183-144">1.5</span><span class="sxs-lookup"><span data-stu-id="fe183-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="fe183-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="fe183-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="fe183-146">作成</span><span class="sxs-lookup"><span data-stu-id="fe183-146">Compose</span></span><br><span data-ttu-id="fe183-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="fe183-147">Read</span></span> | <span data-ttu-id="fe183-148">String</span><span class="sxs-lookup"><span data-stu-id="fe183-148">String</span></span> | [<span data-ttu-id="fe183-149">1.1</span><span class="sxs-lookup"><span data-stu-id="fe183-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="fe183-150">名前空間</span><span class="sxs-lookup"><span data-stu-id="fe183-150">Namespaces</span></span>

<span data-ttu-id="fe183-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="fe183-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="fe183-152">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="fe183-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="fe183-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="fe183-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="fe183-154">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="fe183-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="fe183-155">型</span><span class="sxs-lookup"><span data-stu-id="fe183-155">Type</span></span>

*   <span data-ttu-id="fe183-156">String</span><span class="sxs-lookup"><span data-stu-id="fe183-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fe183-157">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="fe183-157">Properties:</span></span>

|<span data-ttu-id="fe183-158">名前</span><span class="sxs-lookup"><span data-stu-id="fe183-158">Name</span></span>| <span data-ttu-id="fe183-159">種類</span><span class="sxs-lookup"><span data-stu-id="fe183-159">Type</span></span>| <span data-ttu-id="fe183-160">説明</span><span class="sxs-lookup"><span data-stu-id="fe183-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="fe183-161">String</span><span class="sxs-lookup"><span data-stu-id="fe183-161">String</span></span>|<span data-ttu-id="fe183-162">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="fe183-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="fe183-163">String</span><span class="sxs-lookup"><span data-stu-id="fe183-163">String</span></span>|<span data-ttu-id="fe183-164">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="fe183-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fe183-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="fe183-165">Requirements</span></span>

|<span data-ttu-id="fe183-166">要件</span><span class="sxs-lookup"><span data-stu-id="fe183-166">Requirement</span></span>| <span data-ttu-id="fe183-167">値</span><span class="sxs-lookup"><span data-stu-id="fe183-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe183-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fe183-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="fe183-169">1.1</span><span class="sxs-lookup"><span data-stu-id="fe183-169">1.1</span></span>|
|[<span data-ttu-id="fe183-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fe183-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="fe183-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="fe183-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="fe183-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="fe183-172">CoercionType: String</span></span>

<span data-ttu-id="fe183-173">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="fe183-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="fe183-174">型</span><span class="sxs-lookup"><span data-stu-id="fe183-174">Type</span></span>

*   <span data-ttu-id="fe183-175">String</span><span class="sxs-lookup"><span data-stu-id="fe183-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fe183-176">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="fe183-176">Properties:</span></span>

|<span data-ttu-id="fe183-177">名前</span><span class="sxs-lookup"><span data-stu-id="fe183-177">Name</span></span>| <span data-ttu-id="fe183-178">種類</span><span class="sxs-lookup"><span data-stu-id="fe183-178">Type</span></span>| <span data-ttu-id="fe183-179">説明</span><span class="sxs-lookup"><span data-stu-id="fe183-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="fe183-180">String</span><span class="sxs-lookup"><span data-stu-id="fe183-180">String</span></span>|<span data-ttu-id="fe183-181">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="fe183-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="fe183-182">String</span><span class="sxs-lookup"><span data-stu-id="fe183-182">String</span></span>|<span data-ttu-id="fe183-183">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="fe183-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fe183-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="fe183-184">Requirements</span></span>

|<span data-ttu-id="fe183-185">要件</span><span class="sxs-lookup"><span data-stu-id="fe183-185">Requirement</span></span>| <span data-ttu-id="fe183-186">値</span><span class="sxs-lookup"><span data-stu-id="fe183-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe183-187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fe183-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="fe183-188">1.1</span><span class="sxs-lookup"><span data-stu-id="fe183-188">1.1</span></span>|
|[<span data-ttu-id="fe183-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fe183-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="fe183-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="fe183-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="fe183-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="fe183-191">EventType: String</span></span>

<span data-ttu-id="fe183-192">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="fe183-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="fe183-193">型</span><span class="sxs-lookup"><span data-stu-id="fe183-193">Type</span></span>

*   <span data-ttu-id="fe183-194">String</span><span class="sxs-lookup"><span data-stu-id="fe183-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fe183-195">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="fe183-195">Properties:</span></span>

| <span data-ttu-id="fe183-196">名前</span><span class="sxs-lookup"><span data-stu-id="fe183-196">Name</span></span> | <span data-ttu-id="fe183-197">種類</span><span class="sxs-lookup"><span data-stu-id="fe183-197">Type</span></span> | <span data-ttu-id="fe183-198">説明</span><span class="sxs-lookup"><span data-stu-id="fe183-198">Description</span></span> | <span data-ttu-id="fe183-199">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="fe183-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="fe183-200">String</span><span class="sxs-lookup"><span data-stu-id="fe183-200">String</span></span> | <span data-ttu-id="fe183-201">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="fe183-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="fe183-202">1.7</span><span class="sxs-lookup"><span data-stu-id="fe183-202">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="fe183-203">String</span><span class="sxs-lookup"><span data-stu-id="fe183-203">String</span></span> | <span data-ttu-id="fe183-204">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="fe183-204">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="fe183-205">1.5</span><span class="sxs-lookup"><span data-stu-id="fe183-205">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="fe183-206">String</span><span class="sxs-lookup"><span data-stu-id="fe183-206">String</span></span> | <span data-ttu-id="fe183-207">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="fe183-207">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="fe183-208">1.7</span><span class="sxs-lookup"><span data-stu-id="fe183-208">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="fe183-209">String</span><span class="sxs-lookup"><span data-stu-id="fe183-209">String</span></span> | <span data-ttu-id="fe183-210">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="fe183-210">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="fe183-211">1.7</span><span class="sxs-lookup"><span data-stu-id="fe183-211">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fe183-212">Requirements</span><span class="sxs-lookup"><span data-stu-id="fe183-212">Requirements</span></span>

|<span data-ttu-id="fe183-213">要件</span><span class="sxs-lookup"><span data-stu-id="fe183-213">Requirement</span></span>| <span data-ttu-id="fe183-214">値</span><span class="sxs-lookup"><span data-stu-id="fe183-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe183-215">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fe183-215">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="fe183-216">1.5</span><span class="sxs-lookup"><span data-stu-id="fe183-216">1.5</span></span> |
|[<span data-ttu-id="fe183-217">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fe183-217">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="fe183-218">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="fe183-218">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="fe183-219">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="fe183-219">SourceProperty: String</span></span>

<span data-ttu-id="fe183-220">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="fe183-220">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="fe183-221">型</span><span class="sxs-lookup"><span data-stu-id="fe183-221">Type</span></span>

*   <span data-ttu-id="fe183-222">String</span><span class="sxs-lookup"><span data-stu-id="fe183-222">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fe183-223">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="fe183-223">Properties:</span></span>

|<span data-ttu-id="fe183-224">名前</span><span class="sxs-lookup"><span data-stu-id="fe183-224">Name</span></span>| <span data-ttu-id="fe183-225">種類</span><span class="sxs-lookup"><span data-stu-id="fe183-225">Type</span></span>| <span data-ttu-id="fe183-226">説明</span><span class="sxs-lookup"><span data-stu-id="fe183-226">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="fe183-227">String</span><span class="sxs-lookup"><span data-stu-id="fe183-227">String</span></span>|<span data-ttu-id="fe183-228">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="fe183-228">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="fe183-229">String</span><span class="sxs-lookup"><span data-stu-id="fe183-229">String</span></span>|<span data-ttu-id="fe183-230">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="fe183-230">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fe183-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="fe183-231">Requirements</span></span>

|<span data-ttu-id="fe183-232">要件</span><span class="sxs-lookup"><span data-stu-id="fe183-232">Requirement</span></span>| <span data-ttu-id="fe183-233">値</span><span class="sxs-lookup"><span data-stu-id="fe183-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe183-234">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fe183-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="fe183-235">1.1</span><span class="sxs-lookup"><span data-stu-id="fe183-235">1.1</span></span>|
|[<span data-ttu-id="fe183-236">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fe183-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="fe183-237">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="fe183-237">Compose or Read</span></span>|
