---
title: Office 名前空間-要件セット1.7
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 8d22ce8400916dffe12a15bba35f70ceca4db510
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695869"
---
# <a name="office"></a><span data-ttu-id="82606-102">Office</span><span class="sxs-lookup"><span data-stu-id="82606-102">Office</span></span>

<span data-ttu-id="82606-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="82606-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="82606-105">要件</span><span class="sxs-lookup"><span data-stu-id="82606-105">Requirements</span></span>

|<span data-ttu-id="82606-106">要件</span><span class="sxs-lookup"><span data-stu-id="82606-106">Requirement</span></span>| <span data-ttu-id="82606-107">値</span><span class="sxs-lookup"><span data-stu-id="82606-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="82606-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="82606-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="82606-109">1.0</span><span class="sxs-lookup"><span data-stu-id="82606-109">1.0</span></span>|
|[<span data-ttu-id="82606-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="82606-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="82606-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="82606-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="82606-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="82606-112">Members and methods</span></span>

| <span data-ttu-id="82606-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="82606-113">Member</span></span> | <span data-ttu-id="82606-114">種類</span><span class="sxs-lookup"><span data-stu-id="82606-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="82606-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="82606-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="82606-116">Member</span><span class="sxs-lookup"><span data-stu-id="82606-116">Member</span></span> |
| [<span data-ttu-id="82606-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="82606-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="82606-118">Member</span><span class="sxs-lookup"><span data-stu-id="82606-118">Member</span></span> |
| [<span data-ttu-id="82606-119">EventType</span><span class="sxs-lookup"><span data-stu-id="82606-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="82606-120">Member</span><span class="sxs-lookup"><span data-stu-id="82606-120">Member</span></span> |
| [<span data-ttu-id="82606-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="82606-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="82606-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="82606-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="82606-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="82606-123">Namespaces</span></span>

<span data-ttu-id="82606-124">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="82606-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="82606-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.7):、、、、、、などのさまざま`ItemType`な`EntityType`列挙`AttachmentType` `RecipientType` `ResponseType`値が含まれ`ItemNotificationMessageType`ています。</span><span class="sxs-lookup"><span data-stu-id="82606-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.7): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="82606-126">Members</span><span class="sxs-lookup"><span data-stu-id="82606-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="82606-127">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="82606-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="82606-128">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="82606-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="82606-129">型</span><span class="sxs-lookup"><span data-stu-id="82606-129">Type</span></span>

*   <span data-ttu-id="82606-130">String</span><span class="sxs-lookup"><span data-stu-id="82606-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="82606-131">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="82606-131">Properties:</span></span>

|<span data-ttu-id="82606-132">名前</span><span class="sxs-lookup"><span data-stu-id="82606-132">Name</span></span>| <span data-ttu-id="82606-133">種類</span><span class="sxs-lookup"><span data-stu-id="82606-133">Type</span></span>| <span data-ttu-id="82606-134">説明</span><span class="sxs-lookup"><span data-stu-id="82606-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="82606-135">String</span><span class="sxs-lookup"><span data-stu-id="82606-135">String</span></span>|<span data-ttu-id="82606-136">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="82606-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="82606-137">String</span><span class="sxs-lookup"><span data-stu-id="82606-137">String</span></span>|<span data-ttu-id="82606-138">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="82606-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="82606-139">要件</span><span class="sxs-lookup"><span data-stu-id="82606-139">Requirements</span></span>

|<span data-ttu-id="82606-140">要件</span><span class="sxs-lookup"><span data-stu-id="82606-140">Requirement</span></span>| <span data-ttu-id="82606-141">値</span><span class="sxs-lookup"><span data-stu-id="82606-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="82606-142">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="82606-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="82606-143">1.0</span><span class="sxs-lookup"><span data-stu-id="82606-143">1.0</span></span>|
|[<span data-ttu-id="82606-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="82606-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="82606-145">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="82606-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="82606-146">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="82606-146">CoercionType: String</span></span>

<span data-ttu-id="82606-147">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="82606-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="82606-148">型</span><span class="sxs-lookup"><span data-stu-id="82606-148">Type</span></span>

*   <span data-ttu-id="82606-149">String</span><span class="sxs-lookup"><span data-stu-id="82606-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="82606-150">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="82606-150">Properties:</span></span>

|<span data-ttu-id="82606-151">名前</span><span class="sxs-lookup"><span data-stu-id="82606-151">Name</span></span>| <span data-ttu-id="82606-152">種類</span><span class="sxs-lookup"><span data-stu-id="82606-152">Type</span></span>| <span data-ttu-id="82606-153">説明</span><span class="sxs-lookup"><span data-stu-id="82606-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="82606-154">String</span><span class="sxs-lookup"><span data-stu-id="82606-154">String</span></span>|<span data-ttu-id="82606-155">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="82606-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="82606-156">String</span><span class="sxs-lookup"><span data-stu-id="82606-156">String</span></span>|<span data-ttu-id="82606-157">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="82606-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="82606-158">要件</span><span class="sxs-lookup"><span data-stu-id="82606-158">Requirements</span></span>

|<span data-ttu-id="82606-159">要件</span><span class="sxs-lookup"><span data-stu-id="82606-159">Requirement</span></span>| <span data-ttu-id="82606-160">値</span><span class="sxs-lookup"><span data-stu-id="82606-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="82606-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="82606-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="82606-162">1.0</span><span class="sxs-lookup"><span data-stu-id="82606-162">1.0</span></span>|
|[<span data-ttu-id="82606-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="82606-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="82606-164">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="82606-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="82606-165">EventType: String</span><span class="sxs-lookup"><span data-stu-id="82606-165">EventType: String</span></span>

<span data-ttu-id="82606-166">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="82606-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="82606-167">型</span><span class="sxs-lookup"><span data-stu-id="82606-167">Type</span></span>

*   <span data-ttu-id="82606-168">String</span><span class="sxs-lookup"><span data-stu-id="82606-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="82606-169">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="82606-169">Properties:</span></span>

| <span data-ttu-id="82606-170">名前</span><span class="sxs-lookup"><span data-stu-id="82606-170">Name</span></span> | <span data-ttu-id="82606-171">種類</span><span class="sxs-lookup"><span data-stu-id="82606-171">Type</span></span> | <span data-ttu-id="82606-172">説明</span><span class="sxs-lookup"><span data-stu-id="82606-172">Description</span></span> | <span data-ttu-id="82606-173">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="82606-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="82606-174">String</span><span class="sxs-lookup"><span data-stu-id="82606-174">String</span></span> | <span data-ttu-id="82606-175">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="82606-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="82606-176">1.7</span><span class="sxs-lookup"><span data-stu-id="82606-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="82606-177">String</span><span class="sxs-lookup"><span data-stu-id="82606-177">String</span></span> | <span data-ttu-id="82606-178">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="82606-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="82606-179">1.5</span><span class="sxs-lookup"><span data-stu-id="82606-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="82606-180">String</span><span class="sxs-lookup"><span data-stu-id="82606-180">String</span></span> | <span data-ttu-id="82606-181">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="82606-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="82606-182">1.7</span><span class="sxs-lookup"><span data-stu-id="82606-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="82606-183">String</span><span class="sxs-lookup"><span data-stu-id="82606-183">String</span></span> | <span data-ttu-id="82606-184">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="82606-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="82606-185">1.7</span><span class="sxs-lookup"><span data-stu-id="82606-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="82606-186">要件</span><span class="sxs-lookup"><span data-stu-id="82606-186">Requirements</span></span>

|<span data-ttu-id="82606-187">要件</span><span class="sxs-lookup"><span data-stu-id="82606-187">Requirement</span></span>| <span data-ttu-id="82606-188">値</span><span class="sxs-lookup"><span data-stu-id="82606-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="82606-189">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="82606-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="82606-190">1.5</span><span class="sxs-lookup"><span data-stu-id="82606-190">1.5</span></span> |
|[<span data-ttu-id="82606-191">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="82606-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="82606-192">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="82606-192">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="82606-193">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="82606-193">SourceProperty: String</span></span>

<span data-ttu-id="82606-194">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="82606-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="82606-195">型</span><span class="sxs-lookup"><span data-stu-id="82606-195">Type</span></span>

*   <span data-ttu-id="82606-196">String</span><span class="sxs-lookup"><span data-stu-id="82606-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="82606-197">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="82606-197">Properties:</span></span>

|<span data-ttu-id="82606-198">名前</span><span class="sxs-lookup"><span data-stu-id="82606-198">Name</span></span>| <span data-ttu-id="82606-199">種類</span><span class="sxs-lookup"><span data-stu-id="82606-199">Type</span></span>| <span data-ttu-id="82606-200">説明</span><span class="sxs-lookup"><span data-stu-id="82606-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="82606-201">String</span><span class="sxs-lookup"><span data-stu-id="82606-201">String</span></span>|<span data-ttu-id="82606-202">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="82606-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="82606-203">String</span><span class="sxs-lookup"><span data-stu-id="82606-203">String</span></span>|<span data-ttu-id="82606-204">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="82606-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="82606-205">要件</span><span class="sxs-lookup"><span data-stu-id="82606-205">Requirements</span></span>

|<span data-ttu-id="82606-206">要件</span><span class="sxs-lookup"><span data-stu-id="82606-206">Requirement</span></span>| <span data-ttu-id="82606-207">値</span><span class="sxs-lookup"><span data-stu-id="82606-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="82606-208">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="82606-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="82606-209">1.0</span><span class="sxs-lookup"><span data-stu-id="82606-209">1.0</span></span>|
|[<span data-ttu-id="82606-210">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="82606-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="82606-211">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="82606-211">Compose or Read</span></span>|
