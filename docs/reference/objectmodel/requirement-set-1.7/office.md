---
title: Office 名前空間 - 要件セット 1.7
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: d6422e470864d5a02db37e1fef295e8cbb82a213
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067896"
---
# <a name="office"></a><span data-ttu-id="8f061-102">Office</span><span class="sxs-lookup"><span data-stu-id="8f061-102">Office</span></span>

<span data-ttu-id="8f061-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8f061-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8f061-105">要件</span><span class="sxs-lookup"><span data-stu-id="8f061-105">Requirements</span></span>

|<span data-ttu-id="8f061-106">要件</span><span class="sxs-lookup"><span data-stu-id="8f061-106">Requirement</span></span>| <span data-ttu-id="8f061-107">値</span><span class="sxs-lookup"><span data-stu-id="8f061-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f061-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8f061-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8f061-109">1.0</span><span class="sxs-lookup"><span data-stu-id="8f061-109">1.0</span></span>|
|[<span data-ttu-id="8f061-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8f061-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8f061-111">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8f061-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8f061-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="8f061-112">Members and methods</span></span>

| <span data-ttu-id="8f061-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="8f061-113">Member</span></span> | <span data-ttu-id="8f061-114">種類</span><span class="sxs-lookup"><span data-stu-id="8f061-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8f061-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="8f061-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="8f061-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="8f061-116">Member</span></span> |
| [<span data-ttu-id="8f061-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="8f061-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="8f061-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="8f061-118">Member</span></span> |
| [<span data-ttu-id="8f061-119">EventType</span><span class="sxs-lookup"><span data-stu-id="8f061-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="8f061-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="8f061-120">Member</span></span> |
| [<span data-ttu-id="8f061-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="8f061-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="8f061-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="8f061-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="8f061-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="8f061-123">Namespaces</span></span>

<span data-ttu-id="8f061-124">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8f061-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="8f061-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="8f061-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="8f061-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="8f061-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="8f061-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="8f061-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="8f061-128">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="8f061-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="8f061-129">型</span><span class="sxs-lookup"><span data-stu-id="8f061-129">Type</span></span>

*   <span data-ttu-id="8f061-130">String</span><span class="sxs-lookup"><span data-stu-id="8f061-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8f061-131">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8f061-131">Properties:</span></span>

|<span data-ttu-id="8f061-132">名前</span><span class="sxs-lookup"><span data-stu-id="8f061-132">Name</span></span>| <span data-ttu-id="8f061-133">型</span><span class="sxs-lookup"><span data-stu-id="8f061-133">Type</span></span>| <span data-ttu-id="8f061-134">説明</span><span class="sxs-lookup"><span data-stu-id="8f061-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="8f061-135">String</span><span class="sxs-lookup"><span data-stu-id="8f061-135">String</span></span>|<span data-ttu-id="8f061-136">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="8f061-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="8f061-137">String</span><span class="sxs-lookup"><span data-stu-id="8f061-137">String</span></span>|<span data-ttu-id="8f061-138">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="8f061-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8f061-139">要件</span><span class="sxs-lookup"><span data-stu-id="8f061-139">Requirements</span></span>

|<span data-ttu-id="8f061-140">要件</span><span class="sxs-lookup"><span data-stu-id="8f061-140">Requirement</span></span>| <span data-ttu-id="8f061-141">値</span><span class="sxs-lookup"><span data-stu-id="8f061-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f061-142">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8f061-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8f061-143">1.0</span><span class="sxs-lookup"><span data-stu-id="8f061-143">1.0</span></span>|
|[<span data-ttu-id="8f061-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8f061-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8f061-145">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8f061-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="8f061-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="8f061-146">CoercionType :String</span></span>

<span data-ttu-id="8f061-147">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="8f061-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8f061-148">型</span><span class="sxs-lookup"><span data-stu-id="8f061-148">Type</span></span>

*   <span data-ttu-id="8f061-149">String</span><span class="sxs-lookup"><span data-stu-id="8f061-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8f061-150">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8f061-150">Properties:</span></span>

|<span data-ttu-id="8f061-151">名前</span><span class="sxs-lookup"><span data-stu-id="8f061-151">Name</span></span>| <span data-ttu-id="8f061-152">型</span><span class="sxs-lookup"><span data-stu-id="8f061-152">Type</span></span>| <span data-ttu-id="8f061-153">説明</span><span class="sxs-lookup"><span data-stu-id="8f061-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="8f061-154">String</span><span class="sxs-lookup"><span data-stu-id="8f061-154">String</span></span>|<span data-ttu-id="8f061-155">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="8f061-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="8f061-156">String</span><span class="sxs-lookup"><span data-stu-id="8f061-156">String</span></span>|<span data-ttu-id="8f061-157">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="8f061-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8f061-158">要件</span><span class="sxs-lookup"><span data-stu-id="8f061-158">Requirements</span></span>

|<span data-ttu-id="8f061-159">要件</span><span class="sxs-lookup"><span data-stu-id="8f061-159">Requirement</span></span>| <span data-ttu-id="8f061-160">値</span><span class="sxs-lookup"><span data-stu-id="8f061-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f061-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8f061-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8f061-162">1.0</span><span class="sxs-lookup"><span data-stu-id="8f061-162">1.0</span></span>|
|[<span data-ttu-id="8f061-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8f061-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8f061-164">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8f061-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="8f061-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="8f061-165">EventType :String</span></span>

<span data-ttu-id="8f061-166">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="8f061-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="8f061-167">型</span><span class="sxs-lookup"><span data-stu-id="8f061-167">Type</span></span>

*   <span data-ttu-id="8f061-168">String</span><span class="sxs-lookup"><span data-stu-id="8f061-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8f061-169">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8f061-169">Properties:</span></span>

| <span data-ttu-id="8f061-170">名前</span><span class="sxs-lookup"><span data-stu-id="8f061-170">Name</span></span> | <span data-ttu-id="8f061-171">型</span><span class="sxs-lookup"><span data-stu-id="8f061-171">Type</span></span> | <span data-ttu-id="8f061-172">説明</span><span class="sxs-lookup"><span data-stu-id="8f061-172">Description</span></span> | <span data-ttu-id="8f061-173">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="8f061-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="8f061-174">String</span><span class="sxs-lookup"><span data-stu-id="8f061-174">String</span></span> | <span data-ttu-id="8f061-175">選択した予定または一連の予定の日付または時刻が変更された。</span><span class="sxs-lookup"><span data-stu-id="8f061-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="8f061-176">1.7</span><span class="sxs-lookup"><span data-stu-id="8f061-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="8f061-177">String</span><span class="sxs-lookup"><span data-stu-id="8f061-177">String</span></span> | <span data-ttu-id="8f061-178">作業ウィンドウがピン留めされている間、別の Outlook アイテムが選択されて表示されている。</span><span class="sxs-lookup"><span data-stu-id="8f061-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="8f061-179">1.5</span><span class="sxs-lookup"><span data-stu-id="8f061-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="8f061-180">文字列</span><span class="sxs-lookup"><span data-stu-id="8f061-180">String</span></span> | <span data-ttu-id="8f061-181">選択したアイテムまたは予定の場所の受信者リストが変更された。</span><span class="sxs-lookup"><span data-stu-id="8f061-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="8f061-182">1.7</span><span class="sxs-lookup"><span data-stu-id="8f061-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="8f061-183">文字列</span><span class="sxs-lookup"><span data-stu-id="8f061-183">String</span></span> | <span data-ttu-id="8f061-184">選択した一連の予定の定期的なパターンが変更された。</span><span class="sxs-lookup"><span data-stu-id="8f061-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="8f061-185">1.7</span><span class="sxs-lookup"><span data-stu-id="8f061-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8f061-186">要件</span><span class="sxs-lookup"><span data-stu-id="8f061-186">Requirements</span></span>

|<span data-ttu-id="8f061-187">要件</span><span class="sxs-lookup"><span data-stu-id="8f061-187">Requirement</span></span>| <span data-ttu-id="8f061-188">値</span><span class="sxs-lookup"><span data-stu-id="8f061-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f061-189">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8f061-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8f061-190">1.5</span><span class="sxs-lookup"><span data-stu-id="8f061-190">1.5</span></span> |
|[<span data-ttu-id="8f061-191">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8f061-191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8f061-192">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8f061-192">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="8f061-193">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="8f061-193">SourceProperty :String</span></span>

<span data-ttu-id="8f061-194">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="8f061-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8f061-195">型</span><span class="sxs-lookup"><span data-stu-id="8f061-195">Type</span></span>

*   <span data-ttu-id="8f061-196">String</span><span class="sxs-lookup"><span data-stu-id="8f061-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8f061-197">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8f061-197">Properties:</span></span>

|<span data-ttu-id="8f061-198">名前</span><span class="sxs-lookup"><span data-stu-id="8f061-198">Name</span></span>| <span data-ttu-id="8f061-199">型</span><span class="sxs-lookup"><span data-stu-id="8f061-199">Type</span></span>| <span data-ttu-id="8f061-200">説明</span><span class="sxs-lookup"><span data-stu-id="8f061-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="8f061-201">String</span><span class="sxs-lookup"><span data-stu-id="8f061-201">String</span></span>|<span data-ttu-id="8f061-202">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="8f061-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="8f061-203">String</span><span class="sxs-lookup"><span data-stu-id="8f061-203">String</span></span>|<span data-ttu-id="8f061-204">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="8f061-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8f061-205">要件</span><span class="sxs-lookup"><span data-stu-id="8f061-205">Requirements</span></span>

|<span data-ttu-id="8f061-206">要件</span><span class="sxs-lookup"><span data-stu-id="8f061-206">Requirement</span></span>| <span data-ttu-id="8f061-207">値</span><span class="sxs-lookup"><span data-stu-id="8f061-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f061-208">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8f061-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8f061-209">1.0</span><span class="sxs-lookup"><span data-stu-id="8f061-209">1.0</span></span>|
|[<span data-ttu-id="8f061-210">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8f061-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8f061-211">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8f061-211">Compose or Read</span></span>|
