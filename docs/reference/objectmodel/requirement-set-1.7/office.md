---
title: Office 名前空間 - 要件セット 1.7
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 6afaca31dd941b9c6a4b23fa08018de51278cbbd
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457741"
---
# <a name="office"></a><span data-ttu-id="52816-102">Office</span><span class="sxs-lookup"><span data-stu-id="52816-102">Office</span></span>

<span data-ttu-id="52816-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="52816-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="52816-105">要件</span><span class="sxs-lookup"><span data-stu-id="52816-105">Requirements</span></span>

|<span data-ttu-id="52816-106">要件</span><span class="sxs-lookup"><span data-stu-id="52816-106">Requirement</span></span>| <span data-ttu-id="52816-107">値</span><span class="sxs-lookup"><span data-stu-id="52816-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="52816-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="52816-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="52816-109">1.0</span><span class="sxs-lookup"><span data-stu-id="52816-109">1.0</span></span>|
|[<span data-ttu-id="52816-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="52816-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="52816-111">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="52816-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="52816-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="52816-112">Members and methods</span></span>

| <span data-ttu-id="52816-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="52816-113">Member</span></span> | <span data-ttu-id="52816-114">種類</span><span class="sxs-lookup"><span data-stu-id="52816-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="52816-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="52816-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="52816-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="52816-116">Member</span></span> |
| [<span data-ttu-id="52816-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="52816-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="52816-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="52816-118">Member</span></span> |
| [<span data-ttu-id="52816-119">EventType</span><span class="sxs-lookup"><span data-stu-id="52816-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="52816-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="52816-120">Member</span></span> |
| [<span data-ttu-id="52816-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="52816-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="52816-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="52816-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="52816-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="52816-123">Namespaces</span></span>

<span data-ttu-id="52816-124">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="52816-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="52816-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="52816-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="52816-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="52816-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="52816-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="52816-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="52816-128">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="52816-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="52816-129">型:</span><span class="sxs-lookup"><span data-stu-id="52816-129">Type:</span></span>

*   <span data-ttu-id="52816-130">String</span><span class="sxs-lookup"><span data-stu-id="52816-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="52816-131">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="52816-131">Properties:</span></span>

|<span data-ttu-id="52816-132">名前</span><span class="sxs-lookup"><span data-stu-id="52816-132">Name</span></span>| <span data-ttu-id="52816-133">型</span><span class="sxs-lookup"><span data-stu-id="52816-133">Type</span></span>| <span data-ttu-id="52816-134">説明</span><span class="sxs-lookup"><span data-stu-id="52816-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="52816-135">String</span><span class="sxs-lookup"><span data-stu-id="52816-135">String</span></span>|<span data-ttu-id="52816-136">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="52816-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="52816-137">String</span><span class="sxs-lookup"><span data-stu-id="52816-137">String</span></span>|<span data-ttu-id="52816-138">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="52816-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="52816-139">要件</span><span class="sxs-lookup"><span data-stu-id="52816-139">Requirements</span></span>

|<span data-ttu-id="52816-140">要件</span><span class="sxs-lookup"><span data-stu-id="52816-140">Requirement</span></span>| <span data-ttu-id="52816-141">値</span><span class="sxs-lookup"><span data-stu-id="52816-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="52816-142">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="52816-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="52816-143">1.0</span><span class="sxs-lookup"><span data-stu-id="52816-143">1.0</span></span>|
|[<span data-ttu-id="52816-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="52816-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="52816-145">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="52816-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="52816-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="52816-146">CoercionType :String</span></span>

<span data-ttu-id="52816-147">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="52816-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="52816-148">型:</span><span class="sxs-lookup"><span data-stu-id="52816-148">Type:</span></span>

*   <span data-ttu-id="52816-149">String</span><span class="sxs-lookup"><span data-stu-id="52816-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="52816-150">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="52816-150">Properties:</span></span>

|<span data-ttu-id="52816-151">名前</span><span class="sxs-lookup"><span data-stu-id="52816-151">Name</span></span>| <span data-ttu-id="52816-152">型</span><span class="sxs-lookup"><span data-stu-id="52816-152">Type</span></span>| <span data-ttu-id="52816-153">説明</span><span class="sxs-lookup"><span data-stu-id="52816-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="52816-154">String</span><span class="sxs-lookup"><span data-stu-id="52816-154">String</span></span>|<span data-ttu-id="52816-155">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="52816-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="52816-156">String</span><span class="sxs-lookup"><span data-stu-id="52816-156">String</span></span>|<span data-ttu-id="52816-157">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="52816-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="52816-158">要件</span><span class="sxs-lookup"><span data-stu-id="52816-158">Requirements</span></span>

|<span data-ttu-id="52816-159">要件</span><span class="sxs-lookup"><span data-stu-id="52816-159">Requirement</span></span>| <span data-ttu-id="52816-160">値</span><span class="sxs-lookup"><span data-stu-id="52816-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="52816-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="52816-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="52816-162">1.0</span><span class="sxs-lookup"><span data-stu-id="52816-162">1.0</span></span>|
|[<span data-ttu-id="52816-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="52816-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="52816-164">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="52816-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="52816-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="52816-165">EventType :String</span></span>

<span data-ttu-id="52816-166">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="52816-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="52816-167">型:</span><span class="sxs-lookup"><span data-stu-id="52816-167">Type:</span></span>

*   <span data-ttu-id="52816-168">String</span><span class="sxs-lookup"><span data-stu-id="52816-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="52816-169">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="52816-169">Properties:</span></span>

| <span data-ttu-id="52816-170">名前</span><span class="sxs-lookup"><span data-stu-id="52816-170">Name</span></span> | <span data-ttu-id="52816-171">型</span><span class="sxs-lookup"><span data-stu-id="52816-171">Type</span></span> | <span data-ttu-id="52816-172">説明</span><span class="sxs-lookup"><span data-stu-id="52816-172">Description</span></span> | <span data-ttu-id="52816-173">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="52816-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="52816-174">文字列</span><span class="sxs-lookup"><span data-stu-id="52816-174">String</span></span> | <span data-ttu-id="52816-175">選択した予定または一連の予定の日付または時刻が変更された。</span><span class="sxs-lookup"><span data-stu-id="52816-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="52816-176">1.7</span><span class="sxs-lookup"><span data-stu-id="52816-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="52816-177">文字列</span><span class="sxs-lookup"><span data-stu-id="52816-177">String</span></span> | <span data-ttu-id="52816-178">作業ウィンドウがピン留めされている間、別の Outlook アイテムが選択されて表示されている。</span><span class="sxs-lookup"><span data-stu-id="52816-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="52816-179">1.5</span><span class="sxs-lookup"><span data-stu-id="52816-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="52816-180">文字列</span><span class="sxs-lookup"><span data-stu-id="52816-180">String</span></span> | <span data-ttu-id="52816-181">選択したアイテムまたは予定の場所の受信者リストが変更された。</span><span class="sxs-lookup"><span data-stu-id="52816-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="52816-182">1.7</span><span class="sxs-lookup"><span data-stu-id="52816-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="52816-183">文字列</span><span class="sxs-lookup"><span data-stu-id="52816-183">String</span></span> | <span data-ttu-id="52816-184">選択した一連の予定の定期的なパターンが変更された。</span><span class="sxs-lookup"><span data-stu-id="52816-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="52816-185">1.7</span><span class="sxs-lookup"><span data-stu-id="52816-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="52816-186">要件</span><span class="sxs-lookup"><span data-stu-id="52816-186">Requirements</span></span>

|<span data-ttu-id="52816-187">要件</span><span class="sxs-lookup"><span data-stu-id="52816-187">Requirement</span></span>| <span data-ttu-id="52816-188">値</span><span class="sxs-lookup"><span data-stu-id="52816-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="52816-189">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="52816-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="52816-190">1.5</span><span class="sxs-lookup"><span data-stu-id="52816-190">1.5</span></span> |
|[<span data-ttu-id="52816-191">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="52816-191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="52816-192">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="52816-192">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="52816-193">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="52816-193">SourceProperty :String</span></span>

<span data-ttu-id="52816-194">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="52816-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="52816-195">型:</span><span class="sxs-lookup"><span data-stu-id="52816-195">Type:</span></span>

*   <span data-ttu-id="52816-196">String</span><span class="sxs-lookup"><span data-stu-id="52816-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="52816-197">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="52816-197">Properties:</span></span>

|<span data-ttu-id="52816-198">名前</span><span class="sxs-lookup"><span data-stu-id="52816-198">Name</span></span>| <span data-ttu-id="52816-199">型</span><span class="sxs-lookup"><span data-stu-id="52816-199">Type</span></span>| <span data-ttu-id="52816-200">説明</span><span class="sxs-lookup"><span data-stu-id="52816-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="52816-201">String</span><span class="sxs-lookup"><span data-stu-id="52816-201">String</span></span>|<span data-ttu-id="52816-202">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="52816-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="52816-203">String</span><span class="sxs-lookup"><span data-stu-id="52816-203">String</span></span>|<span data-ttu-id="52816-204">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="52816-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="52816-205">要件</span><span class="sxs-lookup"><span data-stu-id="52816-205">Requirements</span></span>

|<span data-ttu-id="52816-206">要件</span><span class="sxs-lookup"><span data-stu-id="52816-206">Requirement</span></span>| <span data-ttu-id="52816-207">値</span><span class="sxs-lookup"><span data-stu-id="52816-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="52816-208">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="52816-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="52816-209">1.0</span><span class="sxs-lookup"><span data-stu-id="52816-209">1.0</span></span>|
|[<span data-ttu-id="52816-210">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="52816-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="52816-211">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="52816-211">Compose or read</span></span>|