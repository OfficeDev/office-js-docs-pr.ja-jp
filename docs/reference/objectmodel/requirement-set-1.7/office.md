---
title: Office 名前空間-要件セット1.7
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 523df189b28fc568ac32e8d17d4a226b52cbd23c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451620"
---
# <a name="office"></a><span data-ttu-id="8afe5-102">Office</span><span class="sxs-lookup"><span data-stu-id="8afe5-102">Office</span></span>

<span data-ttu-id="8afe5-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8afe5-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8afe5-105">要件</span><span class="sxs-lookup"><span data-stu-id="8afe5-105">Requirements</span></span>

|<span data-ttu-id="8afe5-106">要件</span><span class="sxs-lookup"><span data-stu-id="8afe5-106">Requirement</span></span>| <span data-ttu-id="8afe5-107">値</span><span class="sxs-lookup"><span data-stu-id="8afe5-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8afe5-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8afe5-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8afe5-109">1.0</span><span class="sxs-lookup"><span data-stu-id="8afe5-109">1.0</span></span>|
|[<span data-ttu-id="8afe5-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8afe5-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8afe5-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8afe5-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8afe5-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="8afe5-112">Members and methods</span></span>

| <span data-ttu-id="8afe5-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="8afe5-113">Member</span></span> | <span data-ttu-id="8afe5-114">種類</span><span class="sxs-lookup"><span data-stu-id="8afe5-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8afe5-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="8afe5-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="8afe5-116">Member</span><span class="sxs-lookup"><span data-stu-id="8afe5-116">Member</span></span> |
| [<span data-ttu-id="8afe5-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="8afe5-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="8afe5-118">Member</span><span class="sxs-lookup"><span data-stu-id="8afe5-118">Member</span></span> |
| [<span data-ttu-id="8afe5-119">EventType</span><span class="sxs-lookup"><span data-stu-id="8afe5-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="8afe5-120">Member</span><span class="sxs-lookup"><span data-stu-id="8afe5-120">Member</span></span> |
| [<span data-ttu-id="8afe5-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="8afe5-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="8afe5-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="8afe5-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="8afe5-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="8afe5-123">Namespaces</span></span>

<span data-ttu-id="8afe5-124">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8afe5-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="8afe5-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="8afe5-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="8afe5-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="8afe5-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="8afe5-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="8afe5-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="8afe5-128">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="8afe5-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="8afe5-129">型</span><span class="sxs-lookup"><span data-stu-id="8afe5-129">Type</span></span>

*   <span data-ttu-id="8afe5-130">String</span><span class="sxs-lookup"><span data-stu-id="8afe5-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8afe5-131">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8afe5-131">Properties:</span></span>

|<span data-ttu-id="8afe5-132">名前</span><span class="sxs-lookup"><span data-stu-id="8afe5-132">Name</span></span>| <span data-ttu-id="8afe5-133">種類</span><span class="sxs-lookup"><span data-stu-id="8afe5-133">Type</span></span>| <span data-ttu-id="8afe5-134">説明</span><span class="sxs-lookup"><span data-stu-id="8afe5-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="8afe5-135">String</span><span class="sxs-lookup"><span data-stu-id="8afe5-135">String</span></span>|<span data-ttu-id="8afe5-136">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="8afe5-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="8afe5-137">String</span><span class="sxs-lookup"><span data-stu-id="8afe5-137">String</span></span>|<span data-ttu-id="8afe5-138">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="8afe5-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8afe5-139">要件</span><span class="sxs-lookup"><span data-stu-id="8afe5-139">Requirements</span></span>

|<span data-ttu-id="8afe5-140">要件</span><span class="sxs-lookup"><span data-stu-id="8afe5-140">Requirement</span></span>| <span data-ttu-id="8afe5-141">値</span><span class="sxs-lookup"><span data-stu-id="8afe5-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="8afe5-142">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8afe5-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8afe5-143">1.0</span><span class="sxs-lookup"><span data-stu-id="8afe5-143">1.0</span></span>|
|[<span data-ttu-id="8afe5-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8afe5-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8afe5-145">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8afe5-145">Compose or Read</span></span>|

---
---

####  <a name="coerciontype-string"></a><span data-ttu-id="8afe5-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="8afe5-146">CoercionType :String</span></span>

<span data-ttu-id="8afe5-147">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="8afe5-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8afe5-148">型</span><span class="sxs-lookup"><span data-stu-id="8afe5-148">Type</span></span>

*   <span data-ttu-id="8afe5-149">String</span><span class="sxs-lookup"><span data-stu-id="8afe5-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8afe5-150">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8afe5-150">Properties:</span></span>

|<span data-ttu-id="8afe5-151">名前</span><span class="sxs-lookup"><span data-stu-id="8afe5-151">Name</span></span>| <span data-ttu-id="8afe5-152">種類</span><span class="sxs-lookup"><span data-stu-id="8afe5-152">Type</span></span>| <span data-ttu-id="8afe5-153">説明</span><span class="sxs-lookup"><span data-stu-id="8afe5-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="8afe5-154">String</span><span class="sxs-lookup"><span data-stu-id="8afe5-154">String</span></span>|<span data-ttu-id="8afe5-155">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="8afe5-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="8afe5-156">String</span><span class="sxs-lookup"><span data-stu-id="8afe5-156">String</span></span>|<span data-ttu-id="8afe5-157">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="8afe5-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8afe5-158">要件</span><span class="sxs-lookup"><span data-stu-id="8afe5-158">Requirements</span></span>

|<span data-ttu-id="8afe5-159">要件</span><span class="sxs-lookup"><span data-stu-id="8afe5-159">Requirement</span></span>| <span data-ttu-id="8afe5-160">値</span><span class="sxs-lookup"><span data-stu-id="8afe5-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="8afe5-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8afe5-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8afe5-162">1.0</span><span class="sxs-lookup"><span data-stu-id="8afe5-162">1.0</span></span>|
|[<span data-ttu-id="8afe5-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8afe5-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8afe5-164">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8afe5-164">Compose or Read</span></span>|

---
---

####  <a name="eventtype-string"></a><span data-ttu-id="8afe5-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="8afe5-165">EventType :String</span></span>

<span data-ttu-id="8afe5-166">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="8afe5-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="8afe5-167">型</span><span class="sxs-lookup"><span data-stu-id="8afe5-167">Type</span></span>

*   <span data-ttu-id="8afe5-168">String</span><span class="sxs-lookup"><span data-stu-id="8afe5-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8afe5-169">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8afe5-169">Properties:</span></span>

| <span data-ttu-id="8afe5-170">名前</span><span class="sxs-lookup"><span data-stu-id="8afe5-170">Name</span></span> | <span data-ttu-id="8afe5-171">種類</span><span class="sxs-lookup"><span data-stu-id="8afe5-171">Type</span></span> | <span data-ttu-id="8afe5-172">説明</span><span class="sxs-lookup"><span data-stu-id="8afe5-172">Description</span></span> | <span data-ttu-id="8afe5-173">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="8afe5-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="8afe5-174">String</span><span class="sxs-lookup"><span data-stu-id="8afe5-174">String</span></span> | <span data-ttu-id="8afe5-175">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="8afe5-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="8afe5-176">1.7</span><span class="sxs-lookup"><span data-stu-id="8afe5-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="8afe5-177">String</span><span class="sxs-lookup"><span data-stu-id="8afe5-177">String</span></span> | <span data-ttu-id="8afe5-178">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="8afe5-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="8afe5-179">1.5</span><span class="sxs-lookup"><span data-stu-id="8afe5-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="8afe5-180">String</span><span class="sxs-lookup"><span data-stu-id="8afe5-180">String</span></span> | <span data-ttu-id="8afe5-181">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="8afe5-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="8afe5-182">1.7</span><span class="sxs-lookup"><span data-stu-id="8afe5-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="8afe5-183">String</span><span class="sxs-lookup"><span data-stu-id="8afe5-183">String</span></span> | <span data-ttu-id="8afe5-184">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="8afe5-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="8afe5-185">1.7</span><span class="sxs-lookup"><span data-stu-id="8afe5-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8afe5-186">要件</span><span class="sxs-lookup"><span data-stu-id="8afe5-186">Requirements</span></span>

|<span data-ttu-id="8afe5-187">要件</span><span class="sxs-lookup"><span data-stu-id="8afe5-187">Requirement</span></span>| <span data-ttu-id="8afe5-188">値</span><span class="sxs-lookup"><span data-stu-id="8afe5-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="8afe5-189">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8afe5-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8afe5-190">1.5</span><span class="sxs-lookup"><span data-stu-id="8afe5-190">1.5</span></span> |
|[<span data-ttu-id="8afe5-191">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8afe5-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8afe5-192">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8afe5-192">Compose or Read</span></span> |

---
---

####  <a name="sourceproperty-string"></a><span data-ttu-id="8afe5-193">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="8afe5-193">SourceProperty :String</span></span>

<span data-ttu-id="8afe5-194">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="8afe5-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8afe5-195">型</span><span class="sxs-lookup"><span data-stu-id="8afe5-195">Type</span></span>

*   <span data-ttu-id="8afe5-196">String</span><span class="sxs-lookup"><span data-stu-id="8afe5-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8afe5-197">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8afe5-197">Properties:</span></span>

|<span data-ttu-id="8afe5-198">名前</span><span class="sxs-lookup"><span data-stu-id="8afe5-198">Name</span></span>| <span data-ttu-id="8afe5-199">種類</span><span class="sxs-lookup"><span data-stu-id="8afe5-199">Type</span></span>| <span data-ttu-id="8afe5-200">説明</span><span class="sxs-lookup"><span data-stu-id="8afe5-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="8afe5-201">String</span><span class="sxs-lookup"><span data-stu-id="8afe5-201">String</span></span>|<span data-ttu-id="8afe5-202">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="8afe5-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="8afe5-203">String</span><span class="sxs-lookup"><span data-stu-id="8afe5-203">String</span></span>|<span data-ttu-id="8afe5-204">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="8afe5-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8afe5-205">要件</span><span class="sxs-lookup"><span data-stu-id="8afe5-205">Requirements</span></span>

|<span data-ttu-id="8afe5-206">要件</span><span class="sxs-lookup"><span data-stu-id="8afe5-206">Requirement</span></span>| <span data-ttu-id="8afe5-207">値</span><span class="sxs-lookup"><span data-stu-id="8afe5-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="8afe5-208">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8afe5-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8afe5-209">1.0</span><span class="sxs-lookup"><span data-stu-id="8afe5-209">1.0</span></span>|
|[<span data-ttu-id="8afe5-210">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8afe5-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8afe5-211">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8afe5-211">Compose or Read</span></span>|
