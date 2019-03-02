---
title: Office 名前空間 - プレビュー要件セット
description: ''
ms.date: 02/26/2019
localization_priority: Normal
ms.openlocfilehash: 7b27963a85f1dcdaa6f269fce242c45bf1bdd146
ms.sourcegitcommit: f7f3d38ae4430e2218bf0abe7bb2976108de3579
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/01/2019
ms.locfileid: "30359234"
---
# <a name="office"></a><span data-ttu-id="047c8-102">Office</span><span class="sxs-lookup"><span data-stu-id="047c8-102">Office</span></span>

<span data-ttu-id="047c8-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="047c8-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="047c8-105">要件</span><span class="sxs-lookup"><span data-stu-id="047c8-105">Requirements</span></span>

|<span data-ttu-id="047c8-106">要件</span><span class="sxs-lookup"><span data-stu-id="047c8-106">Requirement</span></span>| <span data-ttu-id="047c8-107">値</span><span class="sxs-lookup"><span data-stu-id="047c8-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="047c8-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="047c8-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="047c8-109">1.0</span><span class="sxs-lookup"><span data-stu-id="047c8-109">1.0</span></span>|
|[<span data-ttu-id="047c8-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="047c8-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="047c8-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="047c8-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="047c8-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="047c8-112">Members and methods</span></span>

| <span data-ttu-id="047c8-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="047c8-113">Member</span></span> | <span data-ttu-id="047c8-114">種類</span><span class="sxs-lookup"><span data-stu-id="047c8-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="047c8-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="047c8-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="047c8-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="047c8-116">Member</span></span> |
| [<span data-ttu-id="047c8-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="047c8-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="047c8-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="047c8-118">Member</span></span> |
| [<span data-ttu-id="047c8-119">EventType</span><span class="sxs-lookup"><span data-stu-id="047c8-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="047c8-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="047c8-120">Member</span></span> |
| [<span data-ttu-id="047c8-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="047c8-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="047c8-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="047c8-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="047c8-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="047c8-123">Namespaces</span></span>

<span data-ttu-id="047c8-124">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="047c8-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="047c8-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="047c8-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="047c8-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="047c8-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="047c8-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="047c8-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="047c8-128">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="047c8-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="047c8-129">型</span><span class="sxs-lookup"><span data-stu-id="047c8-129">Type</span></span>

*   <span data-ttu-id="047c8-130">String</span><span class="sxs-lookup"><span data-stu-id="047c8-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="047c8-131">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="047c8-131">Properties:</span></span>

|<span data-ttu-id="047c8-132">名前</span><span class="sxs-lookup"><span data-stu-id="047c8-132">Name</span></span>| <span data-ttu-id="047c8-133">型</span><span class="sxs-lookup"><span data-stu-id="047c8-133">Type</span></span>| <span data-ttu-id="047c8-134">説明</span><span class="sxs-lookup"><span data-stu-id="047c8-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="047c8-135">String</span><span class="sxs-lookup"><span data-stu-id="047c8-135">String</span></span>|<span data-ttu-id="047c8-136">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="047c8-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="047c8-137">String</span><span class="sxs-lookup"><span data-stu-id="047c8-137">String</span></span>|<span data-ttu-id="047c8-138">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="047c8-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="047c8-139">要件</span><span class="sxs-lookup"><span data-stu-id="047c8-139">Requirements</span></span>

|<span data-ttu-id="047c8-140">要件</span><span class="sxs-lookup"><span data-stu-id="047c8-140">Requirement</span></span>| <span data-ttu-id="047c8-141">値</span><span class="sxs-lookup"><span data-stu-id="047c8-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="047c8-142">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="047c8-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="047c8-143">1.0</span><span class="sxs-lookup"><span data-stu-id="047c8-143">1.0</span></span>|
|[<span data-ttu-id="047c8-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="047c8-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="047c8-145">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="047c8-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="047c8-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="047c8-146">CoercionType :String</span></span>

<span data-ttu-id="047c8-147">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="047c8-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="047c8-148">型</span><span class="sxs-lookup"><span data-stu-id="047c8-148">Type</span></span>

*   <span data-ttu-id="047c8-149">String</span><span class="sxs-lookup"><span data-stu-id="047c8-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="047c8-150">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="047c8-150">Properties:</span></span>

|<span data-ttu-id="047c8-151">名前</span><span class="sxs-lookup"><span data-stu-id="047c8-151">Name</span></span>| <span data-ttu-id="047c8-152">型</span><span class="sxs-lookup"><span data-stu-id="047c8-152">Type</span></span>| <span data-ttu-id="047c8-153">説明</span><span class="sxs-lookup"><span data-stu-id="047c8-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="047c8-154">String</span><span class="sxs-lookup"><span data-stu-id="047c8-154">String</span></span>|<span data-ttu-id="047c8-155">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="047c8-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="047c8-156">String</span><span class="sxs-lookup"><span data-stu-id="047c8-156">String</span></span>|<span data-ttu-id="047c8-157">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="047c8-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="047c8-158">要件</span><span class="sxs-lookup"><span data-stu-id="047c8-158">Requirements</span></span>

|<span data-ttu-id="047c8-159">要件</span><span class="sxs-lookup"><span data-stu-id="047c8-159">Requirement</span></span>| <span data-ttu-id="047c8-160">値</span><span class="sxs-lookup"><span data-stu-id="047c8-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="047c8-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="047c8-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="047c8-162">1.0</span><span class="sxs-lookup"><span data-stu-id="047c8-162">1.0</span></span>|
|[<span data-ttu-id="047c8-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="047c8-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="047c8-164">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="047c8-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="047c8-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="047c8-165">EventType :String</span></span>

<span data-ttu-id="047c8-166">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="047c8-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="047c8-167">型</span><span class="sxs-lookup"><span data-stu-id="047c8-167">Type</span></span>

*   <span data-ttu-id="047c8-168">String</span><span class="sxs-lookup"><span data-stu-id="047c8-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="047c8-169">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="047c8-169">Properties:</span></span>

| <span data-ttu-id="047c8-170">名前</span><span class="sxs-lookup"><span data-stu-id="047c8-170">Name</span></span> | <span data-ttu-id="047c8-171">型</span><span class="sxs-lookup"><span data-stu-id="047c8-171">Type</span></span> | <span data-ttu-id="047c8-172">説明</span><span class="sxs-lookup"><span data-stu-id="047c8-172">Description</span></span> | <span data-ttu-id="047c8-173">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="047c8-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="047c8-174">String</span><span class="sxs-lookup"><span data-stu-id="047c8-174">String</span></span> | <span data-ttu-id="047c8-175">選択した予定または一連の予定の日付または時刻が変更された。</span><span class="sxs-lookup"><span data-stu-id="047c8-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="047c8-176">1.7</span><span class="sxs-lookup"><span data-stu-id="047c8-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="047c8-177">String</span><span class="sxs-lookup"><span data-stu-id="047c8-177">String</span></span> | <span data-ttu-id="047c8-178">アイテムに添付ファイルが追加されたか、アイテムから添付ファイルが削除された。</span><span class="sxs-lookup"><span data-stu-id="047c8-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="047c8-179">プレビュー</span><span class="sxs-lookup"><span data-stu-id="047c8-179">Preview</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="047c8-180">String</span><span class="sxs-lookup"><span data-stu-id="047c8-180">String</span></span> | <span data-ttu-id="047c8-181">選択した予定の場所が変更されました。</span><span class="sxs-lookup"><span data-stu-id="047c8-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="047c8-182">プレビュー</span><span class="sxs-lookup"><span data-stu-id="047c8-182">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="047c8-183">String</span><span class="sxs-lookup"><span data-stu-id="047c8-183">String</span></span> | <span data-ttu-id="047c8-184">作業ウィンドウがピン留めされている間、別の Outlook アイテムが選択されて表示されている。</span><span class="sxs-lookup"><span data-stu-id="047c8-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="047c8-185">1.5</span><span class="sxs-lookup"><span data-stu-id="047c8-185">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="047c8-186">文字列</span><span class="sxs-lookup"><span data-stu-id="047c8-186">String</span></span> | <span data-ttu-id="047c8-187">メールボックスの Office テーマが変更された。</span><span class="sxs-lookup"><span data-stu-id="047c8-187">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="047c8-188">プレビュー</span><span class="sxs-lookup"><span data-stu-id="047c8-188">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="047c8-189">String</span><span class="sxs-lookup"><span data-stu-id="047c8-189">String</span></span> | <span data-ttu-id="047c8-190">選択したアイテムまたは予定の場所の受信者リストが変更された。</span><span class="sxs-lookup"><span data-stu-id="047c8-190">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="047c8-191">1.7</span><span class="sxs-lookup"><span data-stu-id="047c8-191">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="047c8-192">文字列</span><span class="sxs-lookup"><span data-stu-id="047c8-192">String</span></span> | <span data-ttu-id="047c8-193">選択した一連の予定の定期的なパターンが変更された。</span><span class="sxs-lookup"><span data-stu-id="047c8-193">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="047c8-194">1.7</span><span class="sxs-lookup"><span data-stu-id="047c8-194">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="047c8-195">要件</span><span class="sxs-lookup"><span data-stu-id="047c8-195">Requirements</span></span>

|<span data-ttu-id="047c8-196">要件</span><span class="sxs-lookup"><span data-stu-id="047c8-196">Requirement</span></span>| <span data-ttu-id="047c8-197">値</span><span class="sxs-lookup"><span data-stu-id="047c8-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="047c8-198">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="047c8-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="047c8-199">1.5</span><span class="sxs-lookup"><span data-stu-id="047c8-199">1.5</span></span> |
|[<span data-ttu-id="047c8-200">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="047c8-200">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="047c8-201">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="047c8-201">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="047c8-202">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="047c8-202">SourceProperty :String</span></span>

<span data-ttu-id="047c8-203">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="047c8-203">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="047c8-204">型</span><span class="sxs-lookup"><span data-stu-id="047c8-204">Type</span></span>

*   <span data-ttu-id="047c8-205">String</span><span class="sxs-lookup"><span data-stu-id="047c8-205">String</span></span>

##### <a name="properties"></a><span data-ttu-id="047c8-206">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="047c8-206">Properties:</span></span>

|<span data-ttu-id="047c8-207">名前</span><span class="sxs-lookup"><span data-stu-id="047c8-207">Name</span></span>| <span data-ttu-id="047c8-208">型</span><span class="sxs-lookup"><span data-stu-id="047c8-208">Type</span></span>| <span data-ttu-id="047c8-209">説明</span><span class="sxs-lookup"><span data-stu-id="047c8-209">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="047c8-210">String</span><span class="sxs-lookup"><span data-stu-id="047c8-210">String</span></span>|<span data-ttu-id="047c8-211">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="047c8-211">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="047c8-212">String</span><span class="sxs-lookup"><span data-stu-id="047c8-212">String</span></span>|<span data-ttu-id="047c8-213">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="047c8-213">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="047c8-214">要件</span><span class="sxs-lookup"><span data-stu-id="047c8-214">Requirements</span></span>

|<span data-ttu-id="047c8-215">要件</span><span class="sxs-lookup"><span data-stu-id="047c8-215">Requirement</span></span>| <span data-ttu-id="047c8-216">値</span><span class="sxs-lookup"><span data-stu-id="047c8-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="047c8-217">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="047c8-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="047c8-218">1.0</span><span class="sxs-lookup"><span data-stu-id="047c8-218">1.0</span></span>|
|[<span data-ttu-id="047c8-219">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="047c8-219">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="047c8-220">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="047c8-220">Compose or Read</span></span>|
