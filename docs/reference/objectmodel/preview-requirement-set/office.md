---
title: Office 名前空間-プレビュー要件セット
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: e426ea87c14c4ad21ebdbfd3df05988ba848b906
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872236"
---
# <a name="office"></a><span data-ttu-id="3b3a9-102">Office</span><span class="sxs-lookup"><span data-stu-id="3b3a9-102">Office</span></span>

<span data-ttu-id="3b3a9-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3b3a9-105">要件</span><span class="sxs-lookup"><span data-stu-id="3b3a9-105">Requirements</span></span>

|<span data-ttu-id="3b3a9-106">要件</span><span class="sxs-lookup"><span data-stu-id="3b3a9-106">Requirement</span></span>| <span data-ttu-id="3b3a9-107">値</span><span class="sxs-lookup"><span data-stu-id="3b3a9-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b3a9-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3b3a9-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b3a9-109">1.0</span><span class="sxs-lookup"><span data-stu-id="3b3a9-109">1.0</span></span>|
|[<span data-ttu-id="3b3a9-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3b3a9-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b3a9-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3b3a9-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3b3a9-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="3b3a9-112">Members and methods</span></span>

| <span data-ttu-id="3b3a9-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="3b3a9-113">Member</span></span> | <span data-ttu-id="3b3a9-114">種類</span><span class="sxs-lookup"><span data-stu-id="3b3a9-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3b3a9-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="3b3a9-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="3b3a9-116">Member</span><span class="sxs-lookup"><span data-stu-id="3b3a9-116">Member</span></span> |
| [<span data-ttu-id="3b3a9-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="3b3a9-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="3b3a9-118">Member</span><span class="sxs-lookup"><span data-stu-id="3b3a9-118">Member</span></span> |
| [<span data-ttu-id="3b3a9-119">EventType</span><span class="sxs-lookup"><span data-stu-id="3b3a9-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="3b3a9-120">Member</span><span class="sxs-lookup"><span data-stu-id="3b3a9-120">Member</span></span> |
| [<span data-ttu-id="3b3a9-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="3b3a9-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="3b3a9-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="3b3a9-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="3b3a9-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="3b3a9-123">Namespaces</span></span>

<span data-ttu-id="3b3a9-124">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="3b3a9-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="3b3a9-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="3b3a9-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="3b3a9-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="3b3a9-128">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="3b3a9-129">型</span><span class="sxs-lookup"><span data-stu-id="3b3a9-129">Type</span></span>

*   <span data-ttu-id="3b3a9-130">String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3b3a9-131">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="3b3a9-131">Properties:</span></span>

|<span data-ttu-id="3b3a9-132">名前</span><span class="sxs-lookup"><span data-stu-id="3b3a9-132">Name</span></span>| <span data-ttu-id="3b3a9-133">種類</span><span class="sxs-lookup"><span data-stu-id="3b3a9-133">Type</span></span>| <span data-ttu-id="3b3a9-134">説明</span><span class="sxs-lookup"><span data-stu-id="3b3a9-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="3b3a9-135">String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-135">String</span></span>|<span data-ttu-id="3b3a9-136">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="3b3a9-137">String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-137">String</span></span>|<span data-ttu-id="3b3a9-138">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b3a9-139">要件</span><span class="sxs-lookup"><span data-stu-id="3b3a9-139">Requirements</span></span>

|<span data-ttu-id="3b3a9-140">要件</span><span class="sxs-lookup"><span data-stu-id="3b3a9-140">Requirement</span></span>| <span data-ttu-id="3b3a9-141">値</span><span class="sxs-lookup"><span data-stu-id="3b3a9-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b3a9-142">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3b3a9-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b3a9-143">1.0</span><span class="sxs-lookup"><span data-stu-id="3b3a9-143">1.0</span></span>|
|[<span data-ttu-id="3b3a9-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3b3a9-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b3a9-145">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3b3a9-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="3b3a9-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-146">CoercionType :String</span></span>

<span data-ttu-id="3b3a9-147">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3b3a9-148">型</span><span class="sxs-lookup"><span data-stu-id="3b3a9-148">Type</span></span>

*   <span data-ttu-id="3b3a9-149">String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3b3a9-150">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="3b3a9-150">Properties:</span></span>

|<span data-ttu-id="3b3a9-151">名前</span><span class="sxs-lookup"><span data-stu-id="3b3a9-151">Name</span></span>| <span data-ttu-id="3b3a9-152">種類</span><span class="sxs-lookup"><span data-stu-id="3b3a9-152">Type</span></span>| <span data-ttu-id="3b3a9-153">説明</span><span class="sxs-lookup"><span data-stu-id="3b3a9-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="3b3a9-154">String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-154">String</span></span>|<span data-ttu-id="3b3a9-155">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="3b3a9-156">String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-156">String</span></span>|<span data-ttu-id="3b3a9-157">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b3a9-158">要件</span><span class="sxs-lookup"><span data-stu-id="3b3a9-158">Requirements</span></span>

|<span data-ttu-id="3b3a9-159">要件</span><span class="sxs-lookup"><span data-stu-id="3b3a9-159">Requirement</span></span>| <span data-ttu-id="3b3a9-160">値</span><span class="sxs-lookup"><span data-stu-id="3b3a9-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b3a9-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3b3a9-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b3a9-162">1.0</span><span class="sxs-lookup"><span data-stu-id="3b3a9-162">1.0</span></span>|
|[<span data-ttu-id="3b3a9-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3b3a9-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b3a9-164">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3b3a9-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="3b3a9-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-165">EventType :String</span></span>

<span data-ttu-id="3b3a9-166">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="3b3a9-167">型</span><span class="sxs-lookup"><span data-stu-id="3b3a9-167">Type</span></span>

*   <span data-ttu-id="3b3a9-168">String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3b3a9-169">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="3b3a9-169">Properties:</span></span>

| <span data-ttu-id="3b3a9-170">名前</span><span class="sxs-lookup"><span data-stu-id="3b3a9-170">Name</span></span> | <span data-ttu-id="3b3a9-171">種類</span><span class="sxs-lookup"><span data-stu-id="3b3a9-171">Type</span></span> | <span data-ttu-id="3b3a9-172">説明</span><span class="sxs-lookup"><span data-stu-id="3b3a9-172">Description</span></span> | <span data-ttu-id="3b3a9-173">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="3b3a9-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="3b3a9-174">String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-174">String</span></span> | <span data-ttu-id="3b3a9-175">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="3b3a9-176">1.7</span><span class="sxs-lookup"><span data-stu-id="3b3a9-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="3b3a9-177">String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-177">String</span></span> | <span data-ttu-id="3b3a9-178">添付ファイルがアイテムに追加またはアイテムから削除されています。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="3b3a9-179">プレビュー</span><span class="sxs-lookup"><span data-stu-id="3b3a9-179">Preview</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="3b3a9-180">String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-180">String</span></span> | <span data-ttu-id="3b3a9-181">選択した予定の場所が変更されました。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="3b3a9-182">プレビュー</span><span class="sxs-lookup"><span data-stu-id="3b3a9-182">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="3b3a9-183">String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-183">String</span></span> | <span data-ttu-id="3b3a9-184">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="3b3a9-185">1.5</span><span class="sxs-lookup"><span data-stu-id="3b3a9-185">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="3b3a9-186">String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-186">String</span></span> | <span data-ttu-id="3b3a9-187">メールボックスの Office テーマが変更されました。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-187">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="3b3a9-188">プレビュー</span><span class="sxs-lookup"><span data-stu-id="3b3a9-188">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="3b3a9-189">String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-189">String</span></span> | <span data-ttu-id="3b3a9-190">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-190">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="3b3a9-191">1.7</span><span class="sxs-lookup"><span data-stu-id="3b3a9-191">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="3b3a9-192">String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-192">String</span></span> | <span data-ttu-id="3b3a9-193">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-193">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="3b3a9-194">1.7</span><span class="sxs-lookup"><span data-stu-id="3b3a9-194">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3b3a9-195">要件</span><span class="sxs-lookup"><span data-stu-id="3b3a9-195">Requirements</span></span>

|<span data-ttu-id="3b3a9-196">要件</span><span class="sxs-lookup"><span data-stu-id="3b3a9-196">Requirement</span></span>| <span data-ttu-id="3b3a9-197">値</span><span class="sxs-lookup"><span data-stu-id="3b3a9-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b3a9-198">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3b3a9-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b3a9-199">1.5</span><span class="sxs-lookup"><span data-stu-id="3b3a9-199">1.5</span></span> |
|[<span data-ttu-id="3b3a9-200">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3b3a9-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b3a9-201">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3b3a9-201">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="3b3a9-202">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-202">SourceProperty :String</span></span>

<span data-ttu-id="3b3a9-203">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-203">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3b3a9-204">型</span><span class="sxs-lookup"><span data-stu-id="3b3a9-204">Type</span></span>

*   <span data-ttu-id="3b3a9-205">String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-205">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3b3a9-206">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="3b3a9-206">Properties:</span></span>

|<span data-ttu-id="3b3a9-207">名前</span><span class="sxs-lookup"><span data-stu-id="3b3a9-207">Name</span></span>| <span data-ttu-id="3b3a9-208">種類</span><span class="sxs-lookup"><span data-stu-id="3b3a9-208">Type</span></span>| <span data-ttu-id="3b3a9-209">説明</span><span class="sxs-lookup"><span data-stu-id="3b3a9-209">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="3b3a9-210">String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-210">String</span></span>|<span data-ttu-id="3b3a9-211">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-211">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="3b3a9-212">String</span><span class="sxs-lookup"><span data-stu-id="3b3a9-212">String</span></span>|<span data-ttu-id="3b3a9-213">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="3b3a9-213">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b3a9-214">要件</span><span class="sxs-lookup"><span data-stu-id="3b3a9-214">Requirements</span></span>

|<span data-ttu-id="3b3a9-215">要件</span><span class="sxs-lookup"><span data-stu-id="3b3a9-215">Requirement</span></span>| <span data-ttu-id="3b3a9-216">値</span><span class="sxs-lookup"><span data-stu-id="3b3a9-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b3a9-217">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3b3a9-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b3a9-218">1.0</span><span class="sxs-lookup"><span data-stu-id="3b3a9-218">1.0</span></span>|
|[<span data-ttu-id="3b3a9-219">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3b3a9-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b3a9-220">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3b3a9-220">Compose or Read</span></span>|
