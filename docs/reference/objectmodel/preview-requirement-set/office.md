---
title: Office 名前空間-プレビュー要件セット
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 7effc930d196aa009c3c779b702e082ae388fada
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838516"
---
# <a name="office"></a><span data-ttu-id="94d8e-102">Office</span><span class="sxs-lookup"><span data-stu-id="94d8e-102">Office</span></span>

<span data-ttu-id="94d8e-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="94d8e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="94d8e-105">要件</span><span class="sxs-lookup"><span data-stu-id="94d8e-105">Requirements</span></span>

|<span data-ttu-id="94d8e-106">要件</span><span class="sxs-lookup"><span data-stu-id="94d8e-106">Requirement</span></span>| <span data-ttu-id="94d8e-107">値</span><span class="sxs-lookup"><span data-stu-id="94d8e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="94d8e-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="94d8e-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="94d8e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="94d8e-109">1.0</span></span>|
|[<span data-ttu-id="94d8e-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="94d8e-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="94d8e-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="94d8e-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="94d8e-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="94d8e-112">Members and methods</span></span>

| <span data-ttu-id="94d8e-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="94d8e-113">Member</span></span> | <span data-ttu-id="94d8e-114">種類</span><span class="sxs-lookup"><span data-stu-id="94d8e-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="94d8e-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="94d8e-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="94d8e-116">Member</span><span class="sxs-lookup"><span data-stu-id="94d8e-116">Member</span></span> |
| [<span data-ttu-id="94d8e-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="94d8e-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="94d8e-118">Member</span><span class="sxs-lookup"><span data-stu-id="94d8e-118">Member</span></span> |
| [<span data-ttu-id="94d8e-119">EventType</span><span class="sxs-lookup"><span data-stu-id="94d8e-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="94d8e-120">Member</span><span class="sxs-lookup"><span data-stu-id="94d8e-120">Member</span></span> |
| [<span data-ttu-id="94d8e-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="94d8e-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="94d8e-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="94d8e-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="94d8e-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="94d8e-123">Namespaces</span></span>

<span data-ttu-id="94d8e-124">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="94d8e-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="94d8e-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="94d8e-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="94d8e-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="94d8e-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="94d8e-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="94d8e-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="94d8e-128">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="94d8e-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="94d8e-129">型</span><span class="sxs-lookup"><span data-stu-id="94d8e-129">Type</span></span>

*   <span data-ttu-id="94d8e-130">String</span><span class="sxs-lookup"><span data-stu-id="94d8e-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="94d8e-131">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="94d8e-131">Properties:</span></span>

|<span data-ttu-id="94d8e-132">名前</span><span class="sxs-lookup"><span data-stu-id="94d8e-132">Name</span></span>| <span data-ttu-id="94d8e-133">種類</span><span class="sxs-lookup"><span data-stu-id="94d8e-133">Type</span></span>| <span data-ttu-id="94d8e-134">説明</span><span class="sxs-lookup"><span data-stu-id="94d8e-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="94d8e-135">文字列</span><span class="sxs-lookup"><span data-stu-id="94d8e-135">String</span></span>|<span data-ttu-id="94d8e-136">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="94d8e-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="94d8e-137">文字列</span><span class="sxs-lookup"><span data-stu-id="94d8e-137">String</span></span>|<span data-ttu-id="94d8e-138">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="94d8e-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="94d8e-139">要件</span><span class="sxs-lookup"><span data-stu-id="94d8e-139">Requirements</span></span>

|<span data-ttu-id="94d8e-140">要件</span><span class="sxs-lookup"><span data-stu-id="94d8e-140">Requirement</span></span>| <span data-ttu-id="94d8e-141">値</span><span class="sxs-lookup"><span data-stu-id="94d8e-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="94d8e-142">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="94d8e-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="94d8e-143">1.0</span><span class="sxs-lookup"><span data-stu-id="94d8e-143">1.0</span></span>|
|[<span data-ttu-id="94d8e-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="94d8e-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="94d8e-145">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="94d8e-145">Compose or Read</span></span>|

---
---

####  <a name="coerciontype-string"></a><span data-ttu-id="94d8e-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="94d8e-146">CoercionType :String</span></span>

<span data-ttu-id="94d8e-147">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="94d8e-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="94d8e-148">型</span><span class="sxs-lookup"><span data-stu-id="94d8e-148">Type</span></span>

*   <span data-ttu-id="94d8e-149">String</span><span class="sxs-lookup"><span data-stu-id="94d8e-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="94d8e-150">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="94d8e-150">Properties:</span></span>

|<span data-ttu-id="94d8e-151">名前</span><span class="sxs-lookup"><span data-stu-id="94d8e-151">Name</span></span>| <span data-ttu-id="94d8e-152">種類</span><span class="sxs-lookup"><span data-stu-id="94d8e-152">Type</span></span>| <span data-ttu-id="94d8e-153">説明</span><span class="sxs-lookup"><span data-stu-id="94d8e-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="94d8e-154">文字列</span><span class="sxs-lookup"><span data-stu-id="94d8e-154">String</span></span>|<span data-ttu-id="94d8e-155">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="94d8e-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="94d8e-156">String</span><span class="sxs-lookup"><span data-stu-id="94d8e-156">String</span></span>|<span data-ttu-id="94d8e-157">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="94d8e-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="94d8e-158">要件</span><span class="sxs-lookup"><span data-stu-id="94d8e-158">Requirements</span></span>

|<span data-ttu-id="94d8e-159">要件</span><span class="sxs-lookup"><span data-stu-id="94d8e-159">Requirement</span></span>| <span data-ttu-id="94d8e-160">値</span><span class="sxs-lookup"><span data-stu-id="94d8e-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="94d8e-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="94d8e-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="94d8e-162">1.0</span><span class="sxs-lookup"><span data-stu-id="94d8e-162">1.0</span></span>|
|[<span data-ttu-id="94d8e-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="94d8e-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="94d8e-164">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="94d8e-164">Compose or Read</span></span>|

---
---

####  <a name="eventtype-string"></a><span data-ttu-id="94d8e-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="94d8e-165">EventType :String</span></span>

<span data-ttu-id="94d8e-166">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="94d8e-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="94d8e-167">型</span><span class="sxs-lookup"><span data-stu-id="94d8e-167">Type</span></span>

*   <span data-ttu-id="94d8e-168">String</span><span class="sxs-lookup"><span data-stu-id="94d8e-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="94d8e-169">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="94d8e-169">Properties:</span></span>

| <span data-ttu-id="94d8e-170">名前</span><span class="sxs-lookup"><span data-stu-id="94d8e-170">Name</span></span> | <span data-ttu-id="94d8e-171">種類</span><span class="sxs-lookup"><span data-stu-id="94d8e-171">Type</span></span> | <span data-ttu-id="94d8e-172">説明</span><span class="sxs-lookup"><span data-stu-id="94d8e-172">Description</span></span> | <span data-ttu-id="94d8e-173">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="94d8e-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="94d8e-174">文字列</span><span class="sxs-lookup"><span data-stu-id="94d8e-174">String</span></span> | <span data-ttu-id="94d8e-175">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="94d8e-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="94d8e-176">1.7</span><span class="sxs-lookup"><span data-stu-id="94d8e-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="94d8e-177">文字列</span><span class="sxs-lookup"><span data-stu-id="94d8e-177">String</span></span> | <span data-ttu-id="94d8e-178">添付ファイルがアイテムに追加またはアイテムから削除されています。</span><span class="sxs-lookup"><span data-stu-id="94d8e-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="94d8e-179">プレビュー</span><span class="sxs-lookup"><span data-stu-id="94d8e-179">Preview</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="94d8e-180">文字列</span><span class="sxs-lookup"><span data-stu-id="94d8e-180">String</span></span> | <span data-ttu-id="94d8e-181">選択した予定の場所が変更されました。</span><span class="sxs-lookup"><span data-stu-id="94d8e-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="94d8e-182">プレビュー</span><span class="sxs-lookup"><span data-stu-id="94d8e-182">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="94d8e-183">文字列</span><span class="sxs-lookup"><span data-stu-id="94d8e-183">String</span></span> | <span data-ttu-id="94d8e-184">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="94d8e-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="94d8e-185">1.5</span><span class="sxs-lookup"><span data-stu-id="94d8e-185">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="94d8e-186">文字列</span><span class="sxs-lookup"><span data-stu-id="94d8e-186">String</span></span> | <span data-ttu-id="94d8e-187">メールボックスの Office テーマが変更されました。</span><span class="sxs-lookup"><span data-stu-id="94d8e-187">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="94d8e-188">プレビュー</span><span class="sxs-lookup"><span data-stu-id="94d8e-188">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="94d8e-189">文字列</span><span class="sxs-lookup"><span data-stu-id="94d8e-189">String</span></span> | <span data-ttu-id="94d8e-190">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="94d8e-190">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="94d8e-191">1.7</span><span class="sxs-lookup"><span data-stu-id="94d8e-191">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="94d8e-192">文字列</span><span class="sxs-lookup"><span data-stu-id="94d8e-192">String</span></span> | <span data-ttu-id="94d8e-193">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="94d8e-193">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="94d8e-194">1.7</span><span class="sxs-lookup"><span data-stu-id="94d8e-194">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="94d8e-195">要件</span><span class="sxs-lookup"><span data-stu-id="94d8e-195">Requirements</span></span>

|<span data-ttu-id="94d8e-196">要件</span><span class="sxs-lookup"><span data-stu-id="94d8e-196">Requirement</span></span>| <span data-ttu-id="94d8e-197">値</span><span class="sxs-lookup"><span data-stu-id="94d8e-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="94d8e-198">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="94d8e-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="94d8e-199">1.5</span><span class="sxs-lookup"><span data-stu-id="94d8e-199">1.5</span></span> |
|[<span data-ttu-id="94d8e-200">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="94d8e-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="94d8e-201">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="94d8e-201">Compose or Read</span></span> |

---
---

####  <a name="sourceproperty-string"></a><span data-ttu-id="94d8e-202">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="94d8e-202">SourceProperty :String</span></span>

<span data-ttu-id="94d8e-203">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="94d8e-203">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="94d8e-204">型</span><span class="sxs-lookup"><span data-stu-id="94d8e-204">Type</span></span>

*   <span data-ttu-id="94d8e-205">String</span><span class="sxs-lookup"><span data-stu-id="94d8e-205">String</span></span>

##### <a name="properties"></a><span data-ttu-id="94d8e-206">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="94d8e-206">Properties:</span></span>

|<span data-ttu-id="94d8e-207">名前</span><span class="sxs-lookup"><span data-stu-id="94d8e-207">Name</span></span>| <span data-ttu-id="94d8e-208">種類</span><span class="sxs-lookup"><span data-stu-id="94d8e-208">Type</span></span>| <span data-ttu-id="94d8e-209">説明</span><span class="sxs-lookup"><span data-stu-id="94d8e-209">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="94d8e-210">文字列</span><span class="sxs-lookup"><span data-stu-id="94d8e-210">String</span></span>|<span data-ttu-id="94d8e-211">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="94d8e-211">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="94d8e-212">文字列</span><span class="sxs-lookup"><span data-stu-id="94d8e-212">String</span></span>|<span data-ttu-id="94d8e-213">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="94d8e-213">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="94d8e-214">要件</span><span class="sxs-lookup"><span data-stu-id="94d8e-214">Requirements</span></span>

|<span data-ttu-id="94d8e-215">要件</span><span class="sxs-lookup"><span data-stu-id="94d8e-215">Requirement</span></span>| <span data-ttu-id="94d8e-216">値</span><span class="sxs-lookup"><span data-stu-id="94d8e-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="94d8e-217">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="94d8e-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="94d8e-218">1.0</span><span class="sxs-lookup"><span data-stu-id="94d8e-218">1.0</span></span>|
|[<span data-ttu-id="94d8e-219">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="94d8e-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="94d8e-220">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="94d8e-220">Compose or Read</span></span>|
