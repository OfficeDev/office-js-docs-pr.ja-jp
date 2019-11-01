---
title: Office 名前空間-プレビュー要件セット
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: eae6f99d166695f24f4a94e89ea4b876bea080ef
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902103"
---
# <a name="office"></a><span data-ttu-id="f2ae6-102">Office</span><span class="sxs-lookup"><span data-stu-id="f2ae6-102">Office</span></span>

<span data-ttu-id="f2ae6-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f2ae6-105">要件</span><span class="sxs-lookup"><span data-stu-id="f2ae6-105">Requirements</span></span>

|<span data-ttu-id="f2ae6-106">要件</span><span class="sxs-lookup"><span data-stu-id="f2ae6-106">Requirement</span></span>| <span data-ttu-id="f2ae6-107">値</span><span class="sxs-lookup"><span data-stu-id="f2ae6-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f2ae6-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f2ae6-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f2ae6-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f2ae6-109">1.0</span></span>|
|[<span data-ttu-id="f2ae6-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f2ae6-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f2ae6-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f2ae6-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f2ae6-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="f2ae6-112">Members and methods</span></span>

| <span data-ttu-id="f2ae6-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="f2ae6-113">Member</span></span> | <span data-ttu-id="f2ae6-114">型</span><span class="sxs-lookup"><span data-stu-id="f2ae6-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f2ae6-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f2ae6-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f2ae6-116">Member</span><span class="sxs-lookup"><span data-stu-id="f2ae6-116">Member</span></span> |
| [<span data-ttu-id="f2ae6-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f2ae6-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f2ae6-118">Member</span><span class="sxs-lookup"><span data-stu-id="f2ae6-118">Member</span></span> |
| [<span data-ttu-id="f2ae6-119">EventType</span><span class="sxs-lookup"><span data-stu-id="f2ae6-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="f2ae6-120">Member</span><span class="sxs-lookup"><span data-stu-id="f2ae6-120">Member</span></span> |
| [<span data-ttu-id="f2ae6-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f2ae6-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f2ae6-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="f2ae6-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="f2ae6-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="f2ae6-123">Namespaces</span></span>

<span data-ttu-id="f2ae6-124">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="f2ae6-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat):、、、、、、などのさまざま`ItemType`な`EntityType`列挙`AttachmentType` `RecipientType` `ResponseType`値が含まれ`ItemNotificationMessageType`ています。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="f2ae6-126">Members</span><span class="sxs-lookup"><span data-stu-id="f2ae6-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="f2ae6-127">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="f2ae6-128">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f2ae6-129">型</span><span class="sxs-lookup"><span data-stu-id="f2ae6-129">Type</span></span>

*   <span data-ttu-id="f2ae6-130">String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f2ae6-131">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="f2ae6-131">Properties:</span></span>

|<span data-ttu-id="f2ae6-132">名前</span><span class="sxs-lookup"><span data-stu-id="f2ae6-132">Name</span></span>| <span data-ttu-id="f2ae6-133">種類</span><span class="sxs-lookup"><span data-stu-id="f2ae6-133">Type</span></span>| <span data-ttu-id="f2ae6-134">説明</span><span class="sxs-lookup"><span data-stu-id="f2ae6-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f2ae6-135">String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-135">String</span></span>|<span data-ttu-id="f2ae6-136">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f2ae6-137">String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-137">String</span></span>|<span data-ttu-id="f2ae6-138">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f2ae6-139">要件</span><span class="sxs-lookup"><span data-stu-id="f2ae6-139">Requirements</span></span>

|<span data-ttu-id="f2ae6-140">要件</span><span class="sxs-lookup"><span data-stu-id="f2ae6-140">Requirement</span></span>| <span data-ttu-id="f2ae6-141">値</span><span class="sxs-lookup"><span data-stu-id="f2ae6-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="f2ae6-142">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f2ae6-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f2ae6-143">1.0</span><span class="sxs-lookup"><span data-stu-id="f2ae6-143">1.0</span></span>|
|[<span data-ttu-id="f2ae6-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f2ae6-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f2ae6-145">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f2ae6-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="f2ae6-146">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-146">CoercionType: String</span></span>

<span data-ttu-id="f2ae6-147">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f2ae6-148">型</span><span class="sxs-lookup"><span data-stu-id="f2ae6-148">Type</span></span>

*   <span data-ttu-id="f2ae6-149">String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f2ae6-150">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="f2ae6-150">Properties:</span></span>

|<span data-ttu-id="f2ae6-151">名前</span><span class="sxs-lookup"><span data-stu-id="f2ae6-151">Name</span></span>| <span data-ttu-id="f2ae6-152">種類</span><span class="sxs-lookup"><span data-stu-id="f2ae6-152">Type</span></span>| <span data-ttu-id="f2ae6-153">説明</span><span class="sxs-lookup"><span data-stu-id="f2ae6-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f2ae6-154">String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-154">String</span></span>|<span data-ttu-id="f2ae6-155">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f2ae6-156">String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-156">String</span></span>|<span data-ttu-id="f2ae6-157">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f2ae6-158">要件</span><span class="sxs-lookup"><span data-stu-id="f2ae6-158">Requirements</span></span>

|<span data-ttu-id="f2ae6-159">要件</span><span class="sxs-lookup"><span data-stu-id="f2ae6-159">Requirement</span></span>| <span data-ttu-id="f2ae6-160">値</span><span class="sxs-lookup"><span data-stu-id="f2ae6-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="f2ae6-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f2ae6-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f2ae6-162">1.0</span><span class="sxs-lookup"><span data-stu-id="f2ae6-162">1.0</span></span>|
|[<span data-ttu-id="f2ae6-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f2ae6-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f2ae6-164">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f2ae6-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="f2ae6-165">EventType: String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-165">EventType: String</span></span>

<span data-ttu-id="f2ae6-166">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="f2ae6-167">型</span><span class="sxs-lookup"><span data-stu-id="f2ae6-167">Type</span></span>

*   <span data-ttu-id="f2ae6-168">String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f2ae6-169">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="f2ae6-169">Properties:</span></span>

| <span data-ttu-id="f2ae6-170">名前</span><span class="sxs-lookup"><span data-stu-id="f2ae6-170">Name</span></span> | <span data-ttu-id="f2ae6-171">種類</span><span class="sxs-lookup"><span data-stu-id="f2ae6-171">Type</span></span> | <span data-ttu-id="f2ae6-172">説明</span><span class="sxs-lookup"><span data-stu-id="f2ae6-172">Description</span></span> | <span data-ttu-id="f2ae6-173">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="f2ae6-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="f2ae6-174">String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-174">String</span></span> | <span data-ttu-id="f2ae6-175">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="f2ae6-176">1.7</span><span class="sxs-lookup"><span data-stu-id="f2ae6-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="f2ae6-177">String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-177">String</span></span> | <span data-ttu-id="f2ae6-178">添付ファイルがアイテムに追加またはアイテムから削除されています。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="f2ae6-179">1.8</span><span class="sxs-lookup"><span data-stu-id="f2ae6-179">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="f2ae6-180">String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-180">String</span></span> | <span data-ttu-id="f2ae6-181">選択した予定の場所が変更されました。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="f2ae6-182">1.8</span><span class="sxs-lookup"><span data-stu-id="f2ae6-182">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="f2ae6-183">String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-183">String</span></span> | <span data-ttu-id="f2ae6-184">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="f2ae6-185">1.5</span><span class="sxs-lookup"><span data-stu-id="f2ae6-185">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="f2ae6-186">String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-186">String</span></span> | <span data-ttu-id="f2ae6-187">メールボックスの Office テーマが変更されました。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-187">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="f2ae6-188">プレビュー</span><span class="sxs-lookup"><span data-stu-id="f2ae6-188">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="f2ae6-189">String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-189">String</span></span> | <span data-ttu-id="f2ae6-190">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-190">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="f2ae6-191">1.7</span><span class="sxs-lookup"><span data-stu-id="f2ae6-191">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="f2ae6-192">String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-192">String</span></span> | <span data-ttu-id="f2ae6-193">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-193">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="f2ae6-194">1.7</span><span class="sxs-lookup"><span data-stu-id="f2ae6-194">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f2ae6-195">要件</span><span class="sxs-lookup"><span data-stu-id="f2ae6-195">Requirements</span></span>

|<span data-ttu-id="f2ae6-196">要件</span><span class="sxs-lookup"><span data-stu-id="f2ae6-196">Requirement</span></span>| <span data-ttu-id="f2ae6-197">値</span><span class="sxs-lookup"><span data-stu-id="f2ae6-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="f2ae6-198">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f2ae6-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f2ae6-199">1.5</span><span class="sxs-lookup"><span data-stu-id="f2ae6-199">1.5</span></span> |
|[<span data-ttu-id="f2ae6-200">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f2ae6-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f2ae6-201">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f2ae6-201">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="f2ae6-202">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-202">SourceProperty: String</span></span>

<span data-ttu-id="f2ae6-203">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-203">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f2ae6-204">型</span><span class="sxs-lookup"><span data-stu-id="f2ae6-204">Type</span></span>

*   <span data-ttu-id="f2ae6-205">String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-205">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f2ae6-206">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="f2ae6-206">Properties:</span></span>

|<span data-ttu-id="f2ae6-207">名前</span><span class="sxs-lookup"><span data-stu-id="f2ae6-207">Name</span></span>| <span data-ttu-id="f2ae6-208">種類</span><span class="sxs-lookup"><span data-stu-id="f2ae6-208">Type</span></span>| <span data-ttu-id="f2ae6-209">説明</span><span class="sxs-lookup"><span data-stu-id="f2ae6-209">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f2ae6-210">String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-210">String</span></span>|<span data-ttu-id="f2ae6-211">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-211">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f2ae6-212">String</span><span class="sxs-lookup"><span data-stu-id="f2ae6-212">String</span></span>|<span data-ttu-id="f2ae6-213">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="f2ae6-213">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f2ae6-214">要件</span><span class="sxs-lookup"><span data-stu-id="f2ae6-214">Requirements</span></span>

|<span data-ttu-id="f2ae6-215">要件</span><span class="sxs-lookup"><span data-stu-id="f2ae6-215">Requirement</span></span>| <span data-ttu-id="f2ae6-216">値</span><span class="sxs-lookup"><span data-stu-id="f2ae6-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="f2ae6-217">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f2ae6-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f2ae6-218">1.0</span><span class="sxs-lookup"><span data-stu-id="f2ae6-218">1.0</span></span>|
|[<span data-ttu-id="f2ae6-219">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f2ae6-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f2ae6-220">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f2ae6-220">Compose or Read</span></span>|
