---
title: Office 名前空間-要件セット1.7
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: be0223e7ed274abf0e742be13f258c14f6dccf91
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395695"
---
# <a name="office"></a><span data-ttu-id="6b3ac-102">Office</span><span class="sxs-lookup"><span data-stu-id="6b3ac-102">Office</span></span>

<span data-ttu-id="6b3ac-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6b3ac-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6b3ac-105">要件</span><span class="sxs-lookup"><span data-stu-id="6b3ac-105">Requirements</span></span>

|<span data-ttu-id="6b3ac-106">要件</span><span class="sxs-lookup"><span data-stu-id="6b3ac-106">Requirement</span></span>| <span data-ttu-id="6b3ac-107">値</span><span class="sxs-lookup"><span data-stu-id="6b3ac-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b3ac-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6b3ac-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b3ac-109">1.0</span><span class="sxs-lookup"><span data-stu-id="6b3ac-109">1.0</span></span>|
|[<span data-ttu-id="6b3ac-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6b3ac-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b3ac-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6b3ac-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="6b3ac-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="6b3ac-112">Members and methods</span></span>

| <span data-ttu-id="6b3ac-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="6b3ac-113">Member</span></span> | <span data-ttu-id="6b3ac-114">種類</span><span class="sxs-lookup"><span data-stu-id="6b3ac-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="6b3ac-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="6b3ac-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="6b3ac-116">Member</span><span class="sxs-lookup"><span data-stu-id="6b3ac-116">Member</span></span> |
| [<span data-ttu-id="6b3ac-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="6b3ac-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="6b3ac-118">Member</span><span class="sxs-lookup"><span data-stu-id="6b3ac-118">Member</span></span> |
| [<span data-ttu-id="6b3ac-119">EventType</span><span class="sxs-lookup"><span data-stu-id="6b3ac-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="6b3ac-120">Member</span><span class="sxs-lookup"><span data-stu-id="6b3ac-120">Member</span></span> |
| [<span data-ttu-id="6b3ac-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="6b3ac-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="6b3ac-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="6b3ac-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="6b3ac-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="6b3ac-123">Namespaces</span></span>

<span data-ttu-id="6b3ac-124">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="6b3ac-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="6b3ac-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.7):、、、、、、などのさまざま`ItemType`な`EntityType`列挙`AttachmentType` `RecipientType` `ResponseType`値が含まれ`ItemNotificationMessageType`ています。</span><span class="sxs-lookup"><span data-stu-id="6b3ac-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.7): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="6b3ac-126">Members</span><span class="sxs-lookup"><span data-stu-id="6b3ac-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="6b3ac-127">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="6b3ac-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="6b3ac-128">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="6b3ac-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="6b3ac-129">型</span><span class="sxs-lookup"><span data-stu-id="6b3ac-129">Type</span></span>

*   <span data-ttu-id="6b3ac-130">String</span><span class="sxs-lookup"><span data-stu-id="6b3ac-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6b3ac-131">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="6b3ac-131">Properties:</span></span>

|<span data-ttu-id="6b3ac-132">名前</span><span class="sxs-lookup"><span data-stu-id="6b3ac-132">Name</span></span>| <span data-ttu-id="6b3ac-133">種類</span><span class="sxs-lookup"><span data-stu-id="6b3ac-133">Type</span></span>| <span data-ttu-id="6b3ac-134">説明</span><span class="sxs-lookup"><span data-stu-id="6b3ac-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="6b3ac-135">String</span><span class="sxs-lookup"><span data-stu-id="6b3ac-135">String</span></span>|<span data-ttu-id="6b3ac-136">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="6b3ac-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="6b3ac-137">String</span><span class="sxs-lookup"><span data-stu-id="6b3ac-137">String</span></span>|<span data-ttu-id="6b3ac-138">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="6b3ac-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6b3ac-139">要件</span><span class="sxs-lookup"><span data-stu-id="6b3ac-139">Requirements</span></span>

|<span data-ttu-id="6b3ac-140">要件</span><span class="sxs-lookup"><span data-stu-id="6b3ac-140">Requirement</span></span>| <span data-ttu-id="6b3ac-141">値</span><span class="sxs-lookup"><span data-stu-id="6b3ac-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b3ac-142">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6b3ac-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b3ac-143">1.0</span><span class="sxs-lookup"><span data-stu-id="6b3ac-143">1.0</span></span>|
|[<span data-ttu-id="6b3ac-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6b3ac-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b3ac-145">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6b3ac-145">Compose or Read</span></span>|

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="6b3ac-146">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="6b3ac-146">CoercionType: String</span></span>

<span data-ttu-id="6b3ac-147">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="6b3ac-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6b3ac-148">型</span><span class="sxs-lookup"><span data-stu-id="6b3ac-148">Type</span></span>

*   <span data-ttu-id="6b3ac-149">String</span><span class="sxs-lookup"><span data-stu-id="6b3ac-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6b3ac-150">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="6b3ac-150">Properties:</span></span>

|<span data-ttu-id="6b3ac-151">名前</span><span class="sxs-lookup"><span data-stu-id="6b3ac-151">Name</span></span>| <span data-ttu-id="6b3ac-152">種類</span><span class="sxs-lookup"><span data-stu-id="6b3ac-152">Type</span></span>| <span data-ttu-id="6b3ac-153">説明</span><span class="sxs-lookup"><span data-stu-id="6b3ac-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="6b3ac-154">String</span><span class="sxs-lookup"><span data-stu-id="6b3ac-154">String</span></span>|<span data-ttu-id="6b3ac-155">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="6b3ac-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="6b3ac-156">String</span><span class="sxs-lookup"><span data-stu-id="6b3ac-156">String</span></span>|<span data-ttu-id="6b3ac-157">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="6b3ac-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6b3ac-158">要件</span><span class="sxs-lookup"><span data-stu-id="6b3ac-158">Requirements</span></span>

|<span data-ttu-id="6b3ac-159">要件</span><span class="sxs-lookup"><span data-stu-id="6b3ac-159">Requirement</span></span>| <span data-ttu-id="6b3ac-160">値</span><span class="sxs-lookup"><span data-stu-id="6b3ac-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b3ac-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6b3ac-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b3ac-162">1.0</span><span class="sxs-lookup"><span data-stu-id="6b3ac-162">1.0</span></span>|
|[<span data-ttu-id="6b3ac-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6b3ac-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b3ac-164">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6b3ac-164">Compose or Read</span></span>|

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="6b3ac-165">EventType: String</span><span class="sxs-lookup"><span data-stu-id="6b3ac-165">EventType: String</span></span>

<span data-ttu-id="6b3ac-166">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="6b3ac-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="6b3ac-167">型</span><span class="sxs-lookup"><span data-stu-id="6b3ac-167">Type</span></span>

*   <span data-ttu-id="6b3ac-168">String</span><span class="sxs-lookup"><span data-stu-id="6b3ac-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6b3ac-169">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="6b3ac-169">Properties:</span></span>

| <span data-ttu-id="6b3ac-170">名前</span><span class="sxs-lookup"><span data-stu-id="6b3ac-170">Name</span></span> | <span data-ttu-id="6b3ac-171">種類</span><span class="sxs-lookup"><span data-stu-id="6b3ac-171">Type</span></span> | <span data-ttu-id="6b3ac-172">説明</span><span class="sxs-lookup"><span data-stu-id="6b3ac-172">Description</span></span> | <span data-ttu-id="6b3ac-173">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="6b3ac-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="6b3ac-174">String</span><span class="sxs-lookup"><span data-stu-id="6b3ac-174">String</span></span> | <span data-ttu-id="6b3ac-175">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="6b3ac-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="6b3ac-176">1.7</span><span class="sxs-lookup"><span data-stu-id="6b3ac-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="6b3ac-177">String</span><span class="sxs-lookup"><span data-stu-id="6b3ac-177">String</span></span> | <span data-ttu-id="6b3ac-178">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="6b3ac-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="6b3ac-179">1.5</span><span class="sxs-lookup"><span data-stu-id="6b3ac-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="6b3ac-180">String</span><span class="sxs-lookup"><span data-stu-id="6b3ac-180">String</span></span> | <span data-ttu-id="6b3ac-181">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="6b3ac-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="6b3ac-182">1.7</span><span class="sxs-lookup"><span data-stu-id="6b3ac-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="6b3ac-183">String</span><span class="sxs-lookup"><span data-stu-id="6b3ac-183">String</span></span> | <span data-ttu-id="6b3ac-184">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="6b3ac-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="6b3ac-185">1.7</span><span class="sxs-lookup"><span data-stu-id="6b3ac-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6b3ac-186">要件</span><span class="sxs-lookup"><span data-stu-id="6b3ac-186">Requirements</span></span>

|<span data-ttu-id="6b3ac-187">要件</span><span class="sxs-lookup"><span data-stu-id="6b3ac-187">Requirement</span></span>| <span data-ttu-id="6b3ac-188">値</span><span class="sxs-lookup"><span data-stu-id="6b3ac-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b3ac-189">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6b3ac-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b3ac-190">1.5</span><span class="sxs-lookup"><span data-stu-id="6b3ac-190">1.5</span></span> |
|[<span data-ttu-id="6b3ac-191">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6b3ac-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b3ac-192">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6b3ac-192">Compose or Read</span></span> |

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="6b3ac-193">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="6b3ac-193">SourceProperty: String</span></span>

<span data-ttu-id="6b3ac-194">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="6b3ac-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6b3ac-195">型</span><span class="sxs-lookup"><span data-stu-id="6b3ac-195">Type</span></span>

*   <span data-ttu-id="6b3ac-196">String</span><span class="sxs-lookup"><span data-stu-id="6b3ac-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6b3ac-197">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="6b3ac-197">Properties:</span></span>

|<span data-ttu-id="6b3ac-198">名前</span><span class="sxs-lookup"><span data-stu-id="6b3ac-198">Name</span></span>| <span data-ttu-id="6b3ac-199">種類</span><span class="sxs-lookup"><span data-stu-id="6b3ac-199">Type</span></span>| <span data-ttu-id="6b3ac-200">説明</span><span class="sxs-lookup"><span data-stu-id="6b3ac-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="6b3ac-201">String</span><span class="sxs-lookup"><span data-stu-id="6b3ac-201">String</span></span>|<span data-ttu-id="6b3ac-202">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="6b3ac-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="6b3ac-203">String</span><span class="sxs-lookup"><span data-stu-id="6b3ac-203">String</span></span>|<span data-ttu-id="6b3ac-204">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="6b3ac-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6b3ac-205">要件</span><span class="sxs-lookup"><span data-stu-id="6b3ac-205">Requirements</span></span>

|<span data-ttu-id="6b3ac-206">要件</span><span class="sxs-lookup"><span data-stu-id="6b3ac-206">Requirement</span></span>| <span data-ttu-id="6b3ac-207">値</span><span class="sxs-lookup"><span data-stu-id="6b3ac-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b3ac-208">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6b3ac-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b3ac-209">1.0</span><span class="sxs-lookup"><span data-stu-id="6b3ac-209">1.0</span></span>|
|[<span data-ttu-id="6b3ac-210">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6b3ac-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b3ac-211">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6b3ac-211">Compose or Read</span></span>|
