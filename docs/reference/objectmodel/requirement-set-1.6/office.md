---
title: Office 名前空間-要件セット1.6
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: e211a3a2983567b79b73a791914f8d4ed1501ab1
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064664"
---
# <a name="office"></a><span data-ttu-id="fb911-102">Office</span><span class="sxs-lookup"><span data-stu-id="fb911-102">Office</span></span>

<span data-ttu-id="fb911-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fb911-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="fb911-105">要件</span><span class="sxs-lookup"><span data-stu-id="fb911-105">Requirements</span></span>

|<span data-ttu-id="fb911-106">要件</span><span class="sxs-lookup"><span data-stu-id="fb911-106">Requirement</span></span>| <span data-ttu-id="fb911-107">値</span><span class="sxs-lookup"><span data-stu-id="fb911-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="fb911-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fb911-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fb911-109">1.0</span><span class="sxs-lookup"><span data-stu-id="fb911-109">1.0</span></span>|
|[<span data-ttu-id="fb911-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fb911-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fb911-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="fb911-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="fb911-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="fb911-112">Members and methods</span></span>

| <span data-ttu-id="fb911-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="fb911-113">Member</span></span> | <span data-ttu-id="fb911-114">種類</span><span class="sxs-lookup"><span data-stu-id="fb911-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="fb911-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="fb911-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="fb911-116">Member</span><span class="sxs-lookup"><span data-stu-id="fb911-116">Member</span></span> |
| [<span data-ttu-id="fb911-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="fb911-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="fb911-118">Member</span><span class="sxs-lookup"><span data-stu-id="fb911-118">Member</span></span> |
| [<span data-ttu-id="fb911-119">EventType</span><span class="sxs-lookup"><span data-stu-id="fb911-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="fb911-120">Member</span><span class="sxs-lookup"><span data-stu-id="fb911-120">Member</span></span> |
| [<span data-ttu-id="fb911-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="fb911-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="fb911-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="fb911-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="fb911-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="fb911-123">Namespaces</span></span>

<span data-ttu-id="fb911-124">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="fb911-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="fb911-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="fb911-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="fb911-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="fb911-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="fb911-127">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="fb911-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="fb911-128">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="fb911-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="fb911-129">型</span><span class="sxs-lookup"><span data-stu-id="fb911-129">Type</span></span>

*   <span data-ttu-id="fb911-130">String</span><span class="sxs-lookup"><span data-stu-id="fb911-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fb911-131">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="fb911-131">Properties:</span></span>

|<span data-ttu-id="fb911-132">名前</span><span class="sxs-lookup"><span data-stu-id="fb911-132">Name</span></span>| <span data-ttu-id="fb911-133">種類</span><span class="sxs-lookup"><span data-stu-id="fb911-133">Type</span></span>| <span data-ttu-id="fb911-134">説明</span><span class="sxs-lookup"><span data-stu-id="fb911-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="fb911-135">String</span><span class="sxs-lookup"><span data-stu-id="fb911-135">String</span></span>|<span data-ttu-id="fb911-136">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="fb911-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="fb911-137">String</span><span class="sxs-lookup"><span data-stu-id="fb911-137">String</span></span>|<span data-ttu-id="fb911-138">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="fb911-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fb911-139">要件</span><span class="sxs-lookup"><span data-stu-id="fb911-139">Requirements</span></span>

|<span data-ttu-id="fb911-140">要件</span><span class="sxs-lookup"><span data-stu-id="fb911-140">Requirement</span></span>| <span data-ttu-id="fb911-141">値</span><span class="sxs-lookup"><span data-stu-id="fb911-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="fb911-142">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fb911-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fb911-143">1.0</span><span class="sxs-lookup"><span data-stu-id="fb911-143">1.0</span></span>|
|[<span data-ttu-id="fb911-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fb911-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fb911-145">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="fb911-145">Compose or Read</span></span>|

---

#### <a name="coerciontype-string"></a><span data-ttu-id="fb911-146">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="fb911-146">CoercionType: String</span></span>

<span data-ttu-id="fb911-147">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="fb911-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="fb911-148">型</span><span class="sxs-lookup"><span data-stu-id="fb911-148">Type</span></span>

*   <span data-ttu-id="fb911-149">String</span><span class="sxs-lookup"><span data-stu-id="fb911-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fb911-150">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="fb911-150">Properties:</span></span>

|<span data-ttu-id="fb911-151">名前</span><span class="sxs-lookup"><span data-stu-id="fb911-151">Name</span></span>| <span data-ttu-id="fb911-152">種類</span><span class="sxs-lookup"><span data-stu-id="fb911-152">Type</span></span>| <span data-ttu-id="fb911-153">説明</span><span class="sxs-lookup"><span data-stu-id="fb911-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="fb911-154">String</span><span class="sxs-lookup"><span data-stu-id="fb911-154">String</span></span>|<span data-ttu-id="fb911-155">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="fb911-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="fb911-156">String</span><span class="sxs-lookup"><span data-stu-id="fb911-156">String</span></span>|<span data-ttu-id="fb911-157">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="fb911-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fb911-158">要件</span><span class="sxs-lookup"><span data-stu-id="fb911-158">Requirements</span></span>

|<span data-ttu-id="fb911-159">要件</span><span class="sxs-lookup"><span data-stu-id="fb911-159">Requirement</span></span>| <span data-ttu-id="fb911-160">値</span><span class="sxs-lookup"><span data-stu-id="fb911-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="fb911-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fb911-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fb911-162">1.0</span><span class="sxs-lookup"><span data-stu-id="fb911-162">1.0</span></span>|
|[<span data-ttu-id="fb911-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fb911-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fb911-164">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="fb911-164">Compose or Read</span></span>|

---

#### <a name="eventtype-string"></a><span data-ttu-id="fb911-165">EventType: String</span><span class="sxs-lookup"><span data-stu-id="fb911-165">EventType: String</span></span>

<span data-ttu-id="fb911-166">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="fb911-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="fb911-167">型</span><span class="sxs-lookup"><span data-stu-id="fb911-167">Type</span></span>

*   <span data-ttu-id="fb911-168">String</span><span class="sxs-lookup"><span data-stu-id="fb911-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fb911-169">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="fb911-169">Properties:</span></span>

| <span data-ttu-id="fb911-170">名前</span><span class="sxs-lookup"><span data-stu-id="fb911-170">Name</span></span> | <span data-ttu-id="fb911-171">種類</span><span class="sxs-lookup"><span data-stu-id="fb911-171">Type</span></span> | <span data-ttu-id="fb911-172">説明</span><span class="sxs-lookup"><span data-stu-id="fb911-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="fb911-173">String</span><span class="sxs-lookup"><span data-stu-id="fb911-173">String</span></span> | <span data-ttu-id="fb911-174">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="fb911-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fb911-175">要件</span><span class="sxs-lookup"><span data-stu-id="fb911-175">Requirements</span></span>

|<span data-ttu-id="fb911-176">要件</span><span class="sxs-lookup"><span data-stu-id="fb911-176">Requirement</span></span>| <span data-ttu-id="fb911-177">値</span><span class="sxs-lookup"><span data-stu-id="fb911-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="fb911-178">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fb911-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fb911-179">1.5</span><span class="sxs-lookup"><span data-stu-id="fb911-179">1.5</span></span> |
|[<span data-ttu-id="fb911-180">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fb911-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fb911-181">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="fb911-181">Compose or Read</span></span> |

---

#### <a name="sourceproperty-string"></a><span data-ttu-id="fb911-182">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="fb911-182">SourceProperty: String</span></span>

<span data-ttu-id="fb911-183">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="fb911-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="fb911-184">型</span><span class="sxs-lookup"><span data-stu-id="fb911-184">Type</span></span>

*   <span data-ttu-id="fb911-185">String</span><span class="sxs-lookup"><span data-stu-id="fb911-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fb911-186">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="fb911-186">Properties:</span></span>

|<span data-ttu-id="fb911-187">名前</span><span class="sxs-lookup"><span data-stu-id="fb911-187">Name</span></span>| <span data-ttu-id="fb911-188">種類</span><span class="sxs-lookup"><span data-stu-id="fb911-188">Type</span></span>| <span data-ttu-id="fb911-189">説明</span><span class="sxs-lookup"><span data-stu-id="fb911-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="fb911-190">String</span><span class="sxs-lookup"><span data-stu-id="fb911-190">String</span></span>|<span data-ttu-id="fb911-191">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="fb911-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="fb911-192">String</span><span class="sxs-lookup"><span data-stu-id="fb911-192">String</span></span>|<span data-ttu-id="fb911-193">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="fb911-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fb911-194">要件</span><span class="sxs-lookup"><span data-stu-id="fb911-194">Requirements</span></span>

|<span data-ttu-id="fb911-195">要件</span><span class="sxs-lookup"><span data-stu-id="fb911-195">Requirement</span></span>| <span data-ttu-id="fb911-196">値</span><span class="sxs-lookup"><span data-stu-id="fb911-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="fb911-197">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fb911-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fb911-198">1.0</span><span class="sxs-lookup"><span data-stu-id="fb911-198">1.0</span></span>|
|[<span data-ttu-id="fb911-199">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fb911-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fb911-200">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="fb911-200">Compose or Read</span></span>|
