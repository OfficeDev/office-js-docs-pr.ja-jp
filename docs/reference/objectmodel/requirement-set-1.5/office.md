---
title: Office 名前空間-要件セット1.5
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 2236dae5421090a571c8cc658cb6f67f2a08d54a
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696079"
---
# <a name="office"></a><span data-ttu-id="700bc-102">Office</span><span class="sxs-lookup"><span data-stu-id="700bc-102">Office</span></span>

<span data-ttu-id="700bc-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="700bc-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="700bc-105">要件</span><span class="sxs-lookup"><span data-stu-id="700bc-105">Requirements</span></span>

|<span data-ttu-id="700bc-106">要件</span><span class="sxs-lookup"><span data-stu-id="700bc-106">Requirement</span></span>| <span data-ttu-id="700bc-107">値</span><span class="sxs-lookup"><span data-stu-id="700bc-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="700bc-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="700bc-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="700bc-109">1.0</span><span class="sxs-lookup"><span data-stu-id="700bc-109">1.0</span></span>|
|[<span data-ttu-id="700bc-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="700bc-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="700bc-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="700bc-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="700bc-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="700bc-112">Members and methods</span></span>

| <span data-ttu-id="700bc-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="700bc-113">Member</span></span> | <span data-ttu-id="700bc-114">種類</span><span class="sxs-lookup"><span data-stu-id="700bc-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="700bc-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="700bc-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="700bc-116">Member</span><span class="sxs-lookup"><span data-stu-id="700bc-116">Member</span></span> |
| [<span data-ttu-id="700bc-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="700bc-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="700bc-118">Member</span><span class="sxs-lookup"><span data-stu-id="700bc-118">Member</span></span> |
| [<span data-ttu-id="700bc-119">EventType</span><span class="sxs-lookup"><span data-stu-id="700bc-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="700bc-120">Member</span><span class="sxs-lookup"><span data-stu-id="700bc-120">Member</span></span> |
| [<span data-ttu-id="700bc-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="700bc-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="700bc-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="700bc-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="700bc-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="700bc-123">Namespaces</span></span>

<span data-ttu-id="700bc-124">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="700bc-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="700bc-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.5):、、、、、、などのさまざま`ItemType`な`EntityType`列挙`AttachmentType` `RecipientType` `ResponseType`値が含まれ`ItemNotificationMessageType`ています。</span><span class="sxs-lookup"><span data-stu-id="700bc-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.5): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="700bc-126">Members</span><span class="sxs-lookup"><span data-stu-id="700bc-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="700bc-127">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="700bc-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="700bc-128">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="700bc-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="700bc-129">型</span><span class="sxs-lookup"><span data-stu-id="700bc-129">Type</span></span>

*   <span data-ttu-id="700bc-130">String</span><span class="sxs-lookup"><span data-stu-id="700bc-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="700bc-131">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="700bc-131">Properties:</span></span>

|<span data-ttu-id="700bc-132">名前</span><span class="sxs-lookup"><span data-stu-id="700bc-132">Name</span></span>| <span data-ttu-id="700bc-133">種類</span><span class="sxs-lookup"><span data-stu-id="700bc-133">Type</span></span>| <span data-ttu-id="700bc-134">説明</span><span class="sxs-lookup"><span data-stu-id="700bc-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="700bc-135">String</span><span class="sxs-lookup"><span data-stu-id="700bc-135">String</span></span>|<span data-ttu-id="700bc-136">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="700bc-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="700bc-137">String</span><span class="sxs-lookup"><span data-stu-id="700bc-137">String</span></span>|<span data-ttu-id="700bc-138">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="700bc-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="700bc-139">要件</span><span class="sxs-lookup"><span data-stu-id="700bc-139">Requirements</span></span>

|<span data-ttu-id="700bc-140">要件</span><span class="sxs-lookup"><span data-stu-id="700bc-140">Requirement</span></span>| <span data-ttu-id="700bc-141">値</span><span class="sxs-lookup"><span data-stu-id="700bc-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="700bc-142">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="700bc-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="700bc-143">1.0</span><span class="sxs-lookup"><span data-stu-id="700bc-143">1.0</span></span>|
|[<span data-ttu-id="700bc-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="700bc-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="700bc-145">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="700bc-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="700bc-146">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="700bc-146">CoercionType: String</span></span>

<span data-ttu-id="700bc-147">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="700bc-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="700bc-148">型</span><span class="sxs-lookup"><span data-stu-id="700bc-148">Type</span></span>

*   <span data-ttu-id="700bc-149">String</span><span class="sxs-lookup"><span data-stu-id="700bc-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="700bc-150">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="700bc-150">Properties:</span></span>

|<span data-ttu-id="700bc-151">名前</span><span class="sxs-lookup"><span data-stu-id="700bc-151">Name</span></span>| <span data-ttu-id="700bc-152">種類</span><span class="sxs-lookup"><span data-stu-id="700bc-152">Type</span></span>| <span data-ttu-id="700bc-153">説明</span><span class="sxs-lookup"><span data-stu-id="700bc-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="700bc-154">String</span><span class="sxs-lookup"><span data-stu-id="700bc-154">String</span></span>|<span data-ttu-id="700bc-155">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="700bc-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="700bc-156">String</span><span class="sxs-lookup"><span data-stu-id="700bc-156">String</span></span>|<span data-ttu-id="700bc-157">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="700bc-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="700bc-158">要件</span><span class="sxs-lookup"><span data-stu-id="700bc-158">Requirements</span></span>

|<span data-ttu-id="700bc-159">要件</span><span class="sxs-lookup"><span data-stu-id="700bc-159">Requirement</span></span>| <span data-ttu-id="700bc-160">値</span><span class="sxs-lookup"><span data-stu-id="700bc-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="700bc-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="700bc-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="700bc-162">1.0</span><span class="sxs-lookup"><span data-stu-id="700bc-162">1.0</span></span>|
|[<span data-ttu-id="700bc-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="700bc-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="700bc-164">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="700bc-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="700bc-165">EventType: String</span><span class="sxs-lookup"><span data-stu-id="700bc-165">EventType: String</span></span>

<span data-ttu-id="700bc-166">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="700bc-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="700bc-167">型</span><span class="sxs-lookup"><span data-stu-id="700bc-167">Type</span></span>

*   <span data-ttu-id="700bc-168">String</span><span class="sxs-lookup"><span data-stu-id="700bc-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="700bc-169">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="700bc-169">Properties:</span></span>

| <span data-ttu-id="700bc-170">名前</span><span class="sxs-lookup"><span data-stu-id="700bc-170">Name</span></span> | <span data-ttu-id="700bc-171">種類</span><span class="sxs-lookup"><span data-stu-id="700bc-171">Type</span></span> | <span data-ttu-id="700bc-172">説明</span><span class="sxs-lookup"><span data-stu-id="700bc-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="700bc-173">String</span><span class="sxs-lookup"><span data-stu-id="700bc-173">String</span></span> | <span data-ttu-id="700bc-174">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="700bc-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="700bc-175">要件</span><span class="sxs-lookup"><span data-stu-id="700bc-175">Requirements</span></span>

|<span data-ttu-id="700bc-176">要件</span><span class="sxs-lookup"><span data-stu-id="700bc-176">Requirement</span></span>| <span data-ttu-id="700bc-177">値</span><span class="sxs-lookup"><span data-stu-id="700bc-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="700bc-178">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="700bc-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="700bc-179">1.5</span><span class="sxs-lookup"><span data-stu-id="700bc-179">1.5</span></span> |
|[<span data-ttu-id="700bc-180">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="700bc-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="700bc-181">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="700bc-181">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="700bc-182">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="700bc-182">SourceProperty: String</span></span>

<span data-ttu-id="700bc-183">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="700bc-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="700bc-184">型</span><span class="sxs-lookup"><span data-stu-id="700bc-184">Type</span></span>

*   <span data-ttu-id="700bc-185">String</span><span class="sxs-lookup"><span data-stu-id="700bc-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="700bc-186">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="700bc-186">Properties:</span></span>

|<span data-ttu-id="700bc-187">名前</span><span class="sxs-lookup"><span data-stu-id="700bc-187">Name</span></span>| <span data-ttu-id="700bc-188">種類</span><span class="sxs-lookup"><span data-stu-id="700bc-188">Type</span></span>| <span data-ttu-id="700bc-189">説明</span><span class="sxs-lookup"><span data-stu-id="700bc-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="700bc-190">String</span><span class="sxs-lookup"><span data-stu-id="700bc-190">String</span></span>|<span data-ttu-id="700bc-191">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="700bc-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="700bc-192">String</span><span class="sxs-lookup"><span data-stu-id="700bc-192">String</span></span>|<span data-ttu-id="700bc-193">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="700bc-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="700bc-194">要件</span><span class="sxs-lookup"><span data-stu-id="700bc-194">Requirements</span></span>

|<span data-ttu-id="700bc-195">要件</span><span class="sxs-lookup"><span data-stu-id="700bc-195">Requirement</span></span>| <span data-ttu-id="700bc-196">値</span><span class="sxs-lookup"><span data-stu-id="700bc-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="700bc-197">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="700bc-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="700bc-198">1.0</span><span class="sxs-lookup"><span data-stu-id="700bc-198">1.0</span></span>|
|[<span data-ttu-id="700bc-199">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="700bc-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="700bc-200">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="700bc-200">Compose or Read</span></span>|
