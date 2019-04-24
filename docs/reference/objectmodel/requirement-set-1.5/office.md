---
title: Office 名前空間-要件セット1.5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d8c51646818681629fa0c184962776beffe22a55
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450262"
---
# <a name="office"></a><span data-ttu-id="ae016-102">Office</span><span class="sxs-lookup"><span data-stu-id="ae016-102">Office</span></span>

<span data-ttu-id="ae016-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ae016-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae016-105">要件</span><span class="sxs-lookup"><span data-stu-id="ae016-105">Requirements</span></span>

|<span data-ttu-id="ae016-106">要件</span><span class="sxs-lookup"><span data-stu-id="ae016-106">Requirement</span></span>| <span data-ttu-id="ae016-107">値</span><span class="sxs-lookup"><span data-stu-id="ae016-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae016-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ae016-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae016-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ae016-109">1.0</span></span>|
|[<span data-ttu-id="ae016-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ae016-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ae016-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ae016-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ae016-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="ae016-112">Members and methods</span></span>

| <span data-ttu-id="ae016-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="ae016-113">Member</span></span> | <span data-ttu-id="ae016-114">種類</span><span class="sxs-lookup"><span data-stu-id="ae016-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ae016-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="ae016-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="ae016-116">Member</span><span class="sxs-lookup"><span data-stu-id="ae016-116">Member</span></span> |
| [<span data-ttu-id="ae016-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="ae016-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="ae016-118">Member</span><span class="sxs-lookup"><span data-stu-id="ae016-118">Member</span></span> |
| [<span data-ttu-id="ae016-119">EventType</span><span class="sxs-lookup"><span data-stu-id="ae016-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="ae016-120">Member</span><span class="sxs-lookup"><span data-stu-id="ae016-120">Member</span></span> |
| [<span data-ttu-id="ae016-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="ae016-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="ae016-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="ae016-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="ae016-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="ae016-123">Namespaces</span></span>

<span data-ttu-id="ae016-124">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ae016-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="ae016-125">[MailboxEnums](/javascript/api/outlook_1_5/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="ae016-125">[MailboxEnums](/javascript/api/outlook_1_5/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="ae016-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="ae016-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="ae016-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="ae016-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="ae016-128">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="ae016-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="ae016-129">型</span><span class="sxs-lookup"><span data-stu-id="ae016-129">Type</span></span>

*   <span data-ttu-id="ae016-130">String</span><span class="sxs-lookup"><span data-stu-id="ae016-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ae016-131">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="ae016-131">Properties:</span></span>

|<span data-ttu-id="ae016-132">名前</span><span class="sxs-lookup"><span data-stu-id="ae016-132">Name</span></span>| <span data-ttu-id="ae016-133">種類</span><span class="sxs-lookup"><span data-stu-id="ae016-133">Type</span></span>| <span data-ttu-id="ae016-134">説明</span><span class="sxs-lookup"><span data-stu-id="ae016-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="ae016-135">String</span><span class="sxs-lookup"><span data-stu-id="ae016-135">String</span></span>|<span data-ttu-id="ae016-136">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="ae016-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="ae016-137">String</span><span class="sxs-lookup"><span data-stu-id="ae016-137">String</span></span>|<span data-ttu-id="ae016-138">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="ae016-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ae016-139">要件</span><span class="sxs-lookup"><span data-stu-id="ae016-139">Requirements</span></span>

|<span data-ttu-id="ae016-140">要件</span><span class="sxs-lookup"><span data-stu-id="ae016-140">Requirement</span></span>| <span data-ttu-id="ae016-141">値</span><span class="sxs-lookup"><span data-stu-id="ae016-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae016-142">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ae016-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae016-143">1.0</span><span class="sxs-lookup"><span data-stu-id="ae016-143">1.0</span></span>|
|[<span data-ttu-id="ae016-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ae016-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ae016-145">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ae016-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="ae016-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="ae016-146">CoercionType :String</span></span>

<span data-ttu-id="ae016-147">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="ae016-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ae016-148">型</span><span class="sxs-lookup"><span data-stu-id="ae016-148">Type</span></span>

*   <span data-ttu-id="ae016-149">String</span><span class="sxs-lookup"><span data-stu-id="ae016-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ae016-150">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="ae016-150">Properties:</span></span>

|<span data-ttu-id="ae016-151">名前</span><span class="sxs-lookup"><span data-stu-id="ae016-151">Name</span></span>| <span data-ttu-id="ae016-152">種類</span><span class="sxs-lookup"><span data-stu-id="ae016-152">Type</span></span>| <span data-ttu-id="ae016-153">説明</span><span class="sxs-lookup"><span data-stu-id="ae016-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="ae016-154">String</span><span class="sxs-lookup"><span data-stu-id="ae016-154">String</span></span>|<span data-ttu-id="ae016-155">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="ae016-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="ae016-156">String</span><span class="sxs-lookup"><span data-stu-id="ae016-156">String</span></span>|<span data-ttu-id="ae016-157">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="ae016-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ae016-158">要件</span><span class="sxs-lookup"><span data-stu-id="ae016-158">Requirements</span></span>

|<span data-ttu-id="ae016-159">要件</span><span class="sxs-lookup"><span data-stu-id="ae016-159">Requirement</span></span>| <span data-ttu-id="ae016-160">値</span><span class="sxs-lookup"><span data-stu-id="ae016-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae016-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ae016-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae016-162">1.0</span><span class="sxs-lookup"><span data-stu-id="ae016-162">1.0</span></span>|
|[<span data-ttu-id="ae016-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ae016-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ae016-164">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ae016-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="ae016-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="ae016-165">EventType :String</span></span>

<span data-ttu-id="ae016-166">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="ae016-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="ae016-167">型</span><span class="sxs-lookup"><span data-stu-id="ae016-167">Type</span></span>

*   <span data-ttu-id="ae016-168">String</span><span class="sxs-lookup"><span data-stu-id="ae016-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ae016-169">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="ae016-169">Properties:</span></span>

| <span data-ttu-id="ae016-170">名前</span><span class="sxs-lookup"><span data-stu-id="ae016-170">Name</span></span> | <span data-ttu-id="ae016-171">種類</span><span class="sxs-lookup"><span data-stu-id="ae016-171">Type</span></span> | <span data-ttu-id="ae016-172">説明</span><span class="sxs-lookup"><span data-stu-id="ae016-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="ae016-173">String</span><span class="sxs-lookup"><span data-stu-id="ae016-173">String</span></span> | <span data-ttu-id="ae016-174">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="ae016-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ae016-175">要件</span><span class="sxs-lookup"><span data-stu-id="ae016-175">Requirements</span></span>

|<span data-ttu-id="ae016-176">要件</span><span class="sxs-lookup"><span data-stu-id="ae016-176">Requirement</span></span>| <span data-ttu-id="ae016-177">値</span><span class="sxs-lookup"><span data-stu-id="ae016-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae016-178">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ae016-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae016-179">1.5</span><span class="sxs-lookup"><span data-stu-id="ae016-179">1.5</span></span> |
|[<span data-ttu-id="ae016-180">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ae016-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ae016-181">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ae016-181">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="ae016-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="ae016-182">SourceProperty :String</span></span>

<span data-ttu-id="ae016-183">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="ae016-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ae016-184">型</span><span class="sxs-lookup"><span data-stu-id="ae016-184">Type</span></span>

*   <span data-ttu-id="ae016-185">String</span><span class="sxs-lookup"><span data-stu-id="ae016-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ae016-186">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="ae016-186">Properties:</span></span>

|<span data-ttu-id="ae016-187">名前</span><span class="sxs-lookup"><span data-stu-id="ae016-187">Name</span></span>| <span data-ttu-id="ae016-188">種類</span><span class="sxs-lookup"><span data-stu-id="ae016-188">Type</span></span>| <span data-ttu-id="ae016-189">説明</span><span class="sxs-lookup"><span data-stu-id="ae016-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="ae016-190">String</span><span class="sxs-lookup"><span data-stu-id="ae016-190">String</span></span>|<span data-ttu-id="ae016-191">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="ae016-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="ae016-192">String</span><span class="sxs-lookup"><span data-stu-id="ae016-192">String</span></span>|<span data-ttu-id="ae016-193">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="ae016-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ae016-194">要件</span><span class="sxs-lookup"><span data-stu-id="ae016-194">Requirements</span></span>

|<span data-ttu-id="ae016-195">要件</span><span class="sxs-lookup"><span data-stu-id="ae016-195">Requirement</span></span>| <span data-ttu-id="ae016-196">値</span><span class="sxs-lookup"><span data-stu-id="ae016-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae016-197">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ae016-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae016-198">1.0</span><span class="sxs-lookup"><span data-stu-id="ae016-198">1.0</span></span>|
|[<span data-ttu-id="ae016-199">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ae016-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ae016-200">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ae016-200">Compose or Read</span></span>|
