---
title: Office 名前空間 - 要件セット 1.5
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 46b21df77456d2392fbc543e45513246a4ad9a10
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433650"
---
# <a name="office"></a><span data-ttu-id="9e35d-102">Office</span><span class="sxs-lookup"><span data-stu-id="9e35d-102">Office</span></span>

<span data-ttu-id="9e35d-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共有 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9e35d-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e35d-105">要件</span><span class="sxs-lookup"><span data-stu-id="9e35d-105">Requirements</span></span>

|<span data-ttu-id="9e35d-106">要件</span><span class="sxs-lookup"><span data-stu-id="9e35d-106">Requirement</span></span>| <span data-ttu-id="9e35d-107">値</span><span class="sxs-lookup"><span data-stu-id="9e35d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e35d-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e35d-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e35d-109">1.0</span><span class="sxs-lookup"><span data-stu-id="9e35d-109">1.0</span></span>|
|[<span data-ttu-id="9e35d-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e35d-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e35d-111">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9e35d-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9e35d-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="9e35d-112">Members and methods</span></span>

| <span data-ttu-id="9e35d-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="9e35d-113">Member</span></span> | <span data-ttu-id="9e35d-114">種類</span><span class="sxs-lookup"><span data-stu-id="9e35d-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9e35d-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="9e35d-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="9e35d-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="9e35d-116">Member</span></span> |
| [<span data-ttu-id="9e35d-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="9e35d-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="9e35d-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="9e35d-118">Member</span></span> |
| [<span data-ttu-id="9e35d-119">EventType</span><span class="sxs-lookup"><span data-stu-id="9e35d-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="9e35d-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="9e35d-120">Member</span></span> |
| [<span data-ttu-id="9e35d-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="9e35d-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="9e35d-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="9e35d-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="9e35d-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="9e35d-123">Namespaces</span></span>

<span data-ttu-id="9e35d-124">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="9e35d-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="9e35d-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="9e35d-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="9e35d-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="9e35d-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="9e35d-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="9e35d-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="9e35d-128">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="9e35d-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="9e35d-129">型:</span><span class="sxs-lookup"><span data-stu-id="9e35d-129">Type:</span></span>

*   <span data-ttu-id="9e35d-130">String</span><span class="sxs-lookup"><span data-stu-id="9e35d-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9e35d-131">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="9e35d-131">Properties:</span></span>

|<span data-ttu-id="9e35d-132">名前</span><span class="sxs-lookup"><span data-stu-id="9e35d-132">Name</span></span>| <span data-ttu-id="9e35d-133">型</span><span class="sxs-lookup"><span data-stu-id="9e35d-133">Type</span></span>| <span data-ttu-id="9e35d-134">説明</span><span class="sxs-lookup"><span data-stu-id="9e35d-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="9e35d-135">String</span><span class="sxs-lookup"><span data-stu-id="9e35d-135">String</span></span>|<span data-ttu-id="9e35d-136">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="9e35d-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="9e35d-137">String</span><span class="sxs-lookup"><span data-stu-id="9e35d-137">String</span></span>|<span data-ttu-id="9e35d-138">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="9e35d-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e35d-139">要件</span><span class="sxs-lookup"><span data-stu-id="9e35d-139">Requirements</span></span>

|<span data-ttu-id="9e35d-140">要件</span><span class="sxs-lookup"><span data-stu-id="9e35d-140">Requirement</span></span>| <span data-ttu-id="9e35d-141">値</span><span class="sxs-lookup"><span data-stu-id="9e35d-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e35d-142">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e35d-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e35d-143">1.0</span><span class="sxs-lookup"><span data-stu-id="9e35d-143">1.0</span></span>|
|[<span data-ttu-id="9e35d-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e35d-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e35d-145">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9e35d-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="9e35d-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="9e35d-146">CoercionType :String</span></span>

<span data-ttu-id="9e35d-147">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="9e35d-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9e35d-148">型:</span><span class="sxs-lookup"><span data-stu-id="9e35d-148">Type:</span></span>

*   <span data-ttu-id="9e35d-149">String</span><span class="sxs-lookup"><span data-stu-id="9e35d-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9e35d-150">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="9e35d-150">Properties:</span></span>

|<span data-ttu-id="9e35d-151">名前</span><span class="sxs-lookup"><span data-stu-id="9e35d-151">Name</span></span>| <span data-ttu-id="9e35d-152">型</span><span class="sxs-lookup"><span data-stu-id="9e35d-152">Type</span></span>| <span data-ttu-id="9e35d-153">説明</span><span class="sxs-lookup"><span data-stu-id="9e35d-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="9e35d-154">String</span><span class="sxs-lookup"><span data-stu-id="9e35d-154">String</span></span>|<span data-ttu-id="9e35d-155">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="9e35d-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="9e35d-156">String</span><span class="sxs-lookup"><span data-stu-id="9e35d-156">String</span></span>|<span data-ttu-id="9e35d-157">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="9e35d-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e35d-158">要件</span><span class="sxs-lookup"><span data-stu-id="9e35d-158">Requirements</span></span>

|<span data-ttu-id="9e35d-159">要件</span><span class="sxs-lookup"><span data-stu-id="9e35d-159">Requirement</span></span>| <span data-ttu-id="9e35d-160">値</span><span class="sxs-lookup"><span data-stu-id="9e35d-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e35d-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e35d-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e35d-162">1.0</span><span class="sxs-lookup"><span data-stu-id="9e35d-162">1.0</span></span>|
|[<span data-ttu-id="9e35d-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e35d-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e35d-164">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9e35d-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="9e35d-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="9e35d-165">EventType :String</span></span>

<span data-ttu-id="9e35d-166">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="9e35d-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="9e35d-167">型:</span><span class="sxs-lookup"><span data-stu-id="9e35d-167">Type:</span></span>

*   <span data-ttu-id="9e35d-168">String</span><span class="sxs-lookup"><span data-stu-id="9e35d-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9e35d-169">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="9e35d-169">Properties:</span></span>

| <span data-ttu-id="9e35d-170">名前</span><span class="sxs-lookup"><span data-stu-id="9e35d-170">Name</span></span> | <span data-ttu-id="9e35d-171">型</span><span class="sxs-lookup"><span data-stu-id="9e35d-171">Type</span></span> | <span data-ttu-id="9e35d-172">説明</span><span class="sxs-lookup"><span data-stu-id="9e35d-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="9e35d-173">文字列</span><span class="sxs-lookup"><span data-stu-id="9e35d-173">String</span></span> | <span data-ttu-id="9e35d-174">作業ウィンドウがピン留めされている間、別の Outlook アイテムが選択されて表示されている。</span><span class="sxs-lookup"><span data-stu-id="9e35d-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9e35d-175">要件</span><span class="sxs-lookup"><span data-stu-id="9e35d-175">Requirements</span></span>

|<span data-ttu-id="9e35d-176">要件</span><span class="sxs-lookup"><span data-stu-id="9e35d-176">Requirement</span></span>| <span data-ttu-id="9e35d-177">値</span><span class="sxs-lookup"><span data-stu-id="9e35d-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e35d-178">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e35d-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e35d-179">1.5</span><span class="sxs-lookup"><span data-stu-id="9e35d-179">1.5</span></span> |
|[<span data-ttu-id="9e35d-180">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e35d-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e35d-181">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9e35d-181">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="9e35d-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="9e35d-182">SourceProperty :String</span></span>

<span data-ttu-id="9e35d-183">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="9e35d-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9e35d-184">型:</span><span class="sxs-lookup"><span data-stu-id="9e35d-184">Type:</span></span>

*   <span data-ttu-id="9e35d-185">String</span><span class="sxs-lookup"><span data-stu-id="9e35d-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9e35d-186">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="9e35d-186">Properties:</span></span>

|<span data-ttu-id="9e35d-187">名前</span><span class="sxs-lookup"><span data-stu-id="9e35d-187">Name</span></span>| <span data-ttu-id="9e35d-188">型</span><span class="sxs-lookup"><span data-stu-id="9e35d-188">Type</span></span>| <span data-ttu-id="9e35d-189">説明</span><span class="sxs-lookup"><span data-stu-id="9e35d-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="9e35d-190">String</span><span class="sxs-lookup"><span data-stu-id="9e35d-190">String</span></span>|<span data-ttu-id="9e35d-191">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="9e35d-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="9e35d-192">String</span><span class="sxs-lookup"><span data-stu-id="9e35d-192">String</span></span>|<span data-ttu-id="9e35d-193">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="9e35d-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e35d-194">要件</span><span class="sxs-lookup"><span data-stu-id="9e35d-194">Requirements</span></span>

|<span data-ttu-id="9e35d-195">要件</span><span class="sxs-lookup"><span data-stu-id="9e35d-195">Requirement</span></span>| <span data-ttu-id="9e35d-196">値</span><span class="sxs-lookup"><span data-stu-id="9e35d-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e35d-197">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e35d-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e35d-198">1.0</span><span class="sxs-lookup"><span data-stu-id="9e35d-198">1.0</span></span>|
|[<span data-ttu-id="9e35d-199">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e35d-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e35d-200">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9e35d-200">Compose or read</span></span>|