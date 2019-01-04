---
title: Office 名前空間 - 要件セット 1.6
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: bf6304515c511eea580a3f37d898b7e80adffaee
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457895"
---
# <a name="office"></a><span data-ttu-id="8e263-102">Office</span><span class="sxs-lookup"><span data-stu-id="8e263-102">Office</span></span>

<span data-ttu-id="8e263-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8e263-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e263-105">要件</span><span class="sxs-lookup"><span data-stu-id="8e263-105">Requirements</span></span>

|<span data-ttu-id="8e263-106">要件</span><span class="sxs-lookup"><span data-stu-id="8e263-106">Requirement</span></span>| <span data-ttu-id="8e263-107">値</span><span class="sxs-lookup"><span data-stu-id="8e263-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e263-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e263-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e263-109">1.0</span><span class="sxs-lookup"><span data-stu-id="8e263-109">1.0</span></span>|
|[<span data-ttu-id="8e263-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e263-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e263-111">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e263-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8e263-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="8e263-112">Members and methods</span></span>

| <span data-ttu-id="8e263-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e263-113">Member</span></span> | <span data-ttu-id="8e263-114">種類</span><span class="sxs-lookup"><span data-stu-id="8e263-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8e263-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="8e263-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="8e263-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e263-116">Member</span></span> |
| [<span data-ttu-id="8e263-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="8e263-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="8e263-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e263-118">Member</span></span> |
| [<span data-ttu-id="8e263-119">EventType</span><span class="sxs-lookup"><span data-stu-id="8e263-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="8e263-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e263-120">Member</span></span> |
| [<span data-ttu-id="8e263-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="8e263-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="8e263-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e263-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="8e263-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="8e263-123">Namespaces</span></span>

<span data-ttu-id="8e263-124">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8e263-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="8e263-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="8e263-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="8e263-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e263-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="8e263-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="8e263-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="8e263-128">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="8e263-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="8e263-129">型:</span><span class="sxs-lookup"><span data-stu-id="8e263-129">Type:</span></span>

*   <span data-ttu-id="8e263-130">String</span><span class="sxs-lookup"><span data-stu-id="8e263-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8e263-131">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8e263-131">Properties:</span></span>

|<span data-ttu-id="8e263-132">名前</span><span class="sxs-lookup"><span data-stu-id="8e263-132">Name</span></span>| <span data-ttu-id="8e263-133">型</span><span class="sxs-lookup"><span data-stu-id="8e263-133">Type</span></span>| <span data-ttu-id="8e263-134">説明</span><span class="sxs-lookup"><span data-stu-id="8e263-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="8e263-135">String</span><span class="sxs-lookup"><span data-stu-id="8e263-135">String</span></span>|<span data-ttu-id="8e263-136">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="8e263-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="8e263-137">String</span><span class="sxs-lookup"><span data-stu-id="8e263-137">String</span></span>|<span data-ttu-id="8e263-138">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="8e263-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e263-139">要件</span><span class="sxs-lookup"><span data-stu-id="8e263-139">Requirements</span></span>

|<span data-ttu-id="8e263-140">要件</span><span class="sxs-lookup"><span data-stu-id="8e263-140">Requirement</span></span>| <span data-ttu-id="8e263-141">値</span><span class="sxs-lookup"><span data-stu-id="8e263-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e263-142">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e263-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e263-143">1.0</span><span class="sxs-lookup"><span data-stu-id="8e263-143">1.0</span></span>|
|[<span data-ttu-id="8e263-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e263-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e263-145">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e263-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="8e263-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="8e263-146">CoercionType :String</span></span>

<span data-ttu-id="8e263-147">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="8e263-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8e263-148">型:</span><span class="sxs-lookup"><span data-stu-id="8e263-148">Type:</span></span>

*   <span data-ttu-id="8e263-149">String</span><span class="sxs-lookup"><span data-stu-id="8e263-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8e263-150">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8e263-150">Properties:</span></span>

|<span data-ttu-id="8e263-151">名前</span><span class="sxs-lookup"><span data-stu-id="8e263-151">Name</span></span>| <span data-ttu-id="8e263-152">型</span><span class="sxs-lookup"><span data-stu-id="8e263-152">Type</span></span>| <span data-ttu-id="8e263-153">説明</span><span class="sxs-lookup"><span data-stu-id="8e263-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="8e263-154">String</span><span class="sxs-lookup"><span data-stu-id="8e263-154">String</span></span>|<span data-ttu-id="8e263-155">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="8e263-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="8e263-156">String</span><span class="sxs-lookup"><span data-stu-id="8e263-156">String</span></span>|<span data-ttu-id="8e263-157">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="8e263-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e263-158">要件</span><span class="sxs-lookup"><span data-stu-id="8e263-158">Requirements</span></span>

|<span data-ttu-id="8e263-159">要件</span><span class="sxs-lookup"><span data-stu-id="8e263-159">Requirement</span></span>| <span data-ttu-id="8e263-160">値</span><span class="sxs-lookup"><span data-stu-id="8e263-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e263-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e263-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e263-162">1.0</span><span class="sxs-lookup"><span data-stu-id="8e263-162">1.0</span></span>|
|[<span data-ttu-id="8e263-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e263-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e263-164">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e263-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="8e263-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="8e263-165">EventType :String</span></span>

<span data-ttu-id="8e263-166">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="8e263-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="8e263-167">型:</span><span class="sxs-lookup"><span data-stu-id="8e263-167">Type:</span></span>

*   <span data-ttu-id="8e263-168">String</span><span class="sxs-lookup"><span data-stu-id="8e263-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8e263-169">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8e263-169">Properties:</span></span>

| <span data-ttu-id="8e263-170">名前</span><span class="sxs-lookup"><span data-stu-id="8e263-170">Name</span></span> | <span data-ttu-id="8e263-171">型</span><span class="sxs-lookup"><span data-stu-id="8e263-171">Type</span></span> | <span data-ttu-id="8e263-172">説明</span><span class="sxs-lookup"><span data-stu-id="8e263-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="8e263-173">文字列</span><span class="sxs-lookup"><span data-stu-id="8e263-173">String</span></span> | <span data-ttu-id="8e263-174">作業ウィンドウがピン留めされている間、別の Outlook アイテムが選択されて表示されている。</span><span class="sxs-lookup"><span data-stu-id="8e263-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8e263-175">要件</span><span class="sxs-lookup"><span data-stu-id="8e263-175">Requirements</span></span>

|<span data-ttu-id="8e263-176">要件</span><span class="sxs-lookup"><span data-stu-id="8e263-176">Requirement</span></span>| <span data-ttu-id="8e263-177">値</span><span class="sxs-lookup"><span data-stu-id="8e263-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e263-178">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e263-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e263-179">1.5</span><span class="sxs-lookup"><span data-stu-id="8e263-179">1.5</span></span> |
|[<span data-ttu-id="8e263-180">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e263-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e263-181">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e263-181">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="8e263-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="8e263-182">SourceProperty :String</span></span>

<span data-ttu-id="8e263-183">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="8e263-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8e263-184">型:</span><span class="sxs-lookup"><span data-stu-id="8e263-184">Type:</span></span>

*   <span data-ttu-id="8e263-185">String</span><span class="sxs-lookup"><span data-stu-id="8e263-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8e263-186">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8e263-186">Properties:</span></span>

|<span data-ttu-id="8e263-187">名前</span><span class="sxs-lookup"><span data-stu-id="8e263-187">Name</span></span>| <span data-ttu-id="8e263-188">型</span><span class="sxs-lookup"><span data-stu-id="8e263-188">Type</span></span>| <span data-ttu-id="8e263-189">説明</span><span class="sxs-lookup"><span data-stu-id="8e263-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="8e263-190">String</span><span class="sxs-lookup"><span data-stu-id="8e263-190">String</span></span>|<span data-ttu-id="8e263-191">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="8e263-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="8e263-192">String</span><span class="sxs-lookup"><span data-stu-id="8e263-192">String</span></span>|<span data-ttu-id="8e263-193">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="8e263-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e263-194">要件</span><span class="sxs-lookup"><span data-stu-id="8e263-194">Requirements</span></span>

|<span data-ttu-id="8e263-195">要件</span><span class="sxs-lookup"><span data-stu-id="8e263-195">Requirement</span></span>| <span data-ttu-id="8e263-196">値</span><span class="sxs-lookup"><span data-stu-id="8e263-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e263-197">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e263-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e263-198">1.0</span><span class="sxs-lookup"><span data-stu-id="8e263-198">1.0</span></span>|
|[<span data-ttu-id="8e263-199">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e263-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e263-200">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e263-200">Compose or read</span></span>|