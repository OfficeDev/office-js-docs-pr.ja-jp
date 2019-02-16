---
title: Office 名前空間 - 要件セット 1.5
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: c9f769550ad2c4994545e51d140b6ea6e67761bc
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067938"
---
# <a name="office"></a><span data-ttu-id="2c662-102">Office</span><span class="sxs-lookup"><span data-stu-id="2c662-102">Office</span></span>

<span data-ttu-id="2c662-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2c662-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="2c662-105">要件</span><span class="sxs-lookup"><span data-stu-id="2c662-105">Requirements</span></span>

|<span data-ttu-id="2c662-106">要件</span><span class="sxs-lookup"><span data-stu-id="2c662-106">Requirement</span></span>| <span data-ttu-id="2c662-107">値</span><span class="sxs-lookup"><span data-stu-id="2c662-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c662-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2c662-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2c662-109">1.0</span><span class="sxs-lookup"><span data-stu-id="2c662-109">1.0</span></span>|
|[<span data-ttu-id="2c662-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2c662-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2c662-111">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="2c662-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2c662-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="2c662-112">Members and methods</span></span>

| <span data-ttu-id="2c662-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="2c662-113">Member</span></span> | <span data-ttu-id="2c662-114">種類</span><span class="sxs-lookup"><span data-stu-id="2c662-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2c662-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="2c662-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="2c662-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="2c662-116">Member</span></span> |
| [<span data-ttu-id="2c662-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="2c662-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="2c662-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="2c662-118">Member</span></span> |
| [<span data-ttu-id="2c662-119">EventType</span><span class="sxs-lookup"><span data-stu-id="2c662-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="2c662-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="2c662-120">Member</span></span> |
| [<span data-ttu-id="2c662-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="2c662-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="2c662-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="2c662-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="2c662-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="2c662-123">Namespaces</span></span>

<span data-ttu-id="2c662-124">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="2c662-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="2c662-125">[MailboxEnums](/javascript/api/outlook_1_5/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="2c662-125">[MailboxEnums](/javascript/api/outlook_1_5/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="2c662-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="2c662-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="2c662-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="2c662-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="2c662-128">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="2c662-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="2c662-129">型</span><span class="sxs-lookup"><span data-stu-id="2c662-129">Type</span></span>

*   <span data-ttu-id="2c662-130">String</span><span class="sxs-lookup"><span data-stu-id="2c662-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2c662-131">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2c662-131">Properties:</span></span>

|<span data-ttu-id="2c662-132">名前</span><span class="sxs-lookup"><span data-stu-id="2c662-132">Name</span></span>| <span data-ttu-id="2c662-133">型</span><span class="sxs-lookup"><span data-stu-id="2c662-133">Type</span></span>| <span data-ttu-id="2c662-134">説明</span><span class="sxs-lookup"><span data-stu-id="2c662-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="2c662-135">String</span><span class="sxs-lookup"><span data-stu-id="2c662-135">String</span></span>|<span data-ttu-id="2c662-136">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="2c662-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="2c662-137">String</span><span class="sxs-lookup"><span data-stu-id="2c662-137">String</span></span>|<span data-ttu-id="2c662-138">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="2c662-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2c662-139">要件</span><span class="sxs-lookup"><span data-stu-id="2c662-139">Requirements</span></span>

|<span data-ttu-id="2c662-140">要件</span><span class="sxs-lookup"><span data-stu-id="2c662-140">Requirement</span></span>| <span data-ttu-id="2c662-141">値</span><span class="sxs-lookup"><span data-stu-id="2c662-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c662-142">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2c662-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2c662-143">1.0</span><span class="sxs-lookup"><span data-stu-id="2c662-143">1.0</span></span>|
|[<span data-ttu-id="2c662-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2c662-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2c662-145">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="2c662-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="2c662-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="2c662-146">CoercionType :String</span></span>

<span data-ttu-id="2c662-147">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="2c662-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2c662-148">型</span><span class="sxs-lookup"><span data-stu-id="2c662-148">Type</span></span>

*   <span data-ttu-id="2c662-149">String</span><span class="sxs-lookup"><span data-stu-id="2c662-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2c662-150">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2c662-150">Properties:</span></span>

|<span data-ttu-id="2c662-151">名前</span><span class="sxs-lookup"><span data-stu-id="2c662-151">Name</span></span>| <span data-ttu-id="2c662-152">型</span><span class="sxs-lookup"><span data-stu-id="2c662-152">Type</span></span>| <span data-ttu-id="2c662-153">説明</span><span class="sxs-lookup"><span data-stu-id="2c662-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="2c662-154">文字列</span><span class="sxs-lookup"><span data-stu-id="2c662-154">String</span></span>|<span data-ttu-id="2c662-155">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="2c662-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="2c662-156">String</span><span class="sxs-lookup"><span data-stu-id="2c662-156">String</span></span>|<span data-ttu-id="2c662-157">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="2c662-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2c662-158">要件</span><span class="sxs-lookup"><span data-stu-id="2c662-158">Requirements</span></span>

|<span data-ttu-id="2c662-159">要件</span><span class="sxs-lookup"><span data-stu-id="2c662-159">Requirement</span></span>| <span data-ttu-id="2c662-160">値</span><span class="sxs-lookup"><span data-stu-id="2c662-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c662-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2c662-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2c662-162">1.0</span><span class="sxs-lookup"><span data-stu-id="2c662-162">1.0</span></span>|
|[<span data-ttu-id="2c662-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2c662-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2c662-164">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="2c662-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="2c662-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="2c662-165">EventType :String</span></span>

<span data-ttu-id="2c662-166">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="2c662-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="2c662-167">型</span><span class="sxs-lookup"><span data-stu-id="2c662-167">Type</span></span>

*   <span data-ttu-id="2c662-168">String</span><span class="sxs-lookup"><span data-stu-id="2c662-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2c662-169">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2c662-169">Properties:</span></span>

| <span data-ttu-id="2c662-170">名前</span><span class="sxs-lookup"><span data-stu-id="2c662-170">Name</span></span> | <span data-ttu-id="2c662-171">型</span><span class="sxs-lookup"><span data-stu-id="2c662-171">Type</span></span> | <span data-ttu-id="2c662-172">説明</span><span class="sxs-lookup"><span data-stu-id="2c662-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="2c662-173">文字列</span><span class="sxs-lookup"><span data-stu-id="2c662-173">String</span></span> | <span data-ttu-id="2c662-174">作業ウィンドウがピン留めされている間、別の Outlook アイテムが選択されて表示されている。</span><span class="sxs-lookup"><span data-stu-id="2c662-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2c662-175">要件</span><span class="sxs-lookup"><span data-stu-id="2c662-175">Requirements</span></span>

|<span data-ttu-id="2c662-176">要件</span><span class="sxs-lookup"><span data-stu-id="2c662-176">Requirement</span></span>| <span data-ttu-id="2c662-177">値</span><span class="sxs-lookup"><span data-stu-id="2c662-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c662-178">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2c662-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2c662-179">1.5</span><span class="sxs-lookup"><span data-stu-id="2c662-179">1.5</span></span> |
|[<span data-ttu-id="2c662-180">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2c662-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2c662-181">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="2c662-181">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="2c662-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="2c662-182">SourceProperty :String</span></span>

<span data-ttu-id="2c662-183">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="2c662-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2c662-184">型</span><span class="sxs-lookup"><span data-stu-id="2c662-184">Type</span></span>

*   <span data-ttu-id="2c662-185">String</span><span class="sxs-lookup"><span data-stu-id="2c662-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2c662-186">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2c662-186">Properties:</span></span>

|<span data-ttu-id="2c662-187">名前</span><span class="sxs-lookup"><span data-stu-id="2c662-187">Name</span></span>| <span data-ttu-id="2c662-188">型</span><span class="sxs-lookup"><span data-stu-id="2c662-188">Type</span></span>| <span data-ttu-id="2c662-189">説明</span><span class="sxs-lookup"><span data-stu-id="2c662-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="2c662-190">文字列</span><span class="sxs-lookup"><span data-stu-id="2c662-190">String</span></span>|<span data-ttu-id="2c662-191">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="2c662-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="2c662-192">String</span><span class="sxs-lookup"><span data-stu-id="2c662-192">String</span></span>|<span data-ttu-id="2c662-193">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="2c662-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2c662-194">要件</span><span class="sxs-lookup"><span data-stu-id="2c662-194">Requirements</span></span>

|<span data-ttu-id="2c662-195">要件</span><span class="sxs-lookup"><span data-stu-id="2c662-195">Requirement</span></span>| <span data-ttu-id="2c662-196">値</span><span class="sxs-lookup"><span data-stu-id="2c662-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c662-197">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2c662-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2c662-198">1.0</span><span class="sxs-lookup"><span data-stu-id="2c662-198">1.0</span></span>|
|[<span data-ttu-id="2c662-199">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2c662-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2c662-200">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="2c662-200">Compose or Read</span></span>|
