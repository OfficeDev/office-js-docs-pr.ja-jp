---
title: Office 名前空間-要件セット1.6
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 84e8fa49e1d4dce4239525badafaa051325bb3ec
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395639"
---
# <a name="office"></a><span data-ttu-id="6efa2-102">Office</span><span class="sxs-lookup"><span data-stu-id="6efa2-102">Office</span></span>

<span data-ttu-id="6efa2-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6efa2-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6efa2-105">要件</span><span class="sxs-lookup"><span data-stu-id="6efa2-105">Requirements</span></span>

|<span data-ttu-id="6efa2-106">要件</span><span class="sxs-lookup"><span data-stu-id="6efa2-106">Requirement</span></span>| <span data-ttu-id="6efa2-107">値</span><span class="sxs-lookup"><span data-stu-id="6efa2-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="6efa2-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6efa2-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6efa2-109">1.0</span><span class="sxs-lookup"><span data-stu-id="6efa2-109">1.0</span></span>|
|[<span data-ttu-id="6efa2-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6efa2-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6efa2-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6efa2-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="6efa2-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="6efa2-112">Members and methods</span></span>

| <span data-ttu-id="6efa2-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="6efa2-113">Member</span></span> | <span data-ttu-id="6efa2-114">種類</span><span class="sxs-lookup"><span data-stu-id="6efa2-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="6efa2-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="6efa2-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="6efa2-116">Member</span><span class="sxs-lookup"><span data-stu-id="6efa2-116">Member</span></span> |
| [<span data-ttu-id="6efa2-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="6efa2-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="6efa2-118">Member</span><span class="sxs-lookup"><span data-stu-id="6efa2-118">Member</span></span> |
| [<span data-ttu-id="6efa2-119">EventType</span><span class="sxs-lookup"><span data-stu-id="6efa2-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="6efa2-120">Member</span><span class="sxs-lookup"><span data-stu-id="6efa2-120">Member</span></span> |
| [<span data-ttu-id="6efa2-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="6efa2-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="6efa2-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="6efa2-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="6efa2-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="6efa2-123">Namespaces</span></span>

<span data-ttu-id="6efa2-124">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="6efa2-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="6efa2-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6):、、、、、、などのさまざま`ItemType`な`EntityType`列挙`AttachmentType` `RecipientType` `ResponseType`値が含まれ`ItemNotificationMessageType`ています。</span><span class="sxs-lookup"><span data-stu-id="6efa2-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="6efa2-126">Members</span><span class="sxs-lookup"><span data-stu-id="6efa2-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="6efa2-127">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="6efa2-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="6efa2-128">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="6efa2-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="6efa2-129">型</span><span class="sxs-lookup"><span data-stu-id="6efa2-129">Type</span></span>

*   <span data-ttu-id="6efa2-130">String</span><span class="sxs-lookup"><span data-stu-id="6efa2-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6efa2-131">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="6efa2-131">Properties:</span></span>

|<span data-ttu-id="6efa2-132">名前</span><span class="sxs-lookup"><span data-stu-id="6efa2-132">Name</span></span>| <span data-ttu-id="6efa2-133">種類</span><span class="sxs-lookup"><span data-stu-id="6efa2-133">Type</span></span>| <span data-ttu-id="6efa2-134">説明</span><span class="sxs-lookup"><span data-stu-id="6efa2-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="6efa2-135">String</span><span class="sxs-lookup"><span data-stu-id="6efa2-135">String</span></span>|<span data-ttu-id="6efa2-136">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="6efa2-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="6efa2-137">String</span><span class="sxs-lookup"><span data-stu-id="6efa2-137">String</span></span>|<span data-ttu-id="6efa2-138">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="6efa2-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6efa2-139">要件</span><span class="sxs-lookup"><span data-stu-id="6efa2-139">Requirements</span></span>

|<span data-ttu-id="6efa2-140">要件</span><span class="sxs-lookup"><span data-stu-id="6efa2-140">Requirement</span></span>| <span data-ttu-id="6efa2-141">値</span><span class="sxs-lookup"><span data-stu-id="6efa2-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="6efa2-142">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6efa2-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6efa2-143">1.0</span><span class="sxs-lookup"><span data-stu-id="6efa2-143">1.0</span></span>|
|[<span data-ttu-id="6efa2-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6efa2-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6efa2-145">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6efa2-145">Compose or Read</span></span>|

---

#### <a name="coerciontype-string"></a><span data-ttu-id="6efa2-146">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="6efa2-146">CoercionType: String</span></span>

<span data-ttu-id="6efa2-147">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="6efa2-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6efa2-148">型</span><span class="sxs-lookup"><span data-stu-id="6efa2-148">Type</span></span>

*   <span data-ttu-id="6efa2-149">String</span><span class="sxs-lookup"><span data-stu-id="6efa2-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6efa2-150">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="6efa2-150">Properties:</span></span>

|<span data-ttu-id="6efa2-151">名前</span><span class="sxs-lookup"><span data-stu-id="6efa2-151">Name</span></span>| <span data-ttu-id="6efa2-152">種類</span><span class="sxs-lookup"><span data-stu-id="6efa2-152">Type</span></span>| <span data-ttu-id="6efa2-153">説明</span><span class="sxs-lookup"><span data-stu-id="6efa2-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="6efa2-154">String</span><span class="sxs-lookup"><span data-stu-id="6efa2-154">String</span></span>|<span data-ttu-id="6efa2-155">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="6efa2-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="6efa2-156">String</span><span class="sxs-lookup"><span data-stu-id="6efa2-156">String</span></span>|<span data-ttu-id="6efa2-157">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="6efa2-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6efa2-158">要件</span><span class="sxs-lookup"><span data-stu-id="6efa2-158">Requirements</span></span>

|<span data-ttu-id="6efa2-159">要件</span><span class="sxs-lookup"><span data-stu-id="6efa2-159">Requirement</span></span>| <span data-ttu-id="6efa2-160">値</span><span class="sxs-lookup"><span data-stu-id="6efa2-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="6efa2-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6efa2-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6efa2-162">1.0</span><span class="sxs-lookup"><span data-stu-id="6efa2-162">1.0</span></span>|
|[<span data-ttu-id="6efa2-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6efa2-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6efa2-164">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6efa2-164">Compose or Read</span></span>|

---

#### <a name="eventtype-string"></a><span data-ttu-id="6efa2-165">EventType: String</span><span class="sxs-lookup"><span data-stu-id="6efa2-165">EventType: String</span></span>

<span data-ttu-id="6efa2-166">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="6efa2-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="6efa2-167">型</span><span class="sxs-lookup"><span data-stu-id="6efa2-167">Type</span></span>

*   <span data-ttu-id="6efa2-168">String</span><span class="sxs-lookup"><span data-stu-id="6efa2-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6efa2-169">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="6efa2-169">Properties:</span></span>

| <span data-ttu-id="6efa2-170">名前</span><span class="sxs-lookup"><span data-stu-id="6efa2-170">Name</span></span> | <span data-ttu-id="6efa2-171">種類</span><span class="sxs-lookup"><span data-stu-id="6efa2-171">Type</span></span> | <span data-ttu-id="6efa2-172">説明</span><span class="sxs-lookup"><span data-stu-id="6efa2-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="6efa2-173">String</span><span class="sxs-lookup"><span data-stu-id="6efa2-173">String</span></span> | <span data-ttu-id="6efa2-174">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="6efa2-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6efa2-175">要件</span><span class="sxs-lookup"><span data-stu-id="6efa2-175">Requirements</span></span>

|<span data-ttu-id="6efa2-176">要件</span><span class="sxs-lookup"><span data-stu-id="6efa2-176">Requirement</span></span>| <span data-ttu-id="6efa2-177">値</span><span class="sxs-lookup"><span data-stu-id="6efa2-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="6efa2-178">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6efa2-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6efa2-179">1.5</span><span class="sxs-lookup"><span data-stu-id="6efa2-179">1.5</span></span> |
|[<span data-ttu-id="6efa2-180">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6efa2-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6efa2-181">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6efa2-181">Compose or Read</span></span> |

---

#### <a name="sourceproperty-string"></a><span data-ttu-id="6efa2-182">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="6efa2-182">SourceProperty: String</span></span>

<span data-ttu-id="6efa2-183">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="6efa2-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6efa2-184">型</span><span class="sxs-lookup"><span data-stu-id="6efa2-184">Type</span></span>

*   <span data-ttu-id="6efa2-185">String</span><span class="sxs-lookup"><span data-stu-id="6efa2-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6efa2-186">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="6efa2-186">Properties:</span></span>

|<span data-ttu-id="6efa2-187">名前</span><span class="sxs-lookup"><span data-stu-id="6efa2-187">Name</span></span>| <span data-ttu-id="6efa2-188">種類</span><span class="sxs-lookup"><span data-stu-id="6efa2-188">Type</span></span>| <span data-ttu-id="6efa2-189">説明</span><span class="sxs-lookup"><span data-stu-id="6efa2-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="6efa2-190">String</span><span class="sxs-lookup"><span data-stu-id="6efa2-190">String</span></span>|<span data-ttu-id="6efa2-191">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="6efa2-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="6efa2-192">String</span><span class="sxs-lookup"><span data-stu-id="6efa2-192">String</span></span>|<span data-ttu-id="6efa2-193">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="6efa2-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6efa2-194">要件</span><span class="sxs-lookup"><span data-stu-id="6efa2-194">Requirements</span></span>

|<span data-ttu-id="6efa2-195">要件</span><span class="sxs-lookup"><span data-stu-id="6efa2-195">Requirement</span></span>| <span data-ttu-id="6efa2-196">値</span><span class="sxs-lookup"><span data-stu-id="6efa2-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="6efa2-197">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6efa2-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6efa2-198">1.0</span><span class="sxs-lookup"><span data-stu-id="6efa2-198">1.0</span></span>|
|[<span data-ttu-id="6efa2-199">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6efa2-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6efa2-200">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6efa2-200">Compose or Read</span></span>|
