---
title: Office 名前空間 - プレビュー要件セット
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: bbec602680da7914666daf33ed36c45751ae69c6
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068323"
---
# <a name="office"></a><span data-ttu-id="7e9d5-102">Office</span><span class="sxs-lookup"><span data-stu-id="7e9d5-102">Office</span></span>

<span data-ttu-id="7e9d5-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="7e9d5-105">要件</span><span class="sxs-lookup"><span data-stu-id="7e9d5-105">Requirements</span></span>

|<span data-ttu-id="7e9d5-106">要件</span><span class="sxs-lookup"><span data-stu-id="7e9d5-106">Requirement</span></span>| <span data-ttu-id="7e9d5-107">値</span><span class="sxs-lookup"><span data-stu-id="7e9d5-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e9d5-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7e9d5-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7e9d5-109">1.0</span><span class="sxs-lookup"><span data-stu-id="7e9d5-109">1.0</span></span>|
|[<span data-ttu-id="7e9d5-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7e9d5-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7e9d5-111">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="7e9d5-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="7e9d5-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="7e9d5-112">Members and methods</span></span>

| <span data-ttu-id="7e9d5-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="7e9d5-113">Member</span></span> | <span data-ttu-id="7e9d5-114">種類</span><span class="sxs-lookup"><span data-stu-id="7e9d5-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="7e9d5-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="7e9d5-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="7e9d5-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="7e9d5-116">Member</span></span> |
| [<span data-ttu-id="7e9d5-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="7e9d5-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="7e9d5-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="7e9d5-118">Member</span></span> |
| [<span data-ttu-id="7e9d5-119">EventType</span><span class="sxs-lookup"><span data-stu-id="7e9d5-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="7e9d5-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="7e9d5-120">Member</span></span> |
| [<span data-ttu-id="7e9d5-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="7e9d5-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="7e9d5-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="7e9d5-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="7e9d5-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="7e9d5-123">Namespaces</span></span>

<span data-ttu-id="7e9d5-124">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="7e9d5-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="7e9d5-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="7e9d5-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="7e9d5-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="7e9d5-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="7e9d5-128">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="7e9d5-129">型</span><span class="sxs-lookup"><span data-stu-id="7e9d5-129">Type</span></span>

*   <span data-ttu-id="7e9d5-130">String</span><span class="sxs-lookup"><span data-stu-id="7e9d5-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7e9d5-131">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="7e9d5-131">Properties:</span></span>

|<span data-ttu-id="7e9d5-132">名前</span><span class="sxs-lookup"><span data-stu-id="7e9d5-132">Name</span></span>| <span data-ttu-id="7e9d5-133">型</span><span class="sxs-lookup"><span data-stu-id="7e9d5-133">Type</span></span>| <span data-ttu-id="7e9d5-134">説明</span><span class="sxs-lookup"><span data-stu-id="7e9d5-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="7e9d5-135">String</span><span class="sxs-lookup"><span data-stu-id="7e9d5-135">String</span></span>|<span data-ttu-id="7e9d5-136">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="7e9d5-137">String</span><span class="sxs-lookup"><span data-stu-id="7e9d5-137">String</span></span>|<span data-ttu-id="7e9d5-138">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7e9d5-139">要件</span><span class="sxs-lookup"><span data-stu-id="7e9d5-139">Requirements</span></span>

|<span data-ttu-id="7e9d5-140">要件</span><span class="sxs-lookup"><span data-stu-id="7e9d5-140">Requirement</span></span>| <span data-ttu-id="7e9d5-141">値</span><span class="sxs-lookup"><span data-stu-id="7e9d5-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e9d5-142">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7e9d5-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7e9d5-143">1.0</span><span class="sxs-lookup"><span data-stu-id="7e9d5-143">1.0</span></span>|
|[<span data-ttu-id="7e9d5-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7e9d5-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7e9d5-145">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="7e9d5-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="7e9d5-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="7e9d5-146">CoercionType :String</span></span>

<span data-ttu-id="7e9d5-147">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7e9d5-148">型</span><span class="sxs-lookup"><span data-stu-id="7e9d5-148">Type</span></span>

*   <span data-ttu-id="7e9d5-149">String</span><span class="sxs-lookup"><span data-stu-id="7e9d5-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7e9d5-150">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="7e9d5-150">Properties:</span></span>

|<span data-ttu-id="7e9d5-151">名前</span><span class="sxs-lookup"><span data-stu-id="7e9d5-151">Name</span></span>| <span data-ttu-id="7e9d5-152">型</span><span class="sxs-lookup"><span data-stu-id="7e9d5-152">Type</span></span>| <span data-ttu-id="7e9d5-153">説明</span><span class="sxs-lookup"><span data-stu-id="7e9d5-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="7e9d5-154">String</span><span class="sxs-lookup"><span data-stu-id="7e9d5-154">String</span></span>|<span data-ttu-id="7e9d5-155">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="7e9d5-156">String</span><span class="sxs-lookup"><span data-stu-id="7e9d5-156">String</span></span>|<span data-ttu-id="7e9d5-157">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7e9d5-158">要件</span><span class="sxs-lookup"><span data-stu-id="7e9d5-158">Requirements</span></span>

|<span data-ttu-id="7e9d5-159">要件</span><span class="sxs-lookup"><span data-stu-id="7e9d5-159">Requirement</span></span>| <span data-ttu-id="7e9d5-160">値</span><span class="sxs-lookup"><span data-stu-id="7e9d5-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e9d5-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7e9d5-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7e9d5-162">1.0</span><span class="sxs-lookup"><span data-stu-id="7e9d5-162">1.0</span></span>|
|[<span data-ttu-id="7e9d5-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7e9d5-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7e9d5-164">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="7e9d5-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="7e9d5-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="7e9d5-165">EventType :String</span></span>

<span data-ttu-id="7e9d5-166">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="7e9d5-167">型</span><span class="sxs-lookup"><span data-stu-id="7e9d5-167">Type</span></span>

*   <span data-ttu-id="7e9d5-168">String</span><span class="sxs-lookup"><span data-stu-id="7e9d5-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7e9d5-169">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="7e9d5-169">Properties:</span></span>

| <span data-ttu-id="7e9d5-170">名前</span><span class="sxs-lookup"><span data-stu-id="7e9d5-170">Name</span></span> | <span data-ttu-id="7e9d5-171">型</span><span class="sxs-lookup"><span data-stu-id="7e9d5-171">Type</span></span> | <span data-ttu-id="7e9d5-172">説明</span><span class="sxs-lookup"><span data-stu-id="7e9d5-172">Description</span></span> | <span data-ttu-id="7e9d5-173">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="7e9d5-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="7e9d5-174">String</span><span class="sxs-lookup"><span data-stu-id="7e9d5-174">String</span></span> | <span data-ttu-id="7e9d5-175">選択した予定または一連の予定の日付または時刻が変更された。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="7e9d5-176">1.7</span><span class="sxs-lookup"><span data-stu-id="7e9d5-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="7e9d5-177">String</span><span class="sxs-lookup"><span data-stu-id="7e9d5-177">String</span></span> | <span data-ttu-id="7e9d5-178">アイテムに添付ファイルが追加されたか、アイテムから添付ファイルが削除された。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="7e9d5-179">プレビュー</span><span class="sxs-lookup"><span data-stu-id="7e9d5-179">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="7e9d5-180">String</span><span class="sxs-lookup"><span data-stu-id="7e9d5-180">String</span></span> | <span data-ttu-id="7e9d5-181">作業ウィンドウがピン留めされている間、別の Outlook アイテムが選択されて表示されている。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-181">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="7e9d5-182">1.5</span><span class="sxs-lookup"><span data-stu-id="7e9d5-182">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="7e9d5-183">文字列</span><span class="sxs-lookup"><span data-stu-id="7e9d5-183">String</span></span> | <span data-ttu-id="7e9d5-184">メールボックスの Office テーマが変更された。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-184">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="7e9d5-185">プレビュー</span><span class="sxs-lookup"><span data-stu-id="7e9d5-185">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="7e9d5-186">String</span><span class="sxs-lookup"><span data-stu-id="7e9d5-186">String</span></span> | <span data-ttu-id="7e9d5-187">選択したアイテムまたは予定の場所の受信者リストが変更された。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-187">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="7e9d5-188">1.7</span><span class="sxs-lookup"><span data-stu-id="7e9d5-188">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="7e9d5-189">文字列</span><span class="sxs-lookup"><span data-stu-id="7e9d5-189">String</span></span> | <span data-ttu-id="7e9d5-190">選択した一連の予定の定期的なパターンが変更された。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-190">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="7e9d5-191">1.7</span><span class="sxs-lookup"><span data-stu-id="7e9d5-191">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7e9d5-192">要件</span><span class="sxs-lookup"><span data-stu-id="7e9d5-192">Requirements</span></span>

|<span data-ttu-id="7e9d5-193">要件</span><span class="sxs-lookup"><span data-stu-id="7e9d5-193">Requirement</span></span>| <span data-ttu-id="7e9d5-194">値</span><span class="sxs-lookup"><span data-stu-id="7e9d5-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e9d5-195">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7e9d5-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7e9d5-196">1.5</span><span class="sxs-lookup"><span data-stu-id="7e9d5-196">1.5</span></span> |
|[<span data-ttu-id="7e9d5-197">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7e9d5-197">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7e9d5-198">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="7e9d5-198">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="7e9d5-199">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="7e9d5-199">SourceProperty :String</span></span>

<span data-ttu-id="7e9d5-200">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-200">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7e9d5-201">型</span><span class="sxs-lookup"><span data-stu-id="7e9d5-201">Type</span></span>

*   <span data-ttu-id="7e9d5-202">String</span><span class="sxs-lookup"><span data-stu-id="7e9d5-202">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7e9d5-203">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="7e9d5-203">Properties:</span></span>

|<span data-ttu-id="7e9d5-204">名前</span><span class="sxs-lookup"><span data-stu-id="7e9d5-204">Name</span></span>| <span data-ttu-id="7e9d5-205">型</span><span class="sxs-lookup"><span data-stu-id="7e9d5-205">Type</span></span>| <span data-ttu-id="7e9d5-206">説明</span><span class="sxs-lookup"><span data-stu-id="7e9d5-206">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="7e9d5-207">String</span><span class="sxs-lookup"><span data-stu-id="7e9d5-207">String</span></span>|<span data-ttu-id="7e9d5-208">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-208">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="7e9d5-209">String</span><span class="sxs-lookup"><span data-stu-id="7e9d5-209">String</span></span>|<span data-ttu-id="7e9d5-210">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="7e9d5-210">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7e9d5-211">要件</span><span class="sxs-lookup"><span data-stu-id="7e9d5-211">Requirements</span></span>

|<span data-ttu-id="7e9d5-212">要件</span><span class="sxs-lookup"><span data-stu-id="7e9d5-212">Requirement</span></span>| <span data-ttu-id="7e9d5-213">値</span><span class="sxs-lookup"><span data-stu-id="7e9d5-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e9d5-214">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7e9d5-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7e9d5-215">1.0</span><span class="sxs-lookup"><span data-stu-id="7e9d5-215">1.0</span></span>|
|[<span data-ttu-id="7e9d5-216">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7e9d5-216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7e9d5-217">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="7e9d5-217">Compose or Read</span></span>|
