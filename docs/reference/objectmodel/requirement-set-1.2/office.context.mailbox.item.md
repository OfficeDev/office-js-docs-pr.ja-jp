---
title: Office. メールボックス-要件セット1.2
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: ab8c55d2f91b250b419c7c9c71fc044b6fa68279
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629210"
---
# <a name="item"></a><span data-ttu-id="d7298-102">item</span><span class="sxs-lookup"><span data-stu-id="d7298-102">item</span></span>

### <span data-ttu-id="d7298-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="d7298-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="d7298-p102">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="d7298-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-107">Requirements</span></span>

|<span data-ttu-id="d7298-108">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-108">Requirement</span></span>| <span data-ttu-id="d7298-109">値</span><span class="sxs-lookup"><span data-stu-id="d7298-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-111">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-111">1.0</span></span>|
|[<span data-ttu-id="d7298-112">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-113">制限あり</span><span class="sxs-lookup"><span data-stu-id="d7298-113">Restricted</span></span>|
|[<span data-ttu-id="d7298-114">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-115">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d7298-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d7298-116">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="d7298-116">Members and methods</span></span>

| <span data-ttu-id="d7298-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="d7298-117">Member</span></span> | <span data-ttu-id="d7298-118">種類</span><span class="sxs-lookup"><span data-stu-id="d7298-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d7298-119">attachments</span><span class="sxs-lookup"><span data-stu-id="d7298-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="d7298-120">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-120">Member</span></span> |
| [<span data-ttu-id="d7298-121">bcc</span><span class="sxs-lookup"><span data-stu-id="d7298-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="d7298-122">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-122">Member</span></span> |
| [<span data-ttu-id="d7298-123">body</span><span class="sxs-lookup"><span data-stu-id="d7298-123">body</span></span>](#body-body) | <span data-ttu-id="d7298-124">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-124">Member</span></span> |
| [<span data-ttu-id="d7298-125">cc</span><span class="sxs-lookup"><span data-stu-id="d7298-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d7298-126">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-126">Member</span></span> |
| [<span data-ttu-id="d7298-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="d7298-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="d7298-128">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-128">Member</span></span> |
| [<span data-ttu-id="d7298-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="d7298-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="d7298-130">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-130">Member</span></span> |
| [<span data-ttu-id="d7298-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="d7298-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="d7298-132">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-132">Member</span></span> |
| [<span data-ttu-id="d7298-133">end</span><span class="sxs-lookup"><span data-stu-id="d7298-133">end</span></span>](#end-datetime) | <span data-ttu-id="d7298-134">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-134">Member</span></span> |
| [<span data-ttu-id="d7298-135">from</span><span class="sxs-lookup"><span data-stu-id="d7298-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="d7298-136">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-136">Member</span></span> |
| [<span data-ttu-id="d7298-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="d7298-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="d7298-138">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-138">Member</span></span> |
| [<span data-ttu-id="d7298-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="d7298-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="d7298-140">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-140">Member</span></span> |
| [<span data-ttu-id="d7298-141">itemId</span><span class="sxs-lookup"><span data-stu-id="d7298-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="d7298-142">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-142">Member</span></span> |
| [<span data-ttu-id="d7298-143">itemType</span><span class="sxs-lookup"><span data-stu-id="d7298-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="d7298-144">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-144">Member</span></span> |
| [<span data-ttu-id="d7298-145">location</span><span class="sxs-lookup"><span data-stu-id="d7298-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="d7298-146">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-146">Member</span></span> |
| [<span data-ttu-id="d7298-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="d7298-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="d7298-148">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-148">Member</span></span> |
| [<span data-ttu-id="d7298-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="d7298-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d7298-150">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-150">Member</span></span> |
| [<span data-ttu-id="d7298-151">organizer</span><span class="sxs-lookup"><span data-stu-id="d7298-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="d7298-152">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-152">Member</span></span> |
| [<span data-ttu-id="d7298-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="d7298-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d7298-154">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-154">Member</span></span> |
| [<span data-ttu-id="d7298-155">sender</span><span class="sxs-lookup"><span data-stu-id="d7298-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="d7298-156">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-156">Member</span></span> |
| [<span data-ttu-id="d7298-157">start</span><span class="sxs-lookup"><span data-stu-id="d7298-157">start</span></span>](#start-datetime) | <span data-ttu-id="d7298-158">Member</span><span class="sxs-lookup"><span data-stu-id="d7298-158">Member</span></span> |
| [<span data-ttu-id="d7298-159">subject</span><span class="sxs-lookup"><span data-stu-id="d7298-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="d7298-160">メンバー</span><span class="sxs-lookup"><span data-stu-id="d7298-160">Member</span></span> |
| [<span data-ttu-id="d7298-161">to</span><span class="sxs-lookup"><span data-stu-id="d7298-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d7298-162">メンバー</span><span class="sxs-lookup"><span data-stu-id="d7298-162">Member</span></span> |
| [<span data-ttu-id="d7298-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d7298-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="d7298-164">メソッド</span><span class="sxs-lookup"><span data-stu-id="d7298-164">Method</span></span> |
| [<span data-ttu-id="d7298-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d7298-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="d7298-166">メソッド</span><span class="sxs-lookup"><span data-stu-id="d7298-166">Method</span></span> |
| [<span data-ttu-id="d7298-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="d7298-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="d7298-168">メソッド</span><span class="sxs-lookup"><span data-stu-id="d7298-168">Method</span></span> |
| [<span data-ttu-id="d7298-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="d7298-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="d7298-170">メソッド</span><span class="sxs-lookup"><span data-stu-id="d7298-170">Method</span></span> |
| [<span data-ttu-id="d7298-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="d7298-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="d7298-172">メソッド</span><span class="sxs-lookup"><span data-stu-id="d7298-172">Method</span></span> |
| [<span data-ttu-id="d7298-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="d7298-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="d7298-174">メソッド</span><span class="sxs-lookup"><span data-stu-id="d7298-174">Method</span></span> |
| [<span data-ttu-id="d7298-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="d7298-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="d7298-176">メソッド</span><span class="sxs-lookup"><span data-stu-id="d7298-176">Method</span></span> |
| [<span data-ttu-id="d7298-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="d7298-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="d7298-178">メソッド</span><span class="sxs-lookup"><span data-stu-id="d7298-178">Method</span></span> |
| [<span data-ttu-id="d7298-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="d7298-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="d7298-180">メソッド</span><span class="sxs-lookup"><span data-stu-id="d7298-180">Method</span></span> |
| [<span data-ttu-id="d7298-181">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="d7298-181">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="d7298-182">メソッド</span><span class="sxs-lookup"><span data-stu-id="d7298-182">Method</span></span> |
| [<span data-ttu-id="d7298-183">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="d7298-183">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="d7298-184">メソッド</span><span class="sxs-lookup"><span data-stu-id="d7298-184">Method</span></span> |
| [<span data-ttu-id="d7298-185">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d7298-185">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="d7298-186">メソッド</span><span class="sxs-lookup"><span data-stu-id="d7298-186">Method</span></span> |
| [<span data-ttu-id="d7298-187">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="d7298-187">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="d7298-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="d7298-188">Method</span></span> |

### <a name="example"></a><span data-ttu-id="d7298-189">例</span><span class="sxs-lookup"><span data-stu-id="d7298-189">Example</span></span>

<span data-ttu-id="d7298-190">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d7298-190">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
  });
};
```

### <a name="members"></a><span data-ttu-id="d7298-191">Members</span><span class="sxs-lookup"><span data-stu-id="d7298-191">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-12"></a><span data-ttu-id="d7298-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="d7298-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

<span data-ttu-id="d7298-p103">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d7298-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d7298-195">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="d7298-195">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="d7298-196">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d7298-196">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="d7298-197">型</span><span class="sxs-lookup"><span data-stu-id="d7298-197">Type</span></span>

*   <span data-ttu-id="d7298-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="d7298-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-199">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-199">Requirements</span></span>

|<span data-ttu-id="d7298-200">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-200">Requirement</span></span>| <span data-ttu-id="d7298-201">値</span><span class="sxs-lookup"><span data-stu-id="d7298-201">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-202">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-202">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-203">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-203">1.0</span></span>|
|[<span data-ttu-id="d7298-204">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-204">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-205">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-205">ReadItem</span></span>|
|[<span data-ttu-id="d7298-206">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-206">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-207">読み取り</span><span class="sxs-lookup"><span data-stu-id="d7298-207">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-208">例</span><span class="sxs-lookup"><span data-stu-id="d7298-208">Example</span></span>

<span data-ttu-id="d7298-209">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="d7298-209">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
var item = Office.context.mailbox.item;
var outputString = "";

if (item.attachments.length > 0) {
  for (i = 0 ; i < item.attachments.length ; i++) {
    var attachment = item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += attachment.name;
    outputString += "<BR>ID: " + attachment.id;
    outputString += "<BR>contentType: " + attachment.contentType;
    outputString += "<BR>size: " + attachment.size;
    outputString += "<BR>attachmentType: " + attachment.attachmentType;
    outputString += "<BR>isInline: " + attachment.isInline;
  }
}

console.log(outputString);
```

<br>

---
---

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="d7298-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="d7298-211">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="d7298-211">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="d7298-212">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d7298-212">Compose mode only.</span></span>

<span data-ttu-id="d7298-213">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d7298-213">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d7298-214">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-214">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d7298-215">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="d7298-215">Get 500 members maximum.</span></span>
- <span data-ttu-id="d7298-216">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="d7298-216">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="d7298-217">型</span><span class="sxs-lookup"><span data-stu-id="d7298-217">Type</span></span>

*   [<span data-ttu-id="d7298-218">受信者</span><span class="sxs-lookup"><span data-stu-id="d7298-218">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="d7298-219">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-219">Requirements</span></span>

|<span data-ttu-id="d7298-220">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-220">Requirement</span></span>| <span data-ttu-id="d7298-221">値</span><span class="sxs-lookup"><span data-stu-id="d7298-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-222">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-222">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-223">1.1</span><span class="sxs-lookup"><span data-stu-id="d7298-223">1.1</span></span>|
|[<span data-ttu-id="d7298-224">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-224">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-225">ReadItem</span></span>|
|[<span data-ttu-id="d7298-226">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-226">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-227">作成</span><span class="sxs-lookup"><span data-stu-id="d7298-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-228">例</span><span class="sxs-lookup"><span data-stu-id="d7298-228">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

<br>

---
---

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-12"></a><span data-ttu-id="d7298-229">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-229">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span></span>

<span data-ttu-id="d7298-230">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="d7298-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d7298-231">型</span><span class="sxs-lookup"><span data-stu-id="d7298-231">Type</span></span>

*   [<span data-ttu-id="d7298-232">Body</span><span class="sxs-lookup"><span data-stu-id="d7298-232">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="d7298-233">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-233">Requirements</span></span>

|<span data-ttu-id="d7298-234">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-234">Requirement</span></span>| <span data-ttu-id="d7298-235">値</span><span class="sxs-lookup"><span data-stu-id="d7298-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-236">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-237">1.1</span><span class="sxs-lookup"><span data-stu-id="d7298-237">1.1</span></span>|
|[<span data-ttu-id="d7298-238">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-239">ReadItem</span></span>|
|[<span data-ttu-id="d7298-240">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-241">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d7298-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-242">例</span><span class="sxs-lookup"><span data-stu-id="d7298-242">Example</span></span>

<span data-ttu-id="d7298-243">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="d7298-243">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="d7298-244">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="d7298-244">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="d7298-245">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-245">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="d7298-246">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d7298-246">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="d7298-247">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="d7298-247">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7298-248">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d7298-248">Read mode</span></span>

<span data-ttu-id="d7298-249">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-249">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="d7298-250">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d7298-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d7298-251">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="d7298-251">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="d7298-252">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d7298-252">Compose mode</span></span>

<span data-ttu-id="d7298-253">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="d7298-254">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d7298-254">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d7298-255">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-255">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d7298-256">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="d7298-256">Get 500 members maximum.</span></span>
- <span data-ttu-id="d7298-257">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="d7298-257">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d7298-258">型</span><span class="sxs-lookup"><span data-stu-id="d7298-258">Type</span></span>

*   <span data-ttu-id="d7298-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-260">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-260">Requirements</span></span>

|<span data-ttu-id="d7298-261">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-261">Requirement</span></span>| <span data-ttu-id="d7298-262">値</span><span class="sxs-lookup"><span data-stu-id="d7298-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-263">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-264">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-264">1.0</span></span>|
|[<span data-ttu-id="d7298-265">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-266">ReadItem</span></span>|
|[<span data-ttu-id="d7298-267">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-268">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d7298-268">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="d7298-269">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="d7298-269">(nullable) conversationId: String</span></span>

<span data-ttu-id="d7298-270">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="d7298-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="d7298-p110">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="d7298-p110">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="d7298-p111">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-p111">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="d7298-275">Type</span><span class="sxs-lookup"><span data-stu-id="d7298-275">Type</span></span>

*   <span data-ttu-id="d7298-276">String</span><span class="sxs-lookup"><span data-stu-id="d7298-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-277">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-277">Requirements</span></span>

|<span data-ttu-id="d7298-278">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-278">Requirement</span></span>| <span data-ttu-id="d7298-279">値</span><span class="sxs-lookup"><span data-stu-id="d7298-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-280">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-281">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-281">1.0</span></span>|
|[<span data-ttu-id="d7298-282">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-283">ReadItem</span></span>|
|[<span data-ttu-id="d7298-284">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-285">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d7298-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-286">例</span><span class="sxs-lookup"><span data-stu-id="d7298-286">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="d7298-287">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="d7298-287">dateTimeCreated: Date</span></span>

<span data-ttu-id="d7298-p112">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d7298-p112">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d7298-290">型</span><span class="sxs-lookup"><span data-stu-id="d7298-290">Type</span></span>

*   <span data-ttu-id="d7298-291">日付</span><span class="sxs-lookup"><span data-stu-id="d7298-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-292">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-292">Requirements</span></span>

|<span data-ttu-id="d7298-293">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-293">Requirement</span></span>| <span data-ttu-id="d7298-294">値</span><span class="sxs-lookup"><span data-stu-id="d7298-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-295">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-296">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-296">1.0</span></span>|
|[<span data-ttu-id="d7298-297">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-298">ReadItem</span></span>|
|[<span data-ttu-id="d7298-299">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-300">読み取り</span><span class="sxs-lookup"><span data-stu-id="d7298-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-301">例</span><span class="sxs-lookup"><span data-stu-id="d7298-301">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="d7298-302">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="d7298-302">dateTimeModified: Date</span></span>

<span data-ttu-id="d7298-p113">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d7298-p113">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d7298-305">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d7298-305">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="d7298-306">種類</span><span class="sxs-lookup"><span data-stu-id="d7298-306">Type</span></span>

*   <span data-ttu-id="d7298-307">日付</span><span class="sxs-lookup"><span data-stu-id="d7298-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-308">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-308">Requirements</span></span>

|<span data-ttu-id="d7298-309">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-309">Requirement</span></span>| <span data-ttu-id="d7298-310">値</span><span class="sxs-lookup"><span data-stu-id="d7298-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-312">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-312">1.0</span></span>|
|[<span data-ttu-id="d7298-313">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-314">ReadItem</span></span>|
|[<span data-ttu-id="d7298-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-316">読み取り</span><span class="sxs-lookup"><span data-stu-id="d7298-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-317">例</span><span class="sxs-lookup"><span data-stu-id="d7298-317">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="d7298-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="d7298-319">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d7298-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="d7298-p114">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="d7298-p114">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7298-322">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d7298-322">Read mode</span></span>

<span data-ttu-id="d7298-323">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-323">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="d7298-324">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d7298-324">Compose mode</span></span>

<span data-ttu-id="d7298-325">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="d7298-326">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d7298-326">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="d7298-327">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="d7298-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="d7298-328">型</span><span class="sxs-lookup"><span data-stu-id="d7298-328">Type</span></span>

*   <span data-ttu-id="d7298-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-330">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-330">Requirements</span></span>

|<span data-ttu-id="d7298-331">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-331">Requirement</span></span>| <span data-ttu-id="d7298-332">値</span><span class="sxs-lookup"><span data-stu-id="d7298-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-333">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-334">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-334">1.0</span></span>|
|[<span data-ttu-id="d7298-335">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-336">ReadItem</span></span>|
|[<span data-ttu-id="d7298-337">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-338">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d7298-338">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="d7298-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="d7298-p115">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d7298-p115">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="d7298-p116">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="d7298-p116">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d7298-344">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="d7298-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="d7298-345">型</span><span class="sxs-lookup"><span data-stu-id="d7298-345">Type</span></span>

*   [<span data-ttu-id="d7298-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d7298-346">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="d7298-347">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-347">Requirements</span></span>

|<span data-ttu-id="d7298-348">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-348">Requirement</span></span>| <span data-ttu-id="d7298-349">値</span><span class="sxs-lookup"><span data-stu-id="d7298-349">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-350">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-350">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-351">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-351">1.0</span></span>|
|[<span data-ttu-id="d7298-352">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-352">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-353">ReadItem</span></span>|
|[<span data-ttu-id="d7298-354">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-354">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-355">読み取り</span><span class="sxs-lookup"><span data-stu-id="d7298-355">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-356">例</span><span class="sxs-lookup"><span data-stu-id="d7298-356">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="d7298-357">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="d7298-357">internetMessageId: String</span></span>

<span data-ttu-id="d7298-p117">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d7298-p117">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d7298-360">Type</span><span class="sxs-lookup"><span data-stu-id="d7298-360">Type</span></span>

*   <span data-ttu-id="d7298-361">String</span><span class="sxs-lookup"><span data-stu-id="d7298-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-362">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-362">Requirements</span></span>

|<span data-ttu-id="d7298-363">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-363">Requirement</span></span>| <span data-ttu-id="d7298-364">値</span><span class="sxs-lookup"><span data-stu-id="d7298-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-365">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-366">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-366">1.0</span></span>|
|[<span data-ttu-id="d7298-367">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-368">ReadItem</span></span>|
|[<span data-ttu-id="d7298-369">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-370">読み取り</span><span class="sxs-lookup"><span data-stu-id="d7298-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-371">例</span><span class="sxs-lookup"><span data-stu-id="d7298-371">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="d7298-372">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="d7298-372">itemClass: String</span></span>

<span data-ttu-id="d7298-p118">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d7298-p118">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="d7298-p119">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="d7298-p119">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="d7298-377">型</span><span class="sxs-lookup"><span data-stu-id="d7298-377">Type</span></span> | <span data-ttu-id="d7298-378">説明</span><span class="sxs-lookup"><span data-stu-id="d7298-378">Description</span></span> | <span data-ttu-id="d7298-379">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="d7298-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="d7298-380">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="d7298-380">Appointment items</span></span> | <span data-ttu-id="d7298-381">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="d7298-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="d7298-382">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="d7298-382">Message items</span></span> | <span data-ttu-id="d7298-383">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d7298-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="d7298-384">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="d7298-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="d7298-385">Type</span><span class="sxs-lookup"><span data-stu-id="d7298-385">Type</span></span>

*   <span data-ttu-id="d7298-386">String</span><span class="sxs-lookup"><span data-stu-id="d7298-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-387">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-387">Requirements</span></span>

|<span data-ttu-id="d7298-388">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-388">Requirement</span></span>| <span data-ttu-id="d7298-389">値</span><span class="sxs-lookup"><span data-stu-id="d7298-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-390">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-391">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-391">1.0</span></span>|
|[<span data-ttu-id="d7298-392">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-393">ReadItem</span></span>|
|[<span data-ttu-id="d7298-394">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-395">読み取り</span><span class="sxs-lookup"><span data-stu-id="d7298-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-396">例</span><span class="sxs-lookup"><span data-stu-id="d7298-396">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="d7298-397">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="d7298-397">(nullable) itemId: String</span></span>

<span data-ttu-id="d7298-p120">現在のアイテムの [Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d7298-p120">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d7298-400">`itemId` プロパティから返される識別子は、[Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) と同じです。</span><span class="sxs-lookup"><span data-stu-id="d7298-400">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="d7298-401">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="d7298-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="d7298-402">この値を使用して REST API を呼び出す前に、を`Office.context.mailbox.convertToRestId`使用して変換する必要があります。これは、要件セット1.3 から開始できます。</span><span class="sxs-lookup"><span data-stu-id="d7298-402">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="d7298-403">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d7298-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="d7298-404">Type</span><span class="sxs-lookup"><span data-stu-id="d7298-404">Type</span></span>

*   <span data-ttu-id="d7298-405">String</span><span class="sxs-lookup"><span data-stu-id="d7298-405">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-406">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-406">Requirements</span></span>

|<span data-ttu-id="d7298-407">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-407">Requirement</span></span>| <span data-ttu-id="d7298-408">値</span><span class="sxs-lookup"><span data-stu-id="d7298-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-409">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-410">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-410">1.0</span></span>|
|[<span data-ttu-id="d7298-411">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-412">ReadItem</span></span>|
|[<span data-ttu-id="d7298-413">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-414">読み取り</span><span class="sxs-lookup"><span data-stu-id="d7298-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-415">例</span><span class="sxs-lookup"><span data-stu-id="d7298-415">Example</span></span>

<span data-ttu-id="d7298-p122">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

<br>

---
---

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-12"></a><span data-ttu-id="d7298-418">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-418">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span></span>

<span data-ttu-id="d7298-419">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="d7298-419">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="d7298-420">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="d7298-420">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d7298-421">型</span><span class="sxs-lookup"><span data-stu-id="d7298-421">Type</span></span>

*   [<span data-ttu-id="d7298-422">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="d7298-422">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="d7298-423">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-423">Requirements</span></span>

|<span data-ttu-id="d7298-424">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-424">Requirement</span></span>| <span data-ttu-id="d7298-425">値</span><span class="sxs-lookup"><span data-stu-id="d7298-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-426">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-427">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-427">1.0</span></span>|
|[<span data-ttu-id="d7298-428">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-429">ReadItem</span></span>|
|[<span data-ttu-id="d7298-430">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-431">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d7298-431">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-432">例</span><span class="sxs-lookup"><span data-stu-id="d7298-432">Example</span></span>

```js
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

<br>

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-12"></a><span data-ttu-id="d7298-433">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-433">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

<span data-ttu-id="d7298-434">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d7298-434">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7298-435">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d7298-435">Read mode</span></span>

<span data-ttu-id="d7298-436">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-436">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="d7298-437">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d7298-437">Compose mode</span></span>

<span data-ttu-id="d7298-438">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-438">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d7298-439">型</span><span class="sxs-lookup"><span data-stu-id="d7298-439">Type</span></span>

*   <span data-ttu-id="d7298-440">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-440">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-441">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-441">Requirements</span></span>

|<span data-ttu-id="d7298-442">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-442">Requirement</span></span>| <span data-ttu-id="d7298-443">値</span><span class="sxs-lookup"><span data-stu-id="d7298-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-444">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-445">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-445">1.0</span></span>|
|[<span data-ttu-id="d7298-446">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-447">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-447">ReadItem</span></span>|
|[<span data-ttu-id="d7298-448">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-449">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d7298-449">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="d7298-450">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="d7298-450">normalizedSubject: String</span></span>

<span data-ttu-id="d7298-p123">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d7298-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="d7298-p124">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="d7298-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="d7298-455">Type</span><span class="sxs-lookup"><span data-stu-id="d7298-455">Type</span></span>

*   <span data-ttu-id="d7298-456">String</span><span class="sxs-lookup"><span data-stu-id="d7298-456">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-457">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-457">Requirements</span></span>

|<span data-ttu-id="d7298-458">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-458">Requirement</span></span>| <span data-ttu-id="d7298-459">値</span><span class="sxs-lookup"><span data-stu-id="d7298-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-460">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-461">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-461">1.0</span></span>|
|[<span data-ttu-id="d7298-462">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-462">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-463">ReadItem</span></span>|
|[<span data-ttu-id="d7298-464">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-464">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-465">読み取り</span><span class="sxs-lookup"><span data-stu-id="d7298-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-466">例</span><span class="sxs-lookup"><span data-stu-id="d7298-466">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="d7298-467">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-467">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="d7298-468">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d7298-468">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="d7298-469">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="d7298-469">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7298-470">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d7298-470">Read mode</span></span>

<span data-ttu-id="d7298-471">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-471">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="d7298-472">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d7298-472">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d7298-473">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="d7298-473">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="d7298-474">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d7298-474">Compose mode</span></span>

<span data-ttu-id="d7298-475">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-475">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="d7298-476">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d7298-476">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d7298-477">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-477">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d7298-478">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="d7298-478">Get 500 members maximum.</span></span>
- <span data-ttu-id="d7298-479">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="d7298-479">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d7298-480">型</span><span class="sxs-lookup"><span data-stu-id="d7298-480">Type</span></span>

*   <span data-ttu-id="d7298-481">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-481">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-482">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-482">Requirements</span></span>

|<span data-ttu-id="d7298-483">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-483">Requirement</span></span>| <span data-ttu-id="d7298-484">値</span><span class="sxs-lookup"><span data-stu-id="d7298-484">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-485">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-485">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-486">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-486">1.0</span></span>|
|[<span data-ttu-id="d7298-487">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-487">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-488">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-488">ReadItem</span></span>|
|[<span data-ttu-id="d7298-489">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-489">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-490">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d7298-490">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="d7298-491">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-491">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="d7298-p128">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d7298-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d7298-494">型</span><span class="sxs-lookup"><span data-stu-id="d7298-494">Type</span></span>

*   [<span data-ttu-id="d7298-495">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d7298-495">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="d7298-496">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-496">Requirements</span></span>

|<span data-ttu-id="d7298-497">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-497">Requirement</span></span>| <span data-ttu-id="d7298-498">値</span><span class="sxs-lookup"><span data-stu-id="d7298-498">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-499">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-499">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-500">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-500">1.0</span></span>|
|[<span data-ttu-id="d7298-501">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-501">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-502">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-502">ReadItem</span></span>|
|[<span data-ttu-id="d7298-503">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-503">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-504">読み取り</span><span class="sxs-lookup"><span data-stu-id="d7298-504">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-505">例</span><span class="sxs-lookup"><span data-stu-id="d7298-505">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="d7298-506">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-506">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="d7298-507">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d7298-507">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="d7298-508">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="d7298-508">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7298-509">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d7298-509">Read mode</span></span>

<span data-ttu-id="d7298-510">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-510">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="d7298-511">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d7298-511">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d7298-512">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="d7298-512">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="d7298-513">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d7298-513">Compose mode</span></span>

<span data-ttu-id="d7298-514">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-514">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="d7298-515">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d7298-515">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d7298-516">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-516">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d7298-517">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="d7298-517">Get 500 members maximum.</span></span>
- <span data-ttu-id="d7298-518">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="d7298-518">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="d7298-519">型</span><span class="sxs-lookup"><span data-stu-id="d7298-519">Type</span></span>

*   <span data-ttu-id="d7298-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-521">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-521">Requirements</span></span>

|<span data-ttu-id="d7298-522">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-522">Requirement</span></span>| <span data-ttu-id="d7298-523">値</span><span class="sxs-lookup"><span data-stu-id="d7298-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-524">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-525">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-525">1.0</span></span>|
|[<span data-ttu-id="d7298-526">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-527">ReadItem</span></span>|
|[<span data-ttu-id="d7298-528">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-529">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d7298-529">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="d7298-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="d7298-p132">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d7298-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="d7298-p133">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="d7298-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d7298-535">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="d7298-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="d7298-536">型</span><span class="sxs-lookup"><span data-stu-id="d7298-536">Type</span></span>

*   [<span data-ttu-id="d7298-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d7298-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="d7298-538">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-538">Requirements</span></span>

|<span data-ttu-id="d7298-539">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-539">Requirement</span></span>| <span data-ttu-id="d7298-540">値</span><span class="sxs-lookup"><span data-stu-id="d7298-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-541">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-542">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-542">1.0</span></span>|
|[<span data-ttu-id="d7298-543">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-544">ReadItem</span></span>|
|[<span data-ttu-id="d7298-545">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-546">読み取り</span><span class="sxs-lookup"><span data-stu-id="d7298-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-547">例</span><span class="sxs-lookup"><span data-stu-id="d7298-547">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="d7298-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="d7298-549">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d7298-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="d7298-p134">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="d7298-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7298-552">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d7298-552">Read mode</span></span>

<span data-ttu-id="d7298-553">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-553">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="d7298-554">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d7298-554">Compose mode</span></span>

<span data-ttu-id="d7298-555">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="d7298-556">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d7298-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="d7298-557">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="d7298-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="d7298-558">型</span><span class="sxs-lookup"><span data-stu-id="d7298-558">Type</span></span>

*   <span data-ttu-id="d7298-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-560">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-560">Requirements</span></span>

|<span data-ttu-id="d7298-561">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-561">Requirement</span></span>| <span data-ttu-id="d7298-562">値</span><span class="sxs-lookup"><span data-stu-id="d7298-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-563">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-564">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-564">1.0</span></span>|
|[<span data-ttu-id="d7298-565">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-566">ReadItem</span></span>|
|[<span data-ttu-id="d7298-567">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-568">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d7298-568">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-12"></a><span data-ttu-id="d7298-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

<span data-ttu-id="d7298-570">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d7298-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="d7298-571">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d7298-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7298-572">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d7298-572">Read mode</span></span>

<span data-ttu-id="d7298-p136">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="d7298-p136">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="d7298-575">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d7298-575">Compose mode</span></span>

<span data-ttu-id="d7298-576">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="d7298-577">型</span><span class="sxs-lookup"><span data-stu-id="d7298-577">Type</span></span>

*   <span data-ttu-id="d7298-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-579">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-579">Requirements</span></span>

|<span data-ttu-id="d7298-580">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-580">Requirement</span></span>| <span data-ttu-id="d7298-581">値</span><span class="sxs-lookup"><span data-stu-id="d7298-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-582">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-583">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-583">1.0</span></span>|
|[<span data-ttu-id="d7298-584">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-585">ReadItem</span></span>|
|[<span data-ttu-id="d7298-586">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-587">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d7298-587">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="d7298-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="d7298-589">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d7298-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="d7298-590">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="d7298-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7298-591">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d7298-591">Read mode</span></span>

<span data-ttu-id="d7298-592">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-592">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="d7298-593">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d7298-593">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d7298-594">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="d7298-594">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="d7298-595">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d7298-595">Compose mode</span></span>

<span data-ttu-id="d7298-596">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-596">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="d7298-597">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d7298-597">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d7298-598">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-598">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d7298-599">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="d7298-599">Get 500 members maximum.</span></span>
- <span data-ttu-id="d7298-600">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="d7298-600">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d7298-601">型</span><span class="sxs-lookup"><span data-stu-id="d7298-601">Type</span></span>

*   <span data-ttu-id="d7298-602">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-602">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-603">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-603">Requirements</span></span>

|<span data-ttu-id="d7298-604">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-604">Requirement</span></span>| <span data-ttu-id="d7298-605">値</span><span class="sxs-lookup"><span data-stu-id="d7298-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-606">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-607">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-607">1.0</span></span>|
|[<span data-ttu-id="d7298-608">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-609">ReadItem</span></span>|
|[<span data-ttu-id="d7298-610">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-611">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d7298-611">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="d7298-612">メソッド</span><span class="sxs-lookup"><span data-stu-id="d7298-612">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="d7298-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d7298-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d7298-614">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="d7298-614">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="d7298-615">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="d7298-615">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="d7298-616">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="d7298-616">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7298-617">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d7298-617">Parameters</span></span>

|<span data-ttu-id="d7298-618">名前</span><span class="sxs-lookup"><span data-stu-id="d7298-618">Name</span></span>| <span data-ttu-id="d7298-619">種類</span><span class="sxs-lookup"><span data-stu-id="d7298-619">Type</span></span>| <span data-ttu-id="d7298-620">属性</span><span class="sxs-lookup"><span data-stu-id="d7298-620">Attributes</span></span>| <span data-ttu-id="d7298-621">説明</span><span class="sxs-lookup"><span data-stu-id="d7298-621">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="d7298-622">String</span><span class="sxs-lookup"><span data-stu-id="d7298-622">String</span></span>||<span data-ttu-id="d7298-p140">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="d7298-p140">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="d7298-625">String</span><span class="sxs-lookup"><span data-stu-id="d7298-625">String</span></span>||<span data-ttu-id="d7298-p141">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d7298-p141">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="d7298-628">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d7298-628">Object</span></span>| <span data-ttu-id="d7298-629">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-629">&lt;optional&gt;</span></span>|<span data-ttu-id="d7298-630">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d7298-630">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d7298-631">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d7298-631">Object</span></span>| <span data-ttu-id="d7298-632">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-632">&lt;optional&gt;</span></span>|<span data-ttu-id="d7298-633">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d7298-633">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d7298-634">function</span><span class="sxs-lookup"><span data-stu-id="d7298-634">function</span></span>| <span data-ttu-id="d7298-635">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-635">&lt;optional&gt;</span></span>|<span data-ttu-id="d7298-636">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-636">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d7298-637">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-637">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d7298-638">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d7298-638">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d7298-639">エラー</span><span class="sxs-lookup"><span data-stu-id="d7298-639">Errors</span></span>

| <span data-ttu-id="d7298-640">エラー コード</span><span class="sxs-lookup"><span data-stu-id="d7298-640">Error code</span></span> | <span data-ttu-id="d7298-641">説明</span><span class="sxs-lookup"><span data-stu-id="d7298-641">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="d7298-642">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="d7298-642">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="d7298-643">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="d7298-643">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="d7298-644">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="d7298-644">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d7298-645">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-645">Requirements</span></span>

|<span data-ttu-id="d7298-646">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-646">Requirement</span></span>| <span data-ttu-id="d7298-647">値</span><span class="sxs-lookup"><span data-stu-id="d7298-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-648">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-648">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-649">1.1</span><span class="sxs-lookup"><span data-stu-id="d7298-649">1.1</span></span>|
|[<span data-ttu-id="d7298-650">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-650">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-651">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d7298-651">ReadWriteItem</span></span>|
|[<span data-ttu-id="d7298-652">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-652">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-653">作成</span><span class="sxs-lookup"><span data-stu-id="d7298-653">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-654">例</span><span class="sxs-lookup"><span data-stu-id="d7298-654">Example</span></span>

```js
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="d7298-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d7298-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d7298-656">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="d7298-656">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="d7298-p142">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="d7298-p142">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="d7298-660">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="d7298-660">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="d7298-661">Office アドインを Outlook on the web で実行している場合、編集中のアイテム以外のアイテムに `addItemAttachmentAsync` メソッドでアイテムを添付できます。ただし、これはサポートされていないため、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="d7298-661">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7298-662">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d7298-662">Parameters</span></span>

|<span data-ttu-id="d7298-663">名前</span><span class="sxs-lookup"><span data-stu-id="d7298-663">Name</span></span>| <span data-ttu-id="d7298-664">型</span><span class="sxs-lookup"><span data-stu-id="d7298-664">Type</span></span>| <span data-ttu-id="d7298-665">属性</span><span class="sxs-lookup"><span data-stu-id="d7298-665">Attributes</span></span>| <span data-ttu-id="d7298-666">説明</span><span class="sxs-lookup"><span data-stu-id="d7298-666">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="d7298-667">String</span><span class="sxs-lookup"><span data-stu-id="d7298-667">String</span></span>||<span data-ttu-id="d7298-p143">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="d7298-p143">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="d7298-670">String</span><span class="sxs-lookup"><span data-stu-id="d7298-670">String</span></span>||<span data-ttu-id="d7298-671">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="d7298-671">The subject of the item to be attached.</span></span> <span data-ttu-id="d7298-672">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d7298-672">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="d7298-673">Object</span><span class="sxs-lookup"><span data-stu-id="d7298-673">Object</span></span>| <span data-ttu-id="d7298-674">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-674">&lt;optional&gt;</span></span>|<span data-ttu-id="d7298-675">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d7298-675">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d7298-676">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d7298-676">Object</span></span>| <span data-ttu-id="d7298-677">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-677">&lt;optional&gt;</span></span>|<span data-ttu-id="d7298-678">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d7298-678">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d7298-679">関数</span><span class="sxs-lookup"><span data-stu-id="d7298-679">function</span></span>| <span data-ttu-id="d7298-680">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-680">&lt;optional&gt;</span></span>|<span data-ttu-id="d7298-681">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-681">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d7298-682">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-682">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d7298-683">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d7298-683">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d7298-684">エラー</span><span class="sxs-lookup"><span data-stu-id="d7298-684">Errors</span></span>

| <span data-ttu-id="d7298-685">エラー コード</span><span class="sxs-lookup"><span data-stu-id="d7298-685">Error code</span></span> | <span data-ttu-id="d7298-686">説明</span><span class="sxs-lookup"><span data-stu-id="d7298-686">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="d7298-687">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="d7298-687">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d7298-688">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-688">Requirements</span></span>

|<span data-ttu-id="d7298-689">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-689">Requirement</span></span>| <span data-ttu-id="d7298-690">値</span><span class="sxs-lookup"><span data-stu-id="d7298-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-691">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-692">1.1</span><span class="sxs-lookup"><span data-stu-id="d7298-692">1.1</span></span>|
|[<span data-ttu-id="d7298-693">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-693">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-694">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d7298-694">ReadWriteItem</span></span>|
|[<span data-ttu-id="d7298-695">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-695">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-696">作成</span><span class="sxs-lookup"><span data-stu-id="d7298-696">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-697">例</span><span class="sxs-lookup"><span data-stu-id="d7298-697">Example</span></span>

<span data-ttu-id="d7298-698">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-698">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach (shortened for readability).
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="d7298-699">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="d7298-699">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="d7298-700">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-700">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d7298-701">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d7298-701">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d7298-702">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-702">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d7298-703">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="d7298-703">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="d7298-p145">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="d7298-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7298-707">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d7298-707">Parameters</span></span>

|<span data-ttu-id="d7298-708">名前</span><span class="sxs-lookup"><span data-stu-id="d7298-708">Name</span></span>| <span data-ttu-id="d7298-709">種類</span><span class="sxs-lookup"><span data-stu-id="d7298-709">Type</span></span>| <span data-ttu-id="d7298-710">説明</span><span class="sxs-lookup"><span data-stu-id="d7298-710">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="d7298-711">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="d7298-711">String &#124; Object</span></span>| |<span data-ttu-id="d7298-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="d7298-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d7298-714">**または**</span><span class="sxs-lookup"><span data-stu-id="d7298-714">**OR**</span></span><br/><span data-ttu-id="d7298-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="d7298-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="d7298-717">String</span><span class="sxs-lookup"><span data-stu-id="d7298-717">String</span></span> | <span data-ttu-id="d7298-718">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-718">&lt;optional&gt;</span></span> | <span data-ttu-id="d7298-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="d7298-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="d7298-721">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-721">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="d7298-722">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-722">&lt;optional&gt;</span></span> | <span data-ttu-id="d7298-723">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="d7298-723">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="d7298-724">String</span><span class="sxs-lookup"><span data-stu-id="d7298-724">String</span></span> | | <span data-ttu-id="d7298-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="d7298-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="d7298-727">String</span><span class="sxs-lookup"><span data-stu-id="d7298-727">String</span></span> | | <span data-ttu-id="d7298-728">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d7298-728">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="d7298-729">文字列</span><span class="sxs-lookup"><span data-stu-id="d7298-729">String</span></span> | | <span data-ttu-id="d7298-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="d7298-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="d7298-732">String</span><span class="sxs-lookup"><span data-stu-id="d7298-732">String</span></span> | | <span data-ttu-id="d7298-p151">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="d7298-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="d7298-736">function</span><span class="sxs-lookup"><span data-stu-id="d7298-736">function</span></span> | <span data-ttu-id="d7298-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-737">&lt;optional&gt;</span></span> | <span data-ttu-id="d7298-738">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-738">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d7298-739">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-739">Requirements</span></span>

|<span data-ttu-id="d7298-740">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-740">Requirement</span></span>| <span data-ttu-id="d7298-741">値</span><span class="sxs-lookup"><span data-stu-id="d7298-741">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-742">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-742">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-743">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-743">1.0</span></span>|
|[<span data-ttu-id="d7298-744">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-744">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-745">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-745">ReadItem</span></span>|
|[<span data-ttu-id="d7298-746">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-746">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-747">読み取り</span><span class="sxs-lookup"><span data-stu-id="d7298-747">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d7298-748">例</span><span class="sxs-lookup"><span data-stu-id="d7298-748">Examples</span></span>

<span data-ttu-id="d7298-749">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="d7298-749">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="d7298-750">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="d7298-750">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="d7298-751">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="d7298-751">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d7298-752">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="d7298-752">Reply with a body and a file attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="d7298-753">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="d7298-753">Reply with a body and an item attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="d7298-754">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="d7298-754">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="d7298-755">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="d7298-755">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="d7298-756">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-756">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d7298-757">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d7298-757">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d7298-758">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-758">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d7298-759">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="d7298-759">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="d7298-p152">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="d7298-p152">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7298-763">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d7298-763">Parameters</span></span>

|<span data-ttu-id="d7298-764">名前</span><span class="sxs-lookup"><span data-stu-id="d7298-764">Name</span></span>| <span data-ttu-id="d7298-765">種類</span><span class="sxs-lookup"><span data-stu-id="d7298-765">Type</span></span>| <span data-ttu-id="d7298-766">説明</span><span class="sxs-lookup"><span data-stu-id="d7298-766">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="d7298-767">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="d7298-767">String &#124; Object</span></span>| | <span data-ttu-id="d7298-p153">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="d7298-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d7298-770">**または**</span><span class="sxs-lookup"><span data-stu-id="d7298-770">**OR**</span></span><br/><span data-ttu-id="d7298-p154">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="d7298-p154">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="d7298-773">String</span><span class="sxs-lookup"><span data-stu-id="d7298-773">String</span></span> | <span data-ttu-id="d7298-774">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-774">&lt;optional&gt;</span></span> | <span data-ttu-id="d7298-p155">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="d7298-p155">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="d7298-777">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-777">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="d7298-778">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-778">&lt;optional&gt;</span></span> | <span data-ttu-id="d7298-779">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="d7298-779">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="d7298-780">String</span><span class="sxs-lookup"><span data-stu-id="d7298-780">String</span></span> | | <span data-ttu-id="d7298-p156">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="d7298-p156">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="d7298-783">String</span><span class="sxs-lookup"><span data-stu-id="d7298-783">String</span></span> | | <span data-ttu-id="d7298-784">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d7298-784">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="d7298-785">文字列</span><span class="sxs-lookup"><span data-stu-id="d7298-785">String</span></span> | | <span data-ttu-id="d7298-p157">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="d7298-p157">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="d7298-788">String</span><span class="sxs-lookup"><span data-stu-id="d7298-788">String</span></span> | | <span data-ttu-id="d7298-p158">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="d7298-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="d7298-792">function</span><span class="sxs-lookup"><span data-stu-id="d7298-792">function</span></span> | <span data-ttu-id="d7298-793">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-793">&lt;optional&gt;</span></span> | <span data-ttu-id="d7298-794">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-794">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d7298-795">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-795">Requirements</span></span>

|<span data-ttu-id="d7298-796">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-796">Requirement</span></span>| <span data-ttu-id="d7298-797">値</span><span class="sxs-lookup"><span data-stu-id="d7298-797">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-798">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-798">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-799">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-799">1.0</span></span>|
|[<span data-ttu-id="d7298-800">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-800">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-801">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-801">ReadItem</span></span>|
|[<span data-ttu-id="d7298-802">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-802">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-803">読み取り</span><span class="sxs-lookup"><span data-stu-id="d7298-803">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d7298-804">例</span><span class="sxs-lookup"><span data-stu-id="d7298-804">Examples</span></span>

<span data-ttu-id="d7298-805">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="d7298-805">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="d7298-806">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="d7298-806">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="d7298-807">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="d7298-807">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d7298-808">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="d7298-808">Reply with a body and a file attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="d7298-809">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="d7298-809">Reply with a body and an item attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="d7298-810">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="d7298-810">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-12"></a><span data-ttu-id="d7298-811">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="d7298-811">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="d7298-812">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="d7298-812">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d7298-813">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d7298-813">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-814">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-814">Requirements</span></span>

|<span data-ttu-id="d7298-815">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-815">Requirement</span></span>| <span data-ttu-id="d7298-816">値</span><span class="sxs-lookup"><span data-stu-id="d7298-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-817">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-818">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-818">1.0</span></span>|
|[<span data-ttu-id="d7298-819">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-820">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-820">ReadItem</span></span>|
|[<span data-ttu-id="d7298-821">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-822">読み取り</span><span class="sxs-lookup"><span data-stu-id="d7298-822">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7298-823">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d7298-823">Returns:</span></span>

<span data-ttu-id="d7298-824">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="d7298-824">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span></span>

##### <a name="example"></a><span data-ttu-id="d7298-825">例</span><span class="sxs-lookup"><span data-stu-id="d7298-825">Example</span></span>

<span data-ttu-id="d7298-826">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="d7298-826">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="d7298-827">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="d7298-827">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="d7298-828">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="d7298-828">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d7298-829">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d7298-829">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7298-830">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d7298-830">Parameters</span></span>

|<span data-ttu-id="d7298-831">名前</span><span class="sxs-lookup"><span data-stu-id="d7298-831">Name</span></span>| <span data-ttu-id="d7298-832">型</span><span class="sxs-lookup"><span data-stu-id="d7298-832">Type</span></span>| <span data-ttu-id="d7298-833">説明</span><span class="sxs-lookup"><span data-stu-id="d7298-833">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="d7298-834">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="d7298-834">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.2)|<span data-ttu-id="d7298-835">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="d7298-835">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7298-836">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-836">Requirements</span></span>

|<span data-ttu-id="d7298-837">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-837">Requirement</span></span>| <span data-ttu-id="d7298-838">値</span><span class="sxs-lookup"><span data-stu-id="d7298-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-839">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-840">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-840">1.0</span></span>|
|[<span data-ttu-id="d7298-841">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-841">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-842">制限あり</span><span class="sxs-lookup"><span data-stu-id="d7298-842">Restricted</span></span>|
|[<span data-ttu-id="d7298-843">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-843">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-844">読み取り</span><span class="sxs-lookup"><span data-stu-id="d7298-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7298-845">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d7298-845">Returns:</span></span>

<span data-ttu-id="d7298-846">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-846">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="d7298-847">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-847">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="d7298-848">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="d7298-848">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="d7298-849">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="d7298-849">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="d7298-850">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="d7298-850">Value of `entityType`</span></span> | <span data-ttu-id="d7298-851">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="d7298-851">Type of objects in returned array</span></span> | <span data-ttu-id="d7298-852">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="d7298-852">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="d7298-853">文字列</span><span class="sxs-lookup"><span data-stu-id="d7298-853">String</span></span> | <span data-ttu-id="d7298-854">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="d7298-854">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="d7298-855">連絡先</span><span class="sxs-lookup"><span data-stu-id="d7298-855">Contact</span></span> | <span data-ttu-id="d7298-856">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d7298-856">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="d7298-857">文字列</span><span class="sxs-lookup"><span data-stu-id="d7298-857">String</span></span> | <span data-ttu-id="d7298-858">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d7298-858">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="d7298-859">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="d7298-859">MeetingSuggestion</span></span> | <span data-ttu-id="d7298-860">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d7298-860">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="d7298-861">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="d7298-861">PhoneNumber</span></span> | <span data-ttu-id="d7298-862">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="d7298-862">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="d7298-863">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="d7298-863">TaskSuggestion</span></span> | <span data-ttu-id="d7298-864">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d7298-864">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="d7298-865">文字列</span><span class="sxs-lookup"><span data-stu-id="d7298-865">String</span></span> | <span data-ttu-id="d7298-866">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="d7298-866">**Restricted**</span></span> |

<span data-ttu-id="d7298-867">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="d7298-867">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

##### <a name="example"></a><span data-ttu-id="d7298-868">例</span><span class="sxs-lookup"><span data-stu-id="d7298-868">Example</span></span>

<span data-ttu-id="d7298-869">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="d7298-869">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="d7298-870">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="d7298-870">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="d7298-871">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-871">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d7298-872">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d7298-872">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d7298-873">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-873">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7298-874">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d7298-874">Parameters</span></span>

|<span data-ttu-id="d7298-875">名前</span><span class="sxs-lookup"><span data-stu-id="d7298-875">Name</span></span>| <span data-ttu-id="d7298-876">型</span><span class="sxs-lookup"><span data-stu-id="d7298-876">Type</span></span>| <span data-ttu-id="d7298-877">説明</span><span class="sxs-lookup"><span data-stu-id="d7298-877">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="d7298-878">String</span><span class="sxs-lookup"><span data-stu-id="d7298-878">String</span></span>|<span data-ttu-id="d7298-879">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="d7298-879">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7298-880">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-880">Requirements</span></span>

|<span data-ttu-id="d7298-881">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-881">Requirement</span></span>| <span data-ttu-id="d7298-882">値</span><span class="sxs-lookup"><span data-stu-id="d7298-882">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-883">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-883">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-884">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-884">1.0</span></span>|
|[<span data-ttu-id="d7298-885">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-885">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-886">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-886">ReadItem</span></span>|
|[<span data-ttu-id="d7298-887">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-887">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-888">読み取り</span><span class="sxs-lookup"><span data-stu-id="d7298-888">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7298-889">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d7298-889">Returns:</span></span>

<span data-ttu-id="d7298-p160">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="d7298-892">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="d7298-892">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="d7298-893">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d7298-893">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="d7298-894">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-894">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d7298-895">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d7298-895">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d7298-p161">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="d7298-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d7298-899">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="d7298-899">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d7298-900">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="d7298-900">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="d7298-p162">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="d7298-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7298-903">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-903">Requirements</span></span>

|<span data-ttu-id="d7298-904">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-904">Requirement</span></span>| <span data-ttu-id="d7298-905">値</span><span class="sxs-lookup"><span data-stu-id="d7298-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-906">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-907">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-907">1.0</span></span>|
|[<span data-ttu-id="d7298-908">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-908">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-909">ReadItem</span></span>|
|[<span data-ttu-id="d7298-910">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-910">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-911">読み取り</span><span class="sxs-lookup"><span data-stu-id="d7298-911">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7298-912">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d7298-912">Returns:</span></span>

<span data-ttu-id="d7298-p163">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="d7298-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="d7298-915">型: Object</span><span class="sxs-lookup"><span data-stu-id="d7298-915">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="d7298-916">例</span><span class="sxs-lookup"><span data-stu-id="d7298-916">Example</span></span>

<span data-ttu-id="d7298-917">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="d7298-917">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="d7298-918">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="d7298-918">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="d7298-919">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-919">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d7298-920">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d7298-920">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d7298-921">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="d7298-921">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="d7298-p164">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="d7298-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7298-924">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d7298-924">Parameters</span></span>

|<span data-ttu-id="d7298-925">名前</span><span class="sxs-lookup"><span data-stu-id="d7298-925">Name</span></span>| <span data-ttu-id="d7298-926">種類</span><span class="sxs-lookup"><span data-stu-id="d7298-926">Type</span></span>| <span data-ttu-id="d7298-927">説明</span><span class="sxs-lookup"><span data-stu-id="d7298-927">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="d7298-928">String</span><span class="sxs-lookup"><span data-stu-id="d7298-928">String</span></span>|<span data-ttu-id="d7298-929">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="d7298-929">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7298-930">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-930">Requirements</span></span>

|<span data-ttu-id="d7298-931">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-931">Requirement</span></span>| <span data-ttu-id="d7298-932">値</span><span class="sxs-lookup"><span data-stu-id="d7298-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-933">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-934">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-934">1.0</span></span>|
|[<span data-ttu-id="d7298-935">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-936">ReadItem</span></span>|
|[<span data-ttu-id="d7298-937">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-938">読み取り</span><span class="sxs-lookup"><span data-stu-id="d7298-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7298-939">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d7298-939">Returns:</span></span>

<span data-ttu-id="d7298-940">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="d7298-940">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="d7298-941">型: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="d7298-941">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="d7298-942">例</span><span class="sxs-lookup"><span data-stu-id="d7298-942">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="d7298-943">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="d7298-943">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="d7298-944">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-944">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="d7298-p165">選択されていない状態でカーソルが本文または件名にある場合、メソッドは選択されたデータに対し空の文字列を返します。本文または件名以外のフィールドが選択されている場合には、メソッドは`InvalidSelection`エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-p165">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7298-947">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d7298-947">Parameters</span></span>

|<span data-ttu-id="d7298-948">名前</span><span class="sxs-lookup"><span data-stu-id="d7298-948">Name</span></span>| <span data-ttu-id="d7298-949">型</span><span class="sxs-lookup"><span data-stu-id="d7298-949">Type</span></span>| <span data-ttu-id="d7298-950">属性</span><span class="sxs-lookup"><span data-stu-id="d7298-950">Attributes</span></span>| <span data-ttu-id="d7298-951">説明</span><span class="sxs-lookup"><span data-stu-id="d7298-951">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="d7298-952">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d7298-952">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="d7298-p166">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="d7298-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="d7298-956">Object</span><span class="sxs-lookup"><span data-stu-id="d7298-956">Object</span></span>| <span data-ttu-id="d7298-957">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-957">&lt;optional&gt;</span></span>|<span data-ttu-id="d7298-958">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d7298-958">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d7298-959">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d7298-959">Object</span></span>| <span data-ttu-id="d7298-960">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-960">&lt;optional&gt;</span></span>|<span data-ttu-id="d7298-961">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d7298-961">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d7298-962">function</span><span class="sxs-lookup"><span data-stu-id="d7298-962">function</span></span>||<span data-ttu-id="d7298-963">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-963">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d7298-964">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="d7298-964">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="d7298-965">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="d7298-965">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7298-966">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-966">Requirements</span></span>

|<span data-ttu-id="d7298-967">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-967">Requirement</span></span>| <span data-ttu-id="d7298-968">値</span><span class="sxs-lookup"><span data-stu-id="d7298-968">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-969">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-969">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-970">1.2</span><span class="sxs-lookup"><span data-stu-id="d7298-970">1.2</span></span>|
|[<span data-ttu-id="d7298-971">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-971">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-972">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-972">ReadItem</span></span>|
|[<span data-ttu-id="d7298-973">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-973">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-974">作成</span><span class="sxs-lookup"><span data-stu-id="d7298-974">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7298-975">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d7298-975">Returns:</span></span>

<span data-ttu-id="d7298-976">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="d7298-976">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="d7298-977">型:String</span><span class="sxs-lookup"><span data-stu-id="d7298-977">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d7298-978">例</span><span class="sxs-lookup"><span data-stu-id="d7298-978">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  console.log("Selected text in " + prop + ": " + text);
}
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="d7298-979">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d7298-979">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="d7298-980">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="d7298-980">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="d7298-p168">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="d7298-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7298-984">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d7298-984">Parameters</span></span>

|<span data-ttu-id="d7298-985">名前</span><span class="sxs-lookup"><span data-stu-id="d7298-985">Name</span></span>| <span data-ttu-id="d7298-986">型</span><span class="sxs-lookup"><span data-stu-id="d7298-986">Type</span></span>| <span data-ttu-id="d7298-987">属性</span><span class="sxs-lookup"><span data-stu-id="d7298-987">Attributes</span></span>| <span data-ttu-id="d7298-988">説明</span><span class="sxs-lookup"><span data-stu-id="d7298-988">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d7298-989">function</span><span class="sxs-lookup"><span data-stu-id="d7298-989">function</span></span>||<span data-ttu-id="d7298-990">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-990">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d7298-991">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-991">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="d7298-992">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="d7298-992">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="d7298-993">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d7298-993">Object</span></span>| <span data-ttu-id="d7298-994">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-994">&lt;optional&gt;</span></span>|<span data-ttu-id="d7298-995">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="d7298-995">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="d7298-996">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="d7298-996">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7298-997">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-997">Requirements</span></span>

|<span data-ttu-id="d7298-998">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-998">Requirement</span></span>| <span data-ttu-id="d7298-999">値</span><span class="sxs-lookup"><span data-stu-id="d7298-999">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-1000">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-1000">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-1001">1.0</span><span class="sxs-lookup"><span data-stu-id="d7298-1001">1.0</span></span>|
|[<span data-ttu-id="d7298-1002">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-1002">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-1003">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7298-1003">ReadItem</span></span>|
|[<span data-ttu-id="d7298-1004">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-1004">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-1005">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d7298-1005">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-1006">例</span><span class="sxs-lookup"><span data-stu-id="d7298-1006">Example</span></span>

<span data-ttu-id="d7298-p171">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="d7298-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="d7298-1010">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d7298-1010">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="d7298-1011">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="d7298-1011">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="d7298-1012">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="d7298-1012">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="d7298-1013">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="d7298-1013">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="d7298-1014">Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="d7298-1014">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="d7298-1015">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="d7298-1015">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7298-1016">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d7298-1016">Parameters</span></span>

|<span data-ttu-id="d7298-1017">名前</span><span class="sxs-lookup"><span data-stu-id="d7298-1017">Name</span></span>| <span data-ttu-id="d7298-1018">型</span><span class="sxs-lookup"><span data-stu-id="d7298-1018">Type</span></span>| <span data-ttu-id="d7298-1019">属性</span><span class="sxs-lookup"><span data-stu-id="d7298-1019">Attributes</span></span>| <span data-ttu-id="d7298-1020">説明</span><span class="sxs-lookup"><span data-stu-id="d7298-1020">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="d7298-1021">String</span><span class="sxs-lookup"><span data-stu-id="d7298-1021">String</span></span>||<span data-ttu-id="d7298-1022">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="d7298-1022">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="d7298-1023">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d7298-1023">Object</span></span>| <span data-ttu-id="d7298-1024">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-1024">&lt;optional&gt;</span></span>|<span data-ttu-id="d7298-1025">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d7298-1025">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d7298-1026">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d7298-1026">Object</span></span>| <span data-ttu-id="d7298-1027">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-1027">&lt;optional&gt;</span></span>|<span data-ttu-id="d7298-1028">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d7298-1028">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d7298-1029">関数</span><span class="sxs-lookup"><span data-stu-id="d7298-1029">function</span></span>| <span data-ttu-id="d7298-1030">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-1030">&lt;optional&gt;</span></span>|<span data-ttu-id="d7298-1031">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-1031">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d7298-1032">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="d7298-1032">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d7298-1033">エラー</span><span class="sxs-lookup"><span data-stu-id="d7298-1033">Errors</span></span>

| <span data-ttu-id="d7298-1034">エラー コード</span><span class="sxs-lookup"><span data-stu-id="d7298-1034">Error code</span></span> | <span data-ttu-id="d7298-1035">説明</span><span class="sxs-lookup"><span data-stu-id="d7298-1035">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="d7298-1036">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="d7298-1036">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d7298-1037">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7298-1037">Requirements</span></span>

|<span data-ttu-id="d7298-1038">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-1038">Requirement</span></span>| <span data-ttu-id="d7298-1039">値</span><span class="sxs-lookup"><span data-stu-id="d7298-1039">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-1040">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-1040">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-1041">1.1</span><span class="sxs-lookup"><span data-stu-id="d7298-1041">1.1</span></span>|
|[<span data-ttu-id="d7298-1042">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-1042">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-1043">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d7298-1043">ReadWriteItem</span></span>|
|[<span data-ttu-id="d7298-1044">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-1044">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-1045">作成</span><span class="sxs-lookup"><span data-stu-id="d7298-1045">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-1046">例</span><span class="sxs-lookup"><span data-stu-id="d7298-1046">Example</span></span>

<span data-ttu-id="d7298-1047">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="d7298-1047">The following code removes an attachment with an identifier of '0'.</span></span>

```js
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="d7298-1048">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="d7298-1048">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="d7298-1049">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="d7298-1049">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="d7298-p173">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="d7298-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7298-1053">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d7298-1053">Parameters</span></span>

|<span data-ttu-id="d7298-1054">名前</span><span class="sxs-lookup"><span data-stu-id="d7298-1054">Name</span></span>| <span data-ttu-id="d7298-1055">型</span><span class="sxs-lookup"><span data-stu-id="d7298-1055">Type</span></span>| <span data-ttu-id="d7298-1056">属性</span><span class="sxs-lookup"><span data-stu-id="d7298-1056">Attributes</span></span>| <span data-ttu-id="d7298-1057">説明</span><span class="sxs-lookup"><span data-stu-id="d7298-1057">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="d7298-1058">String</span><span class="sxs-lookup"><span data-stu-id="d7298-1058">String</span></span>||<span data-ttu-id="d7298-p174">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="d7298-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="d7298-1062">Object</span><span class="sxs-lookup"><span data-stu-id="d7298-1062">Object</span></span>| <span data-ttu-id="d7298-1063">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="d7298-1064">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d7298-1064">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d7298-1065">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d7298-1065">Object</span></span>| <span data-ttu-id="d7298-1066">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="d7298-1067">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d7298-1067">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="d7298-1068">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d7298-1068">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="d7298-1069">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7298-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="d7298-1070">`text` の場合、Outlook on the web とデスクトップ クライアントでは現在のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-1070">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="d7298-1071">フィールドが HTML エディターの場合、データが HTML の場合でもテキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-1071">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="d7298-1072">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Outlook on the web では現在のスタイルが適用され、Outlook デスクトップ クライアントでは既定のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-1072">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="d7298-1073">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-1073">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="d7298-1074">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="d7298-1074">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="d7298-1075">function</span><span class="sxs-lookup"><span data-stu-id="d7298-1075">function</span></span>||<span data-ttu-id="d7298-1076">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="d7298-1076">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d7298-1077">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-1077">Requirements</span></span>

|<span data-ttu-id="d7298-1078">要件</span><span class="sxs-lookup"><span data-stu-id="d7298-1078">Requirement</span></span>| <span data-ttu-id="d7298-1079">値</span><span class="sxs-lookup"><span data-stu-id="d7298-1079">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7298-1080">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d7298-1080">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7298-1081">1.2</span><span class="sxs-lookup"><span data-stu-id="d7298-1081">1.2</span></span>|
|[<span data-ttu-id="d7298-1082">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d7298-1082">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7298-1083">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d7298-1083">ReadWriteItem</span></span>|
|[<span data-ttu-id="d7298-1084">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d7298-1084">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7298-1085">作成</span><span class="sxs-lookup"><span data-stu-id="d7298-1085">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d7298-1086">例</span><span class="sxs-lookup"><span data-stu-id="d7298-1086">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
