---
title: Office. メールボックス-要件セット1.2
description: ''
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 97fa271f500e89c6ce69d82b95a0818f6d5bc7d4
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001608"
---
# <a name="item"></a><span data-ttu-id="07e18-102">item</span><span class="sxs-lookup"><span data-stu-id="07e18-102">item</span></span>

### <span data-ttu-id="07e18-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="07e18-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="07e18-p102">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="07e18-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-107">Requirements</span></span>

|<span data-ttu-id="07e18-108">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-108">Requirement</span></span>| <span data-ttu-id="07e18-109">値</span><span class="sxs-lookup"><span data-stu-id="07e18-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-111">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-111">1.0</span></span>|
|[<span data-ttu-id="07e18-112">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-113">制限あり</span><span class="sxs-lookup"><span data-stu-id="07e18-113">Restricted</span></span>|
|[<span data-ttu-id="07e18-114">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-115">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="07e18-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="07e18-116">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="07e18-116">Members and methods</span></span>

| <span data-ttu-id="07e18-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="07e18-117">Member</span></span> | <span data-ttu-id="07e18-118">種類</span><span class="sxs-lookup"><span data-stu-id="07e18-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="07e18-119">attachments</span><span class="sxs-lookup"><span data-stu-id="07e18-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="07e18-120">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-120">Member</span></span> |
| [<span data-ttu-id="07e18-121">bcc</span><span class="sxs-lookup"><span data-stu-id="07e18-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="07e18-122">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-122">Member</span></span> |
| [<span data-ttu-id="07e18-123">body</span><span class="sxs-lookup"><span data-stu-id="07e18-123">body</span></span>](#body-body) | <span data-ttu-id="07e18-124">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-124">Member</span></span> |
| [<span data-ttu-id="07e18-125">cc</span><span class="sxs-lookup"><span data-stu-id="07e18-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="07e18-126">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-126">Member</span></span> |
| [<span data-ttu-id="07e18-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="07e18-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="07e18-128">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-128">Member</span></span> |
| [<span data-ttu-id="07e18-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="07e18-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="07e18-130">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-130">Member</span></span> |
| [<span data-ttu-id="07e18-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="07e18-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="07e18-132">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-132">Member</span></span> |
| [<span data-ttu-id="07e18-133">end</span><span class="sxs-lookup"><span data-stu-id="07e18-133">end</span></span>](#end-datetime) | <span data-ttu-id="07e18-134">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-134">Member</span></span> |
| [<span data-ttu-id="07e18-135">from</span><span class="sxs-lookup"><span data-stu-id="07e18-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="07e18-136">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-136">Member</span></span> |
| [<span data-ttu-id="07e18-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="07e18-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="07e18-138">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-138">Member</span></span> |
| [<span data-ttu-id="07e18-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="07e18-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="07e18-140">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-140">Member</span></span> |
| [<span data-ttu-id="07e18-141">itemId</span><span class="sxs-lookup"><span data-stu-id="07e18-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="07e18-142">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-142">Member</span></span> |
| [<span data-ttu-id="07e18-143">itemType</span><span class="sxs-lookup"><span data-stu-id="07e18-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="07e18-144">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-144">Member</span></span> |
| [<span data-ttu-id="07e18-145">location</span><span class="sxs-lookup"><span data-stu-id="07e18-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="07e18-146">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-146">Member</span></span> |
| [<span data-ttu-id="07e18-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="07e18-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="07e18-148">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-148">Member</span></span> |
| [<span data-ttu-id="07e18-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="07e18-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="07e18-150">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-150">Member</span></span> |
| [<span data-ttu-id="07e18-151">organizer</span><span class="sxs-lookup"><span data-stu-id="07e18-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="07e18-152">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-152">Member</span></span> |
| [<span data-ttu-id="07e18-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="07e18-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="07e18-154">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-154">Member</span></span> |
| [<span data-ttu-id="07e18-155">sender</span><span class="sxs-lookup"><span data-stu-id="07e18-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="07e18-156">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-156">Member</span></span> |
| [<span data-ttu-id="07e18-157">start</span><span class="sxs-lookup"><span data-stu-id="07e18-157">start</span></span>](#start-datetime) | <span data-ttu-id="07e18-158">Member</span><span class="sxs-lookup"><span data-stu-id="07e18-158">Member</span></span> |
| [<span data-ttu-id="07e18-159">subject</span><span class="sxs-lookup"><span data-stu-id="07e18-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="07e18-160">メンバー</span><span class="sxs-lookup"><span data-stu-id="07e18-160">Member</span></span> |
| [<span data-ttu-id="07e18-161">to</span><span class="sxs-lookup"><span data-stu-id="07e18-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="07e18-162">メンバー</span><span class="sxs-lookup"><span data-stu-id="07e18-162">Member</span></span> |
| [<span data-ttu-id="07e18-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="07e18-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="07e18-164">メソッド</span><span class="sxs-lookup"><span data-stu-id="07e18-164">Method</span></span> |
| [<span data-ttu-id="07e18-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="07e18-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="07e18-166">メソッド</span><span class="sxs-lookup"><span data-stu-id="07e18-166">Method</span></span> |
| [<span data-ttu-id="07e18-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="07e18-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="07e18-168">メソッド</span><span class="sxs-lookup"><span data-stu-id="07e18-168">Method</span></span> |
| [<span data-ttu-id="07e18-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="07e18-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="07e18-170">メソッド</span><span class="sxs-lookup"><span data-stu-id="07e18-170">Method</span></span> |
| [<span data-ttu-id="07e18-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="07e18-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="07e18-172">メソッド</span><span class="sxs-lookup"><span data-stu-id="07e18-172">Method</span></span> |
| [<span data-ttu-id="07e18-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="07e18-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="07e18-174">メソッド</span><span class="sxs-lookup"><span data-stu-id="07e18-174">Method</span></span> |
| [<span data-ttu-id="07e18-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="07e18-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="07e18-176">メソッド</span><span class="sxs-lookup"><span data-stu-id="07e18-176">Method</span></span> |
| [<span data-ttu-id="07e18-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="07e18-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="07e18-178">メソッド</span><span class="sxs-lookup"><span data-stu-id="07e18-178">Method</span></span> |
| [<span data-ttu-id="07e18-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="07e18-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="07e18-180">メソッド</span><span class="sxs-lookup"><span data-stu-id="07e18-180">Method</span></span> |
| [<span data-ttu-id="07e18-181">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="07e18-181">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="07e18-182">メソッド</span><span class="sxs-lookup"><span data-stu-id="07e18-182">Method</span></span> |
| [<span data-ttu-id="07e18-183">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="07e18-183">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="07e18-184">メソッド</span><span class="sxs-lookup"><span data-stu-id="07e18-184">Method</span></span> |
| [<span data-ttu-id="07e18-185">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="07e18-185">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="07e18-186">メソッド</span><span class="sxs-lookup"><span data-stu-id="07e18-186">Method</span></span> |
| [<span data-ttu-id="07e18-187">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="07e18-187">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="07e18-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="07e18-188">Method</span></span> |

### <a name="example"></a><span data-ttu-id="07e18-189">例</span><span class="sxs-lookup"><span data-stu-id="07e18-189">Example</span></span>

<span data-ttu-id="07e18-190">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="07e18-190">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="07e18-191">Members</span><span class="sxs-lookup"><span data-stu-id="07e18-191">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-12"></a><span data-ttu-id="07e18-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="07e18-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

<span data-ttu-id="07e18-p103">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="07e18-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="07e18-195">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="07e18-195">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="07e18-196">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="07e18-196">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="07e18-197">型</span><span class="sxs-lookup"><span data-stu-id="07e18-197">Type</span></span>

*   <span data-ttu-id="07e18-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="07e18-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-199">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-199">Requirements</span></span>

|<span data-ttu-id="07e18-200">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-200">Requirement</span></span>| <span data-ttu-id="07e18-201">値</span><span class="sxs-lookup"><span data-stu-id="07e18-201">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-202">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-202">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-203">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-203">1.0</span></span>|
|[<span data-ttu-id="07e18-204">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-204">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-205">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-205">ReadItem</span></span>|
|[<span data-ttu-id="07e18-206">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-206">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-207">読み取り</span><span class="sxs-lookup"><span data-stu-id="07e18-207">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-208">例</span><span class="sxs-lookup"><span data-stu-id="07e18-208">Example</span></span>

<span data-ttu-id="07e18-209">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="07e18-209">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="07e18-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="07e18-211">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="07e18-211">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="07e18-212">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="07e18-212">Compose mode only.</span></span>

<span data-ttu-id="07e18-213">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="07e18-213">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="07e18-214">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-214">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="07e18-215">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="07e18-215">Get 500 members maximum.</span></span>
- <span data-ttu-id="07e18-216">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="07e18-216">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="07e18-217">型</span><span class="sxs-lookup"><span data-stu-id="07e18-217">Type</span></span>

*   [<span data-ttu-id="07e18-218">受信者</span><span class="sxs-lookup"><span data-stu-id="07e18-218">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="07e18-219">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-219">Requirements</span></span>

|<span data-ttu-id="07e18-220">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-220">Requirement</span></span>| <span data-ttu-id="07e18-221">値</span><span class="sxs-lookup"><span data-stu-id="07e18-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-222">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-222">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-223">1.1</span><span class="sxs-lookup"><span data-stu-id="07e18-223">1.1</span></span>|
|[<span data-ttu-id="07e18-224">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-224">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-225">ReadItem</span></span>|
|[<span data-ttu-id="07e18-226">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-226">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-227">作成</span><span class="sxs-lookup"><span data-stu-id="07e18-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-228">例</span><span class="sxs-lookup"><span data-stu-id="07e18-228">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-12"></a><span data-ttu-id="07e18-229">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-229">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span></span>

<span data-ttu-id="07e18-230">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="07e18-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="07e18-231">型</span><span class="sxs-lookup"><span data-stu-id="07e18-231">Type</span></span>

*   [<span data-ttu-id="07e18-232">Body</span><span class="sxs-lookup"><span data-stu-id="07e18-232">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="07e18-233">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-233">Requirements</span></span>

|<span data-ttu-id="07e18-234">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-234">Requirement</span></span>| <span data-ttu-id="07e18-235">値</span><span class="sxs-lookup"><span data-stu-id="07e18-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-236">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-237">1.1</span><span class="sxs-lookup"><span data-stu-id="07e18-237">1.1</span></span>|
|[<span data-ttu-id="07e18-238">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-239">ReadItem</span></span>|
|[<span data-ttu-id="07e18-240">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-241">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="07e18-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-242">例</span><span class="sxs-lookup"><span data-stu-id="07e18-242">Example</span></span>

<span data-ttu-id="07e18-243">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="07e18-243">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="07e18-244">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="07e18-244">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="07e18-245">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-245">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="07e18-246">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="07e18-246">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="07e18-247">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="07e18-247">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07e18-248">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="07e18-248">Read mode</span></span>

<span data-ttu-id="07e18-249">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-249">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="07e18-250">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="07e18-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="07e18-251">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="07e18-251">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="07e18-252">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="07e18-252">Compose mode</span></span>

<span data-ttu-id="07e18-253">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="07e18-254">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="07e18-254">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="07e18-255">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-255">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="07e18-256">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="07e18-256">Get 500 members maximum.</span></span>
- <span data-ttu-id="07e18-257">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="07e18-257">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="07e18-258">型</span><span class="sxs-lookup"><span data-stu-id="07e18-258">Type</span></span>

*   <span data-ttu-id="07e18-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-260">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-260">Requirements</span></span>

|<span data-ttu-id="07e18-261">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-261">Requirement</span></span>| <span data-ttu-id="07e18-262">値</span><span class="sxs-lookup"><span data-stu-id="07e18-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-263">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-264">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-264">1.0</span></span>|
|[<span data-ttu-id="07e18-265">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-266">ReadItem</span></span>|
|[<span data-ttu-id="07e18-267">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-268">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="07e18-268">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="07e18-269">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="07e18-269">(nullable) conversationId: String</span></span>

<span data-ttu-id="07e18-270">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="07e18-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="07e18-p110">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="07e18-p110">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="07e18-p111">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-p111">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="07e18-275">Type</span><span class="sxs-lookup"><span data-stu-id="07e18-275">Type</span></span>

*   <span data-ttu-id="07e18-276">String</span><span class="sxs-lookup"><span data-stu-id="07e18-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-277">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-277">Requirements</span></span>

|<span data-ttu-id="07e18-278">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-278">Requirement</span></span>| <span data-ttu-id="07e18-279">値</span><span class="sxs-lookup"><span data-stu-id="07e18-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-280">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-281">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-281">1.0</span></span>|
|[<span data-ttu-id="07e18-282">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-283">ReadItem</span></span>|
|[<span data-ttu-id="07e18-284">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-285">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="07e18-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-286">例</span><span class="sxs-lookup"><span data-stu-id="07e18-286">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="07e18-287">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="07e18-287">dateTimeCreated: Date</span></span>

<span data-ttu-id="07e18-p112">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="07e18-p112">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="07e18-290">型</span><span class="sxs-lookup"><span data-stu-id="07e18-290">Type</span></span>

*   <span data-ttu-id="07e18-291">日付</span><span class="sxs-lookup"><span data-stu-id="07e18-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-292">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-292">Requirements</span></span>

|<span data-ttu-id="07e18-293">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-293">Requirement</span></span>| <span data-ttu-id="07e18-294">値</span><span class="sxs-lookup"><span data-stu-id="07e18-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-295">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-296">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-296">1.0</span></span>|
|[<span data-ttu-id="07e18-297">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-298">ReadItem</span></span>|
|[<span data-ttu-id="07e18-299">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-300">読み取り</span><span class="sxs-lookup"><span data-stu-id="07e18-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-301">例</span><span class="sxs-lookup"><span data-stu-id="07e18-301">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="07e18-302">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="07e18-302">dateTimeModified: Date</span></span>

<span data-ttu-id="07e18-p113">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="07e18-p113">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="07e18-305">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="07e18-305">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="07e18-306">種類</span><span class="sxs-lookup"><span data-stu-id="07e18-306">Type</span></span>

*   <span data-ttu-id="07e18-307">日付</span><span class="sxs-lookup"><span data-stu-id="07e18-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-308">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-308">Requirements</span></span>

|<span data-ttu-id="07e18-309">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-309">Requirement</span></span>| <span data-ttu-id="07e18-310">値</span><span class="sxs-lookup"><span data-stu-id="07e18-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-312">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-312">1.0</span></span>|
|[<span data-ttu-id="07e18-313">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-314">ReadItem</span></span>|
|[<span data-ttu-id="07e18-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-316">読み取り</span><span class="sxs-lookup"><span data-stu-id="07e18-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-317">例</span><span class="sxs-lookup"><span data-stu-id="07e18-317">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="07e18-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="07e18-319">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="07e18-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="07e18-p114">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="07e18-p114">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07e18-322">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="07e18-322">Read mode</span></span>

<span data-ttu-id="07e18-323">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-323">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="07e18-324">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="07e18-324">Compose mode</span></span>

<span data-ttu-id="07e18-325">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="07e18-326">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="07e18-326">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="07e18-327">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="07e18-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="07e18-328">型</span><span class="sxs-lookup"><span data-stu-id="07e18-328">Type</span></span>

*   <span data-ttu-id="07e18-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-330">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-330">Requirements</span></span>

|<span data-ttu-id="07e18-331">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-331">Requirement</span></span>| <span data-ttu-id="07e18-332">値</span><span class="sxs-lookup"><span data-stu-id="07e18-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-333">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-334">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-334">1.0</span></span>|
|[<span data-ttu-id="07e18-335">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-336">ReadItem</span></span>|
|[<span data-ttu-id="07e18-337">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-338">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="07e18-338">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="07e18-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="07e18-p115">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="07e18-p115">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="07e18-p116">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="07e18-p116">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="07e18-344">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="07e18-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="07e18-345">型</span><span class="sxs-lookup"><span data-stu-id="07e18-345">Type</span></span>

*   [<span data-ttu-id="07e18-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="07e18-346">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="07e18-347">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-347">Requirements</span></span>

|<span data-ttu-id="07e18-348">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-348">Requirement</span></span>| <span data-ttu-id="07e18-349">値</span><span class="sxs-lookup"><span data-stu-id="07e18-349">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-350">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-350">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-351">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-351">1.0</span></span>|
|[<span data-ttu-id="07e18-352">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-352">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-353">ReadItem</span></span>|
|[<span data-ttu-id="07e18-354">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-354">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-355">読み取り</span><span class="sxs-lookup"><span data-stu-id="07e18-355">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-356">例</span><span class="sxs-lookup"><span data-stu-id="07e18-356">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="07e18-357">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="07e18-357">internetMessageId: String</span></span>

<span data-ttu-id="07e18-p117">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="07e18-p117">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="07e18-360">Type</span><span class="sxs-lookup"><span data-stu-id="07e18-360">Type</span></span>

*   <span data-ttu-id="07e18-361">String</span><span class="sxs-lookup"><span data-stu-id="07e18-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-362">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-362">Requirements</span></span>

|<span data-ttu-id="07e18-363">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-363">Requirement</span></span>| <span data-ttu-id="07e18-364">値</span><span class="sxs-lookup"><span data-stu-id="07e18-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-365">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-366">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-366">1.0</span></span>|
|[<span data-ttu-id="07e18-367">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-368">ReadItem</span></span>|
|[<span data-ttu-id="07e18-369">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-370">読み取り</span><span class="sxs-lookup"><span data-stu-id="07e18-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-371">例</span><span class="sxs-lookup"><span data-stu-id="07e18-371">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="07e18-372">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="07e18-372">itemClass: String</span></span>

<span data-ttu-id="07e18-p118">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="07e18-p118">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="07e18-p119">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="07e18-p119">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="07e18-377">型</span><span class="sxs-lookup"><span data-stu-id="07e18-377">Type</span></span> | <span data-ttu-id="07e18-378">説明</span><span class="sxs-lookup"><span data-stu-id="07e18-378">Description</span></span> | <span data-ttu-id="07e18-379">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="07e18-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="07e18-380">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="07e18-380">Appointment items</span></span> | <span data-ttu-id="07e18-381">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="07e18-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="07e18-382">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="07e18-382">Message items</span></span> | <span data-ttu-id="07e18-383">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="07e18-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="07e18-384">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="07e18-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="07e18-385">Type</span><span class="sxs-lookup"><span data-stu-id="07e18-385">Type</span></span>

*   <span data-ttu-id="07e18-386">String</span><span class="sxs-lookup"><span data-stu-id="07e18-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-387">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-387">Requirements</span></span>

|<span data-ttu-id="07e18-388">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-388">Requirement</span></span>| <span data-ttu-id="07e18-389">値</span><span class="sxs-lookup"><span data-stu-id="07e18-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-390">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-391">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-391">1.0</span></span>|
|[<span data-ttu-id="07e18-392">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-393">ReadItem</span></span>|
|[<span data-ttu-id="07e18-394">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-395">読み取り</span><span class="sxs-lookup"><span data-stu-id="07e18-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-396">例</span><span class="sxs-lookup"><span data-stu-id="07e18-396">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="07e18-397">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="07e18-397">(nullable) itemId: String</span></span>

<span data-ttu-id="07e18-398">現在のアイテムの[Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)を取得します。</span><span class="sxs-lookup"><span data-stu-id="07e18-398">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item.</span></span> <span data-ttu-id="07e18-399">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="07e18-399">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="07e18-400">`itemId`プロパティによって返される識別子は、 [Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)と同じです。</span><span class="sxs-lookup"><span data-stu-id="07e18-400">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="07e18-401">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="07e18-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="07e18-402">この値を使用して REST API を呼び出す前に、を`Office.context.mailbox.convertToRestId`使用して変換する必要があります。これは、要件セット1.3 から開始できます。</span><span class="sxs-lookup"><span data-stu-id="07e18-402">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="07e18-403">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="07e18-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="07e18-404">Type</span><span class="sxs-lookup"><span data-stu-id="07e18-404">Type</span></span>

*   <span data-ttu-id="07e18-405">String</span><span class="sxs-lookup"><span data-stu-id="07e18-405">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-406">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-406">Requirements</span></span>

|<span data-ttu-id="07e18-407">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-407">Requirement</span></span>| <span data-ttu-id="07e18-408">値</span><span class="sxs-lookup"><span data-stu-id="07e18-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-409">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-410">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-410">1.0</span></span>|
|[<span data-ttu-id="07e18-411">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-412">ReadItem</span></span>|
|[<span data-ttu-id="07e18-413">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-414">読み取り</span><span class="sxs-lookup"><span data-stu-id="07e18-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-415">例</span><span class="sxs-lookup"><span data-stu-id="07e18-415">Example</span></span>

<span data-ttu-id="07e18-p122">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-12"></a><span data-ttu-id="07e18-418">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-418">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span></span>

<span data-ttu-id="07e18-419">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="07e18-419">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="07e18-420">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="07e18-420">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="07e18-421">型</span><span class="sxs-lookup"><span data-stu-id="07e18-421">Type</span></span>

*   [<span data-ttu-id="07e18-422">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="07e18-422">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="07e18-423">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-423">Requirements</span></span>

|<span data-ttu-id="07e18-424">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-424">Requirement</span></span>| <span data-ttu-id="07e18-425">値</span><span class="sxs-lookup"><span data-stu-id="07e18-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-426">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-427">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-427">1.0</span></span>|
|[<span data-ttu-id="07e18-428">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-429">ReadItem</span></span>|
|[<span data-ttu-id="07e18-430">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-431">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="07e18-431">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-432">例</span><span class="sxs-lookup"><span data-stu-id="07e18-432">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-12"></a><span data-ttu-id="07e18-433">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-433">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

<span data-ttu-id="07e18-434">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="07e18-434">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07e18-435">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="07e18-435">Read mode</span></span>

<span data-ttu-id="07e18-436">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-436">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="07e18-437">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="07e18-437">Compose mode</span></span>

<span data-ttu-id="07e18-438">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-438">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="07e18-439">型</span><span class="sxs-lookup"><span data-stu-id="07e18-439">Type</span></span>

*   <span data-ttu-id="07e18-440">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-440">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-441">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-441">Requirements</span></span>

|<span data-ttu-id="07e18-442">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-442">Requirement</span></span>| <span data-ttu-id="07e18-443">値</span><span class="sxs-lookup"><span data-stu-id="07e18-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-444">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-445">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-445">1.0</span></span>|
|[<span data-ttu-id="07e18-446">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-447">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-447">ReadItem</span></span>|
|[<span data-ttu-id="07e18-448">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-449">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="07e18-449">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="07e18-450">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="07e18-450">normalizedSubject: String</span></span>

<span data-ttu-id="07e18-p123">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="07e18-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="07e18-p124">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="07e18-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="07e18-455">Type</span><span class="sxs-lookup"><span data-stu-id="07e18-455">Type</span></span>

*   <span data-ttu-id="07e18-456">String</span><span class="sxs-lookup"><span data-stu-id="07e18-456">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-457">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-457">Requirements</span></span>

|<span data-ttu-id="07e18-458">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-458">Requirement</span></span>| <span data-ttu-id="07e18-459">値</span><span class="sxs-lookup"><span data-stu-id="07e18-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-460">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-461">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-461">1.0</span></span>|
|[<span data-ttu-id="07e18-462">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-462">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-463">ReadItem</span></span>|
|[<span data-ttu-id="07e18-464">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-464">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-465">読み取り</span><span class="sxs-lookup"><span data-stu-id="07e18-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-466">例</span><span class="sxs-lookup"><span data-stu-id="07e18-466">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="07e18-467">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-467">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="07e18-468">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="07e18-468">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="07e18-469">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="07e18-469">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07e18-470">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="07e18-470">Read mode</span></span>

<span data-ttu-id="07e18-471">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-471">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="07e18-472">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="07e18-472">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="07e18-473">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="07e18-473">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="07e18-474">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="07e18-474">Compose mode</span></span>

<span data-ttu-id="07e18-475">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-475">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="07e18-476">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="07e18-476">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="07e18-477">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-477">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="07e18-478">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="07e18-478">Get 500 members maximum.</span></span>
- <span data-ttu-id="07e18-479">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="07e18-479">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="07e18-480">型</span><span class="sxs-lookup"><span data-stu-id="07e18-480">Type</span></span>

*   <span data-ttu-id="07e18-481">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-481">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-482">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-482">Requirements</span></span>

|<span data-ttu-id="07e18-483">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-483">Requirement</span></span>| <span data-ttu-id="07e18-484">値</span><span class="sxs-lookup"><span data-stu-id="07e18-484">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-485">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-485">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-486">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-486">1.0</span></span>|
|[<span data-ttu-id="07e18-487">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-487">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-488">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-488">ReadItem</span></span>|
|[<span data-ttu-id="07e18-489">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-489">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-490">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="07e18-490">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="07e18-491">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-491">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="07e18-p128">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="07e18-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="07e18-494">型</span><span class="sxs-lookup"><span data-stu-id="07e18-494">Type</span></span>

*   [<span data-ttu-id="07e18-495">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="07e18-495">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="07e18-496">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-496">Requirements</span></span>

|<span data-ttu-id="07e18-497">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-497">Requirement</span></span>| <span data-ttu-id="07e18-498">値</span><span class="sxs-lookup"><span data-stu-id="07e18-498">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-499">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-499">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-500">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-500">1.0</span></span>|
|[<span data-ttu-id="07e18-501">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-501">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-502">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-502">ReadItem</span></span>|
|[<span data-ttu-id="07e18-503">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-503">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-504">読み取り</span><span class="sxs-lookup"><span data-stu-id="07e18-504">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-505">例</span><span class="sxs-lookup"><span data-stu-id="07e18-505">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="07e18-506">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-506">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="07e18-507">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="07e18-507">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="07e18-508">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="07e18-508">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07e18-509">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="07e18-509">Read mode</span></span>

<span data-ttu-id="07e18-510">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-510">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="07e18-511">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="07e18-511">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="07e18-512">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="07e18-512">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="07e18-513">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="07e18-513">Compose mode</span></span>

<span data-ttu-id="07e18-514">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-514">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="07e18-515">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="07e18-515">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="07e18-516">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-516">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="07e18-517">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="07e18-517">Get 500 members maximum.</span></span>
- <span data-ttu-id="07e18-518">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="07e18-518">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="07e18-519">型</span><span class="sxs-lookup"><span data-stu-id="07e18-519">Type</span></span>

*   <span data-ttu-id="07e18-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-521">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-521">Requirements</span></span>

|<span data-ttu-id="07e18-522">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-522">Requirement</span></span>| <span data-ttu-id="07e18-523">値</span><span class="sxs-lookup"><span data-stu-id="07e18-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-524">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-525">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-525">1.0</span></span>|
|[<span data-ttu-id="07e18-526">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-527">ReadItem</span></span>|
|[<span data-ttu-id="07e18-528">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-529">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="07e18-529">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="07e18-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="07e18-p132">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="07e18-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="07e18-p133">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="07e18-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="07e18-535">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="07e18-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="07e18-536">型</span><span class="sxs-lookup"><span data-stu-id="07e18-536">Type</span></span>

*   [<span data-ttu-id="07e18-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="07e18-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="07e18-538">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-538">Requirements</span></span>

|<span data-ttu-id="07e18-539">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-539">Requirement</span></span>| <span data-ttu-id="07e18-540">値</span><span class="sxs-lookup"><span data-stu-id="07e18-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-541">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-542">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-542">1.0</span></span>|
|[<span data-ttu-id="07e18-543">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-544">ReadItem</span></span>|
|[<span data-ttu-id="07e18-545">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-546">読み取り</span><span class="sxs-lookup"><span data-stu-id="07e18-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-547">例</span><span class="sxs-lookup"><span data-stu-id="07e18-547">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="07e18-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="07e18-549">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="07e18-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="07e18-p134">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="07e18-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07e18-552">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="07e18-552">Read mode</span></span>

<span data-ttu-id="07e18-553">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-553">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="07e18-554">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="07e18-554">Compose mode</span></span>

<span data-ttu-id="07e18-555">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="07e18-556">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="07e18-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="07e18-557">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="07e18-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="07e18-558">型</span><span class="sxs-lookup"><span data-stu-id="07e18-558">Type</span></span>

*   <span data-ttu-id="07e18-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-560">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-560">Requirements</span></span>

|<span data-ttu-id="07e18-561">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-561">Requirement</span></span>| <span data-ttu-id="07e18-562">値</span><span class="sxs-lookup"><span data-stu-id="07e18-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-563">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-564">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-564">1.0</span></span>|
|[<span data-ttu-id="07e18-565">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-566">ReadItem</span></span>|
|[<span data-ttu-id="07e18-567">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-568">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="07e18-568">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-12"></a><span data-ttu-id="07e18-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

<span data-ttu-id="07e18-570">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="07e18-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="07e18-571">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="07e18-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07e18-572">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="07e18-572">Read mode</span></span>

<span data-ttu-id="07e18-p136">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="07e18-p136">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="07e18-575">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="07e18-575">Compose mode</span></span>

<span data-ttu-id="07e18-576">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="07e18-577">型</span><span class="sxs-lookup"><span data-stu-id="07e18-577">Type</span></span>

*   <span data-ttu-id="07e18-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-579">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-579">Requirements</span></span>

|<span data-ttu-id="07e18-580">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-580">Requirement</span></span>| <span data-ttu-id="07e18-581">値</span><span class="sxs-lookup"><span data-stu-id="07e18-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-582">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-583">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-583">1.0</span></span>|
|[<span data-ttu-id="07e18-584">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-585">ReadItem</span></span>|
|[<span data-ttu-id="07e18-586">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-587">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="07e18-587">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="07e18-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="07e18-589">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="07e18-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="07e18-590">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="07e18-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07e18-591">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="07e18-591">Read mode</span></span>

<span data-ttu-id="07e18-592">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-592">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="07e18-593">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="07e18-593">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="07e18-594">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="07e18-594">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="07e18-595">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="07e18-595">Compose mode</span></span>

<span data-ttu-id="07e18-596">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-596">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="07e18-597">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="07e18-597">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="07e18-598">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-598">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="07e18-599">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="07e18-599">Get 500 members maximum.</span></span>
- <span data-ttu-id="07e18-600">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="07e18-600">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="07e18-601">型</span><span class="sxs-lookup"><span data-stu-id="07e18-601">Type</span></span>

*   <span data-ttu-id="07e18-602">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-602">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-603">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-603">Requirements</span></span>

|<span data-ttu-id="07e18-604">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-604">Requirement</span></span>| <span data-ttu-id="07e18-605">値</span><span class="sxs-lookup"><span data-stu-id="07e18-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-606">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-607">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-607">1.0</span></span>|
|[<span data-ttu-id="07e18-608">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-609">ReadItem</span></span>|
|[<span data-ttu-id="07e18-610">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-611">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="07e18-611">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="07e18-612">メソッド</span><span class="sxs-lookup"><span data-stu-id="07e18-612">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="07e18-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="07e18-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="07e18-614">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="07e18-614">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="07e18-615">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="07e18-615">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="07e18-616">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="07e18-616">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07e18-617">パラメーター</span><span class="sxs-lookup"><span data-stu-id="07e18-617">Parameters</span></span>

|<span data-ttu-id="07e18-618">名前</span><span class="sxs-lookup"><span data-stu-id="07e18-618">Name</span></span>| <span data-ttu-id="07e18-619">種類</span><span class="sxs-lookup"><span data-stu-id="07e18-619">Type</span></span>| <span data-ttu-id="07e18-620">属性</span><span class="sxs-lookup"><span data-stu-id="07e18-620">Attributes</span></span>| <span data-ttu-id="07e18-621">説明</span><span class="sxs-lookup"><span data-stu-id="07e18-621">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="07e18-622">String</span><span class="sxs-lookup"><span data-stu-id="07e18-622">String</span></span>||<span data-ttu-id="07e18-p140">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="07e18-p140">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="07e18-625">String</span><span class="sxs-lookup"><span data-stu-id="07e18-625">String</span></span>||<span data-ttu-id="07e18-p141">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="07e18-p141">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="07e18-628">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="07e18-628">Object</span></span>| <span data-ttu-id="07e18-629">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-629">&lt;optional&gt;</span></span>|<span data-ttu-id="07e18-630">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="07e18-630">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="07e18-631">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="07e18-631">Object</span></span>| <span data-ttu-id="07e18-632">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-632">&lt;optional&gt;</span></span>|<span data-ttu-id="07e18-633">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="07e18-633">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="07e18-634">function</span><span class="sxs-lookup"><span data-stu-id="07e18-634">function</span></span>| <span data-ttu-id="07e18-635">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-635">&lt;optional&gt;</span></span>|<span data-ttu-id="07e18-636">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-636">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="07e18-637">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-637">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="07e18-638">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="07e18-638">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="07e18-639">エラー</span><span class="sxs-lookup"><span data-stu-id="07e18-639">Errors</span></span>

| <span data-ttu-id="07e18-640">エラー コード</span><span class="sxs-lookup"><span data-stu-id="07e18-640">Error code</span></span> | <span data-ttu-id="07e18-641">説明</span><span class="sxs-lookup"><span data-stu-id="07e18-641">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="07e18-642">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="07e18-642">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="07e18-643">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="07e18-643">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="07e18-644">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="07e18-644">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="07e18-645">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-645">Requirements</span></span>

|<span data-ttu-id="07e18-646">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-646">Requirement</span></span>| <span data-ttu-id="07e18-647">値</span><span class="sxs-lookup"><span data-stu-id="07e18-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-648">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-648">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-649">1.1</span><span class="sxs-lookup"><span data-stu-id="07e18-649">1.1</span></span>|
|[<span data-ttu-id="07e18-650">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-650">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-651">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07e18-651">ReadWriteItem</span></span>|
|[<span data-ttu-id="07e18-652">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-652">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-653">作成</span><span class="sxs-lookup"><span data-stu-id="07e18-653">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-654">例</span><span class="sxs-lookup"><span data-stu-id="07e18-654">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="07e18-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="07e18-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="07e18-656">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="07e18-656">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="07e18-p142">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="07e18-p142">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="07e18-660">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="07e18-660">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="07e18-661">Office アドインを Outlook on the web で実行している場合、編集中のアイテム以外のアイテムに `addItemAttachmentAsync` メソッドでアイテムを添付できます。ただし、これはサポートされていないため、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="07e18-661">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07e18-662">パラメーター</span><span class="sxs-lookup"><span data-stu-id="07e18-662">Parameters</span></span>

|<span data-ttu-id="07e18-663">名前</span><span class="sxs-lookup"><span data-stu-id="07e18-663">Name</span></span>| <span data-ttu-id="07e18-664">型</span><span class="sxs-lookup"><span data-stu-id="07e18-664">Type</span></span>| <span data-ttu-id="07e18-665">属性</span><span class="sxs-lookup"><span data-stu-id="07e18-665">Attributes</span></span>| <span data-ttu-id="07e18-666">説明</span><span class="sxs-lookup"><span data-stu-id="07e18-666">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="07e18-667">String</span><span class="sxs-lookup"><span data-stu-id="07e18-667">String</span></span>||<span data-ttu-id="07e18-p143">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="07e18-p143">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="07e18-670">String</span><span class="sxs-lookup"><span data-stu-id="07e18-670">String</span></span>||<span data-ttu-id="07e18-671">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="07e18-671">The subject of the item to be attached.</span></span> <span data-ttu-id="07e18-672">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="07e18-672">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="07e18-673">Object</span><span class="sxs-lookup"><span data-stu-id="07e18-673">Object</span></span>| <span data-ttu-id="07e18-674">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-674">&lt;optional&gt;</span></span>|<span data-ttu-id="07e18-675">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="07e18-675">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="07e18-676">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="07e18-676">Object</span></span>| <span data-ttu-id="07e18-677">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-677">&lt;optional&gt;</span></span>|<span data-ttu-id="07e18-678">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="07e18-678">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="07e18-679">関数</span><span class="sxs-lookup"><span data-stu-id="07e18-679">function</span></span>| <span data-ttu-id="07e18-680">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-680">&lt;optional&gt;</span></span>|<span data-ttu-id="07e18-681">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-681">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="07e18-682">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-682">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="07e18-683">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="07e18-683">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="07e18-684">エラー</span><span class="sxs-lookup"><span data-stu-id="07e18-684">Errors</span></span>

| <span data-ttu-id="07e18-685">エラー コード</span><span class="sxs-lookup"><span data-stu-id="07e18-685">Error code</span></span> | <span data-ttu-id="07e18-686">説明</span><span class="sxs-lookup"><span data-stu-id="07e18-686">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="07e18-687">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="07e18-687">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="07e18-688">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-688">Requirements</span></span>

|<span data-ttu-id="07e18-689">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-689">Requirement</span></span>| <span data-ttu-id="07e18-690">値</span><span class="sxs-lookup"><span data-stu-id="07e18-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-691">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-692">1.1</span><span class="sxs-lookup"><span data-stu-id="07e18-692">1.1</span></span>|
|[<span data-ttu-id="07e18-693">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-693">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-694">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07e18-694">ReadWriteItem</span></span>|
|[<span data-ttu-id="07e18-695">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-695">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-696">作成</span><span class="sxs-lookup"><span data-stu-id="07e18-696">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-697">例</span><span class="sxs-lookup"><span data-stu-id="07e18-697">Example</span></span>

<span data-ttu-id="07e18-698">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-698">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="07e18-699">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="07e18-699">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="07e18-700">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-700">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="07e18-701">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="07e18-701">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="07e18-702">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-702">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="07e18-703">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="07e18-703">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="07e18-p145">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="07e18-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07e18-707">パラメーター</span><span class="sxs-lookup"><span data-stu-id="07e18-707">Parameters</span></span>

|<span data-ttu-id="07e18-708">名前</span><span class="sxs-lookup"><span data-stu-id="07e18-708">Name</span></span>| <span data-ttu-id="07e18-709">種類</span><span class="sxs-lookup"><span data-stu-id="07e18-709">Type</span></span>| <span data-ttu-id="07e18-710">説明</span><span class="sxs-lookup"><span data-stu-id="07e18-710">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="07e18-711">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="07e18-711">String &#124; Object</span></span>| |<span data-ttu-id="07e18-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="07e18-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="07e18-714">**または**</span><span class="sxs-lookup"><span data-stu-id="07e18-714">**OR**</span></span><br/><span data-ttu-id="07e18-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="07e18-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="07e18-717">String</span><span class="sxs-lookup"><span data-stu-id="07e18-717">String</span></span> | <span data-ttu-id="07e18-718">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-718">&lt;optional&gt;</span></span> | <span data-ttu-id="07e18-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="07e18-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="07e18-721">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-721">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="07e18-722">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-722">&lt;optional&gt;</span></span> | <span data-ttu-id="07e18-723">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="07e18-723">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="07e18-724">String</span><span class="sxs-lookup"><span data-stu-id="07e18-724">String</span></span> | | <span data-ttu-id="07e18-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="07e18-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="07e18-727">String</span><span class="sxs-lookup"><span data-stu-id="07e18-727">String</span></span> | | <span data-ttu-id="07e18-728">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="07e18-728">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="07e18-729">文字列</span><span class="sxs-lookup"><span data-stu-id="07e18-729">String</span></span> | | <span data-ttu-id="07e18-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="07e18-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="07e18-732">String</span><span class="sxs-lookup"><span data-stu-id="07e18-732">String</span></span> | | <span data-ttu-id="07e18-p151">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="07e18-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="07e18-736">function</span><span class="sxs-lookup"><span data-stu-id="07e18-736">function</span></span> | <span data-ttu-id="07e18-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-737">&lt;optional&gt;</span></span> | <span data-ttu-id="07e18-738">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-738">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="07e18-739">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-739">Requirements</span></span>

|<span data-ttu-id="07e18-740">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-740">Requirement</span></span>| <span data-ttu-id="07e18-741">値</span><span class="sxs-lookup"><span data-stu-id="07e18-741">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-742">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-742">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-743">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-743">1.0</span></span>|
|[<span data-ttu-id="07e18-744">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-744">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-745">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-745">ReadItem</span></span>|
|[<span data-ttu-id="07e18-746">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-746">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-747">読み取り</span><span class="sxs-lookup"><span data-stu-id="07e18-747">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="07e18-748">例</span><span class="sxs-lookup"><span data-stu-id="07e18-748">Examples</span></span>

<span data-ttu-id="07e18-749">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="07e18-749">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="07e18-750">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="07e18-750">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="07e18-751">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="07e18-751">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="07e18-752">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="07e18-752">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="07e18-753">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="07e18-753">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="07e18-754">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="07e18-754">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="07e18-755">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="07e18-755">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="07e18-756">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-756">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="07e18-757">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="07e18-757">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="07e18-758">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-758">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="07e18-759">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="07e18-759">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="07e18-p152">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="07e18-p152">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07e18-763">パラメーター</span><span class="sxs-lookup"><span data-stu-id="07e18-763">Parameters</span></span>

|<span data-ttu-id="07e18-764">名前</span><span class="sxs-lookup"><span data-stu-id="07e18-764">Name</span></span>| <span data-ttu-id="07e18-765">種類</span><span class="sxs-lookup"><span data-stu-id="07e18-765">Type</span></span>| <span data-ttu-id="07e18-766">説明</span><span class="sxs-lookup"><span data-stu-id="07e18-766">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="07e18-767">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="07e18-767">String &#124; Object</span></span>| | <span data-ttu-id="07e18-p153">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="07e18-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="07e18-770">**または**</span><span class="sxs-lookup"><span data-stu-id="07e18-770">**OR**</span></span><br/><span data-ttu-id="07e18-p154">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="07e18-p154">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="07e18-773">String</span><span class="sxs-lookup"><span data-stu-id="07e18-773">String</span></span> | <span data-ttu-id="07e18-774">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-774">&lt;optional&gt;</span></span> | <span data-ttu-id="07e18-p155">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="07e18-p155">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="07e18-777">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-777">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="07e18-778">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-778">&lt;optional&gt;</span></span> | <span data-ttu-id="07e18-779">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="07e18-779">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="07e18-780">String</span><span class="sxs-lookup"><span data-stu-id="07e18-780">String</span></span> | | <span data-ttu-id="07e18-p156">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="07e18-p156">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="07e18-783">String</span><span class="sxs-lookup"><span data-stu-id="07e18-783">String</span></span> | | <span data-ttu-id="07e18-784">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="07e18-784">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="07e18-785">文字列</span><span class="sxs-lookup"><span data-stu-id="07e18-785">String</span></span> | | <span data-ttu-id="07e18-p157">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="07e18-p157">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="07e18-788">String</span><span class="sxs-lookup"><span data-stu-id="07e18-788">String</span></span> | | <span data-ttu-id="07e18-p158">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="07e18-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="07e18-792">function</span><span class="sxs-lookup"><span data-stu-id="07e18-792">function</span></span> | <span data-ttu-id="07e18-793">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-793">&lt;optional&gt;</span></span> | <span data-ttu-id="07e18-794">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-794">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="07e18-795">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-795">Requirements</span></span>

|<span data-ttu-id="07e18-796">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-796">Requirement</span></span>| <span data-ttu-id="07e18-797">値</span><span class="sxs-lookup"><span data-stu-id="07e18-797">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-798">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-798">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-799">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-799">1.0</span></span>|
|[<span data-ttu-id="07e18-800">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-800">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-801">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-801">ReadItem</span></span>|
|[<span data-ttu-id="07e18-802">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-802">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-803">読み取り</span><span class="sxs-lookup"><span data-stu-id="07e18-803">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="07e18-804">例</span><span class="sxs-lookup"><span data-stu-id="07e18-804">Examples</span></span>

<span data-ttu-id="07e18-805">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="07e18-805">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="07e18-806">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="07e18-806">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="07e18-807">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="07e18-807">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="07e18-808">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="07e18-808">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="07e18-809">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="07e18-809">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="07e18-810">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="07e18-810">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-12"></a><span data-ttu-id="07e18-811">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="07e18-811">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="07e18-812">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="07e18-812">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="07e18-813">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="07e18-813">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-814">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-814">Requirements</span></span>

|<span data-ttu-id="07e18-815">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-815">Requirement</span></span>| <span data-ttu-id="07e18-816">値</span><span class="sxs-lookup"><span data-stu-id="07e18-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-817">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-818">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-818">1.0</span></span>|
|[<span data-ttu-id="07e18-819">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-820">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-820">ReadItem</span></span>|
|[<span data-ttu-id="07e18-821">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-822">読み取り</span><span class="sxs-lookup"><span data-stu-id="07e18-822">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="07e18-823">戻り値:</span><span class="sxs-lookup"><span data-stu-id="07e18-823">Returns:</span></span>

<span data-ttu-id="07e18-824">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="07e18-824">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span></span>

##### <a name="example"></a><span data-ttu-id="07e18-825">例</span><span class="sxs-lookup"><span data-stu-id="07e18-825">Example</span></span>

<span data-ttu-id="07e18-826">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="07e18-826">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="07e18-827">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="07e18-827">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="07e18-828">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="07e18-828">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="07e18-829">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="07e18-829">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07e18-830">パラメーター</span><span class="sxs-lookup"><span data-stu-id="07e18-830">Parameters</span></span>

|<span data-ttu-id="07e18-831">名前</span><span class="sxs-lookup"><span data-stu-id="07e18-831">Name</span></span>| <span data-ttu-id="07e18-832">型</span><span class="sxs-lookup"><span data-stu-id="07e18-832">Type</span></span>| <span data-ttu-id="07e18-833">説明</span><span class="sxs-lookup"><span data-stu-id="07e18-833">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="07e18-834">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="07e18-834">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.2)|<span data-ttu-id="07e18-835">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="07e18-835">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07e18-836">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-836">Requirements</span></span>

|<span data-ttu-id="07e18-837">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-837">Requirement</span></span>| <span data-ttu-id="07e18-838">値</span><span class="sxs-lookup"><span data-stu-id="07e18-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-839">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-840">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-840">1.0</span></span>|
|[<span data-ttu-id="07e18-841">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-841">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-842">制限あり</span><span class="sxs-lookup"><span data-stu-id="07e18-842">Restricted</span></span>|
|[<span data-ttu-id="07e18-843">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-843">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-844">読み取り</span><span class="sxs-lookup"><span data-stu-id="07e18-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="07e18-845">戻り値:</span><span class="sxs-lookup"><span data-stu-id="07e18-845">Returns:</span></span>

<span data-ttu-id="07e18-846">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-846">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="07e18-847">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-847">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="07e18-848">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="07e18-848">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="07e18-849">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="07e18-849">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="07e18-850">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="07e18-850">Value of `entityType`</span></span> | <span data-ttu-id="07e18-851">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="07e18-851">Type of objects in returned array</span></span> | <span data-ttu-id="07e18-852">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="07e18-852">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="07e18-853">文字列</span><span class="sxs-lookup"><span data-stu-id="07e18-853">String</span></span> | <span data-ttu-id="07e18-854">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="07e18-854">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="07e18-855">連絡先</span><span class="sxs-lookup"><span data-stu-id="07e18-855">Contact</span></span> | <span data-ttu-id="07e18-856">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="07e18-856">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="07e18-857">文字列</span><span class="sxs-lookup"><span data-stu-id="07e18-857">String</span></span> | <span data-ttu-id="07e18-858">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="07e18-858">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="07e18-859">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="07e18-859">MeetingSuggestion</span></span> | <span data-ttu-id="07e18-860">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="07e18-860">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="07e18-861">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="07e18-861">PhoneNumber</span></span> | <span data-ttu-id="07e18-862">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="07e18-862">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="07e18-863">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="07e18-863">TaskSuggestion</span></span> | <span data-ttu-id="07e18-864">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="07e18-864">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="07e18-865">文字列</span><span class="sxs-lookup"><span data-stu-id="07e18-865">String</span></span> | <span data-ttu-id="07e18-866">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="07e18-866">**Restricted**</span></span> |

<span data-ttu-id="07e18-867">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="07e18-867">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

##### <a name="example"></a><span data-ttu-id="07e18-868">例</span><span class="sxs-lookup"><span data-stu-id="07e18-868">Example</span></span>

<span data-ttu-id="07e18-869">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="07e18-869">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="07e18-870">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="07e18-870">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="07e18-871">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-871">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="07e18-872">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="07e18-872">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="07e18-873">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-873">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07e18-874">パラメーター</span><span class="sxs-lookup"><span data-stu-id="07e18-874">Parameters</span></span>

|<span data-ttu-id="07e18-875">名前</span><span class="sxs-lookup"><span data-stu-id="07e18-875">Name</span></span>| <span data-ttu-id="07e18-876">型</span><span class="sxs-lookup"><span data-stu-id="07e18-876">Type</span></span>| <span data-ttu-id="07e18-877">説明</span><span class="sxs-lookup"><span data-stu-id="07e18-877">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="07e18-878">String</span><span class="sxs-lookup"><span data-stu-id="07e18-878">String</span></span>|<span data-ttu-id="07e18-879">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="07e18-879">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07e18-880">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-880">Requirements</span></span>

|<span data-ttu-id="07e18-881">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-881">Requirement</span></span>| <span data-ttu-id="07e18-882">値</span><span class="sxs-lookup"><span data-stu-id="07e18-882">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-883">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-883">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-884">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-884">1.0</span></span>|
|[<span data-ttu-id="07e18-885">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-885">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-886">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-886">ReadItem</span></span>|
|[<span data-ttu-id="07e18-887">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-887">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-888">読み取り</span><span class="sxs-lookup"><span data-stu-id="07e18-888">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="07e18-889">戻り値:</span><span class="sxs-lookup"><span data-stu-id="07e18-889">Returns:</span></span>

<span data-ttu-id="07e18-p160">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="07e18-892">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="07e18-892">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="07e18-893">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="07e18-893">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="07e18-894">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-894">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="07e18-895">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="07e18-895">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="07e18-p161">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="07e18-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="07e18-899">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="07e18-899">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="07e18-900">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="07e18-900">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="07e18-p162">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="07e18-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="07e18-903">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-903">Requirements</span></span>

|<span data-ttu-id="07e18-904">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-904">Requirement</span></span>| <span data-ttu-id="07e18-905">値</span><span class="sxs-lookup"><span data-stu-id="07e18-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-906">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-907">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-907">1.0</span></span>|
|[<span data-ttu-id="07e18-908">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-908">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-909">ReadItem</span></span>|
|[<span data-ttu-id="07e18-910">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-910">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-911">読み取り</span><span class="sxs-lookup"><span data-stu-id="07e18-911">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="07e18-912">戻り値:</span><span class="sxs-lookup"><span data-stu-id="07e18-912">Returns:</span></span>

<span data-ttu-id="07e18-p163">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="07e18-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="07e18-915">型: Object</span><span class="sxs-lookup"><span data-stu-id="07e18-915">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="07e18-916">例</span><span class="sxs-lookup"><span data-stu-id="07e18-916">Example</span></span>

<span data-ttu-id="07e18-917">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="07e18-917">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="07e18-918">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="07e18-918">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="07e18-919">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-919">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="07e18-920">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="07e18-920">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="07e18-921">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="07e18-921">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="07e18-p164">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="07e18-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07e18-924">パラメーター</span><span class="sxs-lookup"><span data-stu-id="07e18-924">Parameters</span></span>

|<span data-ttu-id="07e18-925">名前</span><span class="sxs-lookup"><span data-stu-id="07e18-925">Name</span></span>| <span data-ttu-id="07e18-926">種類</span><span class="sxs-lookup"><span data-stu-id="07e18-926">Type</span></span>| <span data-ttu-id="07e18-927">説明</span><span class="sxs-lookup"><span data-stu-id="07e18-927">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="07e18-928">String</span><span class="sxs-lookup"><span data-stu-id="07e18-928">String</span></span>|<span data-ttu-id="07e18-929">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="07e18-929">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07e18-930">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-930">Requirements</span></span>

|<span data-ttu-id="07e18-931">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-931">Requirement</span></span>| <span data-ttu-id="07e18-932">値</span><span class="sxs-lookup"><span data-stu-id="07e18-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-933">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-934">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-934">1.0</span></span>|
|[<span data-ttu-id="07e18-935">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-936">ReadItem</span></span>|
|[<span data-ttu-id="07e18-937">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-938">読み取り</span><span class="sxs-lookup"><span data-stu-id="07e18-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="07e18-939">戻り値:</span><span class="sxs-lookup"><span data-stu-id="07e18-939">Returns:</span></span>

<span data-ttu-id="07e18-940">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="07e18-940">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="07e18-941">型: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="07e18-941">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="07e18-942">例</span><span class="sxs-lookup"><span data-stu-id="07e18-942">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="07e18-943">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="07e18-943">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="07e18-944">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-944">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="07e18-p165">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="07e18-947">Web 上の Outlook では、テキストが選択されておらず、カーソルが本文にある場合、このメソッドは文字列 "null" を返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-947">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="07e18-948">このような状況を確認するには、次のようなコードを含めます。</span><span class="sxs-lookup"><span data-stu-id="07e18-948">To check for this situation, include code similar to the following:</span></span>
>
> `var selectedText = (asyncResult.value.endPosition === asyncResult.value.startPosition) ? "" : asyncResult.value.data;`

##### <a name="parameters"></a><span data-ttu-id="07e18-949">パラメーター</span><span class="sxs-lookup"><span data-stu-id="07e18-949">Parameters</span></span>

|<span data-ttu-id="07e18-950">名前</span><span class="sxs-lookup"><span data-stu-id="07e18-950">Name</span></span>| <span data-ttu-id="07e18-951">型</span><span class="sxs-lookup"><span data-stu-id="07e18-951">Type</span></span>| <span data-ttu-id="07e18-952">属性</span><span class="sxs-lookup"><span data-stu-id="07e18-952">Attributes</span></span>| <span data-ttu-id="07e18-953">説明</span><span class="sxs-lookup"><span data-stu-id="07e18-953">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="07e18-954">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="07e18-954">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="07e18-p167">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="07e18-p167">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="07e18-958">Object</span><span class="sxs-lookup"><span data-stu-id="07e18-958">Object</span></span>| <span data-ttu-id="07e18-959">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-959">&lt;optional&gt;</span></span>|<span data-ttu-id="07e18-960">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="07e18-960">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="07e18-961">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="07e18-961">Object</span></span>| <span data-ttu-id="07e18-962">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-962">&lt;optional&gt;</span></span>|<span data-ttu-id="07e18-963">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="07e18-963">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="07e18-964">function</span><span class="sxs-lookup"><span data-stu-id="07e18-964">function</span></span>||<span data-ttu-id="07e18-965">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-965">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="07e18-966">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="07e18-966">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="07e18-967">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="07e18-967">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07e18-968">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-968">Requirements</span></span>

|<span data-ttu-id="07e18-969">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-969">Requirement</span></span>| <span data-ttu-id="07e18-970">値</span><span class="sxs-lookup"><span data-stu-id="07e18-970">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-971">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-971">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-972">1.2</span><span class="sxs-lookup"><span data-stu-id="07e18-972">1.2</span></span>|
|[<span data-ttu-id="07e18-973">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-973">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-974">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-974">ReadItem</span></span>|
|[<span data-ttu-id="07e18-975">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-975">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-976">作成</span><span class="sxs-lookup"><span data-stu-id="07e18-976">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="07e18-977">戻り値:</span><span class="sxs-lookup"><span data-stu-id="07e18-977">Returns:</span></span>

<span data-ttu-id="07e18-978">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="07e18-978">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="07e18-979">型:String</span><span class="sxs-lookup"><span data-stu-id="07e18-979">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="07e18-980">例</span><span class="sxs-lookup"><span data-stu-id="07e18-980">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
  // Check for errors.
}
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="07e18-981">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="07e18-981">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="07e18-982">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="07e18-982">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="07e18-p169">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="07e18-p169">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07e18-986">パラメーター</span><span class="sxs-lookup"><span data-stu-id="07e18-986">Parameters</span></span>

|<span data-ttu-id="07e18-987">名前</span><span class="sxs-lookup"><span data-stu-id="07e18-987">Name</span></span>| <span data-ttu-id="07e18-988">型</span><span class="sxs-lookup"><span data-stu-id="07e18-988">Type</span></span>| <span data-ttu-id="07e18-989">属性</span><span class="sxs-lookup"><span data-stu-id="07e18-989">Attributes</span></span>| <span data-ttu-id="07e18-990">説明</span><span class="sxs-lookup"><span data-stu-id="07e18-990">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="07e18-991">function</span><span class="sxs-lookup"><span data-stu-id="07e18-991">function</span></span>||<span data-ttu-id="07e18-992">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="07e18-993">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-993">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="07e18-994">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="07e18-994">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="07e18-995">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="07e18-995">Object</span></span>| <span data-ttu-id="07e18-996">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-996">&lt;optional&gt;</span></span>|<span data-ttu-id="07e18-997">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="07e18-997">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="07e18-998">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="07e18-998">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07e18-999">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-999">Requirements</span></span>

|<span data-ttu-id="07e18-1000">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-1000">Requirement</span></span>| <span data-ttu-id="07e18-1001">値</span><span class="sxs-lookup"><span data-stu-id="07e18-1001">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-1002">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-1002">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-1003">1.0</span><span class="sxs-lookup"><span data-stu-id="07e18-1003">1.0</span></span>|
|[<span data-ttu-id="07e18-1004">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-1004">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-1005">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07e18-1005">ReadItem</span></span>|
|[<span data-ttu-id="07e18-1006">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-1006">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-1007">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="07e18-1007">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-1008">例</span><span class="sxs-lookup"><span data-stu-id="07e18-1008">Example</span></span>

<span data-ttu-id="07e18-p172">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="07e18-p172">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="07e18-1012">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="07e18-1012">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="07e18-1013">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="07e18-1013">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="07e18-1014">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="07e18-1014">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="07e18-1015">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="07e18-1015">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="07e18-1016">Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="07e18-1016">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="07e18-1017">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="07e18-1017">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07e18-1018">パラメーター</span><span class="sxs-lookup"><span data-stu-id="07e18-1018">Parameters</span></span>

|<span data-ttu-id="07e18-1019">名前</span><span class="sxs-lookup"><span data-stu-id="07e18-1019">Name</span></span>| <span data-ttu-id="07e18-1020">型</span><span class="sxs-lookup"><span data-stu-id="07e18-1020">Type</span></span>| <span data-ttu-id="07e18-1021">属性</span><span class="sxs-lookup"><span data-stu-id="07e18-1021">Attributes</span></span>| <span data-ttu-id="07e18-1022">説明</span><span class="sxs-lookup"><span data-stu-id="07e18-1022">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="07e18-1023">String</span><span class="sxs-lookup"><span data-stu-id="07e18-1023">String</span></span>||<span data-ttu-id="07e18-1024">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="07e18-1024">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="07e18-1025">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="07e18-1025">Object</span></span>| <span data-ttu-id="07e18-1026">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-1026">&lt;optional&gt;</span></span>|<span data-ttu-id="07e18-1027">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="07e18-1027">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="07e18-1028">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="07e18-1028">Object</span></span>| <span data-ttu-id="07e18-1029">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-1029">&lt;optional&gt;</span></span>|<span data-ttu-id="07e18-1030">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="07e18-1030">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="07e18-1031">関数</span><span class="sxs-lookup"><span data-stu-id="07e18-1031">function</span></span>| <span data-ttu-id="07e18-1032">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-1032">&lt;optional&gt;</span></span>|<span data-ttu-id="07e18-1033">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-1033">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="07e18-1034">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="07e18-1034">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="07e18-1035">エラー</span><span class="sxs-lookup"><span data-stu-id="07e18-1035">Errors</span></span>

| <span data-ttu-id="07e18-1036">エラー コード</span><span class="sxs-lookup"><span data-stu-id="07e18-1036">Error code</span></span> | <span data-ttu-id="07e18-1037">説明</span><span class="sxs-lookup"><span data-stu-id="07e18-1037">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="07e18-1038">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="07e18-1038">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="07e18-1039">Requirements</span><span class="sxs-lookup"><span data-stu-id="07e18-1039">Requirements</span></span>

|<span data-ttu-id="07e18-1040">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-1040">Requirement</span></span>| <span data-ttu-id="07e18-1041">値</span><span class="sxs-lookup"><span data-stu-id="07e18-1041">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-1042">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-1042">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-1043">1.1</span><span class="sxs-lookup"><span data-stu-id="07e18-1043">1.1</span></span>|
|[<span data-ttu-id="07e18-1044">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-1044">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-1045">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07e18-1045">ReadWriteItem</span></span>|
|[<span data-ttu-id="07e18-1046">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-1046">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-1047">作成</span><span class="sxs-lookup"><span data-stu-id="07e18-1047">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-1048">例</span><span class="sxs-lookup"><span data-stu-id="07e18-1048">Example</span></span>

<span data-ttu-id="07e18-1049">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="07e18-1049">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="07e18-1050">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="07e18-1050">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="07e18-1051">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="07e18-1051">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="07e18-p174">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="07e18-p174">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07e18-1055">パラメーター</span><span class="sxs-lookup"><span data-stu-id="07e18-1055">Parameters</span></span>

|<span data-ttu-id="07e18-1056">名前</span><span class="sxs-lookup"><span data-stu-id="07e18-1056">Name</span></span>| <span data-ttu-id="07e18-1057">型</span><span class="sxs-lookup"><span data-stu-id="07e18-1057">Type</span></span>| <span data-ttu-id="07e18-1058">属性</span><span class="sxs-lookup"><span data-stu-id="07e18-1058">Attributes</span></span>| <span data-ttu-id="07e18-1059">説明</span><span class="sxs-lookup"><span data-stu-id="07e18-1059">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="07e18-1060">String</span><span class="sxs-lookup"><span data-stu-id="07e18-1060">String</span></span>||<span data-ttu-id="07e18-p175">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="07e18-p175">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="07e18-1064">Object</span><span class="sxs-lookup"><span data-stu-id="07e18-1064">Object</span></span>| <span data-ttu-id="07e18-1065">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-1065">&lt;optional&gt;</span></span>|<span data-ttu-id="07e18-1066">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="07e18-1066">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="07e18-1067">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="07e18-1067">Object</span></span>| <span data-ttu-id="07e18-1068">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-1068">&lt;optional&gt;</span></span>|<span data-ttu-id="07e18-1069">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="07e18-1069">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="07e18-1070">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="07e18-1070">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="07e18-1071">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07e18-1071">&lt;optional&gt;</span></span>|<span data-ttu-id="07e18-1072">`text` の場合、Outlook on the web とデスクトップ クライアントでは現在のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-1072">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="07e18-1073">フィールドが HTML エディターの場合、データが HTML の場合でもテキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-1073">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="07e18-1074">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Outlook on the web では現在のスタイルが適用され、Outlook デスクトップ クライアントでは既定のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-1074">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="07e18-1075">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-1075">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="07e18-1076">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="07e18-1076">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="07e18-1077">function</span><span class="sxs-lookup"><span data-stu-id="07e18-1077">function</span></span>||<span data-ttu-id="07e18-1078">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="07e18-1078">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="07e18-1079">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-1079">Requirements</span></span>

|<span data-ttu-id="07e18-1080">要件</span><span class="sxs-lookup"><span data-stu-id="07e18-1080">Requirement</span></span>| <span data-ttu-id="07e18-1081">値</span><span class="sxs-lookup"><span data-stu-id="07e18-1081">Value</span></span>|
|---|---|
|[<span data-ttu-id="07e18-1082">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="07e18-1082">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07e18-1083">1.2</span><span class="sxs-lookup"><span data-stu-id="07e18-1083">1.2</span></span>|
|[<span data-ttu-id="07e18-1084">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="07e18-1084">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07e18-1085">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07e18-1085">ReadWriteItem</span></span>|
|[<span data-ttu-id="07e18-1086">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="07e18-1086">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07e18-1087">作成</span><span class="sxs-lookup"><span data-stu-id="07e18-1087">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="07e18-1088">例</span><span class="sxs-lookup"><span data-stu-id="07e18-1088">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
