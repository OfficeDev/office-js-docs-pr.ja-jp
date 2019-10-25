---
title: Office. メールボックス-要件セット1.3
description: ''
ms.date: 10/23/2019
localization_priority: Normal
ms.openlocfilehash: e2e91dc196e0c67eed3a358e9f0d864885a01945
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/24/2019
ms.locfileid: "37682649"
---
# <a name="item"></a><span data-ttu-id="87d5c-102">item</span><span class="sxs-lookup"><span data-stu-id="87d5c-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="87d5c-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="87d5c-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="87d5c-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-106">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-106">Requirements</span></span>

|<span data-ttu-id="87d5c-107">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-107">Requirement</span></span>| <span data-ttu-id="87d5c-108">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-110">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-110">1.0</span></span>|
|[<span data-ttu-id="87d5c-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="87d5c-112">Restricted</span></span>|
|[<span data-ttu-id="87d5c-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="87d5c-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="87d5c-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="87d5c-115">Members and methods</span></span>

| <span data-ttu-id="87d5c-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-116">Member</span></span> | <span data-ttu-id="87d5c-117">種類</span><span class="sxs-lookup"><span data-stu-id="87d5c-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="87d5c-118">attachments</span><span class="sxs-lookup"><span data-stu-id="87d5c-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="87d5c-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-119">Member</span></span> |
| [<span data-ttu-id="87d5c-120">bcc</span><span class="sxs-lookup"><span data-stu-id="87d5c-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="87d5c-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-121">Member</span></span> |
| [<span data-ttu-id="87d5c-122">body</span><span class="sxs-lookup"><span data-stu-id="87d5c-122">body</span></span>](#body-body) | <span data-ttu-id="87d5c-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-123">Member</span></span> |
| [<span data-ttu-id="87d5c-124">cc</span><span class="sxs-lookup"><span data-stu-id="87d5c-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="87d5c-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-125">Member</span></span> |
| [<span data-ttu-id="87d5c-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="87d5c-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="87d5c-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-127">Member</span></span> |
| [<span data-ttu-id="87d5c-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="87d5c-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="87d5c-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-129">Member</span></span> |
| [<span data-ttu-id="87d5c-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="87d5c-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="87d5c-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-131">Member</span></span> |
| [<span data-ttu-id="87d5c-132">end</span><span class="sxs-lookup"><span data-stu-id="87d5c-132">end</span></span>](#end-datetime) | <span data-ttu-id="87d5c-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-133">Member</span></span> |
| [<span data-ttu-id="87d5c-134">from</span><span class="sxs-lookup"><span data-stu-id="87d5c-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="87d5c-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-135">Member</span></span> |
| [<span data-ttu-id="87d5c-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="87d5c-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="87d5c-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-137">Member</span></span> |
| [<span data-ttu-id="87d5c-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="87d5c-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="87d5c-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-139">Member</span></span> |
| [<span data-ttu-id="87d5c-140">itemId</span><span class="sxs-lookup"><span data-stu-id="87d5c-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="87d5c-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-141">Member</span></span> |
| [<span data-ttu-id="87d5c-142">itemType</span><span class="sxs-lookup"><span data-stu-id="87d5c-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="87d5c-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-143">Member</span></span> |
| [<span data-ttu-id="87d5c-144">location</span><span class="sxs-lookup"><span data-stu-id="87d5c-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="87d5c-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-145">Member</span></span> |
| [<span data-ttu-id="87d5c-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="87d5c-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="87d5c-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-147">Member</span></span> |
| [<span data-ttu-id="87d5c-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="87d5c-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="87d5c-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-149">Member</span></span> |
| [<span data-ttu-id="87d5c-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="87d5c-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="87d5c-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-151">Member</span></span> |
| [<span data-ttu-id="87d5c-152">organizer</span><span class="sxs-lookup"><span data-stu-id="87d5c-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="87d5c-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-153">Member</span></span> |
| [<span data-ttu-id="87d5c-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="87d5c-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="87d5c-155">Member</span><span class="sxs-lookup"><span data-stu-id="87d5c-155">Member</span></span> |
| [<span data-ttu-id="87d5c-156">sender</span><span class="sxs-lookup"><span data-stu-id="87d5c-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="87d5c-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-157">Member</span></span> |
| [<span data-ttu-id="87d5c-158">start</span><span class="sxs-lookup"><span data-stu-id="87d5c-158">start</span></span>](#start-datetime) | <span data-ttu-id="87d5c-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-159">Member</span></span> |
| [<span data-ttu-id="87d5c-160">subject</span><span class="sxs-lookup"><span data-stu-id="87d5c-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="87d5c-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-161">Member</span></span> |
| [<span data-ttu-id="87d5c-162">to</span><span class="sxs-lookup"><span data-stu-id="87d5c-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="87d5c-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="87d5c-163">Member</span></span> |
| [<span data-ttu-id="87d5c-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="87d5c-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="87d5c-165">メソッド</span><span class="sxs-lookup"><span data-stu-id="87d5c-165">Method</span></span> |
| [<span data-ttu-id="87d5c-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="87d5c-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="87d5c-167">メソッド</span><span class="sxs-lookup"><span data-stu-id="87d5c-167">Method</span></span> |
| [<span data-ttu-id="87d5c-168">close</span><span class="sxs-lookup"><span data-stu-id="87d5c-168">close</span></span>](#close) | <span data-ttu-id="87d5c-169">メソッド</span><span class="sxs-lookup"><span data-stu-id="87d5c-169">Method</span></span> |
| [<span data-ttu-id="87d5c-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="87d5c-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="87d5c-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="87d5c-171">Method</span></span> |
| [<span data-ttu-id="87d5c-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="87d5c-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="87d5c-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="87d5c-173">Method</span></span> |
| [<span data-ttu-id="87d5c-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="87d5c-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="87d5c-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="87d5c-175">Method</span></span> |
| [<span data-ttu-id="87d5c-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="87d5c-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="87d5c-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="87d5c-177">Method</span></span> |
| [<span data-ttu-id="87d5c-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="87d5c-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="87d5c-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="87d5c-179">Method</span></span> |
| [<span data-ttu-id="87d5c-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="87d5c-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="87d5c-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="87d5c-181">Method</span></span> |
| [<span data-ttu-id="87d5c-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="87d5c-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="87d5c-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="87d5c-183">Method</span></span> |
| [<span data-ttu-id="87d5c-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="87d5c-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="87d5c-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="87d5c-185">Method</span></span> |
| [<span data-ttu-id="87d5c-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="87d5c-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="87d5c-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="87d5c-187">Method</span></span> |
| [<span data-ttu-id="87d5c-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="87d5c-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="87d5c-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="87d5c-189">Method</span></span> |
| [<span data-ttu-id="87d5c-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="87d5c-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="87d5c-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="87d5c-191">Method</span></span> |
| [<span data-ttu-id="87d5c-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="87d5c-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="87d5c-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="87d5c-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="87d5c-194">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-194">Example</span></span>

<span data-ttu-id="87d5c-195">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="87d5c-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="87d5c-196">Members</span><span class="sxs-lookup"><span data-stu-id="87d5c-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-13"></a><span data-ttu-id="87d5c-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span><span class="sxs-lookup"><span data-stu-id="87d5c-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span></span>

<span data-ttu-id="87d5c-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="87d5c-200">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="87d5c-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="87d5c-201">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="87d5c-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="87d5c-202">型</span><span class="sxs-lookup"><span data-stu-id="87d5c-202">Type</span></span>

*   <span data-ttu-id="87d5c-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span><span class="sxs-lookup"><span data-stu-id="87d5c-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-204">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-204">Requirements</span></span>

|<span data-ttu-id="87d5c-205">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-205">Requirement</span></span>| <span data-ttu-id="87d5c-206">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-207">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-208">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-208">1.0</span></span>|
|[<span data-ttu-id="87d5c-209">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-210">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-211">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-212">読み取り</span><span class="sxs-lookup"><span data-stu-id="87d5c-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-213">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-213">Example</span></span>

<span data-ttu-id="87d5c-214">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="87d5c-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="87d5c-216">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="87d5c-217">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="87d5c-217">Compose mode only.</span></span>

<span data-ttu-id="87d5c-218">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="87d5c-218">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="87d5c-219">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-219">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="87d5c-220">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-220">Get 500 members maximum.</span></span>
- <span data-ttu-id="87d5c-221">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="87d5c-221">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="87d5c-222">Type</span><span class="sxs-lookup"><span data-stu-id="87d5c-222">Type</span></span>

*   [<span data-ttu-id="87d5c-223">受信者</span><span class="sxs-lookup"><span data-stu-id="87d5c-223">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="87d5c-224">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-224">Requirements</span></span>

|<span data-ttu-id="87d5c-225">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-225">Requirement</span></span>| <span data-ttu-id="87d5c-226">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-227">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-228">1.1</span><span class="sxs-lookup"><span data-stu-id="87d5c-228">1.1</span></span>|
|[<span data-ttu-id="87d5c-229">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-230">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-231">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-232">作成</span><span class="sxs-lookup"><span data-stu-id="87d5c-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-233">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-233">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-13"></a><span data-ttu-id="87d5c-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3)</span></span>

<span data-ttu-id="87d5c-235">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="87d5c-236">Type</span><span class="sxs-lookup"><span data-stu-id="87d5c-236">Type</span></span>

*   [<span data-ttu-id="87d5c-237">Body</span><span class="sxs-lookup"><span data-stu-id="87d5c-237">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="87d5c-238">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-238">Requirements</span></span>

|<span data-ttu-id="87d5c-239">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-239">Requirement</span></span>| <span data-ttu-id="87d5c-240">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-241">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-242">1.1</span><span class="sxs-lookup"><span data-stu-id="87d5c-242">1.1</span></span>|
|[<span data-ttu-id="87d5c-243">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-244">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-246">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="87d5c-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-247">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-247">Example</span></span>

<span data-ttu-id="87d5c-248">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-248">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="87d5c-249">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="87d5c-249">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="87d5c-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="87d5c-251">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="87d5c-252">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="87d5c-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="87d5c-253">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="87d5c-253">Read mode</span></span>

<span data-ttu-id="87d5c-254">`cc` プロパティは、メッセージの `EmailAddressDetails` 行にある各受信者について、\*\*\*\* オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-254">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="87d5c-255">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="87d5c-255">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="87d5c-256">ただし、Windows および Mac では、最大500メンバーを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-256">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="87d5c-257">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="87d5c-257">Compose mode</span></span>

<span data-ttu-id="87d5c-258">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-258">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="87d5c-259">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="87d5c-259">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="87d5c-260">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-260">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="87d5c-261">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-261">Get 500 members maximum.</span></span>
- <span data-ttu-id="87d5c-262">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="87d5c-262">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

<br>

---
---

##### <a name="type"></a><span data-ttu-id="87d5c-263">Type</span><span class="sxs-lookup"><span data-stu-id="87d5c-263">Type</span></span>

*   <span data-ttu-id="87d5c-264">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-264">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-265">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-265">Requirements</span></span>

|<span data-ttu-id="87d5c-266">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-266">Requirement</span></span>| <span data-ttu-id="87d5c-267">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-268">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-269">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-269">1.0</span></span>|
|[<span data-ttu-id="87d5c-270">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-270">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-271">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-272">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-272">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-273">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="87d5c-273">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="87d5c-274">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="87d5c-274">(nullable) conversationId: String</span></span>

<span data-ttu-id="87d5c-275">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-275">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="87d5c-p109">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="87d5c-p110">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="87d5c-280">Type</span><span class="sxs-lookup"><span data-stu-id="87d5c-280">Type</span></span>

*   <span data-ttu-id="87d5c-281">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-281">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-282">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-282">Requirements</span></span>

|<span data-ttu-id="87d5c-283">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-283">Requirement</span></span>| <span data-ttu-id="87d5c-284">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-285">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-285">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-286">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-286">1.0</span></span>|
|[<span data-ttu-id="87d5c-287">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-287">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-288">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-289">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-289">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-290">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="87d5c-290">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-291">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-291">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="87d5c-292">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="87d5c-292">dateTimeCreated: Date</span></span>

<span data-ttu-id="87d5c-p111">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="87d5c-295">型</span><span class="sxs-lookup"><span data-stu-id="87d5c-295">Type</span></span>

*   <span data-ttu-id="87d5c-296">日付</span><span class="sxs-lookup"><span data-stu-id="87d5c-296">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-297">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-297">Requirements</span></span>

|<span data-ttu-id="87d5c-298">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-298">Requirement</span></span>| <span data-ttu-id="87d5c-299">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-300">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-300">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-301">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-301">1.0</span></span>|
|[<span data-ttu-id="87d5c-302">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-302">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-303">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-303">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-304">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-304">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-305">読み取り</span><span class="sxs-lookup"><span data-stu-id="87d5c-305">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-306">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-306">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="87d5c-307">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="87d5c-307">dateTimeModified: Date</span></span>

<span data-ttu-id="87d5c-p112">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="87d5c-310">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="87d5c-310">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="87d5c-311">種類</span><span class="sxs-lookup"><span data-stu-id="87d5c-311">Type</span></span>

*   <span data-ttu-id="87d5c-312">日付</span><span class="sxs-lookup"><span data-stu-id="87d5c-312">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-313">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-313">Requirements</span></span>

|<span data-ttu-id="87d5c-314">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-314">Requirement</span></span>| <span data-ttu-id="87d5c-315">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-316">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-317">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-317">1.0</span></span>|
|[<span data-ttu-id="87d5c-318">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-319">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-320">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-321">読み取り</span><span class="sxs-lookup"><span data-stu-id="87d5c-321">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-322">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-322">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-13"></a><span data-ttu-id="87d5c-323">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-323">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

<span data-ttu-id="87d5c-324">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-324">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="87d5c-p113">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="87d5c-327">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="87d5c-327">Read mode</span></span>

<span data-ttu-id="87d5c-328">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-328">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="87d5c-329">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="87d5c-329">Compose mode</span></span>

<span data-ttu-id="87d5c-330">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-330">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="87d5c-331">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="87d5c-331">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="87d5c-332">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-332">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="87d5c-333">Type</span><span class="sxs-lookup"><span data-stu-id="87d5c-333">Type</span></span>

*   <span data-ttu-id="87d5c-334">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-334">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-335">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-335">Requirements</span></span>

|<span data-ttu-id="87d5c-336">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-336">Requirement</span></span>| <span data-ttu-id="87d5c-337">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-338">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-339">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-339">1.0</span></span>|
|[<span data-ttu-id="87d5c-340">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-341">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-342">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-343">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="87d5c-343">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13"></a><span data-ttu-id="87d5c-344">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-344">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span></span>

<span data-ttu-id="87d5c-p114">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p114">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="87d5c-p115">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p115">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="87d5c-349">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="87d5c-349">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="87d5c-350">Type</span><span class="sxs-lookup"><span data-stu-id="87d5c-350">Type</span></span>

*   [<span data-ttu-id="87d5c-351">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="87d5c-351">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="87d5c-352">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-352">Requirements</span></span>

|<span data-ttu-id="87d5c-353">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-353">Requirement</span></span>| <span data-ttu-id="87d5c-354">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-354">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-355">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-356">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-356">1.0</span></span>|
|[<span data-ttu-id="87d5c-357">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-357">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-358">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-359">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-359">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-360">読み取り</span><span class="sxs-lookup"><span data-stu-id="87d5c-360">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-361">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-361">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="87d5c-362">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="87d5c-362">internetMessageId: String</span></span>

<span data-ttu-id="87d5c-p116">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="87d5c-365">Type</span><span class="sxs-lookup"><span data-stu-id="87d5c-365">Type</span></span>

*   <span data-ttu-id="87d5c-366">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-366">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-367">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-367">Requirements</span></span>

|<span data-ttu-id="87d5c-368">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-368">Requirement</span></span>| <span data-ttu-id="87d5c-369">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-369">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-370">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-370">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-371">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-371">1.0</span></span>|
|[<span data-ttu-id="87d5c-372">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-372">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-373">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-374">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-374">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-375">読み取り</span><span class="sxs-lookup"><span data-stu-id="87d5c-375">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-376">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-376">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="87d5c-377">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="87d5c-377">itemClass: String</span></span>

<span data-ttu-id="87d5c-p117">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="87d5c-p118">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="87d5c-382">種類</span><span class="sxs-lookup"><span data-stu-id="87d5c-382">Type</span></span> | <span data-ttu-id="87d5c-383">説明</span><span class="sxs-lookup"><span data-stu-id="87d5c-383">Description</span></span> | <span data-ttu-id="87d5c-384">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="87d5c-384">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="87d5c-385">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="87d5c-385">Appointment items</span></span> | <span data-ttu-id="87d5c-386">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="87d5c-386">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="87d5c-387">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="87d5c-387">Message items</span></span> | <span data-ttu-id="87d5c-388">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-388">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="87d5c-389">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-389">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="87d5c-390">Type</span><span class="sxs-lookup"><span data-stu-id="87d5c-390">Type</span></span>

*   <span data-ttu-id="87d5c-391">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-391">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-392">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-392">Requirements</span></span>

|<span data-ttu-id="87d5c-393">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-393">Requirement</span></span>| <span data-ttu-id="87d5c-394">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-395">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-396">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-396">1.0</span></span>|
|[<span data-ttu-id="87d5c-397">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-397">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-398">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-399">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-399">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-400">読み取り</span><span class="sxs-lookup"><span data-stu-id="87d5c-400">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-401">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-401">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="87d5c-402">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="87d5c-402">(nullable) itemId: String</span></span>

<span data-ttu-id="87d5c-p119">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p119">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="87d5c-405">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="87d5c-405">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="87d5c-406">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="87d5c-406">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="87d5c-407">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="87d5c-407">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="87d5c-408">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="87d5c-408">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="87d5c-p121">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="87d5c-411">Type</span><span class="sxs-lookup"><span data-stu-id="87d5c-411">Type</span></span>

*   <span data-ttu-id="87d5c-412">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-412">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-413">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-413">Requirements</span></span>

|<span data-ttu-id="87d5c-414">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-414">Requirement</span></span>| <span data-ttu-id="87d5c-415">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-415">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-416">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-416">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-417">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-417">1.0</span></span>|
|[<span data-ttu-id="87d5c-418">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-418">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-419">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-419">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-420">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-420">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-421">読み取り</span><span class="sxs-lookup"><span data-stu-id="87d5c-421">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-422">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-422">Example</span></span>

<span data-ttu-id="87d5c-p122">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-13"></a><span data-ttu-id="87d5c-425">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-425">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3)</span></span>

<span data-ttu-id="87d5c-426">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-426">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="87d5c-427">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="87d5c-427">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="87d5c-428">型</span><span class="sxs-lookup"><span data-stu-id="87d5c-428">Type</span></span>

*   [<span data-ttu-id="87d5c-429">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="87d5c-429">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="87d5c-430">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-430">Requirements</span></span>

|<span data-ttu-id="87d5c-431">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-431">Requirement</span></span>| <span data-ttu-id="87d5c-432">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-432">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-433">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-433">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-434">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-434">1.0</span></span>|
|[<span data-ttu-id="87d5c-435">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-435">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-436">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-436">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-437">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-437">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-438">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="87d5c-438">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-439">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-439">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-13"></a><span data-ttu-id="87d5c-440">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-440">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span></span>

<span data-ttu-id="87d5c-441">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-441">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="87d5c-442">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="87d5c-442">Read mode</span></span>

<span data-ttu-id="87d5c-443">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-443">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="87d5c-444">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="87d5c-444">Compose mode</span></span>

<span data-ttu-id="87d5c-445">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-445">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="87d5c-446">型</span><span class="sxs-lookup"><span data-stu-id="87d5c-446">Type</span></span>

*   <span data-ttu-id="87d5c-447">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-447">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-448">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-448">Requirements</span></span>

|<span data-ttu-id="87d5c-449">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-449">Requirement</span></span>| <span data-ttu-id="87d5c-450">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-451">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-451">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-452">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-452">1.0</span></span>|
|[<span data-ttu-id="87d5c-453">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-453">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-454">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-455">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-455">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-456">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="87d5c-456">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="87d5c-457">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="87d5c-457">normalizedSubject: String</span></span>

<span data-ttu-id="87d5c-p123">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="87d5c-p124">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="87d5c-462">Type</span><span class="sxs-lookup"><span data-stu-id="87d5c-462">Type</span></span>

*   <span data-ttu-id="87d5c-463">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-463">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-464">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-464">Requirements</span></span>

|<span data-ttu-id="87d5c-465">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-465">Requirement</span></span>| <span data-ttu-id="87d5c-466">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-467">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-468">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-468">1.0</span></span>|
|[<span data-ttu-id="87d5c-469">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-470">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-471">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-472">読み取り</span><span class="sxs-lookup"><span data-stu-id="87d5c-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-473">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-473">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-13"></a><span data-ttu-id="87d5c-474">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-474">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3)</span></span>

<span data-ttu-id="87d5c-475">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-475">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="87d5c-476">型</span><span class="sxs-lookup"><span data-stu-id="87d5c-476">Type</span></span>

*   [<span data-ttu-id="87d5c-477">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="87d5c-477">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="87d5c-478">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-478">Requirements</span></span>

|<span data-ttu-id="87d5c-479">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-479">Requirement</span></span>| <span data-ttu-id="87d5c-480">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-481">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-482">1.3</span><span class="sxs-lookup"><span data-stu-id="87d5c-482">1.3</span></span>|
|[<span data-ttu-id="87d5c-483">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-484">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-485">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-486">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="87d5c-486">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-487">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-487">Example</span></span>

```js
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="87d5c-488">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-488">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="87d5c-489">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-489">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="87d5c-490">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="87d5c-490">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="87d5c-491">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="87d5c-491">Read mode</span></span>

<span data-ttu-id="87d5c-492">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-492">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="87d5c-493">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="87d5c-493">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="87d5c-494">ただし、Windows および Mac では、最大500メンバーを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-494">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="87d5c-495">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="87d5c-495">Compose mode</span></span>

<span data-ttu-id="87d5c-496">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-496">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="87d5c-497">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="87d5c-497">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="87d5c-498">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-498">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="87d5c-499">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-499">Get 500 members maximum.</span></span>
- <span data-ttu-id="87d5c-500">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="87d5c-500">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="87d5c-501">Type</span><span class="sxs-lookup"><span data-stu-id="87d5c-501">Type</span></span>

*   <span data-ttu-id="87d5c-502">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-502">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-503">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-503">Requirements</span></span>

|<span data-ttu-id="87d5c-504">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-504">Requirement</span></span>| <span data-ttu-id="87d5c-505">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-506">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-507">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-507">1.0</span></span>|
|[<span data-ttu-id="87d5c-508">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-508">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-509">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-510">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-510">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-511">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="87d5c-511">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13"></a><span data-ttu-id="87d5c-512">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-512">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span></span>

<span data-ttu-id="87d5c-p128">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="87d5c-515">型</span><span class="sxs-lookup"><span data-stu-id="87d5c-515">Type</span></span>

*   [<span data-ttu-id="87d5c-516">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="87d5c-516">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="87d5c-517">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-517">Requirements</span></span>

|<span data-ttu-id="87d5c-518">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-518">Requirement</span></span>| <span data-ttu-id="87d5c-519">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-520">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-521">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-521">1.0</span></span>|
|[<span data-ttu-id="87d5c-522">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-523">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-524">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-525">読み取り</span><span class="sxs-lookup"><span data-stu-id="87d5c-525">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-526">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-526">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="87d5c-527">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-527">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="87d5c-528">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-528">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="87d5c-529">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="87d5c-529">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="87d5c-530">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="87d5c-530">Read mode</span></span>

<span data-ttu-id="87d5c-531">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-531">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="87d5c-532">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="87d5c-532">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="87d5c-533">ただし、Windows および Mac では、最大500メンバーを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-533">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="87d5c-534">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="87d5c-534">Compose mode</span></span>

<span data-ttu-id="87d5c-535">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-535">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="87d5c-536">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="87d5c-536">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="87d5c-537">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-537">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="87d5c-538">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-538">Get 500 members maximum.</span></span>
- <span data-ttu-id="87d5c-539">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="87d5c-539">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="87d5c-540">Type</span><span class="sxs-lookup"><span data-stu-id="87d5c-540">Type</span></span>

*   <span data-ttu-id="87d5c-541">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-541">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-542">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-542">Requirements</span></span>

|<span data-ttu-id="87d5c-543">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-543">Requirement</span></span>| <span data-ttu-id="87d5c-544">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-545">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-546">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-546">1.0</span></span>|
|[<span data-ttu-id="87d5c-547">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-548">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-549">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-550">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="87d5c-550">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13"></a><span data-ttu-id="87d5c-551">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-551">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span></span>

<span data-ttu-id="87d5c-p132">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="87d5c-p133">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="87d5c-556">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="87d5c-556">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="87d5c-557">Type</span><span class="sxs-lookup"><span data-stu-id="87d5c-557">Type</span></span>

*   [<span data-ttu-id="87d5c-558">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="87d5c-558">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="87d5c-559">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-559">Requirements</span></span>

|<span data-ttu-id="87d5c-560">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-560">Requirement</span></span>| <span data-ttu-id="87d5c-561">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-561">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-562">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-562">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-563">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-563">1.0</span></span>|
|[<span data-ttu-id="87d5c-564">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-564">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-565">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-566">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-566">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-567">読み取り</span><span class="sxs-lookup"><span data-stu-id="87d5c-567">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-568">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-568">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-13"></a><span data-ttu-id="87d5c-569">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-569">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

<span data-ttu-id="87d5c-570">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-570">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="87d5c-p134">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="87d5c-573">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="87d5c-573">Read mode</span></span>

<span data-ttu-id="87d5c-574">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-574">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="87d5c-575">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="87d5c-575">Compose mode</span></span>

<span data-ttu-id="87d5c-576">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-576">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="87d5c-577">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="87d5c-577">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="87d5c-578">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-578">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="87d5c-579">型</span><span class="sxs-lookup"><span data-stu-id="87d5c-579">Type</span></span>

*   <span data-ttu-id="87d5c-580">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-580">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-581">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-581">Requirements</span></span>

|<span data-ttu-id="87d5c-582">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-582">Requirement</span></span>| <span data-ttu-id="87d5c-583">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-583">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-584">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-584">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-585">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-585">1.0</span></span>|
|[<span data-ttu-id="87d5c-586">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-586">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-587">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-587">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-588">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-588">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-589">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="87d5c-589">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-13"></a><span data-ttu-id="87d5c-590">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-590">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span></span>

<span data-ttu-id="87d5c-591">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-591">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="87d5c-592">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-592">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="87d5c-593">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="87d5c-593">Read mode</span></span>

<span data-ttu-id="87d5c-p135">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="87d5c-596">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="87d5c-596">Compose mode</span></span>

<span data-ttu-id="87d5c-597">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-597">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="87d5c-598">型</span><span class="sxs-lookup"><span data-stu-id="87d5c-598">Type</span></span>

*   <span data-ttu-id="87d5c-599">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-599">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-600">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-600">Requirements</span></span>

|<span data-ttu-id="87d5c-601">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-601">Requirement</span></span>| <span data-ttu-id="87d5c-602">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-603">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-603">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-604">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-604">1.0</span></span>|
|[<span data-ttu-id="87d5c-605">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-605">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-606">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-606">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-607">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-607">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-608">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="87d5c-608">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="87d5c-609">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-609">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="87d5c-610">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-610">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="87d5c-611">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="87d5c-611">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="87d5c-612">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="87d5c-612">Read mode</span></span>

<span data-ttu-id="87d5c-613">`to` プロパティは、メッセージの `EmailAddressDetails` 行にある各受信者について、\*\*\*\* オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-613">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="87d5c-614">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="87d5c-614">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="87d5c-615">ただし、Windows および Mac では、最大500メンバーを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-615">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="87d5c-616">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="87d5c-616">Compose mode</span></span>

<span data-ttu-id="87d5c-617">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-617">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="87d5c-618">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="87d5c-618">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="87d5c-619">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-619">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="87d5c-620">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-620">Get 500 members maximum.</span></span>
- <span data-ttu-id="87d5c-621">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="87d5c-621">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="87d5c-622">Type</span><span class="sxs-lookup"><span data-stu-id="87d5c-622">Type</span></span>

*   <span data-ttu-id="87d5c-623">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-623">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-624">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-624">Requirements</span></span>

|<span data-ttu-id="87d5c-625">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-625">Requirement</span></span>| <span data-ttu-id="87d5c-626">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-627">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-628">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-628">1.0</span></span>|
|[<span data-ttu-id="87d5c-629">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-629">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-630">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-631">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-631">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-632">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="87d5c-632">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="87d5c-633">メソッド</span><span class="sxs-lookup"><span data-stu-id="87d5c-633">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="87d5c-634">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="87d5c-634">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="87d5c-635">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-635">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="87d5c-636">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-636">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="87d5c-637">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-637">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="87d5c-638">パラメーター</span><span class="sxs-lookup"><span data-stu-id="87d5c-638">Parameters</span></span>

|<span data-ttu-id="87d5c-639">名前</span><span class="sxs-lookup"><span data-stu-id="87d5c-639">Name</span></span>| <span data-ttu-id="87d5c-640">種類</span><span class="sxs-lookup"><span data-stu-id="87d5c-640">Type</span></span>| <span data-ttu-id="87d5c-641">属性</span><span class="sxs-lookup"><span data-stu-id="87d5c-641">Attributes</span></span>| <span data-ttu-id="87d5c-642">説明</span><span class="sxs-lookup"><span data-stu-id="87d5c-642">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="87d5c-643">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-643">String</span></span>||<span data-ttu-id="87d5c-p139">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="87d5c-646">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-646">String</span></span>||<span data-ttu-id="87d5c-p140">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="87d5c-649">Object</span><span class="sxs-lookup"><span data-stu-id="87d5c-649">Object</span></span>| <span data-ttu-id="87d5c-650">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-650">&lt;optional&gt;</span></span>|<span data-ttu-id="87d5c-651">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="87d5c-651">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="87d5c-652">Object</span><span class="sxs-lookup"><span data-stu-id="87d5c-652">Object</span></span>| <span data-ttu-id="87d5c-653">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-653">&lt;optional&gt;</span></span>|<span data-ttu-id="87d5c-654">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-654">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="87d5c-655">function</span><span class="sxs-lookup"><span data-stu-id="87d5c-655">function</span></span>| <span data-ttu-id="87d5c-656">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-656">&lt;optional&gt;</span></span>|<span data-ttu-id="87d5c-657">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-657">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="87d5c-658">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-658">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="87d5c-659">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-659">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="87d5c-660">エラー</span><span class="sxs-lookup"><span data-stu-id="87d5c-660">Errors</span></span>

| <span data-ttu-id="87d5c-661">エラー コード</span><span class="sxs-lookup"><span data-stu-id="87d5c-661">Error code</span></span> | <span data-ttu-id="87d5c-662">説明</span><span class="sxs-lookup"><span data-stu-id="87d5c-662">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="87d5c-663">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="87d5c-663">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="87d5c-664">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="87d5c-664">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="87d5c-665">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-665">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="87d5c-666">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-666">Requirements</span></span>

|<span data-ttu-id="87d5c-667">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-667">Requirement</span></span>| <span data-ttu-id="87d5c-668">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-669">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-670">1.1</span><span class="sxs-lookup"><span data-stu-id="87d5c-670">1.1</span></span>|
|[<span data-ttu-id="87d5c-671">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-671">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-672">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-672">ReadWriteItem</span></span>|
|[<span data-ttu-id="87d5c-673">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-673">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-674">作成</span><span class="sxs-lookup"><span data-stu-id="87d5c-674">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-675">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-675">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="87d5c-676">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="87d5c-676">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="87d5c-677">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-677">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="87d5c-p141">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="87d5c-681">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-681">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="87d5c-682">Office アドインを Outlook on the web で実行している場合、編集中のアイテム以外のアイテムに `addItemAttachmentAsync` メソッドでアイテムを添付できます。ただし、これはサポートされていないため、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="87d5c-682">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="87d5c-683">パラメーター</span><span class="sxs-lookup"><span data-stu-id="87d5c-683">Parameters</span></span>

|<span data-ttu-id="87d5c-684">名前</span><span class="sxs-lookup"><span data-stu-id="87d5c-684">Name</span></span>| <span data-ttu-id="87d5c-685">種類</span><span class="sxs-lookup"><span data-stu-id="87d5c-685">Type</span></span>| <span data-ttu-id="87d5c-686">属性</span><span class="sxs-lookup"><span data-stu-id="87d5c-686">Attributes</span></span>| <span data-ttu-id="87d5c-687">説明</span><span class="sxs-lookup"><span data-stu-id="87d5c-687">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="87d5c-688">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-688">String</span></span>||<span data-ttu-id="87d5c-p142">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="87d5c-691">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-691">String</span></span>||<span data-ttu-id="87d5c-692">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="87d5c-692">The subject of the item to be attached.</span></span> <span data-ttu-id="87d5c-693">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="87d5c-693">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="87d5c-694">Object</span><span class="sxs-lookup"><span data-stu-id="87d5c-694">Object</span></span>| <span data-ttu-id="87d5c-695">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-695">&lt;optional&gt;</span></span>|<span data-ttu-id="87d5c-696">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="87d5c-696">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="87d5c-697">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="87d5c-697">Object</span></span>| <span data-ttu-id="87d5c-698">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-698">&lt;optional&gt;</span></span>|<span data-ttu-id="87d5c-699">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-699">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="87d5c-700">関数</span><span class="sxs-lookup"><span data-stu-id="87d5c-700">function</span></span>| <span data-ttu-id="87d5c-701">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-701">&lt;optional&gt;</span></span>|<span data-ttu-id="87d5c-702">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-702">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="87d5c-703">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-703">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="87d5c-704">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-704">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="87d5c-705">エラー</span><span class="sxs-lookup"><span data-stu-id="87d5c-705">Errors</span></span>

| <span data-ttu-id="87d5c-706">エラー コード</span><span class="sxs-lookup"><span data-stu-id="87d5c-706">Error code</span></span> | <span data-ttu-id="87d5c-707">説明</span><span class="sxs-lookup"><span data-stu-id="87d5c-707">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="87d5c-708">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-708">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="87d5c-709">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-709">Requirements</span></span>

|<span data-ttu-id="87d5c-710">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-710">Requirement</span></span>| <span data-ttu-id="87d5c-711">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-711">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-712">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-712">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-713">1.1</span><span class="sxs-lookup"><span data-stu-id="87d5c-713">1.1</span></span>|
|[<span data-ttu-id="87d5c-714">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-714">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-715">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-715">ReadWriteItem</span></span>|
|[<span data-ttu-id="87d5c-716">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-716">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-717">作成</span><span class="sxs-lookup"><span data-stu-id="87d5c-717">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-718">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-718">Example</span></span>

<span data-ttu-id="87d5c-719">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-719">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="87d5c-720">close()</span><span class="sxs-lookup"><span data-stu-id="87d5c-720">close()</span></span>

<span data-ttu-id="87d5c-721">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-721">Closes the current item that is being composed.</span></span>

<span data-ttu-id="87d5c-p144">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="87d5c-724">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-724">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="87d5c-725">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="87d5c-725">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-726">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-726">Requirements</span></span>

|<span data-ttu-id="87d5c-727">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-727">Requirement</span></span>| <span data-ttu-id="87d5c-728">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-729">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-730">1.3</span><span class="sxs-lookup"><span data-stu-id="87d5c-730">1.3</span></span>|
|[<span data-ttu-id="87d5c-731">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-731">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-732">制限あり</span><span class="sxs-lookup"><span data-stu-id="87d5c-732">Restricted</span></span>|
|[<span data-ttu-id="87d5c-733">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-733">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-734">新規作成</span><span class="sxs-lookup"><span data-stu-id="87d5c-734">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="87d5c-735">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="87d5c-735">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="87d5c-736">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-736">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="87d5c-737">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="87d5c-737">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="87d5c-738">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-738">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="87d5c-739">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="87d5c-739">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="87d5c-p145">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="87d5c-743">パラメーター</span><span class="sxs-lookup"><span data-stu-id="87d5c-743">Parameters</span></span>

|<span data-ttu-id="87d5c-744">名前</span><span class="sxs-lookup"><span data-stu-id="87d5c-744">Name</span></span>| <span data-ttu-id="87d5c-745">種類</span><span class="sxs-lookup"><span data-stu-id="87d5c-745">Type</span></span>| <span data-ttu-id="87d5c-746">説明</span><span class="sxs-lookup"><span data-stu-id="87d5c-746">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="87d5c-747">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="87d5c-747">String &#124; Object</span></span>| |<span data-ttu-id="87d5c-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="87d5c-750">**または**</span><span class="sxs-lookup"><span data-stu-id="87d5c-750">**OR**</span></span><br/><span data-ttu-id="87d5c-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="87d5c-753">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-753">String</span></span> | <span data-ttu-id="87d5c-754">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-754">&lt;optional&gt;</span></span> | <span data-ttu-id="87d5c-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="87d5c-757">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-757">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="87d5c-758">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-758">&lt;optional&gt;</span></span> | <span data-ttu-id="87d5c-759">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="87d5c-759">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="87d5c-760">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-760">String</span></span> | | <span data-ttu-id="87d5c-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="87d5c-763">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-763">String</span></span> | | <span data-ttu-id="87d5c-764">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="87d5c-764">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="87d5c-765">文字列</span><span class="sxs-lookup"><span data-stu-id="87d5c-765">String</span></span> | | <span data-ttu-id="87d5c-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="87d5c-768">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-768">String</span></span> | | <span data-ttu-id="87d5c-p151">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="87d5c-772">function</span><span class="sxs-lookup"><span data-stu-id="87d5c-772">function</span></span> | <span data-ttu-id="87d5c-773">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-773">&lt;optional&gt;</span></span> | <span data-ttu-id="87d5c-774">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-774">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="87d5c-775">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-775">Requirements</span></span>

|<span data-ttu-id="87d5c-776">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-776">Requirement</span></span>| <span data-ttu-id="87d5c-777">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-777">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-778">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-778">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-779">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-779">1.0</span></span>|
|[<span data-ttu-id="87d5c-780">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-780">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-781">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-781">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-782">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-782">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-783">読み取り</span><span class="sxs-lookup"><span data-stu-id="87d5c-783">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="87d5c-784">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-784">Examples</span></span>

<span data-ttu-id="87d5c-785">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-785">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="87d5c-786">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-786">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="87d5c-787">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-787">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="87d5c-788">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-788">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="87d5c-789">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-789">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="87d5c-790">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-790">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="87d5c-791">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="87d5c-791">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="87d5c-792">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-792">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="87d5c-793">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="87d5c-793">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="87d5c-794">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-794">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="87d5c-795">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="87d5c-795">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="87d5c-p152">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p152">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="87d5c-799">パラメーター</span><span class="sxs-lookup"><span data-stu-id="87d5c-799">Parameters</span></span>

|<span data-ttu-id="87d5c-800">名前</span><span class="sxs-lookup"><span data-stu-id="87d5c-800">Name</span></span>| <span data-ttu-id="87d5c-801">種類</span><span class="sxs-lookup"><span data-stu-id="87d5c-801">Type</span></span>| <span data-ttu-id="87d5c-802">説明</span><span class="sxs-lookup"><span data-stu-id="87d5c-802">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="87d5c-803">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="87d5c-803">String &#124; Object</span></span>| | <span data-ttu-id="87d5c-p153">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="87d5c-806">**または**</span><span class="sxs-lookup"><span data-stu-id="87d5c-806">**OR**</span></span><br/><span data-ttu-id="87d5c-p154">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p154">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="87d5c-809">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-809">String</span></span> | <span data-ttu-id="87d5c-810">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-810">&lt;optional&gt;</span></span> | <span data-ttu-id="87d5c-p155">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p155">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="87d5c-813">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-813">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="87d5c-814">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-814">&lt;optional&gt;</span></span> | <span data-ttu-id="87d5c-815">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="87d5c-815">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="87d5c-816">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-816">String</span></span> | | <span data-ttu-id="87d5c-p156">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p156">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="87d5c-819">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-819">String</span></span> | | <span data-ttu-id="87d5c-820">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="87d5c-820">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="87d5c-821">文字列</span><span class="sxs-lookup"><span data-stu-id="87d5c-821">String</span></span> | | <span data-ttu-id="87d5c-p157">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p157">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="87d5c-824">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-824">String</span></span> | | <span data-ttu-id="87d5c-p158">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="87d5c-828">function</span><span class="sxs-lookup"><span data-stu-id="87d5c-828">function</span></span> | <span data-ttu-id="87d5c-829">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-829">&lt;optional&gt;</span></span> | <span data-ttu-id="87d5c-830">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-830">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="87d5c-831">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-831">Requirements</span></span>

|<span data-ttu-id="87d5c-832">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-832">Requirement</span></span>| <span data-ttu-id="87d5c-833">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-833">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-834">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-834">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-835">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-835">1.0</span></span>|
|[<span data-ttu-id="87d5c-836">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-836">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-837">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-837">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-838">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-838">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-839">読み取り</span><span class="sxs-lookup"><span data-stu-id="87d5c-839">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="87d5c-840">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-840">Examples</span></span>

<span data-ttu-id="87d5c-841">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-841">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="87d5c-842">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-842">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="87d5c-843">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-843">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="87d5c-844">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-844">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="87d5c-845">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-845">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="87d5c-846">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-846">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-13"></a><span data-ttu-id="87d5c-847">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)}</span><span class="sxs-lookup"><span data-stu-id="87d5c-847">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)}</span></span>

<span data-ttu-id="87d5c-848">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-848">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="87d5c-849">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="87d5c-849">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-850">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-850">Requirements</span></span>

|<span data-ttu-id="87d5c-851">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-851">Requirement</span></span>| <span data-ttu-id="87d5c-852">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-852">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-853">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-853">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-854">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-854">1.0</span></span>|
|[<span data-ttu-id="87d5c-855">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-855">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-856">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-856">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-857">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-857">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-858">読み取り</span><span class="sxs-lookup"><span data-stu-id="87d5c-858">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="87d5c-859">戻り値:</span><span class="sxs-lookup"><span data-stu-id="87d5c-859">Returns:</span></span>

<span data-ttu-id="87d5c-860">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="87d5c-860">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)</span></span>

##### <a name="example"></a><span data-ttu-id="87d5c-861">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-861">Example</span></span>

<span data-ttu-id="87d5c-862">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="87d5c-862">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-13meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-13phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-13tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-13"></a><span data-ttu-id="87d5c-863">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span><span class="sxs-lookup"><span data-stu-id="87d5c-863">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span></span>

<span data-ttu-id="87d5c-864">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-864">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="87d5c-865">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="87d5c-865">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="87d5c-866">パラメーター</span><span class="sxs-lookup"><span data-stu-id="87d5c-866">Parameters</span></span>

|<span data-ttu-id="87d5c-867">名前</span><span class="sxs-lookup"><span data-stu-id="87d5c-867">Name</span></span>| <span data-ttu-id="87d5c-868">種類</span><span class="sxs-lookup"><span data-stu-id="87d5c-868">Type</span></span>| <span data-ttu-id="87d5c-869">説明</span><span class="sxs-lookup"><span data-stu-id="87d5c-869">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="87d5c-870">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="87d5c-870">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.3)|<span data-ttu-id="87d5c-871">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="87d5c-871">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="87d5c-872">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-872">Requirements</span></span>

|<span data-ttu-id="87d5c-873">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-873">Requirement</span></span>| <span data-ttu-id="87d5c-874">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-874">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-875">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-875">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-876">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-876">1.0</span></span>|
|[<span data-ttu-id="87d5c-877">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-877">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-878">制限あり</span><span class="sxs-lookup"><span data-stu-id="87d5c-878">Restricted</span></span>|
|[<span data-ttu-id="87d5c-879">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-879">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-880">読み取り</span><span class="sxs-lookup"><span data-stu-id="87d5c-880">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="87d5c-881">戻り値:</span><span class="sxs-lookup"><span data-stu-id="87d5c-881">Returns:</span></span>

<span data-ttu-id="87d5c-882">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-882">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="87d5c-883">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-883">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="87d5c-884">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="87d5c-884">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="87d5c-885">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="87d5c-885">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="87d5c-886">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="87d5c-886">Value of `entityType`</span></span> | <span data-ttu-id="87d5c-887">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="87d5c-887">Type of objects in returned array</span></span> | <span data-ttu-id="87d5c-888">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-888">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="87d5c-889">文字列</span><span class="sxs-lookup"><span data-stu-id="87d5c-889">String</span></span> | <span data-ttu-id="87d5c-890">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="87d5c-890">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="87d5c-891">連絡先</span><span class="sxs-lookup"><span data-stu-id="87d5c-891">Contact</span></span> | <span data-ttu-id="87d5c-892">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="87d5c-892">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="87d5c-893">文字列</span><span class="sxs-lookup"><span data-stu-id="87d5c-893">String</span></span> | <span data-ttu-id="87d5c-894">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="87d5c-894">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="87d5c-895">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="87d5c-895">MeetingSuggestion</span></span> | <span data-ttu-id="87d5c-896">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="87d5c-896">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="87d5c-897">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="87d5c-897">PhoneNumber</span></span> | <span data-ttu-id="87d5c-898">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="87d5c-898">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="87d5c-899">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="87d5c-899">TaskSuggestion</span></span> | <span data-ttu-id="87d5c-900">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="87d5c-900">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="87d5c-901">文字列</span><span class="sxs-lookup"><span data-stu-id="87d5c-901">String</span></span> | <span data-ttu-id="87d5c-902">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="87d5c-902">**Restricted**</span></span> |

<span data-ttu-id="87d5c-903">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span><span class="sxs-lookup"><span data-stu-id="87d5c-903">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span></span>

##### <a name="example"></a><span data-ttu-id="87d5c-904">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-904">Example</span></span>

<span data-ttu-id="87d5c-905">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-905">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-13meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-13phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-13tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-13"></a><span data-ttu-id="87d5c-906">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span><span class="sxs-lookup"><span data-stu-id="87d5c-906">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span></span>

<span data-ttu-id="87d5c-907">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-907">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="87d5c-908">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="87d5c-908">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="87d5c-909">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-909">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="87d5c-910">パラメーター</span><span class="sxs-lookup"><span data-stu-id="87d5c-910">Parameters</span></span>

|<span data-ttu-id="87d5c-911">名前</span><span class="sxs-lookup"><span data-stu-id="87d5c-911">Name</span></span>| <span data-ttu-id="87d5c-912">種類</span><span class="sxs-lookup"><span data-stu-id="87d5c-912">Type</span></span>| <span data-ttu-id="87d5c-913">説明</span><span class="sxs-lookup"><span data-stu-id="87d5c-913">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="87d5c-914">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-914">String</span></span>|<span data-ttu-id="87d5c-915">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="87d5c-915">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="87d5c-916">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-916">Requirements</span></span>

|<span data-ttu-id="87d5c-917">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-917">Requirement</span></span>| <span data-ttu-id="87d5c-918">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-918">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-919">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-919">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-920">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-920">1.0</span></span>|
|[<span data-ttu-id="87d5c-921">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-921">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-922">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-922">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-923">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-923">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-924">読み取り</span><span class="sxs-lookup"><span data-stu-id="87d5c-924">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="87d5c-925">戻り値:</span><span class="sxs-lookup"><span data-stu-id="87d5c-925">Returns:</span></span>

<span data-ttu-id="87d5c-p160">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="87d5c-928">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span><span class="sxs-lookup"><span data-stu-id="87d5c-928">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="87d5c-929">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="87d5c-929">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="87d5c-930">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-930">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="87d5c-931">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="87d5c-931">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="87d5c-p161">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="87d5c-935">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="87d5c-935">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="87d5c-936">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="87d5c-936">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="87d5c-p162">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="87d5c-940">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-940">Requirements</span></span>

|<span data-ttu-id="87d5c-941">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-941">Requirement</span></span>| <span data-ttu-id="87d5c-942">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-942">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-943">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-943">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-944">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-944">1.0</span></span>|
|[<span data-ttu-id="87d5c-945">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-945">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-946">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-946">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-947">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-947">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-948">読み取り</span><span class="sxs-lookup"><span data-stu-id="87d5c-948">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="87d5c-949">戻り値:</span><span class="sxs-lookup"><span data-stu-id="87d5c-949">Returns:</span></span>

<span data-ttu-id="87d5c-p163">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="87d5c-952">型: Object</span><span class="sxs-lookup"><span data-stu-id="87d5c-952">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="87d5c-953">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-953">Example</span></span>

<span data-ttu-id="87d5c-954">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="87d5c-954">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="87d5c-955">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="87d5c-955">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="87d5c-956">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-956">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="87d5c-957">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="87d5c-957">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="87d5c-958">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="87d5c-958">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="87d5c-p164">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="87d5c-961">パラメーター</span><span class="sxs-lookup"><span data-stu-id="87d5c-961">Parameters</span></span>

|<span data-ttu-id="87d5c-962">名前</span><span class="sxs-lookup"><span data-stu-id="87d5c-962">Name</span></span>| <span data-ttu-id="87d5c-963">種類</span><span class="sxs-lookup"><span data-stu-id="87d5c-963">Type</span></span>| <span data-ttu-id="87d5c-964">説明</span><span class="sxs-lookup"><span data-stu-id="87d5c-964">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="87d5c-965">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-965">String</span></span>|<span data-ttu-id="87d5c-966">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="87d5c-966">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="87d5c-967">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-967">Requirements</span></span>

|<span data-ttu-id="87d5c-968">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-968">Requirement</span></span>| <span data-ttu-id="87d5c-969">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-969">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-970">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-970">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-971">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-971">1.0</span></span>|
|[<span data-ttu-id="87d5c-972">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-972">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-973">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-973">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-974">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-974">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-975">読み取り</span><span class="sxs-lookup"><span data-stu-id="87d5c-975">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="87d5c-976">戻り値:</span><span class="sxs-lookup"><span data-stu-id="87d5c-976">Returns:</span></span>

<span data-ttu-id="87d5c-977">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="87d5c-977">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="87d5c-978">型: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="87d5c-978">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="87d5c-979">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-979">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="87d5c-980">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="87d5c-980">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="87d5c-981">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-981">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="87d5c-p165">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="87d5c-984">パラメーター</span><span class="sxs-lookup"><span data-stu-id="87d5c-984">Parameters</span></span>

|<span data-ttu-id="87d5c-985">名前</span><span class="sxs-lookup"><span data-stu-id="87d5c-985">Name</span></span>| <span data-ttu-id="87d5c-986">種類</span><span class="sxs-lookup"><span data-stu-id="87d5c-986">Type</span></span>| <span data-ttu-id="87d5c-987">属性</span><span class="sxs-lookup"><span data-stu-id="87d5c-987">Attributes</span></span>| <span data-ttu-id="87d5c-988">説明</span><span class="sxs-lookup"><span data-stu-id="87d5c-988">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="87d5c-989">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="87d5c-989">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="87d5c-p166">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="87d5c-993">Object</span><span class="sxs-lookup"><span data-stu-id="87d5c-993">Object</span></span>| <span data-ttu-id="87d5c-994">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-994">&lt;optional&gt;</span></span>|<span data-ttu-id="87d5c-995">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="87d5c-995">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="87d5c-996">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="87d5c-996">Object</span></span>| <span data-ttu-id="87d5c-997">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-997">&lt;optional&gt;</span></span>|<span data-ttu-id="87d5c-998">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-998">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="87d5c-999">function</span><span class="sxs-lookup"><span data-stu-id="87d5c-999">function</span></span>||<span data-ttu-id="87d5c-1000">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1000">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="87d5c-1001">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1001">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="87d5c-1002">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1002">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="87d5c-1003">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-1003">Requirements</span></span>

|<span data-ttu-id="87d5c-1004">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-1004">Requirement</span></span>| <span data-ttu-id="87d5c-1005">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-1006">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-1006">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-1007">1.2</span><span class="sxs-lookup"><span data-stu-id="87d5c-1007">1.2</span></span>|
|[<span data-ttu-id="87d5c-1008">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-1008">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-1009">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-1009">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-1010">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-1010">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-1011">作成</span><span class="sxs-lookup"><span data-stu-id="87d5c-1011">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="87d5c-1012">戻り値:</span><span class="sxs-lookup"><span data-stu-id="87d5c-1012">Returns:</span></span>

<span data-ttu-id="87d5c-1013">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1013">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="87d5c-1014">型:String</span><span class="sxs-lookup"><span data-stu-id="87d5c-1014">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="87d5c-1015">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-1015">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="87d5c-1016">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="87d5c-1016">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="87d5c-1017">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1017">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="87d5c-p168">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="87d5c-1021">パラメーター</span><span class="sxs-lookup"><span data-stu-id="87d5c-1021">Parameters</span></span>

|<span data-ttu-id="87d5c-1022">名前</span><span class="sxs-lookup"><span data-stu-id="87d5c-1022">Name</span></span>| <span data-ttu-id="87d5c-1023">種類</span><span class="sxs-lookup"><span data-stu-id="87d5c-1023">Type</span></span>| <span data-ttu-id="87d5c-1024">属性</span><span class="sxs-lookup"><span data-stu-id="87d5c-1024">Attributes</span></span>| <span data-ttu-id="87d5c-1025">説明</span><span class="sxs-lookup"><span data-stu-id="87d5c-1025">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="87d5c-1026">function</span><span class="sxs-lookup"><span data-stu-id="87d5c-1026">function</span></span>||<span data-ttu-id="87d5c-1027">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1027">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="87d5c-1028">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.3) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1028">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.3) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="87d5c-1029">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1029">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="87d5c-1030">Object</span><span class="sxs-lookup"><span data-stu-id="87d5c-1030">Object</span></span>| <span data-ttu-id="87d5c-1031">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-1031">&lt;optional&gt;</span></span>|<span data-ttu-id="87d5c-1032">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1032">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="87d5c-1033">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1033">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="87d5c-1034">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-1034">Requirements</span></span>

|<span data-ttu-id="87d5c-1035">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-1035">Requirement</span></span>| <span data-ttu-id="87d5c-1036">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-1036">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-1037">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-1037">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-1038">1.0</span><span class="sxs-lookup"><span data-stu-id="87d5c-1038">1.0</span></span>|
|[<span data-ttu-id="87d5c-1039">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-1039">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-1040">ReadItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-1040">ReadItem</span></span>|
|[<span data-ttu-id="87d5c-1041">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-1041">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-1042">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="87d5c-1042">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-1043">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-1043">Example</span></span>

<span data-ttu-id="87d5c-p171">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="87d5c-1047">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="87d5c-1047">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="87d5c-1048">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1048">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="87d5c-1049">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1049">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="87d5c-1050">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1050">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="87d5c-1051">Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1051">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="87d5c-1052">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1052">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="87d5c-1053">パラメーター</span><span class="sxs-lookup"><span data-stu-id="87d5c-1053">Parameters</span></span>

|<span data-ttu-id="87d5c-1054">名前</span><span class="sxs-lookup"><span data-stu-id="87d5c-1054">Name</span></span>| <span data-ttu-id="87d5c-1055">種類</span><span class="sxs-lookup"><span data-stu-id="87d5c-1055">Type</span></span>| <span data-ttu-id="87d5c-1056">属性</span><span class="sxs-lookup"><span data-stu-id="87d5c-1056">Attributes</span></span>| <span data-ttu-id="87d5c-1057">説明</span><span class="sxs-lookup"><span data-stu-id="87d5c-1057">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="87d5c-1058">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-1058">String</span></span>||<span data-ttu-id="87d5c-1059">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1059">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="87d5c-1060">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="87d5c-1060">Object</span></span>| <span data-ttu-id="87d5c-1061">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-1061">&lt;optional&gt;</span></span>|<span data-ttu-id="87d5c-1062">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1062">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="87d5c-1063">Object</span><span class="sxs-lookup"><span data-stu-id="87d5c-1063">Object</span></span>| <span data-ttu-id="87d5c-1064">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-1064">&lt;optional&gt;</span></span>|<span data-ttu-id="87d5c-1065">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1065">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="87d5c-1066">function</span><span class="sxs-lookup"><span data-stu-id="87d5c-1066">function</span></span>| <span data-ttu-id="87d5c-1067">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="87d5c-1068">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1068">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="87d5c-1069">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1069">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="87d5c-1070">エラー</span><span class="sxs-lookup"><span data-stu-id="87d5c-1070">Errors</span></span>

| <span data-ttu-id="87d5c-1071">エラー コード</span><span class="sxs-lookup"><span data-stu-id="87d5c-1071">Error code</span></span> | <span data-ttu-id="87d5c-1072">説明</span><span class="sxs-lookup"><span data-stu-id="87d5c-1072">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="87d5c-1073">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1073">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="87d5c-1074">Requirements</span><span class="sxs-lookup"><span data-stu-id="87d5c-1074">Requirements</span></span>

|<span data-ttu-id="87d5c-1075">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-1075">Requirement</span></span>| <span data-ttu-id="87d5c-1076">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-1076">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-1077">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-1077">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-1078">1.1</span><span class="sxs-lookup"><span data-stu-id="87d5c-1078">1.1</span></span>|
|[<span data-ttu-id="87d5c-1079">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-1079">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-1080">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-1080">ReadWriteItem</span></span>|
|[<span data-ttu-id="87d5c-1081">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-1081">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-1082">作成</span><span class="sxs-lookup"><span data-stu-id="87d5c-1082">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-1083">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-1083">Example</span></span>

<span data-ttu-id="87d5c-1084">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1084">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="87d5c-1085">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="87d5c-1085">saveAsync([options], callback)</span></span>

<span data-ttu-id="87d5c-1086">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1086">Asynchronously saves an item.</span></span>

<span data-ttu-id="87d5c-1087">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1087">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="87d5c-1088">Outlook on the web またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1088">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="87d5c-1089">キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1089">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="87d5c-1090">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1090">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="87d5c-1091">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1091">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="87d5c-p175">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="87d5c-1095">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1095">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="87d5c-1096">Outlook on Mac では、会議の保存はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1096">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="87d5c-1097">`saveAsync` メソッドは、作成モードの会議から呼び出されると失敗します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1097">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="87d5c-1098">回避策については、「[Office JS API を使用して Outlook for Mac で会議を下書きとして保存できない](https://support.microsoft.com/help/4505745)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1098">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="87d5c-1099">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1099">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="87d5c-1100">パラメーター</span><span class="sxs-lookup"><span data-stu-id="87d5c-1100">Parameters</span></span>

|<span data-ttu-id="87d5c-1101">名前</span><span class="sxs-lookup"><span data-stu-id="87d5c-1101">Name</span></span>| <span data-ttu-id="87d5c-1102">種類</span><span class="sxs-lookup"><span data-stu-id="87d5c-1102">Type</span></span>| <span data-ttu-id="87d5c-1103">属性</span><span class="sxs-lookup"><span data-stu-id="87d5c-1103">Attributes</span></span>| <span data-ttu-id="87d5c-1104">説明</span><span class="sxs-lookup"><span data-stu-id="87d5c-1104">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="87d5c-1105">Object</span><span class="sxs-lookup"><span data-stu-id="87d5c-1105">Object</span></span>| <span data-ttu-id="87d5c-1106">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-1106">&lt;optional&gt;</span></span>|<span data-ttu-id="87d5c-1107">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1107">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="87d5c-1108">Object</span><span class="sxs-lookup"><span data-stu-id="87d5c-1108">Object</span></span>| <span data-ttu-id="87d5c-1109">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-1109">&lt;optional&gt;</span></span>|<span data-ttu-id="87d5c-1110">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1110">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="87d5c-1111">function</span><span class="sxs-lookup"><span data-stu-id="87d5c-1111">function</span></span>||<span data-ttu-id="87d5c-1112">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1112">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="87d5c-1113">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1113">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="87d5c-1114">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-1114">Requirements</span></span>

|<span data-ttu-id="87d5c-1115">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-1115">Requirement</span></span>| <span data-ttu-id="87d5c-1116">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-1116">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-1117">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-1117">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-1118">1.3</span><span class="sxs-lookup"><span data-stu-id="87d5c-1118">1.3</span></span>|
|[<span data-ttu-id="87d5c-1119">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-1119">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-1120">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-1120">ReadWriteItem</span></span>|
|[<span data-ttu-id="87d5c-1121">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-1121">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-1122">作成</span><span class="sxs-lookup"><span data-stu-id="87d5c-1122">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="87d5c-1123">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-1123">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="87d5c-p177">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="87d5c-1126">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="87d5c-1126">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="87d5c-1127">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1127">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="87d5c-p178">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="87d5c-1131">パラメーター</span><span class="sxs-lookup"><span data-stu-id="87d5c-1131">Parameters</span></span>

|<span data-ttu-id="87d5c-1132">名前</span><span class="sxs-lookup"><span data-stu-id="87d5c-1132">Name</span></span>| <span data-ttu-id="87d5c-1133">種類</span><span class="sxs-lookup"><span data-stu-id="87d5c-1133">Type</span></span>| <span data-ttu-id="87d5c-1134">属性</span><span class="sxs-lookup"><span data-stu-id="87d5c-1134">Attributes</span></span>| <span data-ttu-id="87d5c-1135">説明</span><span class="sxs-lookup"><span data-stu-id="87d5c-1135">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="87d5c-1136">String</span><span class="sxs-lookup"><span data-stu-id="87d5c-1136">String</span></span>||<span data-ttu-id="87d5c-p179">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="87d5c-1140">Object</span><span class="sxs-lookup"><span data-stu-id="87d5c-1140">Object</span></span>| <span data-ttu-id="87d5c-1141">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-1141">&lt;optional&gt;</span></span>|<span data-ttu-id="87d5c-1142">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1142">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="87d5c-1143">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="87d5c-1143">Object</span></span>| <span data-ttu-id="87d5c-1144">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-1144">&lt;optional&gt;</span></span>|<span data-ttu-id="87d5c-1145">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1145">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="87d5c-1146">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="87d5c-1146">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="87d5c-1147">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="87d5c-1147">&lt;optional&gt;</span></span>|<span data-ttu-id="87d5c-1148">`text` の場合、Outlook on the web とデスクトップ クライアントでは現在のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1148">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="87d5c-1149">フィールドが HTML エディターの場合、データが HTML の場合でもテキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1149">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="87d5c-1150">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Outlook on the web では現在のスタイルが適用され、Outlook デスクトップ クライアントでは既定のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1150">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="87d5c-1151">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1151">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="87d5c-1152">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1152">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="87d5c-1153">function</span><span class="sxs-lookup"><span data-stu-id="87d5c-1153">function</span></span>||<span data-ttu-id="87d5c-1154">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="87d5c-1154">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="87d5c-1155">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-1155">Requirements</span></span>

|<span data-ttu-id="87d5c-1156">要件</span><span class="sxs-lookup"><span data-stu-id="87d5c-1156">Requirement</span></span>| <span data-ttu-id="87d5c-1157">値</span><span class="sxs-lookup"><span data-stu-id="87d5c-1157">Value</span></span>|
|---|---|
|[<span data-ttu-id="87d5c-1158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="87d5c-1158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87d5c-1159">1.2</span><span class="sxs-lookup"><span data-stu-id="87d5c-1159">1.2</span></span>|
|[<span data-ttu-id="87d5c-1160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="87d5c-1160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="87d5c-1161">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="87d5c-1161">ReadWriteItem</span></span>|
|[<span data-ttu-id="87d5c-1162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="87d5c-1162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="87d5c-1163">作成</span><span class="sxs-lookup"><span data-stu-id="87d5c-1163">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="87d5c-1164">例</span><span class="sxs-lookup"><span data-stu-id="87d5c-1164">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
