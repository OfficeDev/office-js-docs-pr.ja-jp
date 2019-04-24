---
title: Office. メールボックス-要件セット1.6
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: cc7897f791c5a07ed5c17a686b6601a1a7633f00
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451774"
---
# <a name="item"></a><span data-ttu-id="58d85-102">item</span><span class="sxs-lookup"><span data-stu-id="58d85-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="58d85-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="58d85-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="58d85-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="58d85-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-106">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-106">Requirements</span></span>

|<span data-ttu-id="58d85-107">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-107">Requirement</span></span>| <span data-ttu-id="58d85-108">値</span><span class="sxs-lookup"><span data-stu-id="58d85-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-110">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-110">1.0</span></span>|
|[<span data-ttu-id="58d85-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="58d85-112">Restricted</span></span>|
|[<span data-ttu-id="58d85-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58d85-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="58d85-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-115">Members and methods</span></span>

| <span data-ttu-id="58d85-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-116">Member</span></span> | <span data-ttu-id="58d85-117">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="58d85-118">attachments</span><span class="sxs-lookup"><span data-stu-id="58d85-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="58d85-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-119">Member</span></span> |
| [<span data-ttu-id="58d85-120">bcc</span><span class="sxs-lookup"><span data-stu-id="58d85-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="58d85-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-121">Member</span></span> |
| [<span data-ttu-id="58d85-122">body</span><span class="sxs-lookup"><span data-stu-id="58d85-122">body</span></span>](#body-body) | <span data-ttu-id="58d85-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-123">Member</span></span> |
| [<span data-ttu-id="58d85-124">cc</span><span class="sxs-lookup"><span data-stu-id="58d85-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="58d85-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-125">Member</span></span> |
| [<span data-ttu-id="58d85-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="58d85-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="58d85-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-127">Member</span></span> |
| [<span data-ttu-id="58d85-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="58d85-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="58d85-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-129">Member</span></span> |
| [<span data-ttu-id="58d85-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="58d85-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="58d85-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-131">Member</span></span> |
| [<span data-ttu-id="58d85-132">end</span><span class="sxs-lookup"><span data-stu-id="58d85-132">end</span></span>](#end-datetime) | <span data-ttu-id="58d85-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-133">Member</span></span> |
| [<span data-ttu-id="58d85-134">from</span><span class="sxs-lookup"><span data-stu-id="58d85-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="58d85-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-135">Member</span></span> |
| [<span data-ttu-id="58d85-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="58d85-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="58d85-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-137">Member</span></span> |
| [<span data-ttu-id="58d85-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="58d85-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="58d85-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-139">Member</span></span> |
| [<span data-ttu-id="58d85-140">itemId</span><span class="sxs-lookup"><span data-stu-id="58d85-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="58d85-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-141">Member</span></span> |
| [<span data-ttu-id="58d85-142">itemType</span><span class="sxs-lookup"><span data-stu-id="58d85-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="58d85-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-143">Member</span></span> |
| [<span data-ttu-id="58d85-144">location</span><span class="sxs-lookup"><span data-stu-id="58d85-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="58d85-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-145">Member</span></span> |
| [<span data-ttu-id="58d85-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="58d85-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="58d85-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-147">Member</span></span> |
| [<span data-ttu-id="58d85-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="58d85-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="58d85-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-149">Member</span></span> |
| [<span data-ttu-id="58d85-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="58d85-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="58d85-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-151">Member</span></span> |
| [<span data-ttu-id="58d85-152">organizer</span><span class="sxs-lookup"><span data-stu-id="58d85-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="58d85-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-153">Member</span></span> |
| [<span data-ttu-id="58d85-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="58d85-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="58d85-155">Member</span><span class="sxs-lookup"><span data-stu-id="58d85-155">Member</span></span> |
| [<span data-ttu-id="58d85-156">sender</span><span class="sxs-lookup"><span data-stu-id="58d85-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="58d85-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-157">Member</span></span> |
| [<span data-ttu-id="58d85-158">start</span><span class="sxs-lookup"><span data-stu-id="58d85-158">start</span></span>](#start-datetime) | <span data-ttu-id="58d85-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-159">Member</span></span> |
| [<span data-ttu-id="58d85-160">subject</span><span class="sxs-lookup"><span data-stu-id="58d85-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="58d85-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-161">Member</span></span> |
| [<span data-ttu-id="58d85-162">to</span><span class="sxs-lookup"><span data-stu-id="58d85-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="58d85-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-163">Member</span></span> |
| [<span data-ttu-id="58d85-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="58d85-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="58d85-165">メソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-165">Method</span></span> |
| [<span data-ttu-id="58d85-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="58d85-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="58d85-167">メソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-167">Method</span></span> |
| [<span data-ttu-id="58d85-168">close</span><span class="sxs-lookup"><span data-stu-id="58d85-168">close</span></span>](#close) | <span data-ttu-id="58d85-169">メソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-169">Method</span></span> |
| [<span data-ttu-id="58d85-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="58d85-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="58d85-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-171">Method</span></span> |
| [<span data-ttu-id="58d85-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="58d85-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="58d85-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-173">Method</span></span> |
| [<span data-ttu-id="58d85-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="58d85-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="58d85-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-175">Method</span></span> |
| [<span data-ttu-id="58d85-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="58d85-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="58d85-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-177">Method</span></span> |
| [<span data-ttu-id="58d85-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="58d85-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="58d85-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-179">Method</span></span> |
| [<span data-ttu-id="58d85-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="58d85-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="58d85-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-181">Method</span></span> |
| [<span data-ttu-id="58d85-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="58d85-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="58d85-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-183">Method</span></span> |
| [<span data-ttu-id="58d85-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="58d85-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="58d85-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-185">Method</span></span> |
| [<span data-ttu-id="58d85-186">office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="58d85-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="58d85-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-187">Method</span></span> |
| [<span data-ttu-id="58d85-188">office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="58d85-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="58d85-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-189">Method</span></span> |
| [<span data-ttu-id="58d85-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="58d85-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="58d85-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-191">Method</span></span> |
| [<span data-ttu-id="58d85-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="58d85-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="58d85-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-193">Method</span></span> |
| [<span data-ttu-id="58d85-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="58d85-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="58d85-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-195">Method</span></span> |
| [<span data-ttu-id="58d85-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="58d85-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="58d85-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="58d85-198">例</span><span class="sxs-lookup"><span data-stu-id="58d85-198">Example</span></span>

<span data-ttu-id="58d85-199">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="58d85-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
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

### <a name="members"></a><span data-ttu-id="58d85-200">メンバー</span><span class="sxs-lookup"><span data-stu-id="58d85-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="58d85-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="58d85-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="58d85-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="58d85-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="58d85-204">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="58d85-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="58d85-205">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="58d85-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="58d85-206">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-206">Type</span></span>

*   <span data-ttu-id="58d85-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="58d85-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-208">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-208">Requirements</span></span>

|<span data-ttu-id="58d85-209">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-209">Requirement</span></span>| <span data-ttu-id="58d85-210">値</span><span class="sxs-lookup"><span data-stu-id="58d85-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-211">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-212">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-212">1.0</span></span>|
|[<span data-ttu-id="58d85-213">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-214">ReadItem</span></span>|
|[<span data-ttu-id="58d85-215">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-216">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58d85-217">例</span><span class="sxs-lookup"><span data-stu-id="58d85-217">Example</span></span>

<span data-ttu-id="58d85-218">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="58d85-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="58d85-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="58d85-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="58d85-220">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="58d85-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="58d85-221">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="58d85-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="58d85-222">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-222">Type</span></span>

*   [<span data-ttu-id="58d85-223">受信者</span><span class="sxs-lookup"><span data-stu-id="58d85-223">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="58d85-224">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-224">Requirements</span></span>

|<span data-ttu-id="58d85-225">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-225">Requirement</span></span>| <span data-ttu-id="58d85-226">値</span><span class="sxs-lookup"><span data-stu-id="58d85-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-227">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-228">1.1</span><span class="sxs-lookup"><span data-stu-id="58d85-228">1.1</span></span>|
|[<span data-ttu-id="58d85-229">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-230">ReadItem</span></span>|
|[<span data-ttu-id="58d85-231">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-232">作成</span><span class="sxs-lookup"><span data-stu-id="58d85-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="58d85-233">例</span><span class="sxs-lookup"><span data-stu-id="58d85-233">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="58d85-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="58d85-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="58d85-235">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="58d85-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="58d85-236">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-236">Type</span></span>

*   [<span data-ttu-id="58d85-237">Body</span><span class="sxs-lookup"><span data-stu-id="58d85-237">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="58d85-238">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-238">Requirements</span></span>

|<span data-ttu-id="58d85-239">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-239">Requirement</span></span>| <span data-ttu-id="58d85-240">値</span><span class="sxs-lookup"><span data-stu-id="58d85-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-241">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-242">1.1</span><span class="sxs-lookup"><span data-stu-id="58d85-242">1.1</span></span>|
|[<span data-ttu-id="58d85-243">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-244">ReadItem</span></span>|
|[<span data-ttu-id="58d85-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-246">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58d85-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58d85-247">例</span><span class="sxs-lookup"><span data-stu-id="58d85-247">Example</span></span>

<span data-ttu-id="58d85-248">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="58d85-248">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="58d85-249">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="58d85-249">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="58d85-250">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="58d85-250">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="58d85-251">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="58d85-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="58d85-252">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="58d85-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="58d85-253">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="58d85-253">Read mode</span></span>

<span data-ttu-id="58d85-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="58d85-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="58d85-256">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="58d85-256">Compose mode</span></span>

<span data-ttu-id="58d85-257">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-257">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="58d85-258">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-258">Type</span></span>

*   <span data-ttu-id="58d85-259">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="58d85-259">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-260">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-260">Requirements</span></span>

|<span data-ttu-id="58d85-261">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-261">Requirement</span></span>| <span data-ttu-id="58d85-262">値</span><span class="sxs-lookup"><span data-stu-id="58d85-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-263">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-264">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-264">1.0</span></span>|
|[<span data-ttu-id="58d85-265">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-266">ReadItem</span></span>|
|[<span data-ttu-id="58d85-267">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-268">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58d85-268">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="58d85-269">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="58d85-269">(nullable) conversationId :String</span></span>

<span data-ttu-id="58d85-270">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="58d85-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="58d85-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="58d85-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="58d85-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="58d85-275">Type</span><span class="sxs-lookup"><span data-stu-id="58d85-275">Type</span></span>

*   <span data-ttu-id="58d85-276">String</span><span class="sxs-lookup"><span data-stu-id="58d85-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-277">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-277">Requirements</span></span>

|<span data-ttu-id="58d85-278">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-278">Requirement</span></span>| <span data-ttu-id="58d85-279">値</span><span class="sxs-lookup"><span data-stu-id="58d85-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-280">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-281">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-281">1.0</span></span>|
|[<span data-ttu-id="58d85-282">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-283">ReadItem</span></span>|
|[<span data-ttu-id="58d85-284">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-285">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58d85-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58d85-286">例</span><span class="sxs-lookup"><span data-stu-id="58d85-286">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="58d85-287">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="58d85-287">dateTimeCreated :Date</span></span>

<span data-ttu-id="58d85-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="58d85-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="58d85-290">型</span><span class="sxs-lookup"><span data-stu-id="58d85-290">Type</span></span>

*   <span data-ttu-id="58d85-291">日付</span><span class="sxs-lookup"><span data-stu-id="58d85-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-292">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-292">Requirements</span></span>

|<span data-ttu-id="58d85-293">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-293">Requirement</span></span>| <span data-ttu-id="58d85-294">値</span><span class="sxs-lookup"><span data-stu-id="58d85-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-295">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-296">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-296">1.0</span></span>|
|[<span data-ttu-id="58d85-297">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-298">ReadItem</span></span>|
|[<span data-ttu-id="58d85-299">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-300">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58d85-301">例</span><span class="sxs-lookup"><span data-stu-id="58d85-301">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="58d85-302">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="58d85-302">dateTimeModified :Date</span></span>

<span data-ttu-id="58d85-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="58d85-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="58d85-305">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="58d85-305">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="58d85-306">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-306">Type</span></span>

*   <span data-ttu-id="58d85-307">日付</span><span class="sxs-lookup"><span data-stu-id="58d85-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-308">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-308">Requirements</span></span>

|<span data-ttu-id="58d85-309">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-309">Requirement</span></span>| <span data-ttu-id="58d85-310">値</span><span class="sxs-lookup"><span data-stu-id="58d85-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-312">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-312">1.0</span></span>|
|[<span data-ttu-id="58d85-313">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-314">ReadItem</span></span>|
|[<span data-ttu-id="58d85-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-316">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58d85-317">例</span><span class="sxs-lookup"><span data-stu-id="58d85-317">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="58d85-318">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="58d85-318">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="58d85-319">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="58d85-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="58d85-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="58d85-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="58d85-322">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="58d85-322">Read mode</span></span>

<span data-ttu-id="58d85-323">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-323">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="58d85-324">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="58d85-324">Compose mode</span></span>

<span data-ttu-id="58d85-325">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="58d85-326">[`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="58d85-326">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="58d85-327">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="58d85-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

##### <a name="type"></a><span data-ttu-id="58d85-328">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-328">Type</span></span>

*   <span data-ttu-id="58d85-329">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="58d85-329">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-330">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-330">Requirements</span></span>

|<span data-ttu-id="58d85-331">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-331">Requirement</span></span>| <span data-ttu-id="58d85-332">値</span><span class="sxs-lookup"><span data-stu-id="58d85-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-333">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-334">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-334">1.0</span></span>|
|[<span data-ttu-id="58d85-335">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-336">ReadItem</span></span>|
|[<span data-ttu-id="58d85-337">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-338">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58d85-338">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="58d85-339">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="58d85-339">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="58d85-p112">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="58d85-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="58d85-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="58d85-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="58d85-344">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="58d85-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="58d85-345">型</span><span class="sxs-lookup"><span data-stu-id="58d85-345">Type</span></span>

*   [<span data-ttu-id="58d85-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="58d85-346">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="example"></a><span data-ttu-id="58d85-347">例</span><span class="sxs-lookup"><span data-stu-id="58d85-347">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="58d85-348">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-348">Requirements</span></span>

|<span data-ttu-id="58d85-349">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-349">Requirement</span></span>| <span data-ttu-id="58d85-350">値</span><span class="sxs-lookup"><span data-stu-id="58d85-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-351">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-352">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-352">1.0</span></span>|
|[<span data-ttu-id="58d85-353">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-354">ReadItem</span></span>|
|[<span data-ttu-id="58d85-355">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-356">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-356">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="58d85-357">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="58d85-357">internetMessageId :String</span></span>

<span data-ttu-id="58d85-p114">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="58d85-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="58d85-360">Type</span><span class="sxs-lookup"><span data-stu-id="58d85-360">Type</span></span>

*   <span data-ttu-id="58d85-361">String</span><span class="sxs-lookup"><span data-stu-id="58d85-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-362">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-362">Requirements</span></span>

|<span data-ttu-id="58d85-363">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-363">Requirement</span></span>| <span data-ttu-id="58d85-364">値</span><span class="sxs-lookup"><span data-stu-id="58d85-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-365">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-366">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-366">1.0</span></span>|
|[<span data-ttu-id="58d85-367">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-368">ReadItem</span></span>|
|[<span data-ttu-id="58d85-369">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-370">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58d85-371">例</span><span class="sxs-lookup"><span data-stu-id="58d85-371">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="58d85-372">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="58d85-372">itemClass :String</span></span>

<span data-ttu-id="58d85-p115">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="58d85-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="58d85-p116">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="58d85-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="58d85-377">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-377">Type</span></span> | <span data-ttu-id="58d85-378">説明</span><span class="sxs-lookup"><span data-stu-id="58d85-378">Description</span></span> | <span data-ttu-id="58d85-379">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="58d85-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="58d85-380">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="58d85-380">Appointment items</span></span> | <span data-ttu-id="58d85-381">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="58d85-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="58d85-382">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="58d85-382">Message items</span></span> | <span data-ttu-id="58d85-383">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="58d85-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="58d85-384">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="58d85-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="58d85-385">Type</span><span class="sxs-lookup"><span data-stu-id="58d85-385">Type</span></span>

*   <span data-ttu-id="58d85-386">String</span><span class="sxs-lookup"><span data-stu-id="58d85-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-387">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-387">Requirements</span></span>

|<span data-ttu-id="58d85-388">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-388">Requirement</span></span>| <span data-ttu-id="58d85-389">値</span><span class="sxs-lookup"><span data-stu-id="58d85-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-390">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-391">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-391">1.0</span></span>|
|[<span data-ttu-id="58d85-392">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-393">ReadItem</span></span>|
|[<span data-ttu-id="58d85-394">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-395">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58d85-396">例</span><span class="sxs-lookup"><span data-stu-id="58d85-396">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="58d85-397">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="58d85-397">(nullable) itemId :String</span></span>

<span data-ttu-id="58d85-p117">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="58d85-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="58d85-400">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="58d85-400">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="58d85-401">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="58d85-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="58d85-402">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="58d85-402">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="58d85-403">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="58d85-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="58d85-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="58d85-406">型</span><span class="sxs-lookup"><span data-stu-id="58d85-406">Type</span></span>

*   <span data-ttu-id="58d85-407">String</span><span class="sxs-lookup"><span data-stu-id="58d85-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-408">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-408">Requirements</span></span>

|<span data-ttu-id="58d85-409">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-409">Requirement</span></span>| <span data-ttu-id="58d85-410">値</span><span class="sxs-lookup"><span data-stu-id="58d85-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-411">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-412">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-412">1.0</span></span>|
|[<span data-ttu-id="58d85-413">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-414">ReadItem</span></span>|
|[<span data-ttu-id="58d85-415">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-416">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58d85-417">例</span><span class="sxs-lookup"><span data-stu-id="58d85-417">Example</span></span>

<span data-ttu-id="58d85-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="58d85-420">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="58d85-420">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="58d85-421">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="58d85-421">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="58d85-422">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="58d85-422">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="58d85-423">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-423">Type</span></span>

*   [<span data-ttu-id="58d85-424">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="58d85-424">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="58d85-425">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-425">Requirements</span></span>

|<span data-ttu-id="58d85-426">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-426">Requirement</span></span>| <span data-ttu-id="58d85-427">値</span><span class="sxs-lookup"><span data-stu-id="58d85-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-428">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-429">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-429">1.0</span></span>|
|[<span data-ttu-id="58d85-430">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-431">ReadItem</span></span>|
|[<span data-ttu-id="58d85-432">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-433">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58d85-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58d85-434">例</span><span class="sxs-lookup"><span data-stu-id="58d85-434">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="58d85-435">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="58d85-435">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="58d85-436">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="58d85-436">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="58d85-437">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="58d85-437">Read mode</span></span>

<span data-ttu-id="58d85-438">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-438">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="58d85-439">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="58d85-439">Compose mode</span></span>

<span data-ttu-id="58d85-440">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-440">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="58d85-441">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-441">Type</span></span>

*   <span data-ttu-id="58d85-442">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="58d85-442">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-443">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-443">Requirements</span></span>

|<span data-ttu-id="58d85-444">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-444">Requirement</span></span>| <span data-ttu-id="58d85-445">値</span><span class="sxs-lookup"><span data-stu-id="58d85-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-446">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-447">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-447">1.0</span></span>|
|[<span data-ttu-id="58d85-448">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-448">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-449">ReadItem</span></span>|
|[<span data-ttu-id="58d85-450">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-450">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-451">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58d85-451">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="58d85-452">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="58d85-452">normalizedSubject :String</span></span>

<span data-ttu-id="58d85-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="58d85-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="58d85-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="58d85-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="58d85-457">Type</span><span class="sxs-lookup"><span data-stu-id="58d85-457">Type</span></span>

*   <span data-ttu-id="58d85-458">String</span><span class="sxs-lookup"><span data-stu-id="58d85-458">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-459">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-459">Requirements</span></span>

|<span data-ttu-id="58d85-460">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-460">Requirement</span></span>| <span data-ttu-id="58d85-461">値</span><span class="sxs-lookup"><span data-stu-id="58d85-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-462">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-463">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-463">1.0</span></span>|
|[<span data-ttu-id="58d85-464">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-465">ReadItem</span></span>|
|[<span data-ttu-id="58d85-466">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-467">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58d85-468">例</span><span class="sxs-lookup"><span data-stu-id="58d85-468">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="58d85-469">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="58d85-469">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="58d85-470">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="58d85-470">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="58d85-471">型</span><span class="sxs-lookup"><span data-stu-id="58d85-471">Type</span></span>

*   [<span data-ttu-id="58d85-472">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="58d85-472">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="58d85-473">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-473">Requirements</span></span>

|<span data-ttu-id="58d85-474">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-474">Requirement</span></span>| <span data-ttu-id="58d85-475">値</span><span class="sxs-lookup"><span data-stu-id="58d85-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-476">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-476">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-477">1.3</span><span class="sxs-lookup"><span data-stu-id="58d85-477">1.3</span></span>|
|[<span data-ttu-id="58d85-478">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-478">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-479">ReadItem</span></span>|
|[<span data-ttu-id="58d85-480">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-480">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-481">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58d85-481">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58d85-482">例</span><span class="sxs-lookup"><span data-stu-id="58d85-482">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="58d85-483">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="58d85-483">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="58d85-484">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="58d85-484">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="58d85-485">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="58d85-485">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="58d85-486">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="58d85-486">Read mode</span></span>

<span data-ttu-id="58d85-487">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-487">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="58d85-488">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="58d85-488">Compose mode</span></span>

<span data-ttu-id="58d85-489">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-489">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="58d85-490">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-490">Type</span></span>

*   <span data-ttu-id="58d85-491">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="58d85-491">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-492">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-492">Requirements</span></span>

|<span data-ttu-id="58d85-493">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-493">Requirement</span></span>| <span data-ttu-id="58d85-494">値</span><span class="sxs-lookup"><span data-stu-id="58d85-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-495">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-496">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-496">1.0</span></span>|
|[<span data-ttu-id="58d85-497">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-498">ReadItem</span></span>|
|[<span data-ttu-id="58d85-499">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-500">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58d85-500">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="58d85-501">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="58d85-501">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="58d85-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="58d85-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="58d85-504">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-504">Type</span></span>

*   [<span data-ttu-id="58d85-505">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="58d85-505">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="58d85-506">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-506">Requirements</span></span>

|<span data-ttu-id="58d85-507">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-507">Requirement</span></span>| <span data-ttu-id="58d85-508">値</span><span class="sxs-lookup"><span data-stu-id="58d85-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-509">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-510">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-510">1.0</span></span>|
|[<span data-ttu-id="58d85-511">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-512">ReadItem</span></span>|
|[<span data-ttu-id="58d85-513">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-514">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-514">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58d85-515">例</span><span class="sxs-lookup"><span data-stu-id="58d85-515">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="58d85-516">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="58d85-516">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="58d85-517">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="58d85-517">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="58d85-518">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="58d85-518">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="58d85-519">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="58d85-519">Read mode</span></span>

<span data-ttu-id="58d85-520">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-520">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="58d85-521">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="58d85-521">Compose mode</span></span>

<span data-ttu-id="58d85-522">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-522">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="58d85-523">型</span><span class="sxs-lookup"><span data-stu-id="58d85-523">Type</span></span>

*   <span data-ttu-id="58d85-524">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="58d85-524">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-525">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-525">Requirements</span></span>

|<span data-ttu-id="58d85-526">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-526">Requirement</span></span>| <span data-ttu-id="58d85-527">値</span><span class="sxs-lookup"><span data-stu-id="58d85-527">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-528">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-528">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-529">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-529">1.0</span></span>|
|[<span data-ttu-id="58d85-530">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-530">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-531">ReadItem</span></span>|
|[<span data-ttu-id="58d85-532">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-532">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-533">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58d85-533">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="58d85-534">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="58d85-534">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="58d85-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="58d85-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="58d85-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="58d85-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="58d85-539">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="58d85-539">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="58d85-540">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-540">Type</span></span>

*   [<span data-ttu-id="58d85-541">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="58d85-541">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="58d85-542">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-542">Requirements</span></span>

|<span data-ttu-id="58d85-543">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-543">Requirement</span></span>| <span data-ttu-id="58d85-544">値</span><span class="sxs-lookup"><span data-stu-id="58d85-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-545">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-546">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-546">1.0</span></span>|
|[<span data-ttu-id="58d85-547">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-548">ReadItem</span></span>|
|[<span data-ttu-id="58d85-549">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-550">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-550">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58d85-551">例</span><span class="sxs-lookup"><span data-stu-id="58d85-551">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="58d85-552">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="58d85-552">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="58d85-553">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="58d85-553">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="58d85-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="58d85-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="58d85-556">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="58d85-556">Read mode</span></span>

<span data-ttu-id="58d85-557">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-557">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="58d85-558">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="58d85-558">Compose mode</span></span>

<span data-ttu-id="58d85-559">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-559">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="58d85-560">[`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="58d85-560">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="58d85-561">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="58d85-561">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

##### <a name="type"></a><span data-ttu-id="58d85-562">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-562">Type</span></span>

*   <span data-ttu-id="58d85-563">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="58d85-563">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-564">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-564">Requirements</span></span>

|<span data-ttu-id="58d85-565">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-565">Requirement</span></span>| <span data-ttu-id="58d85-566">値</span><span class="sxs-lookup"><span data-stu-id="58d85-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-567">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-568">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-568">1.0</span></span>|
|[<span data-ttu-id="58d85-569">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-570">ReadItem</span></span>|
|[<span data-ttu-id="58d85-571">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-572">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58d85-572">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="58d85-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="58d85-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="58d85-574">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="58d85-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="58d85-575">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="58d85-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="58d85-576">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="58d85-576">Read mode</span></span>

<span data-ttu-id="58d85-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="58d85-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="58d85-579">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="58d85-579">Compose mode</span></span>

<span data-ttu-id="58d85-580">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="58d85-581">型</span><span class="sxs-lookup"><span data-stu-id="58d85-581">Type</span></span>

*   <span data-ttu-id="58d85-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="58d85-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-583">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-583">Requirements</span></span>

|<span data-ttu-id="58d85-584">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-584">Requirement</span></span>| <span data-ttu-id="58d85-585">値</span><span class="sxs-lookup"><span data-stu-id="58d85-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-586">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-587">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-587">1.0</span></span>|
|[<span data-ttu-id="58d85-588">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-588">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-589">ReadItem</span></span>|
|[<span data-ttu-id="58d85-590">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-590">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-591">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58d85-591">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="58d85-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="58d85-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="58d85-593">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="58d85-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="58d85-594">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="58d85-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="58d85-595">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="58d85-595">Read mode</span></span>

<span data-ttu-id="58d85-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="58d85-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="58d85-598">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="58d85-598">Compose mode</span></span>

<span data-ttu-id="58d85-599">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="58d85-600">型</span><span class="sxs-lookup"><span data-stu-id="58d85-600">Type</span></span>

*   <span data-ttu-id="58d85-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="58d85-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-602">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-602">Requirements</span></span>

|<span data-ttu-id="58d85-603">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-603">Requirement</span></span>| <span data-ttu-id="58d85-604">値</span><span class="sxs-lookup"><span data-stu-id="58d85-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-605">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-606">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-606">1.0</span></span>|
|[<span data-ttu-id="58d85-607">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-608">ReadItem</span></span>|
|[<span data-ttu-id="58d85-609">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-610">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58d85-610">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="58d85-611">メソッド</span><span class="sxs-lookup"><span data-stu-id="58d85-611">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="58d85-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="58d85-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="58d85-613">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="58d85-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="58d85-614">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="58d85-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="58d85-615">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="58d85-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58d85-616">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58d85-616">Parameters</span></span>

|<span data-ttu-id="58d85-617">名前</span><span class="sxs-lookup"><span data-stu-id="58d85-617">Name</span></span>| <span data-ttu-id="58d85-618">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-618">Type</span></span>| <span data-ttu-id="58d85-619">属性</span><span class="sxs-lookup"><span data-stu-id="58d85-619">Attributes</span></span>| <span data-ttu-id="58d85-620">説明</span><span class="sxs-lookup"><span data-stu-id="58d85-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="58d85-621">String</span><span class="sxs-lookup"><span data-stu-id="58d85-621">String</span></span>||<span data-ttu-id="58d85-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="58d85-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="58d85-624">String</span><span class="sxs-lookup"><span data-stu-id="58d85-624">String</span></span>||<span data-ttu-id="58d85-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="58d85-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="58d85-627">Object</span><span class="sxs-lookup"><span data-stu-id="58d85-627">Object</span></span>| <span data-ttu-id="58d85-628">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-628">&lt;optional&gt;</span></span>|<span data-ttu-id="58d85-629">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="58d85-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="58d85-630">Object</span><span class="sxs-lookup"><span data-stu-id="58d85-630">Object</span></span> | <span data-ttu-id="58d85-631">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-631">&lt;optional&gt;</span></span> | <span data-ttu-id="58d85-632">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="58d85-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="58d85-633">Boolean</span><span class="sxs-lookup"><span data-stu-id="58d85-633">Boolean</span></span> | <span data-ttu-id="58d85-634">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-634">&lt;optional&gt;</span></span> | <span data-ttu-id="58d85-635">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="58d85-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="58d85-636">function</span><span class="sxs-lookup"><span data-stu-id="58d85-636">function</span></span>| <span data-ttu-id="58d85-637">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-637">&lt;optional&gt;</span></span>|<span data-ttu-id="58d85-638">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="58d85-639">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="58d85-640">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="58d85-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="58d85-641">エラー</span><span class="sxs-lookup"><span data-stu-id="58d85-641">Errors</span></span>

| <span data-ttu-id="58d85-642">エラー コード</span><span class="sxs-lookup"><span data-stu-id="58d85-642">Error code</span></span> | <span data-ttu-id="58d85-643">説明</span><span class="sxs-lookup"><span data-stu-id="58d85-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="58d85-644">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="58d85-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="58d85-645">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="58d85-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="58d85-646">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="58d85-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="58d85-647">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-647">Requirements</span></span>

|<span data-ttu-id="58d85-648">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-648">Requirement</span></span>| <span data-ttu-id="58d85-649">値</span><span class="sxs-lookup"><span data-stu-id="58d85-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-650">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-651">1.1</span><span class="sxs-lookup"><span data-stu-id="58d85-651">1.1</span></span>|
|[<span data-ttu-id="58d85-652">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="58d85-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="58d85-654">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-655">作成</span><span class="sxs-lookup"><span data-stu-id="58d85-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="58d85-656">例</span><span class="sxs-lookup"><span data-stu-id="58d85-656">Examples</span></span>

```javascript
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

<span data-ttu-id="58d85-657">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="58d85-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```javascript
Office.context.mailbox.item.addFileAttachmentAsync(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        // Do something here.
      });
  });
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="58d85-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="58d85-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="58d85-659">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="58d85-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="58d85-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="58d85-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="58d85-663">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="58d85-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="58d85-664">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="58d85-664">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58d85-665">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58d85-665">Parameters</span></span>

|<span data-ttu-id="58d85-666">名前</span><span class="sxs-lookup"><span data-stu-id="58d85-666">Name</span></span>| <span data-ttu-id="58d85-667">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-667">Type</span></span>| <span data-ttu-id="58d85-668">属性</span><span class="sxs-lookup"><span data-stu-id="58d85-668">Attributes</span></span>| <span data-ttu-id="58d85-669">説明</span><span class="sxs-lookup"><span data-stu-id="58d85-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="58d85-670">String</span><span class="sxs-lookup"><span data-stu-id="58d85-670">String</span></span>||<span data-ttu-id="58d85-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="58d85-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="58d85-673">String</span><span class="sxs-lookup"><span data-stu-id="58d85-673">String</span></span>||<span data-ttu-id="58d85-674">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="58d85-674">The subject of the item to be attached.</span></span> <span data-ttu-id="58d85-675">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="58d85-675">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="58d85-676">Object</span><span class="sxs-lookup"><span data-stu-id="58d85-676">Object</span></span>| <span data-ttu-id="58d85-677">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-677">&lt;optional&gt;</span></span>|<span data-ttu-id="58d85-678">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="58d85-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="58d85-679">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="58d85-679">Object</span></span>| <span data-ttu-id="58d85-680">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-680">&lt;optional&gt;</span></span>|<span data-ttu-id="58d85-681">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="58d85-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="58d85-682">関数</span><span class="sxs-lookup"><span data-stu-id="58d85-682">function</span></span>| <span data-ttu-id="58d85-683">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-683">&lt;optional&gt;</span></span>|<span data-ttu-id="58d85-684">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="58d85-685">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="58d85-686">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="58d85-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="58d85-687">エラー</span><span class="sxs-lookup"><span data-stu-id="58d85-687">Errors</span></span>

| <span data-ttu-id="58d85-688">エラー コード</span><span class="sxs-lookup"><span data-stu-id="58d85-688">Error code</span></span> | <span data-ttu-id="58d85-689">説明</span><span class="sxs-lookup"><span data-stu-id="58d85-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="58d85-690">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="58d85-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="58d85-691">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-691">Requirements</span></span>

|<span data-ttu-id="58d85-692">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-692">Requirement</span></span>| <span data-ttu-id="58d85-693">値</span><span class="sxs-lookup"><span data-stu-id="58d85-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-694">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-695">1.1</span><span class="sxs-lookup"><span data-stu-id="58d85-695">1.1</span></span>|
|[<span data-ttu-id="58d85-696">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="58d85-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="58d85-698">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-699">作成</span><span class="sxs-lookup"><span data-stu-id="58d85-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="58d85-700">例</span><span class="sxs-lookup"><span data-stu-id="58d85-700">Example</span></span>

<span data-ttu-id="58d85-701">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```javascript
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

####  <a name="close"></a><span data-ttu-id="58d85-702">close()</span><span class="sxs-lookup"><span data-stu-id="58d85-702">close()</span></span>

<span data-ttu-id="58d85-703">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="58d85-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="58d85-p137">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="58d85-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="58d85-706">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="58d85-707">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="58d85-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-708">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-708">Requirements</span></span>

|<span data-ttu-id="58d85-709">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-709">Requirement</span></span>| <span data-ttu-id="58d85-710">値</span><span class="sxs-lookup"><span data-stu-id="58d85-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-711">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-712">1.3</span><span class="sxs-lookup"><span data-stu-id="58d85-712">1.3</span></span>|
|[<span data-ttu-id="58d85-713">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-714">制限あり</span><span class="sxs-lookup"><span data-stu-id="58d85-714">Restricted</span></span>|
|[<span data-ttu-id="58d85-715">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-716">新規作成</span><span class="sxs-lookup"><span data-stu-id="58d85-716">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="58d85-717">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="58d85-717">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="58d85-718">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="58d85-719">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="58d85-719">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="58d85-720">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="58d85-721">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="58d85-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="58d85-p138">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="58d85-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58d85-725">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58d85-725">Parameters</span></span>

| <span data-ttu-id="58d85-726">名前</span><span class="sxs-lookup"><span data-stu-id="58d85-726">Name</span></span> | <span data-ttu-id="58d85-727">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-727">Type</span></span> | <span data-ttu-id="58d85-728">属性</span><span class="sxs-lookup"><span data-stu-id="58d85-728">Attributes</span></span> | <span data-ttu-id="58d85-729">説明</span><span class="sxs-lookup"><span data-stu-id="58d85-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="58d85-730">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="58d85-730">String &#124; Object</span></span>| |<span data-ttu-id="58d85-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="58d85-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="58d85-733">**または**</span><span class="sxs-lookup"><span data-stu-id="58d85-733">**OR**</span></span><br/><span data-ttu-id="58d85-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="58d85-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="58d85-736">String</span><span class="sxs-lookup"><span data-stu-id="58d85-736">String</span></span> | <span data-ttu-id="58d85-737">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-737">&lt;optional&gt;</span></span> | <span data-ttu-id="58d85-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="58d85-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="58d85-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="58d85-741">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-741">&lt;optional&gt;</span></span> | <span data-ttu-id="58d85-742">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="58d85-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="58d85-743">String</span><span class="sxs-lookup"><span data-stu-id="58d85-743">String</span></span> | | <span data-ttu-id="58d85-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="58d85-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="58d85-746">String</span><span class="sxs-lookup"><span data-stu-id="58d85-746">String</span></span> | | <span data-ttu-id="58d85-747">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="58d85-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="58d85-748">文字列</span><span class="sxs-lookup"><span data-stu-id="58d85-748">String</span></span> | | <span data-ttu-id="58d85-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="58d85-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="58d85-751">ブール値</span><span class="sxs-lookup"><span data-stu-id="58d85-751">Boolean</span></span> | | <span data-ttu-id="58d85-p144">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="58d85-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="58d85-754">String</span><span class="sxs-lookup"><span data-stu-id="58d85-754">String</span></span> | | <span data-ttu-id="58d85-p145">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="58d85-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="58d85-758">function</span><span class="sxs-lookup"><span data-stu-id="58d85-758">function</span></span> | <span data-ttu-id="58d85-759">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-759">&lt;optional&gt;</span></span> | <span data-ttu-id="58d85-760">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="58d85-761">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-761">Requirements</span></span>

|<span data-ttu-id="58d85-762">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-762">Requirement</span></span>| <span data-ttu-id="58d85-763">値</span><span class="sxs-lookup"><span data-stu-id="58d85-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-764">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-765">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-765">1.0</span></span>|
|[<span data-ttu-id="58d85-766">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-767">ReadItem</span></span>|
|[<span data-ttu-id="58d85-768">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-769">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="58d85-770">例</span><span class="sxs-lookup"><span data-stu-id="58d85-770">Examples</span></span>

<span data-ttu-id="58d85-771">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="58d85-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="58d85-772">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="58d85-772">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="58d85-773">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="58d85-773">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="58d85-774">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="58d85-774">Reply with a body and a file attachment.</span></span>

```javascript
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

<span data-ttu-id="58d85-775">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="58d85-775">Reply with a body and an item attachment.</span></span>

```javascript
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

<span data-ttu-id="58d85-776">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="58d85-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```javascript
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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="58d85-777">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="58d85-777">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="58d85-778">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="58d85-779">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="58d85-779">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="58d85-780">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="58d85-781">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="58d85-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="58d85-p146">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="58d85-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58d85-785">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58d85-785">Parameters</span></span>

| <span data-ttu-id="58d85-786">名前</span><span class="sxs-lookup"><span data-stu-id="58d85-786">Name</span></span> | <span data-ttu-id="58d85-787">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-787">Type</span></span> | <span data-ttu-id="58d85-788">属性</span><span class="sxs-lookup"><span data-stu-id="58d85-788">Attributes</span></span> | <span data-ttu-id="58d85-789">説明</span><span class="sxs-lookup"><span data-stu-id="58d85-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="58d85-790">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="58d85-790">String &#124; Object</span></span>| | <span data-ttu-id="58d85-p147">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="58d85-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="58d85-793">**または**</span><span class="sxs-lookup"><span data-stu-id="58d85-793">**OR**</span></span><br/><span data-ttu-id="58d85-p148">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="58d85-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="58d85-796">String</span><span class="sxs-lookup"><span data-stu-id="58d85-796">String</span></span> | <span data-ttu-id="58d85-797">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-797">&lt;optional&gt;</span></span> | <span data-ttu-id="58d85-p149">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="58d85-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="58d85-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="58d85-801">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-801">&lt;optional&gt;</span></span> | <span data-ttu-id="58d85-802">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="58d85-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="58d85-803">String</span><span class="sxs-lookup"><span data-stu-id="58d85-803">String</span></span> | | <span data-ttu-id="58d85-p150">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="58d85-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="58d85-806">String</span><span class="sxs-lookup"><span data-stu-id="58d85-806">String</span></span> | | <span data-ttu-id="58d85-807">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="58d85-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="58d85-808">文字列</span><span class="sxs-lookup"><span data-stu-id="58d85-808">String</span></span> | | <span data-ttu-id="58d85-p151">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="58d85-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="58d85-811">ブール値</span><span class="sxs-lookup"><span data-stu-id="58d85-811">Boolean</span></span> | | <span data-ttu-id="58d85-p152">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="58d85-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="58d85-814">String</span><span class="sxs-lookup"><span data-stu-id="58d85-814">String</span></span> | | <span data-ttu-id="58d85-p153">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="58d85-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="58d85-818">function</span><span class="sxs-lookup"><span data-stu-id="58d85-818">function</span></span> | <span data-ttu-id="58d85-819">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-819">&lt;optional&gt;</span></span> | <span data-ttu-id="58d85-820">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="58d85-821">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-821">Requirements</span></span>

|<span data-ttu-id="58d85-822">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-822">Requirement</span></span>| <span data-ttu-id="58d85-823">値</span><span class="sxs-lookup"><span data-stu-id="58d85-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-824">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-825">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-825">1.0</span></span>|
|[<span data-ttu-id="58d85-826">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-826">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-827">ReadItem</span></span>|
|[<span data-ttu-id="58d85-828">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-828">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-829">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="58d85-830">例</span><span class="sxs-lookup"><span data-stu-id="58d85-830">Examples</span></span>

<span data-ttu-id="58d85-831">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="58d85-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="58d85-832">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="58d85-832">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="58d85-833">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="58d85-833">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="58d85-834">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="58d85-834">Reply with a body and a file attachment.</span></span>

```javascript
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

<span data-ttu-id="58d85-835">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="58d85-835">Reply with a body and an item attachment.</span></span>

```javascript
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

<span data-ttu-id="58d85-836">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="58d85-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```javascript
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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="58d85-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="58d85-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="58d85-838">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="58d85-838">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="58d85-839">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="58d85-839">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-840">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-840">Requirements</span></span>

|<span data-ttu-id="58d85-841">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-841">Requirement</span></span>| <span data-ttu-id="58d85-842">値</span><span class="sxs-lookup"><span data-stu-id="58d85-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-843">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-844">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-844">1.0</span></span>|
|[<span data-ttu-id="58d85-845">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-846">ReadItem</span></span>|
|[<span data-ttu-id="58d85-847">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-848">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="58d85-849">戻り値:</span><span class="sxs-lookup"><span data-stu-id="58d85-849">Returns:</span></span>

<span data-ttu-id="58d85-850">型:[Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="58d85-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="58d85-851">例</span><span class="sxs-lookup"><span data-stu-id="58d85-851">Example</span></span>

<span data-ttu-id="58d85-852">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="58d85-852">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="58d85-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="58d85-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="58d85-854">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="58d85-854">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="58d85-855">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="58d85-855">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58d85-856">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58d85-856">Parameters</span></span>

|<span data-ttu-id="58d85-857">名前</span><span class="sxs-lookup"><span data-stu-id="58d85-857">Name</span></span>| <span data-ttu-id="58d85-858">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-858">Type</span></span>| <span data-ttu-id="58d85-859">説明</span><span class="sxs-lookup"><span data-stu-id="58d85-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="58d85-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="58d85-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="58d85-861">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="58d85-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="58d85-862">Requirements</span><span class="sxs-lookup"><span data-stu-id="58d85-862">Requirements</span></span>

|<span data-ttu-id="58d85-863">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-863">Requirement</span></span>| <span data-ttu-id="58d85-864">値</span><span class="sxs-lookup"><span data-stu-id="58d85-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-865">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-866">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-866">1.0</span></span>|
|[<span data-ttu-id="58d85-867">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-867">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-868">制限あり</span><span class="sxs-lookup"><span data-stu-id="58d85-868">Restricted</span></span>|
|[<span data-ttu-id="58d85-869">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-869">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-870">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="58d85-871">戻り値:</span><span class="sxs-lookup"><span data-stu-id="58d85-871">Returns:</span></span>

<span data-ttu-id="58d85-872">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="58d85-873">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-873">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="58d85-874">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="58d85-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="58d85-875">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="58d85-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="58d85-876">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="58d85-876">Value of `entityType`</span></span> | <span data-ttu-id="58d85-877">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="58d85-877">Type of objects in returned array</span></span> | <span data-ttu-id="58d85-878">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="58d85-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="58d85-879">文字列</span><span class="sxs-lookup"><span data-stu-id="58d85-879">String</span></span> | <span data-ttu-id="58d85-880">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="58d85-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="58d85-881">連絡先</span><span class="sxs-lookup"><span data-stu-id="58d85-881">Contact</span></span> | <span data-ttu-id="58d85-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="58d85-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="58d85-883">文字列</span><span class="sxs-lookup"><span data-stu-id="58d85-883">String</span></span> | <span data-ttu-id="58d85-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="58d85-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="58d85-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="58d85-885">MeetingSuggestion</span></span> | <span data-ttu-id="58d85-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="58d85-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="58d85-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="58d85-887">PhoneNumber</span></span> | <span data-ttu-id="58d85-888">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="58d85-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="58d85-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="58d85-889">TaskSuggestion</span></span> | <span data-ttu-id="58d85-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="58d85-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="58d85-891">文字列</span><span class="sxs-lookup"><span data-stu-id="58d85-891">String</span></span> | <span data-ttu-id="58d85-892">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="58d85-892">**Restricted**</span></span> |

<span data-ttu-id="58d85-893">型:Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="58d85-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="58d85-894">例</span><span class="sxs-lookup"><span data-stu-id="58d85-894">Example</span></span>

<span data-ttu-id="58d85-895">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="58d85-895">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```javascript
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="58d85-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="58d85-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="58d85-897">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="58d85-898">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="58d85-898">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="58d85-899">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58d85-900">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58d85-900">Parameters</span></span>

|<span data-ttu-id="58d85-901">名前</span><span class="sxs-lookup"><span data-stu-id="58d85-901">Name</span></span>| <span data-ttu-id="58d85-902">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-902">Type</span></span>| <span data-ttu-id="58d85-903">説明</span><span class="sxs-lookup"><span data-stu-id="58d85-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="58d85-904">String</span><span class="sxs-lookup"><span data-stu-id="58d85-904">String</span></span>|<span data-ttu-id="58d85-905">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="58d85-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="58d85-906">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-906">Requirements</span></span>

|<span data-ttu-id="58d85-907">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-907">Requirement</span></span>| <span data-ttu-id="58d85-908">値</span><span class="sxs-lookup"><span data-stu-id="58d85-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-909">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-910">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-910">1.0</span></span>|
|[<span data-ttu-id="58d85-911">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-911">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-912">ReadItem</span></span>|
|[<span data-ttu-id="58d85-913">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-913">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-914">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="58d85-915">戻り値:</span><span class="sxs-lookup"><span data-stu-id="58d85-915">Returns:</span></span>

<span data-ttu-id="58d85-p155">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="58d85-918">型:Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="58d85-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="58d85-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="58d85-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="58d85-920">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="58d85-921">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="58d85-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="58d85-p156">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="58d85-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="58d85-925">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="58d85-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="58d85-926">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="58d85-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="58d85-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="58d85-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-930">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-930">Requirements</span></span>

|<span data-ttu-id="58d85-931">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-931">Requirement</span></span>| <span data-ttu-id="58d85-932">値</span><span class="sxs-lookup"><span data-stu-id="58d85-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-933">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-934">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-934">1.0</span></span>|
|[<span data-ttu-id="58d85-935">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-936">ReadItem</span></span>|
|[<span data-ttu-id="58d85-937">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-938">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="58d85-939">戻り値:</span><span class="sxs-lookup"><span data-stu-id="58d85-939">Returns:</span></span>

<span data-ttu-id="58d85-p158">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="58d85-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="58d85-942">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="58d85-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="58d85-943">Object</span><span class="sxs-lookup"><span data-stu-id="58d85-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="58d85-944">例</span><span class="sxs-lookup"><span data-stu-id="58d85-944">Example</span></span>

<span data-ttu-id="58d85-945">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="58d85-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="58d85-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="58d85-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="58d85-947">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="58d85-948">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="58d85-948">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="58d85-949">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="58d85-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="58d85-p159">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="58d85-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58d85-952">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58d85-952">Parameters</span></span>

|<span data-ttu-id="58d85-953">名前</span><span class="sxs-lookup"><span data-stu-id="58d85-953">Name</span></span>| <span data-ttu-id="58d85-954">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-954">Type</span></span>| <span data-ttu-id="58d85-955">説明</span><span class="sxs-lookup"><span data-stu-id="58d85-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="58d85-956">String</span><span class="sxs-lookup"><span data-stu-id="58d85-956">String</span></span>|<span data-ttu-id="58d85-957">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="58d85-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="58d85-958">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-958">Requirements</span></span>

|<span data-ttu-id="58d85-959">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-959">Requirement</span></span>| <span data-ttu-id="58d85-960">値</span><span class="sxs-lookup"><span data-stu-id="58d85-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-961">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-961">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-962">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-962">1.0</span></span>|
|[<span data-ttu-id="58d85-963">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-963">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-964">ReadItem</span></span>|
|[<span data-ttu-id="58d85-965">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-965">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-966">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="58d85-967">戻り値:</span><span class="sxs-lookup"><span data-stu-id="58d85-967">Returns:</span></span>

<span data-ttu-id="58d85-968">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="58d85-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="58d85-969">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="58d85-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="58d85-970">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="58d85-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="58d85-971">例</span><span class="sxs-lookup"><span data-stu-id="58d85-971">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="58d85-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="58d85-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="58d85-973">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="58d85-p160">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58d85-976">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58d85-976">Parameters</span></span>

|<span data-ttu-id="58d85-977">名前</span><span class="sxs-lookup"><span data-stu-id="58d85-977">Name</span></span>| <span data-ttu-id="58d85-978">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-978">Type</span></span>| <span data-ttu-id="58d85-979">属性</span><span class="sxs-lookup"><span data-stu-id="58d85-979">Attributes</span></span>| <span data-ttu-id="58d85-980">説明</span><span class="sxs-lookup"><span data-stu-id="58d85-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="58d85-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="58d85-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="58d85-p161">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="58d85-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="58d85-985">Object</span><span class="sxs-lookup"><span data-stu-id="58d85-985">Object</span></span>| <span data-ttu-id="58d85-986">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-986">&lt;optional&gt;</span></span>|<span data-ttu-id="58d85-987">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="58d85-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="58d85-988">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="58d85-988">Object</span></span>| <span data-ttu-id="58d85-989">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-989">&lt;optional&gt;</span></span>|<span data-ttu-id="58d85-990">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="58d85-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="58d85-991">function</span><span class="sxs-lookup"><span data-stu-id="58d85-991">function</span></span>||<span data-ttu-id="58d85-992">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="58d85-993">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="58d85-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="58d85-994">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="58d85-994">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="58d85-995">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-995">Requirements</span></span>

|<span data-ttu-id="58d85-996">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-996">Requirement</span></span>| <span data-ttu-id="58d85-997">値</span><span class="sxs-lookup"><span data-stu-id="58d85-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-998">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-998">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-999">1.2</span><span class="sxs-lookup"><span data-stu-id="58d85-999">1.2</span></span>|
|[<span data-ttu-id="58d85-1000">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-1000">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="58d85-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="58d85-1002">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-1002">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-1003">作成</span><span class="sxs-lookup"><span data-stu-id="58d85-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="58d85-1004">戻り値:</span><span class="sxs-lookup"><span data-stu-id="58d85-1004">Returns:</span></span>

<span data-ttu-id="58d85-1005">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="58d85-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="58d85-1006">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="58d85-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="58d85-1007">String</span><span class="sxs-lookup"><span data-stu-id="58d85-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="58d85-1008">例</span><span class="sxs-lookup"><span data-stu-id="58d85-1008">Example</span></span>

```javascript
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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="58d85-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="58d85-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="58d85-1010">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="58d85-1010">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="58d85-1011">強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1011">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="58d85-1012">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="58d85-1012">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-1013">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-1013">Requirements</span></span>

|<span data-ttu-id="58d85-1014">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-1014">Requirement</span></span>| <span data-ttu-id="58d85-1015">値</span><span class="sxs-lookup"><span data-stu-id="58d85-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-1016">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-1016">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="58d85-1017">1.6</span></span> |
|[<span data-ttu-id="58d85-1018">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-1018">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-1019">ReadItem</span></span>|
|[<span data-ttu-id="58d85-1020">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-1020">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-1021">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="58d85-1022">戻り値:</span><span class="sxs-lookup"><span data-stu-id="58d85-1022">Returns:</span></span>

<span data-ttu-id="58d85-1023">型:[Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="58d85-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="58d85-1024">例</span><span class="sxs-lookup"><span data-stu-id="58d85-1024">Example</span></span>

<span data-ttu-id="58d85-1025">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="58d85-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="58d85-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="58d85-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="58d85-p164">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="58d85-1029">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="58d85-1029">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="58d85-p165">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="58d85-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="58d85-1033">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="58d85-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="58d85-1034">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="58d85-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="58d85-p166">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="58d85-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="58d85-1038">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-1038">Requirements</span></span>

|<span data-ttu-id="58d85-1039">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-1039">Requirement</span></span>| <span data-ttu-id="58d85-1040">値</span><span class="sxs-lookup"><span data-stu-id="58d85-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-1041">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="58d85-1042">1.6</span></span> |
|[<span data-ttu-id="58d85-1043">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-1043">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-1044">ReadItem</span></span>|
|[<span data-ttu-id="58d85-1045">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-1045">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-1046">読み取り</span><span class="sxs-lookup"><span data-stu-id="58d85-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="58d85-1047">戻り値:</span><span class="sxs-lookup"><span data-stu-id="58d85-1047">Returns:</span></span>

<span data-ttu-id="58d85-p167">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="58d85-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="58d85-1050">例</span><span class="sxs-lookup"><span data-stu-id="58d85-1050">Example</span></span>

<span data-ttu-id="58d85-1051">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="58d85-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="58d85-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="58d85-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="58d85-1053">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="58d85-p168">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="58d85-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58d85-1057">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58d85-1057">Parameters</span></span>

|<span data-ttu-id="58d85-1058">名前</span><span class="sxs-lookup"><span data-stu-id="58d85-1058">Name</span></span>| <span data-ttu-id="58d85-1059">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-1059">Type</span></span>| <span data-ttu-id="58d85-1060">属性</span><span class="sxs-lookup"><span data-stu-id="58d85-1060">Attributes</span></span>| <span data-ttu-id="58d85-1061">説明</span><span class="sxs-lookup"><span data-stu-id="58d85-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="58d85-1062">function</span><span class="sxs-lookup"><span data-stu-id="58d85-1062">function</span></span>||<span data-ttu-id="58d85-1063">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="58d85-1064">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="58d85-1065">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1065">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="58d85-1066">Object</span><span class="sxs-lookup"><span data-stu-id="58d85-1066">Object</span></span>| <span data-ttu-id="58d85-1067">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="58d85-1068">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1068">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="58d85-1069">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="58d85-1070">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-1070">Requirements</span></span>

|<span data-ttu-id="58d85-1071">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-1071">Requirement</span></span>| <span data-ttu-id="58d85-1072">値</span><span class="sxs-lookup"><span data-stu-id="58d85-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-1073">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-1073">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="58d85-1074">1.0</span></span>|
|[<span data-ttu-id="58d85-1075">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-1075">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58d85-1076">ReadItem</span></span>|
|[<span data-ttu-id="58d85-1077">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-1077">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-1078">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58d85-1078">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58d85-1079">例</span><span class="sxs-lookup"><span data-stu-id="58d85-1079">Example</span></span>

<span data-ttu-id="58d85-p171">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="58d85-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
};

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="58d85-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="58d85-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="58d85-1084">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="58d85-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="58d85-p172">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="58d85-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58d85-1089">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58d85-1089">Parameters</span></span>

|<span data-ttu-id="58d85-1090">名前</span><span class="sxs-lookup"><span data-stu-id="58d85-1090">Name</span></span>| <span data-ttu-id="58d85-1091">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-1091">Type</span></span>| <span data-ttu-id="58d85-1092">属性</span><span class="sxs-lookup"><span data-stu-id="58d85-1092">Attributes</span></span>| <span data-ttu-id="58d85-1093">説明</span><span class="sxs-lookup"><span data-stu-id="58d85-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="58d85-1094">String</span><span class="sxs-lookup"><span data-stu-id="58d85-1094">String</span></span>||<span data-ttu-id="58d85-1095">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="58d85-1095">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="58d85-1096">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="58d85-1096">Object</span></span>| <span data-ttu-id="58d85-1097">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="58d85-1098">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="58d85-1098">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="58d85-1099">Object</span><span class="sxs-lookup"><span data-stu-id="58d85-1099">Object</span></span>| <span data-ttu-id="58d85-1100">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-1100">&lt;optional&gt;</span></span>|<span data-ttu-id="58d85-1101">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1101">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="58d85-1102">function</span><span class="sxs-lookup"><span data-stu-id="58d85-1102">function</span></span>| <span data-ttu-id="58d85-1103">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-1103">&lt;optional&gt;</span></span>|<span data-ttu-id="58d85-1104">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1104">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="58d85-1105">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1105">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="58d85-1106">エラー</span><span class="sxs-lookup"><span data-stu-id="58d85-1106">Errors</span></span>

| <span data-ttu-id="58d85-1107">エラー コード</span><span class="sxs-lookup"><span data-stu-id="58d85-1107">Error code</span></span> | <span data-ttu-id="58d85-1108">説明</span><span class="sxs-lookup"><span data-stu-id="58d85-1108">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="58d85-1109">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="58d85-1109">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="58d85-1110">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-1110">Requirements</span></span>

|<span data-ttu-id="58d85-1111">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-1111">Requirement</span></span>| <span data-ttu-id="58d85-1112">値</span><span class="sxs-lookup"><span data-stu-id="58d85-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-1113">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-1114">1.1</span><span class="sxs-lookup"><span data-stu-id="58d85-1114">1.1</span></span>|
|[<span data-ttu-id="58d85-1115">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-1115">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-1116">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="58d85-1116">ReadWriteItem</span></span>|
|[<span data-ttu-id="58d85-1117">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-1117">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-1118">作成</span><span class="sxs-lookup"><span data-stu-id="58d85-1118">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="58d85-1119">例</span><span class="sxs-lookup"><span data-stu-id="58d85-1119">Example</span></span>

<span data-ttu-id="58d85-1120">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="58d85-1120">The following code removes an attachment with an identifier of '0'.</span></span>

```javascript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="58d85-1121">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="58d85-1121">saveAsync([options], callback)</span></span>

<span data-ttu-id="58d85-1122">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="58d85-1122">Asynchronously saves an item.</span></span>

<span data-ttu-id="58d85-p173">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-p173">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="58d85-1126">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="58d85-1126">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="58d85-1127">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1127">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="58d85-p175">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="58d85-1131">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="58d85-1131">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="58d85-1132">Mac Outlook では、新規作成モードの会議で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="58d85-1132">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="58d85-1133">Mac Outlook では、会議で `saveAsync` を呼び出すとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1133">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="58d85-1134">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1134">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58d85-1135">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58d85-1135">Parameters</span></span>

|<span data-ttu-id="58d85-1136">名前</span><span class="sxs-lookup"><span data-stu-id="58d85-1136">Name</span></span>| <span data-ttu-id="58d85-1137">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-1137">Type</span></span>| <span data-ttu-id="58d85-1138">属性</span><span class="sxs-lookup"><span data-stu-id="58d85-1138">Attributes</span></span>| <span data-ttu-id="58d85-1139">説明</span><span class="sxs-lookup"><span data-stu-id="58d85-1139">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="58d85-1140">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="58d85-1140">Object</span></span>| <span data-ttu-id="58d85-1141">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-1141">&lt;optional&gt;</span></span>|<span data-ttu-id="58d85-1142">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="58d85-1142">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="58d85-1143">Object</span><span class="sxs-lookup"><span data-stu-id="58d85-1143">Object</span></span>| <span data-ttu-id="58d85-1144">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-1144">&lt;optional&gt;</span></span>|<span data-ttu-id="58d85-1145">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1145">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="58d85-1146">関数</span><span class="sxs-lookup"><span data-stu-id="58d85-1146">function</span></span>||<span data-ttu-id="58d85-1147">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1147">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="58d85-1148">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1148">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="58d85-1149">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-1149">Requirements</span></span>

|<span data-ttu-id="58d85-1150">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-1150">Requirement</span></span>| <span data-ttu-id="58d85-1151">値</span><span class="sxs-lookup"><span data-stu-id="58d85-1151">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-1152">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-1152">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-1153">1.3</span><span class="sxs-lookup"><span data-stu-id="58d85-1153">1.3</span></span>|
|[<span data-ttu-id="58d85-1154">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-1154">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-1155">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="58d85-1155">ReadWriteItem</span></span>|
|[<span data-ttu-id="58d85-1156">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-1156">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-1157">作成</span><span class="sxs-lookup"><span data-stu-id="58d85-1157">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="58d85-1158">例</span><span class="sxs-lookup"><span data-stu-id="58d85-1158">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="58d85-p177">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="58d85-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="58d85-1161">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="58d85-1161">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="58d85-1162">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="58d85-1162">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="58d85-p178">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="58d85-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58d85-1166">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58d85-1166">Parameters</span></span>

|<span data-ttu-id="58d85-1167">名前</span><span class="sxs-lookup"><span data-stu-id="58d85-1167">Name</span></span>| <span data-ttu-id="58d85-1168">種類</span><span class="sxs-lookup"><span data-stu-id="58d85-1168">Type</span></span>| <span data-ttu-id="58d85-1169">属性</span><span class="sxs-lookup"><span data-stu-id="58d85-1169">Attributes</span></span>| <span data-ttu-id="58d85-1170">説明</span><span class="sxs-lookup"><span data-stu-id="58d85-1170">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="58d85-1171">String</span><span class="sxs-lookup"><span data-stu-id="58d85-1171">String</span></span>||<span data-ttu-id="58d85-p179">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="58d85-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="58d85-1175">Object</span><span class="sxs-lookup"><span data-stu-id="58d85-1175">Object</span></span>| <span data-ttu-id="58d85-1176">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-1176">&lt;optional&gt;</span></span>|<span data-ttu-id="58d85-1177">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="58d85-1177">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="58d85-1178">Object</span><span class="sxs-lookup"><span data-stu-id="58d85-1178">Object</span></span>| <span data-ttu-id="58d85-1179">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-1179">&lt;optional&gt;</span></span>|<span data-ttu-id="58d85-1180">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1180">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="58d85-1181">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="58d85-1181">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="58d85-1182">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="58d85-1182">&lt;optional&gt;</span></span>|<span data-ttu-id="58d85-p180">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-p180">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="58d85-p181">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-p181">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="58d85-1187">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1187">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="58d85-1188">function</span><span class="sxs-lookup"><span data-stu-id="58d85-1188">function</span></span>||<span data-ttu-id="58d85-1189">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="58d85-1189">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="58d85-1190">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-1190">Requirements</span></span>

|<span data-ttu-id="58d85-1191">要件</span><span class="sxs-lookup"><span data-stu-id="58d85-1191">Requirement</span></span>| <span data-ttu-id="58d85-1192">値</span><span class="sxs-lookup"><span data-stu-id="58d85-1192">Value</span></span>|
|---|---|
|[<span data-ttu-id="58d85-1193">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58d85-1193">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58d85-1194">1.2</span><span class="sxs-lookup"><span data-stu-id="58d85-1194">1.2</span></span>|
|[<span data-ttu-id="58d85-1195">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58d85-1195">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58d85-1196">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="58d85-1196">ReadWriteItem</span></span>|
|[<span data-ttu-id="58d85-1197">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58d85-1197">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58d85-1198">作成</span><span class="sxs-lookup"><span data-stu-id="58d85-1198">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="58d85-1199">例</span><span class="sxs-lookup"><span data-stu-id="58d85-1199">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
