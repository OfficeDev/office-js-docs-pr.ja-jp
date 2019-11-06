---
title: Office. メールボックス-要件セット1.8
description: ''
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 065ea3c74580555c0df1af7b495127a25493b612
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001573"
---
# <a name="item"></a><span data-ttu-id="57e53-102">item</span><span class="sxs-lookup"><span data-stu-id="57e53-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="57e53-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="57e53-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="57e53-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-106">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-106">Requirements</span></span>

|<span data-ttu-id="57e53-107">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-107">Requirement</span></span>|<span data-ttu-id="57e53-108">値</span><span class="sxs-lookup"><span data-stu-id="57e53-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-110">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-110">1.0</span></span>|
|[<span data-ttu-id="57e53-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="57e53-112">Restricted</span></span>|
|[<span data-ttu-id="57e53-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="57e53-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-115">Members and methods</span></span>

| <span data-ttu-id="57e53-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-116">Member</span></span> | <span data-ttu-id="57e53-117">種類</span><span class="sxs-lookup"><span data-stu-id="57e53-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="57e53-118">attachments</span><span class="sxs-lookup"><span data-stu-id="57e53-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="57e53-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-119">Member</span></span> |
| [<span data-ttu-id="57e53-120">bcc</span><span class="sxs-lookup"><span data-stu-id="57e53-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="57e53-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-121">Member</span></span> |
| [<span data-ttu-id="57e53-122">body</span><span class="sxs-lookup"><span data-stu-id="57e53-122">body</span></span>](#body-body) | <span data-ttu-id="57e53-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-123">Member</span></span> |
| [<span data-ttu-id="57e53-124">categories</span><span class="sxs-lookup"><span data-stu-id="57e53-124">categories</span></span>](#categories-categories) | <span data-ttu-id="57e53-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-125">Member</span></span> |
| [<span data-ttu-id="57e53-126">cc</span><span class="sxs-lookup"><span data-stu-id="57e53-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="57e53-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-127">Member</span></span> |
| [<span data-ttu-id="57e53-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="57e53-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="57e53-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-129">Member</span></span> |
| [<span data-ttu-id="57e53-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="57e53-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="57e53-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-131">Member</span></span> |
| [<span data-ttu-id="57e53-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="57e53-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="57e53-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-133">Member</span></span> |
| [<span data-ttu-id="57e53-134">end</span><span class="sxs-lookup"><span data-stu-id="57e53-134">end</span></span>](#end-datetime) | <span data-ttu-id="57e53-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-135">Member</span></span> |
| [<span data-ttu-id="57e53-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="57e53-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="57e53-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-137">Member</span></span> |
| [<span data-ttu-id="57e53-138">from</span><span class="sxs-lookup"><span data-stu-id="57e53-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="57e53-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-139">Member</span></span> |
| [<span data-ttu-id="57e53-140">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="57e53-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="57e53-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-141">Member</span></span> |
| [<span data-ttu-id="57e53-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="57e53-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="57e53-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-143">Member</span></span> |
| [<span data-ttu-id="57e53-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="57e53-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="57e53-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-145">Member</span></span> |
| [<span data-ttu-id="57e53-146">itemId</span><span class="sxs-lookup"><span data-stu-id="57e53-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="57e53-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-147">Member</span></span> |
| [<span data-ttu-id="57e53-148">itemType</span><span class="sxs-lookup"><span data-stu-id="57e53-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="57e53-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-149">Member</span></span> |
| [<span data-ttu-id="57e53-150">location</span><span class="sxs-lookup"><span data-stu-id="57e53-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="57e53-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-151">Member</span></span> |
| [<span data-ttu-id="57e53-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="57e53-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="57e53-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-153">Member</span></span> |
| [<span data-ttu-id="57e53-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="57e53-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="57e53-155">Member</span><span class="sxs-lookup"><span data-stu-id="57e53-155">Member</span></span> |
| [<span data-ttu-id="57e53-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="57e53-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="57e53-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-157">Member</span></span> |
| [<span data-ttu-id="57e53-158">organizer</span><span class="sxs-lookup"><span data-stu-id="57e53-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="57e53-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-159">Member</span></span> |
| [<span data-ttu-id="57e53-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="57e53-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="57e53-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-161">Member</span></span> |
| [<span data-ttu-id="57e53-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="57e53-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="57e53-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-163">Member</span></span> |
| [<span data-ttu-id="57e53-164">sender</span><span class="sxs-lookup"><span data-stu-id="57e53-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="57e53-165">Member</span><span class="sxs-lookup"><span data-stu-id="57e53-165">Member</span></span> |
| [<span data-ttu-id="57e53-166">系列 Id</span><span class="sxs-lookup"><span data-stu-id="57e53-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="57e53-167">Member</span><span class="sxs-lookup"><span data-stu-id="57e53-167">Member</span></span> |
| [<span data-ttu-id="57e53-168">start</span><span class="sxs-lookup"><span data-stu-id="57e53-168">start</span></span>](#start-datetime) | <span data-ttu-id="57e53-169">Member</span><span class="sxs-lookup"><span data-stu-id="57e53-169">Member</span></span> |
| [<span data-ttu-id="57e53-170">subject</span><span class="sxs-lookup"><span data-stu-id="57e53-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="57e53-171">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-171">Member</span></span> |
| [<span data-ttu-id="57e53-172">to</span><span class="sxs-lookup"><span data-stu-id="57e53-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="57e53-173">メンバー</span><span class="sxs-lookup"><span data-stu-id="57e53-173">Member</span></span> |
| [<span data-ttu-id="57e53-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="57e53-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="57e53-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-175">Method</span></span> |
| [<span data-ttu-id="57e53-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="57e53-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="57e53-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-177">Method</span></span> |
| [<span data-ttu-id="57e53-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="57e53-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="57e53-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-179">Method</span></span> |
| [<span data-ttu-id="57e53-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="57e53-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="57e53-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-181">Method</span></span> |
| [<span data-ttu-id="57e53-182">close</span><span class="sxs-lookup"><span data-stu-id="57e53-182">close</span></span>](#close) | <span data-ttu-id="57e53-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-183">Method</span></span> |
| [<span data-ttu-id="57e53-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="57e53-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="57e53-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-185">Method</span></span> |
| [<span data-ttu-id="57e53-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="57e53-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="57e53-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-187">Method</span></span> |
| [<span data-ttu-id="57e53-188">getAllInternetHeadersAsync</span><span class="sxs-lookup"><span data-stu-id="57e53-188">getAllInternetHeadersAsync</span></span>](#getallinternetheadersasyncoptions-callback) | <span data-ttu-id="57e53-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-189">Method</span></span> |
| [<span data-ttu-id="57e53-190">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="57e53-190">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="57e53-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-191">Method</span></span> |
| [<span data-ttu-id="57e53-192">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="57e53-192">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="57e53-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-193">Method</span></span> |
| [<span data-ttu-id="57e53-194">getEntities</span><span class="sxs-lookup"><span data-stu-id="57e53-194">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="57e53-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-195">Method</span></span> |
| [<span data-ttu-id="57e53-196">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="57e53-196">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="57e53-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-197">Method</span></span> |
| [<span data-ttu-id="57e53-198">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="57e53-198">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="57e53-199">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-199">Method</span></span> |
| [<span data-ttu-id="57e53-200">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="57e53-200">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="57e53-201">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-201">Method</span></span> |
| [<span data-ttu-id="57e53-202">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="57e53-202">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="57e53-203">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-203">Method</span></span> |
| [<span data-ttu-id="57e53-204">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="57e53-204">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="57e53-205">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-205">Method</span></span> |
| [<span data-ttu-id="57e53-206">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="57e53-206">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="57e53-207">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-207">Method</span></span> |
| [<span data-ttu-id="57e53-208">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="57e53-208">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="57e53-209">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-209">Method</span></span> |
| [<span data-ttu-id="57e53-210">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="57e53-210">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="57e53-211">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-211">Method</span></span> |
| [<span data-ttu-id="57e53-212">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="57e53-212">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="57e53-213">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-213">Method</span></span> |
| [<span data-ttu-id="57e53-214">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="57e53-214">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="57e53-215">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-215">Method</span></span> |
| [<span data-ttu-id="57e53-216">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="57e53-216">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="57e53-217">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-217">Method</span></span> |
| [<span data-ttu-id="57e53-218">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="57e53-218">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="57e53-219">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-219">Method</span></span> |
| [<span data-ttu-id="57e53-220">saveAsync</span><span class="sxs-lookup"><span data-stu-id="57e53-220">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="57e53-221">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-221">Method</span></span> |
| [<span data-ttu-id="57e53-222">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="57e53-222">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="57e53-223">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-223">Method</span></span> |

### <a name="example"></a><span data-ttu-id="57e53-224">例</span><span class="sxs-lookup"><span data-stu-id="57e53-224">Example</span></span>

<span data-ttu-id="57e53-225">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="57e53-225">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="57e53-226">Members</span><span class="sxs-lookup"><span data-stu-id="57e53-226">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-18"></a><span data-ttu-id="57e53-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="57e53-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

<span data-ttu-id="57e53-228">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-228">Gets the item's attachments as an array.</span></span> <span data-ttu-id="57e53-229">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="57e53-229">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-230">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="57e53-230">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="57e53-231">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57e53-231">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="57e53-232">型</span><span class="sxs-lookup"><span data-stu-id="57e53-232">Type</span></span>

*   <span data-ttu-id="57e53-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="57e53-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-234">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-234">Requirements</span></span>

|<span data-ttu-id="57e53-235">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-235">Requirement</span></span>|<span data-ttu-id="57e53-236">値</span><span class="sxs-lookup"><span data-stu-id="57e53-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-237">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-238">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-238">1.0</span></span>|
|[<span data-ttu-id="57e53-239">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-240">ReadItem</span></span>|
|[<span data-ttu-id="57e53-241">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-242">読み取り</span><span class="sxs-lookup"><span data-stu-id="57e53-242">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-243">例</span><span class="sxs-lookup"><span data-stu-id="57e53-243">Example</span></span>

<span data-ttu-id="57e53-244">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="57e53-244">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="57e53-245">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-245">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-246">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-246">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="57e53-247">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="57e53-247">Compose mode only.</span></span>

<span data-ttu-id="57e53-248">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="57e53-248">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="57e53-249">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-249">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="57e53-250">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-250">Get 500 members maximum.</span></span>
- <span data-ttu-id="57e53-251">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="57e53-251">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="57e53-252">型</span><span class="sxs-lookup"><span data-stu-id="57e53-252">Type</span></span>

*   [<span data-ttu-id="57e53-253">受信者</span><span class="sxs-lookup"><span data-stu-id="57e53-253">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="57e53-254">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-254">Requirements</span></span>

|<span data-ttu-id="57e53-255">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-255">Requirement</span></span>|<span data-ttu-id="57e53-256">値</span><span class="sxs-lookup"><span data-stu-id="57e53-256">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-257">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-257">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-258">1.1</span><span class="sxs-lookup"><span data-stu-id="57e53-258">1.1</span></span>|
|[<span data-ttu-id="57e53-259">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-259">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-260">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-260">ReadItem</span></span>|
|[<span data-ttu-id="57e53-261">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-261">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-262">作成</span><span class="sxs-lookup"><span data-stu-id="57e53-262">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-263">例</span><span class="sxs-lookup"><span data-stu-id="57e53-263">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-18"></a><span data-ttu-id="57e53-264">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-264">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-265">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-265">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="57e53-266">型</span><span class="sxs-lookup"><span data-stu-id="57e53-266">Type</span></span>

*   [<span data-ttu-id="57e53-267">Body</span><span class="sxs-lookup"><span data-stu-id="57e53-267">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="57e53-268">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-268">Requirements</span></span>

|<span data-ttu-id="57e53-269">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-269">Requirement</span></span>|<span data-ttu-id="57e53-270">値</span><span class="sxs-lookup"><span data-stu-id="57e53-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-271">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-272">1.1</span><span class="sxs-lookup"><span data-stu-id="57e53-272">1.1</span></span>|
|[<span data-ttu-id="57e53-273">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-274">ReadItem</span></span>|
|[<span data-ttu-id="57e53-275">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-276">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-277">例</span><span class="sxs-lookup"><span data-stu-id="57e53-277">Example</span></span>

<span data-ttu-id="57e53-278">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-278">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="57e53-279">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="57e53-279">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="categories-categoriesjavascriptapioutlookofficecategoriesviewoutlook-js-18"></a><span data-ttu-id="57e53-280">カテゴリ:[カテゴリ](/javascript/api/outlook/office.categories?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-280">categories: [Categories](/javascript/api/outlook/office.categories?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-281">アイテムのカテゴリを管理するためのメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-281">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-282">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="57e53-282">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="57e53-283">種類</span><span class="sxs-lookup"><span data-stu-id="57e53-283">Type</span></span>

*   [<span data-ttu-id="57e53-284">Categories</span><span class="sxs-lookup"><span data-stu-id="57e53-284">Categories</span></span>](/javascript/api/outlook/office.categories?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="57e53-285">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-285">Requirements</span></span>

|<span data-ttu-id="57e53-286">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-286">Requirement</span></span>|<span data-ttu-id="57e53-287">値</span><span class="sxs-lookup"><span data-stu-id="57e53-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-288">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-289">1.8</span><span class="sxs-lookup"><span data-stu-id="57e53-289">1.8</span></span>|
|[<span data-ttu-id="57e53-290">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-290">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-291">ReadItem</span></span>|
|[<span data-ttu-id="57e53-292">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-293">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-293">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-294">例</span><span class="sxs-lookup"><span data-stu-id="57e53-294">Example</span></span>

<span data-ttu-id="57e53-295">この例では、アイテムのカテゴリを取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-295">This example gets the item's categories.</span></span>

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="57e53-296">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-296">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-297">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="57e53-297">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="57e53-298">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="57e53-298">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57e53-299">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="57e53-299">Read mode</span></span>

<span data-ttu-id="57e53-300">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-300">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="57e53-301">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="57e53-301">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="57e53-302">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-302">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="57e53-303">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="57e53-303">Compose mode</span></span>

<span data-ttu-id="57e53-304">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-304">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="57e53-305">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="57e53-305">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="57e53-306">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-306">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="57e53-307">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-307">Get 500 members maximum.</span></span>
- <span data-ttu-id="57e53-308">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="57e53-308">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="57e53-309">型</span><span class="sxs-lookup"><span data-stu-id="57e53-309">Type</span></span>

*   <span data-ttu-id="57e53-310">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-310">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-311">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-311">Requirements</span></span>

|<span data-ttu-id="57e53-312">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-312">Requirement</span></span>|<span data-ttu-id="57e53-313">値</span><span class="sxs-lookup"><span data-stu-id="57e53-313">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-314">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-314">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-315">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-315">1.0</span></span>|
|[<span data-ttu-id="57e53-316">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-316">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-317">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-317">ReadItem</span></span>|
|[<span data-ttu-id="57e53-318">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-318">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-319">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-319">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="57e53-320">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="57e53-320">(nullable) conversationId: String</span></span>

<span data-ttu-id="57e53-321">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-321">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="57e53-p109">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="57e53-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="57e53-p110">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="57e53-326">Type</span><span class="sxs-lookup"><span data-stu-id="57e53-326">Type</span></span>

*   <span data-ttu-id="57e53-327">String</span><span class="sxs-lookup"><span data-stu-id="57e53-327">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-328">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-328">Requirements</span></span>

|<span data-ttu-id="57e53-329">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-329">Requirement</span></span>|<span data-ttu-id="57e53-330">値</span><span class="sxs-lookup"><span data-stu-id="57e53-330">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-331">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-331">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-332">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-332">1.0</span></span>|
|[<span data-ttu-id="57e53-333">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-333">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-334">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-334">ReadItem</span></span>|
|[<span data-ttu-id="57e53-335">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-335">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-336">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-336">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-337">例</span><span class="sxs-lookup"><span data-stu-id="57e53-337">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="57e53-338">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="57e53-338">dateTimeCreated: Date</span></span>

<span data-ttu-id="57e53-p111">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="57e53-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="57e53-341">型</span><span class="sxs-lookup"><span data-stu-id="57e53-341">Type</span></span>

*   <span data-ttu-id="57e53-342">日付</span><span class="sxs-lookup"><span data-stu-id="57e53-342">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-343">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-343">Requirements</span></span>

|<span data-ttu-id="57e53-344">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-344">Requirement</span></span>|<span data-ttu-id="57e53-345">値</span><span class="sxs-lookup"><span data-stu-id="57e53-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-346">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-347">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-347">1.0</span></span>|
|[<span data-ttu-id="57e53-348">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-349">ReadItem</span></span>|
|[<span data-ttu-id="57e53-350">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-351">読み取り</span><span class="sxs-lookup"><span data-stu-id="57e53-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-352">例</span><span class="sxs-lookup"><span data-stu-id="57e53-352">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="57e53-353">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="57e53-353">dateTimeModified: Date</span></span>

<span data-ttu-id="57e53-p112">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="57e53-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-356">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="57e53-356">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="57e53-357">種類</span><span class="sxs-lookup"><span data-stu-id="57e53-357">Type</span></span>

*   <span data-ttu-id="57e53-358">日付</span><span class="sxs-lookup"><span data-stu-id="57e53-358">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-359">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-359">Requirements</span></span>

|<span data-ttu-id="57e53-360">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-360">Requirement</span></span>|<span data-ttu-id="57e53-361">値</span><span class="sxs-lookup"><span data-stu-id="57e53-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-362">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-363">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-363">1.0</span></span>|
|[<span data-ttu-id="57e53-364">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-365">ReadItem</span></span>|
|[<span data-ttu-id="57e53-366">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-367">読み取り</span><span class="sxs-lookup"><span data-stu-id="57e53-367">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-368">例</span><span class="sxs-lookup"><span data-stu-id="57e53-368">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-18"></a><span data-ttu-id="57e53-369">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-369">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-370">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="57e53-370">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="57e53-p113">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="57e53-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57e53-373">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="57e53-373">Read mode</span></span>

<span data-ttu-id="57e53-374">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-374">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="57e53-375">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="57e53-375">Compose mode</span></span>

<span data-ttu-id="57e53-376">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-376">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="57e53-377">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="57e53-377">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="57e53-378">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="57e53-378">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="57e53-379">型</span><span class="sxs-lookup"><span data-stu-id="57e53-379">Type</span></span>

*   <span data-ttu-id="57e53-380">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-380">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-381">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-381">Requirements</span></span>

|<span data-ttu-id="57e53-382">必要条件</span><span class="sxs-lookup"><span data-stu-id="57e53-382">Requirement</span></span>|<span data-ttu-id="57e53-383">値</span><span class="sxs-lookup"><span data-stu-id="57e53-383">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-384">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-384">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-385">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-385">1.0</span></span>|
|[<span data-ttu-id="57e53-386">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-386">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-387">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-387">ReadItem</span></span>|
|[<span data-ttu-id="57e53-388">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-388">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-389">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-389">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocationviewoutlook-js-18"></a><span data-ttu-id="57e53-390">enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-390">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-391">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="57e53-391">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57e53-392">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="57e53-392">Read mode</span></span>

<span data-ttu-id="57e53-393">この`enhancedLocation`プロパティは、予定に関連付けられている場所 ( [locationdetails](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8)オブジェクトで表される) のセットを取得できる[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-393">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="57e53-394">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="57e53-394">Compose mode</span></span>

<span data-ttu-id="57e53-395">この`enhancedLocation`プロパティは、予定の場所を取得、削除、または追加するためのメソッドを提供する[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-395">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="57e53-396">型</span><span class="sxs-lookup"><span data-stu-id="57e53-396">Type</span></span>

*   [<span data-ttu-id="57e53-397">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="57e53-397">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="57e53-398">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-398">Requirements</span></span>

|<span data-ttu-id="57e53-399">必要条件</span><span class="sxs-lookup"><span data-stu-id="57e53-399">Requirement</span></span>|<span data-ttu-id="57e53-400">値</span><span class="sxs-lookup"><span data-stu-id="57e53-400">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-401">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-401">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-402">1.8</span><span class="sxs-lookup"><span data-stu-id="57e53-402">1.8</span></span>|
|[<span data-ttu-id="57e53-403">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-403">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-404">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-404">ReadItem</span></span>|
|[<span data-ttu-id="57e53-405">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-405">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-406">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-406">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-407">例</span><span class="sxs-lookup"><span data-stu-id="57e53-407">Example</span></span>

<span data-ttu-id="57e53-408">次の例では、予定に関連付けられている現在の場所を取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-408">The following example gets the current locations associated with the appointment.</span></span>

```js
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}
```

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18fromjavascriptapioutlookofficefromviewoutlook-js-18"></a><span data-ttu-id="57e53-409">from: [emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[from](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-409">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-410">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-410">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="57e53-p114">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="57e53-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-413">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="57e53-413">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57e53-414">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="57e53-414">Read mode</span></span>

<span data-ttu-id="57e53-415">プロパティ`from`は`EmailAddressDetails`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-415">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="57e53-416">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="57e53-416">Compose mode</span></span>

<span data-ttu-id="57e53-417">プロパティ`from`は、from `From`値を取得するメソッドを提供するオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-417">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="57e53-418">型</span><span class="sxs-lookup"><span data-stu-id="57e53-418">Type</span></span>

*   <span data-ttu-id="57e53-419">[電子メールアドレス](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [の](/javascript/api/outlook/office.from?view=outlook-js-1.8)詳細</span><span class="sxs-lookup"><span data-stu-id="57e53-419">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-420">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-420">Requirements</span></span>

|<span data-ttu-id="57e53-421">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-421">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="57e53-422">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-422">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-423">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-423">1.0</span></span>|<span data-ttu-id="57e53-424">1.7</span><span class="sxs-lookup"><span data-stu-id="57e53-424">1.7</span></span>|
|[<span data-ttu-id="57e53-425">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-425">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-426">ReadItem</span></span>|<span data-ttu-id="57e53-427">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="57e53-427">ReadWriteItem</span></span>|
|[<span data-ttu-id="57e53-428">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-429">Read</span><span class="sxs-lookup"><span data-stu-id="57e53-429">Read</span></span>|<span data-ttu-id="57e53-430">Compose</span><span class="sxs-lookup"><span data-stu-id="57e53-430">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheadersviewoutlook-js-18"></a><span data-ttu-id="57e53-431">internetHeaders: [internetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-431">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-432">メッセージのカスタムインターネットヘッダーを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="57e53-432">Gets or sets custom internet headers on a message.</span></span> <span data-ttu-id="57e53-433">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="57e53-433">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="57e53-434">型</span><span class="sxs-lookup"><span data-stu-id="57e53-434">Type</span></span>

*   [<span data-ttu-id="57e53-435">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="57e53-435">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="57e53-436">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-436">Requirements</span></span>

|<span data-ttu-id="57e53-437">必要条件</span><span class="sxs-lookup"><span data-stu-id="57e53-437">Requirement</span></span>|<span data-ttu-id="57e53-438">値</span><span class="sxs-lookup"><span data-stu-id="57e53-438">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-439">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-439">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-440">1.8</span><span class="sxs-lookup"><span data-stu-id="57e53-440">1.8</span></span>|
|[<span data-ttu-id="57e53-441">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-441">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-442">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-442">ReadItem</span></span>|
|[<span data-ttu-id="57e53-443">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-443">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-444">作成</span><span class="sxs-lookup"><span data-stu-id="57e53-444">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-445">例</span><span class="sxs-lookup"><span data-stu-id="57e53-445">Example</span></span>

```js
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="57e53-446">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="57e53-446">internetMessageId: String</span></span>

<span data-ttu-id="57e53-p116">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="57e53-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="57e53-449">Type</span><span class="sxs-lookup"><span data-stu-id="57e53-449">Type</span></span>

*   <span data-ttu-id="57e53-450">String</span><span class="sxs-lookup"><span data-stu-id="57e53-450">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-451">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-451">Requirements</span></span>

|<span data-ttu-id="57e53-452">必要条件</span><span class="sxs-lookup"><span data-stu-id="57e53-452">Requirement</span></span>|<span data-ttu-id="57e53-453">値</span><span class="sxs-lookup"><span data-stu-id="57e53-453">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-454">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-454">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-455">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-455">1.0</span></span>|
|[<span data-ttu-id="57e53-456">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-456">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-457">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-457">ReadItem</span></span>|
|[<span data-ttu-id="57e53-458">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-458">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-459">読み取り</span><span class="sxs-lookup"><span data-stu-id="57e53-459">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-460">例</span><span class="sxs-lookup"><span data-stu-id="57e53-460">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="57e53-461">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="57e53-461">itemClass: String</span></span>

<span data-ttu-id="57e53-p117">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="57e53-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="57e53-p118">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="57e53-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="57e53-466">型</span><span class="sxs-lookup"><span data-stu-id="57e53-466">Type</span></span>|<span data-ttu-id="57e53-467">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-467">Description</span></span>|<span data-ttu-id="57e53-468">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="57e53-468">item class</span></span>|
|---|---|---|
|<span data-ttu-id="57e53-469">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="57e53-469">Appointment items</span></span>|<span data-ttu-id="57e53-470">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="57e53-470">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="57e53-471">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="57e53-471">Message items</span></span>|<span data-ttu-id="57e53-472">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="57e53-472">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="57e53-473">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-473">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="57e53-474">Type</span><span class="sxs-lookup"><span data-stu-id="57e53-474">Type</span></span>

*   <span data-ttu-id="57e53-475">String</span><span class="sxs-lookup"><span data-stu-id="57e53-475">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-476">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-476">Requirements</span></span>

|<span data-ttu-id="57e53-477">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-477">Requirement</span></span>|<span data-ttu-id="57e53-478">値</span><span class="sxs-lookup"><span data-stu-id="57e53-478">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-479">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-479">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-480">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-480">1.0</span></span>|
|[<span data-ttu-id="57e53-481">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-481">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-482">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-482">ReadItem</span></span>|
|[<span data-ttu-id="57e53-483">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-483">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-484">読み取り</span><span class="sxs-lookup"><span data-stu-id="57e53-484">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-485">例</span><span class="sxs-lookup"><span data-stu-id="57e53-485">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="57e53-486">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="57e53-486">(nullable) itemId: String</span></span>

<span data-ttu-id="57e53-487">現在のアイテムの[Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)を取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-487">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item.</span></span> <span data-ttu-id="57e53-488">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="57e53-488">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-489">`itemId`プロパティによって返される識別子は、 [Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)と同じです。</span><span class="sxs-lookup"><span data-stu-id="57e53-489">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="57e53-490">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="57e53-490">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="57e53-491">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="57e53-491">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="57e53-492">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57e53-492">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="57e53-p121">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="57e53-495">Type</span><span class="sxs-lookup"><span data-stu-id="57e53-495">Type</span></span>

*   <span data-ttu-id="57e53-496">String</span><span class="sxs-lookup"><span data-stu-id="57e53-496">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-497">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-497">Requirements</span></span>

|<span data-ttu-id="57e53-498">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-498">Requirement</span></span>|<span data-ttu-id="57e53-499">値</span><span class="sxs-lookup"><span data-stu-id="57e53-499">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-500">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-500">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-501">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-501">1.0</span></span>|
|[<span data-ttu-id="57e53-502">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-502">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-503">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-503">ReadItem</span></span>|
|[<span data-ttu-id="57e53-504">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-504">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-505">読み取り</span><span class="sxs-lookup"><span data-stu-id="57e53-505">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-506">例</span><span class="sxs-lookup"><span data-stu-id="57e53-506">Example</span></span>

<span data-ttu-id="57e53-p122">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-18"></a><span data-ttu-id="57e53-509">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-509">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-510">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-510">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="57e53-511">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="57e53-511">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="57e53-512">型</span><span class="sxs-lookup"><span data-stu-id="57e53-512">Type</span></span>

*   [<span data-ttu-id="57e53-513">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="57e53-513">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="57e53-514">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-514">Requirements</span></span>

|<span data-ttu-id="57e53-515">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-515">Requirement</span></span>|<span data-ttu-id="57e53-516">値</span><span class="sxs-lookup"><span data-stu-id="57e53-516">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-517">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-517">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-518">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-518">1.0</span></span>|
|[<span data-ttu-id="57e53-519">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-519">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-520">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-520">ReadItem</span></span>|
|[<span data-ttu-id="57e53-521">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-521">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-522">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-522">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-523">例</span><span class="sxs-lookup"><span data-stu-id="57e53-523">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-18"></a><span data-ttu-id="57e53-524">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-524">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-525">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="57e53-525">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57e53-526">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="57e53-526">Read mode</span></span>

<span data-ttu-id="57e53-527">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-527">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="57e53-528">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="57e53-528">Compose mode</span></span>

<span data-ttu-id="57e53-529">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-529">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="57e53-530">型</span><span class="sxs-lookup"><span data-stu-id="57e53-530">Type</span></span>

*   <span data-ttu-id="57e53-531">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-531">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-532">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-532">Requirements</span></span>

|<span data-ttu-id="57e53-533">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-533">Requirement</span></span>|<span data-ttu-id="57e53-534">値</span><span class="sxs-lookup"><span data-stu-id="57e53-534">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-535">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-535">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-536">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-536">1.0</span></span>|
|[<span data-ttu-id="57e53-537">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-537">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-538">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-538">ReadItem</span></span>|
|[<span data-ttu-id="57e53-539">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-539">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-540">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-540">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="57e53-541">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="57e53-541">normalizedSubject: String</span></span>

<span data-ttu-id="57e53-p123">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="57e53-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="57e53-p124">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="57e53-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="57e53-546">型</span><span class="sxs-lookup"><span data-stu-id="57e53-546">Type</span></span>

*   <span data-ttu-id="57e53-547">String</span><span class="sxs-lookup"><span data-stu-id="57e53-547">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-548">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-548">Requirements</span></span>

|<span data-ttu-id="57e53-549">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-549">Requirement</span></span>|<span data-ttu-id="57e53-550">値</span><span class="sxs-lookup"><span data-stu-id="57e53-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-551">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-552">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-552">1.0</span></span>|
|[<span data-ttu-id="57e53-553">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-553">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-554">ReadItem</span></span>|
|[<span data-ttu-id="57e53-555">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-555">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-556">読み取り</span><span class="sxs-lookup"><span data-stu-id="57e53-556">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-557">例</span><span class="sxs-lookup"><span data-stu-id="57e53-557">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-18"></a><span data-ttu-id="57e53-558">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-558">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-559">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-559">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="57e53-560">型</span><span class="sxs-lookup"><span data-stu-id="57e53-560">Type</span></span>

*   [<span data-ttu-id="57e53-561">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="57e53-561">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="57e53-562">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-562">Requirements</span></span>

|<span data-ttu-id="57e53-563">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-563">Requirement</span></span>|<span data-ttu-id="57e53-564">値</span><span class="sxs-lookup"><span data-stu-id="57e53-564">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-565">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-565">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-566">1.3</span><span class="sxs-lookup"><span data-stu-id="57e53-566">1.3</span></span>|
|[<span data-ttu-id="57e53-567">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-567">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-568">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-568">ReadItem</span></span>|
|[<span data-ttu-id="57e53-569">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-569">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-570">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-570">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-571">例</span><span class="sxs-lookup"><span data-stu-id="57e53-571">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="57e53-572">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-572">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-573">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="57e53-573">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="57e53-574">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="57e53-574">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57e53-575">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="57e53-575">Read mode</span></span>

<span data-ttu-id="57e53-576">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-576">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="57e53-577">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="57e53-577">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="57e53-578">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-578">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="57e53-579">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="57e53-579">Compose mode</span></span>

<span data-ttu-id="57e53-580">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-580">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="57e53-581">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="57e53-581">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="57e53-582">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-582">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="57e53-583">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-583">Get 500 members maximum.</span></span>
- <span data-ttu-id="57e53-584">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="57e53-584">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="57e53-585">型</span><span class="sxs-lookup"><span data-stu-id="57e53-585">Type</span></span>

*   <span data-ttu-id="57e53-586">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-586">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-587">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-587">Requirements</span></span>

|<span data-ttu-id="57e53-588">必要条件</span><span class="sxs-lookup"><span data-stu-id="57e53-588">Requirement</span></span>|<span data-ttu-id="57e53-589">値</span><span class="sxs-lookup"><span data-stu-id="57e53-589">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-590">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-590">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-591">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-591">1.0</span></span>|
|[<span data-ttu-id="57e53-592">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-592">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-593">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-593">ReadItem</span></span>|
|[<span data-ttu-id="57e53-594">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-594">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-595">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-595">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18organizerjavascriptapioutlookofficeorganizerviewoutlook-js-18"></a><span data-ttu-id="57e53-596">開催者: [emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[開催者](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-596">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-597">指定した会議の開催者の電子メールアドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-597">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57e53-598">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="57e53-598">Read mode</span></span>

<span data-ttu-id="57e53-599">プロパティ`organizer`は、会議の開催者を表す[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-599">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="57e53-600">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="57e53-600">Compose mode</span></span>

<span data-ttu-id="57e53-601">プロパティ`organizer`は、開催者の値を取得するためのメソッドを提供する[オーガナイザー](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-601">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.8) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="57e53-602">型</span><span class="sxs-lookup"><span data-stu-id="57e53-602">Type</span></span>

*   <span data-ttu-id="57e53-603">[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [開催者](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-603">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-604">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-604">Requirements</span></span>

|<span data-ttu-id="57e53-605">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-605">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="57e53-606">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-607">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-607">1.0</span></span>|<span data-ttu-id="57e53-608">1.7</span><span class="sxs-lookup"><span data-stu-id="57e53-608">1.7</span></span>|
|[<span data-ttu-id="57e53-609">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-610">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-610">ReadItem</span></span>|<span data-ttu-id="57e53-611">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="57e53-611">ReadWriteItem</span></span>|
|[<span data-ttu-id="57e53-612">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-613">Read</span><span class="sxs-lookup"><span data-stu-id="57e53-613">Read</span></span>|<span data-ttu-id="57e53-614">Compose</span><span class="sxs-lookup"><span data-stu-id="57e53-614">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-18"></a><span data-ttu-id="57e53-615">(nullable) 定期的なスケジュール:[定期的](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)なアイテム</span><span class="sxs-lookup"><span data-stu-id="57e53-615">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-616">予定の定期的なパターンを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="57e53-616">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="57e53-617">会議出席依頼の定期的なパターンを取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-617">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="57e53-618">予定アイテムの読み取りおよび作成モード。</span><span class="sxs-lookup"><span data-stu-id="57e53-618">Read and compose modes for appointment items.</span></span> <span data-ttu-id="57e53-619">会議出席依頼アイテムの閲覧モード。</span><span class="sxs-lookup"><span data-stu-id="57e53-619">Read mode for meeting request items.</span></span>

<span data-ttu-id="57e53-620">この`recurrence`プロパティは、アイテムが series または series 内のインスタンスの場合、定期的な予定または会議出席依頼に対して[定期的](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-620">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="57e53-621">`null`は、単一の予定および1つの予定の会議出席依頼に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-621">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="57e53-622">`undefined`は、会議出席依頼ではないメッセージに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-622">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="57e53-623">注: 会議出席依頼に`itemClass`は、IPM という値があります。出席依頼。</span><span class="sxs-lookup"><span data-stu-id="57e53-623">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="57e53-624">注: 定期的なオブジェクトが`null`の場合は、そのオブジェクトが単一の予定または1つの予定の会議出席依頼であり、データ系列の一部ではないことを示します。</span><span class="sxs-lookup"><span data-stu-id="57e53-624">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57e53-625">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="57e53-625">Read mode</span></span>

<span data-ttu-id="57e53-626">この`recurrence`プロパティは、定期的な予定を表す[定期的](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-626">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) object that represents the appointment recurrence.</span></span> <span data-ttu-id="57e53-627">これは、予定および会議出席依頼に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-627">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="57e53-628">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="57e53-628">Compose mode</span></span>

<span data-ttu-id="57e53-629">この`recurrence`プロパティは、予定の繰り返しを管理するためのメソッドを提供する[定期的](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-629">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="57e53-630">これは予定に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-630">This is available for appointments.</span></span>

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var recurrence = asyncResult.value;
  if (!recurrence) {
    console.log("One-time appointment or meeting");
  } else {
    console.log(JSON.stringify(recurrence));
  }
}

// The following example shows the results of the getAsync call that retrieves the recurrence for a series.
// NOTE: In this example, seriesTimeObject is a placeholder for the JSON representing the
// recurrence.seriesTime property. You should use the SeriesTime object's methods to get the
// recurrence date and time properties.
Recurrence = {
  "recurrenceType": "weekly",
  "recurrenceProperties": {"interval": 2, "days": ["mon","thu","fri"], "firstDayOfWeek": "sun"},
  "seriesTime": {seriesTimeObject},
  "recurrenceTimeZone": {"name": "Pacific Standard Time", "offset": -480}
}
```

##### <a name="type"></a><span data-ttu-id="57e53-631">型</span><span class="sxs-lookup"><span data-stu-id="57e53-631">Type</span></span>

* [<span data-ttu-id="57e53-632">繰り返さ</span><span class="sxs-lookup"><span data-stu-id="57e53-632">Recurrence</span></span>](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)

|<span data-ttu-id="57e53-633">必要条件</span><span class="sxs-lookup"><span data-stu-id="57e53-633">Requirement</span></span>|<span data-ttu-id="57e53-634">値</span><span class="sxs-lookup"><span data-stu-id="57e53-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-635">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-636">1.7</span><span class="sxs-lookup"><span data-stu-id="57e53-636">1.7</span></span>|
|[<span data-ttu-id="57e53-637">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-638">ReadItem</span></span>|
|[<span data-ttu-id="57e53-639">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-640">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-640">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="57e53-641">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-641">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-642">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="57e53-642">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="57e53-643">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="57e53-643">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57e53-644">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="57e53-644">Read mode</span></span>

<span data-ttu-id="57e53-645">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-645">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="57e53-646">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="57e53-646">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="57e53-647">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-647">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="57e53-648">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="57e53-648">Compose mode</span></span>

<span data-ttu-id="57e53-649">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-649">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="57e53-650">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="57e53-650">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="57e53-651">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-651">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="57e53-652">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-652">Get 500 members maximum.</span></span>
- <span data-ttu-id="57e53-653">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="57e53-653">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="57e53-654">型</span><span class="sxs-lookup"><span data-stu-id="57e53-654">Type</span></span>

*   <span data-ttu-id="57e53-655">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-655">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-656">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-656">Requirements</span></span>

|<span data-ttu-id="57e53-657">必要条件</span><span class="sxs-lookup"><span data-stu-id="57e53-657">Requirement</span></span>|<span data-ttu-id="57e53-658">値</span><span class="sxs-lookup"><span data-stu-id="57e53-658">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-659">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-659">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-660">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-660">1.0</span></span>|
|[<span data-ttu-id="57e53-661">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-661">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-662">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-662">ReadItem</span></span>|
|[<span data-ttu-id="57e53-663">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-663">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-664">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-664">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18"></a><span data-ttu-id="57e53-665">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-665">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-p135">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="57e53-p135">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="57e53-p136">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsfrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="57e53-p136">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-670">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="57e53-670">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="57e53-671">型</span><span class="sxs-lookup"><span data-stu-id="57e53-671">Type</span></span>

*   [<span data-ttu-id="57e53-672">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="57e53-672">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="57e53-673">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-673">Requirements</span></span>

|<span data-ttu-id="57e53-674">必要条件</span><span class="sxs-lookup"><span data-stu-id="57e53-674">Requirement</span></span>|<span data-ttu-id="57e53-675">値</span><span class="sxs-lookup"><span data-stu-id="57e53-675">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-676">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-676">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-677">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-677">1.0</span></span>|
|[<span data-ttu-id="57e53-678">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-678">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-679">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-679">ReadItem</span></span>|
|[<span data-ttu-id="57e53-680">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-680">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-681">読み取り</span><span class="sxs-lookup"><span data-stu-id="57e53-681">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-682">例</span><span class="sxs-lookup"><span data-stu-id="57e53-682">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="57e53-683">(nullable) 系列 Id: String</span><span class="sxs-lookup"><span data-stu-id="57e53-683">(nullable) seriesId: String</span></span>

<span data-ttu-id="57e53-684">インスタンスが属する系列の id を取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-684">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="57e53-685">Web 上の Outlook およびデスクトップクライアントでは、 `seriesId`は、このアイテムが属する親 (シリーズ) アイテムの Exchange web サービス (EWS) ID を返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-685">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="57e53-686">ただし、iOS と Android では、 `seriesId`は親アイテムの REST ID を返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-686">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-687">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="57e53-687">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="57e53-688">`seriesId`プロパティが OUTLOOK REST API で使用される outlook id と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="57e53-688">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="57e53-689">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="57e53-689">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="57e53-690">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57e53-690">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="57e53-691">この`seriesId`プロパティは`null` 、単一の予定、系列のアイテム、会議出席依頼などの親アイテムを持たないアイテムに`undefined`対して、会議出席依頼以外のアイテムに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-691">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="57e53-692">Type</span><span class="sxs-lookup"><span data-stu-id="57e53-692">Type</span></span>

* <span data-ttu-id="57e53-693">文字列</span><span class="sxs-lookup"><span data-stu-id="57e53-693">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-694">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-694">Requirements</span></span>

|<span data-ttu-id="57e53-695">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-695">Requirement</span></span>|<span data-ttu-id="57e53-696">値</span><span class="sxs-lookup"><span data-stu-id="57e53-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-697">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-698">1.7</span><span class="sxs-lookup"><span data-stu-id="57e53-698">1.7</span></span>|
|[<span data-ttu-id="57e53-699">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-699">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-700">ReadItem</span></span>|
|[<span data-ttu-id="57e53-701">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-701">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-702">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-702">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-703">例</span><span class="sxs-lookup"><span data-stu-id="57e53-703">Example</span></span>

```js
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-18"></a><span data-ttu-id="57e53-704">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-704">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-705">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="57e53-705">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="57e53-p139">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="57e53-p139">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57e53-708">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="57e53-708">Read mode</span></span>

<span data-ttu-id="57e53-709">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-709">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="57e53-710">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="57e53-710">Compose mode</span></span>

<span data-ttu-id="57e53-711">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-711">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="57e53-712">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="57e53-712">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="57e53-713">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="57e53-713">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="57e53-714">型</span><span class="sxs-lookup"><span data-stu-id="57e53-714">Type</span></span>

*   <span data-ttu-id="57e53-715">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-715">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-716">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-716">Requirements</span></span>

|<span data-ttu-id="57e53-717">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-717">Requirement</span></span>|<span data-ttu-id="57e53-718">値</span><span class="sxs-lookup"><span data-stu-id="57e53-718">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-719">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-719">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-720">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-720">1.0</span></span>|
|[<span data-ttu-id="57e53-721">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-721">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-722">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-722">ReadItem</span></span>|
|[<span data-ttu-id="57e53-723">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-723">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-724">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-724">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-18"></a><span data-ttu-id="57e53-725">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-725">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-726">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="57e53-726">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="57e53-727">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="57e53-727">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57e53-728">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="57e53-728">Read mode</span></span>

<span data-ttu-id="57e53-p140">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-p140">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="57e53-731">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="57e53-731">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="57e53-732">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="57e53-732">Compose mode</span></span>
<span data-ttu-id="57e53-733">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-733">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="57e53-734">型</span><span class="sxs-lookup"><span data-stu-id="57e53-734">Type</span></span>

*   <span data-ttu-id="57e53-735">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-735">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-736">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-736">Requirements</span></span>

|<span data-ttu-id="57e53-737">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-737">Requirement</span></span>|<span data-ttu-id="57e53-738">値</span><span class="sxs-lookup"><span data-stu-id="57e53-738">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-739">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-739">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-740">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-740">1.0</span></span>|
|[<span data-ttu-id="57e53-741">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-741">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-742">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-742">ReadItem</span></span>|
|[<span data-ttu-id="57e53-743">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-743">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-744">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-744">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="57e53-745">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-745">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-746">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="57e53-746">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="57e53-747">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="57e53-747">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57e53-748">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="57e53-748">Read mode</span></span>

<span data-ttu-id="57e53-749">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-749">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="57e53-750">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="57e53-750">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="57e53-751">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-751">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="57e53-752">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="57e53-752">Compose mode</span></span>

<span data-ttu-id="57e53-753">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-753">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="57e53-754">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="57e53-754">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="57e53-755">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-755">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="57e53-756">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-756">Get 500 members maximum.</span></span>
- <span data-ttu-id="57e53-757">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="57e53-757">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="57e53-758">型</span><span class="sxs-lookup"><span data-stu-id="57e53-758">Type</span></span>

*   <span data-ttu-id="57e53-759">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-759">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-760">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-760">Requirements</span></span>

|<span data-ttu-id="57e53-761">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-761">Requirement</span></span>|<span data-ttu-id="57e53-762">値</span><span class="sxs-lookup"><span data-stu-id="57e53-762">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-763">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-763">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-764">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-764">1.0</span></span>|
|[<span data-ttu-id="57e53-765">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-765">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-766">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-766">ReadItem</span></span>|
|[<span data-ttu-id="57e53-767">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-767">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-768">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-768">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="57e53-769">メソッド</span><span class="sxs-lookup"><span data-stu-id="57e53-769">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="57e53-770">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="57e53-770">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="57e53-771">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="57e53-771">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="57e53-772">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="57e53-772">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="57e53-773">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-773">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-774">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-774">Parameters</span></span>
|<span data-ttu-id="57e53-775">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-775">Name</span></span>|<span data-ttu-id="57e53-776">種類</span><span class="sxs-lookup"><span data-stu-id="57e53-776">Type</span></span>|<span data-ttu-id="57e53-777">属性</span><span class="sxs-lookup"><span data-stu-id="57e53-777">Attributes</span></span>|<span data-ttu-id="57e53-778">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-778">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="57e53-779">String</span><span class="sxs-lookup"><span data-stu-id="57e53-779">String</span></span>||<span data-ttu-id="57e53-p144">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="57e53-p144">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="57e53-782">String</span><span class="sxs-lookup"><span data-stu-id="57e53-782">String</span></span>||<span data-ttu-id="57e53-p145">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="57e53-p145">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="57e53-785">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-785">Object</span></span>|<span data-ttu-id="57e53-786">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-786">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-787">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="57e53-787">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57e53-788">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-788">Object</span></span>|<span data-ttu-id="57e53-789">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-789">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-790">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-790">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="57e53-791">Boolean</span><span class="sxs-lookup"><span data-stu-id="57e53-791">Boolean</span></span>|<span data-ttu-id="57e53-792">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-792">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-793">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="57e53-793">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="57e53-794">function</span><span class="sxs-lookup"><span data-stu-id="57e53-794">function</span></span>|<span data-ttu-id="57e53-795">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-795">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-796">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-796">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="57e53-797">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-797">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="57e53-798">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="57e53-798">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="57e53-799">エラー</span><span class="sxs-lookup"><span data-stu-id="57e53-799">Errors</span></span>

|<span data-ttu-id="57e53-800">エラー コード</span><span class="sxs-lookup"><span data-stu-id="57e53-800">Error code</span></span>|<span data-ttu-id="57e53-801">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-801">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="57e53-802">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="57e53-802">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="57e53-803">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="57e53-803">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="57e53-804">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="57e53-804">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-805">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-805">Requirements</span></span>

|<span data-ttu-id="57e53-806">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-806">Requirement</span></span>|<span data-ttu-id="57e53-807">値</span><span class="sxs-lookup"><span data-stu-id="57e53-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-808">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-809">1.1</span><span class="sxs-lookup"><span data-stu-id="57e53-809">1.1</span></span>|
|[<span data-ttu-id="57e53-810">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-811">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="57e53-811">ReadWriteItem</span></span>|
|[<span data-ttu-id="57e53-812">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-813">作成</span><span class="sxs-lookup"><span data-stu-id="57e53-813">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="57e53-814">例</span><span class="sxs-lookup"><span data-stu-id="57e53-814">Examples</span></span>

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

<span data-ttu-id="57e53-815">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="57e53-815">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
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

<br>

---
---

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="57e53-816">addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="57e53-816">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="57e53-817">Base64 エンコードのファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="57e53-817">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="57e53-818">この`addFileAttachmentFromBase64Async`メソッドは、base64 エンコードからファイルをアップロードし、新規作成フォームのアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="57e53-818">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="57e53-819">このメソッドは、AsyncResult オブジェクトの添付ファイル識別子を返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-819">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="57e53-820">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-820">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-821">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-821">Parameters</span></span>

|<span data-ttu-id="57e53-822">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-822">Name</span></span>|<span data-ttu-id="57e53-823">型</span><span class="sxs-lookup"><span data-stu-id="57e53-823">Type</span></span>|<span data-ttu-id="57e53-824">属性</span><span class="sxs-lookup"><span data-stu-id="57e53-824">Attributes</span></span>|<span data-ttu-id="57e53-825">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-825">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="57e53-826">String</span><span class="sxs-lookup"><span data-stu-id="57e53-826">String</span></span>||<span data-ttu-id="57e53-827">電子メールまたはイベントに追加する画像またはファイルの、base64 でエンコードされたコンテンツ。</span><span class="sxs-lookup"><span data-stu-id="57e53-827">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="57e53-828">String</span><span class="sxs-lookup"><span data-stu-id="57e53-828">String</span></span>||<span data-ttu-id="57e53-p147">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="57e53-p147">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="57e53-831">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-831">Object</span></span>|<span data-ttu-id="57e53-832">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-832">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-833">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="57e53-833">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57e53-834">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-834">Object</span></span>|<span data-ttu-id="57e53-835">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-835">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-836">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-836">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="57e53-837">Boolean</span><span class="sxs-lookup"><span data-stu-id="57e53-837">Boolean</span></span>|<span data-ttu-id="57e53-838">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-838">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-839">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="57e53-839">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="57e53-840">function</span><span class="sxs-lookup"><span data-stu-id="57e53-840">function</span></span>|<span data-ttu-id="57e53-841">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-841">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-842">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-842">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="57e53-843">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-843">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="57e53-844">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="57e53-844">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="57e53-845">エラー</span><span class="sxs-lookup"><span data-stu-id="57e53-845">Errors</span></span>

|<span data-ttu-id="57e53-846">エラー コード</span><span class="sxs-lookup"><span data-stu-id="57e53-846">Error code</span></span>|<span data-ttu-id="57e53-847">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-847">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="57e53-848">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="57e53-848">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="57e53-849">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="57e53-849">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="57e53-850">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="57e53-850">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-851">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-851">Requirements</span></span>

|<span data-ttu-id="57e53-852">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-852">Requirement</span></span>|<span data-ttu-id="57e53-853">値</span><span class="sxs-lookup"><span data-stu-id="57e53-853">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-854">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-854">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-855">1.8</span><span class="sxs-lookup"><span data-stu-id="57e53-855">1.8</span></span>|
|[<span data-ttu-id="57e53-856">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-856">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-857">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="57e53-857">ReadWriteItem</span></span>|
|[<span data-ttu-id="57e53-858">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-858">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-859">作成</span><span class="sxs-lookup"><span data-stu-id="57e53-859">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="57e53-860">例</span><span class="sxs-lookup"><span data-stu-id="57e53-860">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
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

<br>

---
---

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="57e53-861">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="57e53-861">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="57e53-862">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="57e53-862">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="57e53-863">現在、サポートされて`Office.EventType.AttachmentsChanged`いる`Office.EventType.AppointmentTimeChanged`イベント`Office.EventType.EnhancedLocationsChanged`の`Office.EventType.RecipientsChanged`種類は`Office.EventType.RecurrenceChanged`、、、、、です。</span><span class="sxs-lookup"><span data-stu-id="57e53-863">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-864">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-864">Parameters</span></span>

| <span data-ttu-id="57e53-865">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-865">Name</span></span> | <span data-ttu-id="57e53-866">型</span><span class="sxs-lookup"><span data-stu-id="57e53-866">Type</span></span> | <span data-ttu-id="57e53-867">属性</span><span class="sxs-lookup"><span data-stu-id="57e53-867">Attributes</span></span> | <span data-ttu-id="57e53-868">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-868">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="57e53-869">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="57e53-869">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="57e53-870">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="57e53-870">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="57e53-871">Function</span><span class="sxs-lookup"><span data-stu-id="57e53-871">Function</span></span> || <span data-ttu-id="57e53-p148">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="57e53-p148">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="57e53-875">Object</span><span class="sxs-lookup"><span data-stu-id="57e53-875">Object</span></span> | <span data-ttu-id="57e53-876">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-876">&lt;optional&gt;</span></span> | <span data-ttu-id="57e53-877">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="57e53-877">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="57e53-878">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-878">Object</span></span> | <span data-ttu-id="57e53-879">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-879">&lt;optional&gt;</span></span> | <span data-ttu-id="57e53-880">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-880">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="57e53-881">function</span><span class="sxs-lookup"><span data-stu-id="57e53-881">function</span></span>| <span data-ttu-id="57e53-882">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-882">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-883">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-883">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-884">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-884">Requirements</span></span>

|<span data-ttu-id="57e53-885">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-885">Requirement</span></span>| <span data-ttu-id="57e53-886">値</span><span class="sxs-lookup"><span data-stu-id="57e53-886">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-887">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-887">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="57e53-888">1.7</span><span class="sxs-lookup"><span data-stu-id="57e53-888">1.7</span></span> |
|[<span data-ttu-id="57e53-889">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-889">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="57e53-890">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-890">ReadItem</span></span> |
|[<span data-ttu-id="57e53-891">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-891">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57e53-892">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-892">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="57e53-893">例</span><span class="sxs-lookup"><span data-stu-id="57e53-893">Example</span></span>

```js
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="57e53-894">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="57e53-894">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="57e53-895">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="57e53-895">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="57e53-p149">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="57e53-p149">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="57e53-899">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-899">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="57e53-900">Office アドインを Outlook on the web で実行している場合、編集中のアイテム以外のアイテムに `addItemAttachmentAsync` メソッドでアイテムを添付できます。ただし、これはサポートされていないため、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="57e53-900">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-901">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-901">Parameters</span></span>

|<span data-ttu-id="57e53-902">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-902">Name</span></span>|<span data-ttu-id="57e53-903">型</span><span class="sxs-lookup"><span data-stu-id="57e53-903">Type</span></span>|<span data-ttu-id="57e53-904">属性</span><span class="sxs-lookup"><span data-stu-id="57e53-904">Attributes</span></span>|<span data-ttu-id="57e53-905">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-905">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="57e53-906">String</span><span class="sxs-lookup"><span data-stu-id="57e53-906">String</span></span>||<span data-ttu-id="57e53-p150">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="57e53-p150">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="57e53-909">String</span><span class="sxs-lookup"><span data-stu-id="57e53-909">String</span></span>||<span data-ttu-id="57e53-910">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="57e53-910">The subject of the item to be attached.</span></span> <span data-ttu-id="57e53-911">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="57e53-911">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="57e53-912">Object</span><span class="sxs-lookup"><span data-stu-id="57e53-912">Object</span></span>|<span data-ttu-id="57e53-913">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-913">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-914">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="57e53-914">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57e53-915">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-915">Object</span></span>|<span data-ttu-id="57e53-916">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-916">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-917">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-917">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="57e53-918">関数</span><span class="sxs-lookup"><span data-stu-id="57e53-918">function</span></span>|<span data-ttu-id="57e53-919">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-919">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-920">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-920">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="57e53-921">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-921">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="57e53-922">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="57e53-922">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="57e53-923">エラー</span><span class="sxs-lookup"><span data-stu-id="57e53-923">Errors</span></span>

|<span data-ttu-id="57e53-924">エラー コード</span><span class="sxs-lookup"><span data-stu-id="57e53-924">Error code</span></span>|<span data-ttu-id="57e53-925">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-925">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="57e53-926">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="57e53-926">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-927">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-927">Requirements</span></span>

|<span data-ttu-id="57e53-928">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-928">Requirement</span></span>|<span data-ttu-id="57e53-929">値</span><span class="sxs-lookup"><span data-stu-id="57e53-929">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-930">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-930">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-931">1.1</span><span class="sxs-lookup"><span data-stu-id="57e53-931">1.1</span></span>|
|[<span data-ttu-id="57e53-932">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-932">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-933">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="57e53-933">ReadWriteItem</span></span>|
|[<span data-ttu-id="57e53-934">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-934">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-935">作成</span><span class="sxs-lookup"><span data-stu-id="57e53-935">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-936">例</span><span class="sxs-lookup"><span data-stu-id="57e53-936">Example</span></span>

<span data-ttu-id="57e53-937">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-937">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="57e53-938">close()</span><span class="sxs-lookup"><span data-stu-id="57e53-938">close()</span></span>

<span data-ttu-id="57e53-939">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="57e53-939">Closes the current item that is being composed.</span></span>

<span data-ttu-id="57e53-p152">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="57e53-p152">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-942">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-942">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="57e53-943">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="57e53-943">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-944">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-944">Requirements</span></span>

|<span data-ttu-id="57e53-945">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-945">Requirement</span></span>|<span data-ttu-id="57e53-946">値</span><span class="sxs-lookup"><span data-stu-id="57e53-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-947">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-948">1.3</span><span class="sxs-lookup"><span data-stu-id="57e53-948">1.3</span></span>|
|[<span data-ttu-id="57e53-949">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-949">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-950">制限あり</span><span class="sxs-lookup"><span data-stu-id="57e53-950">Restricted</span></span>|
|[<span data-ttu-id="57e53-951">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-951">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-952">新規作成</span><span class="sxs-lookup"><span data-stu-id="57e53-952">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="57e53-953">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="57e53-953">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="57e53-954">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-954">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-955">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="57e53-955">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="57e53-956">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-956">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="57e53-957">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="57e53-957">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="57e53-p153">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="57e53-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-961">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-961">Parameters</span></span>

|<span data-ttu-id="57e53-962">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-962">Name</span></span>|<span data-ttu-id="57e53-963">型</span><span class="sxs-lookup"><span data-stu-id="57e53-963">Type</span></span>|<span data-ttu-id="57e53-964">属性</span><span class="sxs-lookup"><span data-stu-id="57e53-964">Attributes</span></span>|<span data-ttu-id="57e53-965">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-965">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="57e53-966">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="57e53-966">String &#124; Object</span></span>||<span data-ttu-id="57e53-p154">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="57e53-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="57e53-969">**または**</span><span class="sxs-lookup"><span data-stu-id="57e53-969">**OR**</span></span><br/><span data-ttu-id="57e53-p155">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="57e53-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="57e53-972">String</span><span class="sxs-lookup"><span data-stu-id="57e53-972">String</span></span>|<span data-ttu-id="57e53-973">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-973">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-p156">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="57e53-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="57e53-976">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-976">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="57e53-977">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-977">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-978">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="57e53-978">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="57e53-979">String</span><span class="sxs-lookup"><span data-stu-id="57e53-979">String</span></span>||<span data-ttu-id="57e53-p157">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="57e53-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="57e53-982">String</span><span class="sxs-lookup"><span data-stu-id="57e53-982">String</span></span>||<span data-ttu-id="57e53-983">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="57e53-983">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="57e53-984">文字列</span><span class="sxs-lookup"><span data-stu-id="57e53-984">String</span></span>||<span data-ttu-id="57e53-p158">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="57e53-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="57e53-987">ブール値</span><span class="sxs-lookup"><span data-stu-id="57e53-987">Boolean</span></span>||<span data-ttu-id="57e53-p159">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="57e53-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="57e53-990">String</span><span class="sxs-lookup"><span data-stu-id="57e53-990">String</span></span>||<span data-ttu-id="57e53-p160">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="57e53-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="57e53-994">function</span><span class="sxs-lookup"><span data-stu-id="57e53-994">function</span></span>|<span data-ttu-id="57e53-995">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-995">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-996">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-996">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-997">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-997">Requirements</span></span>

|<span data-ttu-id="57e53-998">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-998">Requirement</span></span>|<span data-ttu-id="57e53-999">値</span><span class="sxs-lookup"><span data-stu-id="57e53-999">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1000">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1000">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1001">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-1001">1.0</span></span>|
|[<span data-ttu-id="57e53-1002">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1002">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1003">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1003">ReadItem</span></span>|
|[<span data-ttu-id="57e53-1004">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1004">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1005">読み取り</span><span class="sxs-lookup"><span data-stu-id="57e53-1005">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="57e53-1006">例</span><span class="sxs-lookup"><span data-stu-id="57e53-1006">Examples</span></span>

<span data-ttu-id="57e53-1007">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1007">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="57e53-1008">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1008">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="57e53-1009">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1009">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="57e53-1010">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1010">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="57e53-1011">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1011">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="57e53-1012">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1012">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="57e53-1013">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="57e53-1013">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="57e53-1014">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1014">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-1015">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="57e53-1015">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="57e53-1016">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1016">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="57e53-1017">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="57e53-1017">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="57e53-p161">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="57e53-p161">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-1021">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-1021">Parameters</span></span>

|<span data-ttu-id="57e53-1022">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-1022">Name</span></span>|<span data-ttu-id="57e53-1023">型</span><span class="sxs-lookup"><span data-stu-id="57e53-1023">Type</span></span>|<span data-ttu-id="57e53-1024">属性</span><span class="sxs-lookup"><span data-stu-id="57e53-1024">Attributes</span></span>|<span data-ttu-id="57e53-1025">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-1025">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="57e53-1026">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="57e53-1026">String &#124; Object</span></span>||<span data-ttu-id="57e53-p162">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="57e53-p162">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="57e53-1029">**または**</span><span class="sxs-lookup"><span data-stu-id="57e53-1029">**OR**</span></span><br/><span data-ttu-id="57e53-p163">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="57e53-p163">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="57e53-1032">String</span><span class="sxs-lookup"><span data-stu-id="57e53-1032">String</span></span>|<span data-ttu-id="57e53-1033">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1033">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-p164">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="57e53-p164">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="57e53-1036">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1036">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="57e53-1037">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1037">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1038">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="57e53-1038">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="57e53-1039">String</span><span class="sxs-lookup"><span data-stu-id="57e53-1039">String</span></span>||<span data-ttu-id="57e53-p165">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="57e53-p165">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="57e53-1042">String</span><span class="sxs-lookup"><span data-stu-id="57e53-1042">String</span></span>||<span data-ttu-id="57e53-1043">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="57e53-1043">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="57e53-1044">文字列</span><span class="sxs-lookup"><span data-stu-id="57e53-1044">String</span></span>||<span data-ttu-id="57e53-p166">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="57e53-p166">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="57e53-1047">ブール値</span><span class="sxs-lookup"><span data-stu-id="57e53-1047">Boolean</span></span>||<span data-ttu-id="57e53-p167">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="57e53-p167">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="57e53-1050">String</span><span class="sxs-lookup"><span data-stu-id="57e53-1050">String</span></span>||<span data-ttu-id="57e53-p168">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="57e53-p168">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="57e53-1054">function</span><span class="sxs-lookup"><span data-stu-id="57e53-1054">function</span></span>|<span data-ttu-id="57e53-1055">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1055">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1056">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1056">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-1057">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1057">Requirements</span></span>

|<span data-ttu-id="57e53-1058">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1058">Requirement</span></span>|<span data-ttu-id="57e53-1059">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1059">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1060">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1060">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1061">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-1061">1.0</span></span>|
|[<span data-ttu-id="57e53-1062">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1062">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1063">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1063">ReadItem</span></span>|
|[<span data-ttu-id="57e53-1064">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1064">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1065">読み取り</span><span class="sxs-lookup"><span data-stu-id="57e53-1065">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="57e53-1066">例</span><span class="sxs-lookup"><span data-stu-id="57e53-1066">Examples</span></span>

<span data-ttu-id="57e53-1067">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1067">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="57e53-1068">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1068">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="57e53-1069">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1069">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="57e53-1070">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1070">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="57e53-1071">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1071">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="57e53-1072">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1072">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getallinternetheadersasyncoptions-callback"></a><span data-ttu-id="57e53-1073">getAllInternetHeadersAsync ([オプション], [callback])</span><span class="sxs-lookup"><span data-stu-id="57e53-1073">getAllInternetHeadersAsync([options], [callback])</span></span>

<span data-ttu-id="57e53-1074">メッセージのすべてのインターネットヘッダーを文字列として取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1074">Gets all the internet headers for the message as a string.</span></span> <span data-ttu-id="57e53-1075">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="57e53-1075">Read mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-1076">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-1076">Parameters</span></span>

|<span data-ttu-id="57e53-1077">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-1077">Name</span></span>|<span data-ttu-id="57e53-1078">型</span><span class="sxs-lookup"><span data-stu-id="57e53-1078">Type</span></span>|<span data-ttu-id="57e53-1079">属性</span><span class="sxs-lookup"><span data-stu-id="57e53-1079">Attributes</span></span>|<span data-ttu-id="57e53-1080">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-1080">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="57e53-1081">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1081">Object</span></span>|<span data-ttu-id="57e53-1082">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1082">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1083">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="57e53-1083">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57e53-1084">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1084">Object</span></span>|<span data-ttu-id="57e53-1085">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1085">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1086">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1086">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="57e53-1087">関数</span><span class="sxs-lookup"><span data-stu-id="57e53-1087">function</span></span>|<span data-ttu-id="57e53-1088">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1088">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1089">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1089">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="57e53-1090">成功した場合、インターネットヘッダーデータは、文字列として asyncResult プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1090">On success, the internet headers data is provided in the asyncResult.value property as a string.</span></span> <span data-ttu-id="57e53-1091">返される文字列値の書式情報については、 [RFC 2183](https://tools.ietf.org/html/rfc2183)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57e53-1091">Refer to [RFC 2183](https://tools.ietf.org/html/rfc2183) for the formatting information of the returned string value.</span></span> <span data-ttu-id="57e53-1092">呼び出しが失敗した場合、asyncResult. error プロパティには、エラーの理由と共にエラーコードが含まれます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1092">If the call fails, the asyncResult.error property will contain an error code with the reason for the failure.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-1093">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-1093">Requirements</span></span>

|<span data-ttu-id="57e53-1094">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1094">Requirement</span></span>|<span data-ttu-id="57e53-1095">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1095">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1096">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1096">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1097">1.8</span><span class="sxs-lookup"><span data-stu-id="57e53-1097">1.8</span></span>|
|[<span data-ttu-id="57e53-1098">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1098">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1099">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1099">ReadItem</span></span>|
|[<span data-ttu-id="57e53-1100">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1100">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1101">読み取り</span><span class="sxs-lookup"><span data-stu-id="57e53-1101">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="57e53-1102">戻り値:</span><span class="sxs-lookup"><span data-stu-id="57e53-1102">Returns:</span></span>

<span data-ttu-id="57e53-1103">[RFC 2183](https://tools.ietf.org/html/rfc2183)に従って書式設定された文字列としてのインターネットヘッダーデータ。</span><span class="sxs-lookup"><span data-stu-id="57e53-1103">The internet headers data as a string formatted according to [RFC 2183](https://tools.ietf.org/html/rfc2183).</span></span>

<span data-ttu-id="57e53-1104">型:String</span><span class="sxs-lookup"><span data-stu-id="57e53-1104">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="57e53-1105">例</span><span class="sxs-lookup"><span data-stu-id="57e53-1105">Example</span></span>

```js
// Get the internet headers related to the mail.
Office.context.mailbox.item.getAllInternetHeadersAsync(
  function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(asyncResult.value);
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

<br>

---
---

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontentviewoutlook-js-18"></a><span data-ttu-id="57e53-1106">getAttachmentContentAsync (attachmentId, [options], [callback]) > [Attachmentcontent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-1106">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span></span>

<span data-ttu-id="57e53-1107">メッセージまたは予定から指定された添付ファイルを取得し`AttachmentContent` 、それをオブジェクトとして返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1107">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="57e53-1108">メソッド`getAttachmentContentAsync`は、指定された id の添付ファイルをアイテムから取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1108">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="57e53-1109">ベストプラクティスとして、識別子を使用して、または`getAttachmentsAsync` `item.attachments`の呼び出しで attachmentIds を取得したのと同じセッションの添付ファイルを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="57e53-1109">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="57e53-1110">Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="57e53-1110">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="57e53-1111">ユーザーがアプリを閉じたとき、またはインラインフォームの作成が開始されたときに、別のウィンドウで続行するためにフォームをポップアウトした後、セッションが終了します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1111">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-1112">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-1112">Parameters</span></span>

|<span data-ttu-id="57e53-1113">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-1113">Name</span></span>|<span data-ttu-id="57e53-1114">型</span><span class="sxs-lookup"><span data-stu-id="57e53-1114">Type</span></span>|<span data-ttu-id="57e53-1115">属性</span><span class="sxs-lookup"><span data-stu-id="57e53-1115">Attributes</span></span>|<span data-ttu-id="57e53-1116">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-1116">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="57e53-1117">String</span><span class="sxs-lookup"><span data-stu-id="57e53-1117">String</span></span>||<span data-ttu-id="57e53-1118">取得する添付ファイルの識別子を指定します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1118">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="57e53-1119">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1119">Object</span></span>|<span data-ttu-id="57e53-1120">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1120">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1121">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="57e53-1121">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57e53-1122">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1122">Object</span></span>|<span data-ttu-id="57e53-1123">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1123">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1124">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1124">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="57e53-1125">関数</span><span class="sxs-lookup"><span data-stu-id="57e53-1125">function</span></span>|<span data-ttu-id="57e53-1126">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1126">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1127">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1127">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-1128">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1128">Requirements</span></span>

|<span data-ttu-id="57e53-1129">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1129">Requirement</span></span>|<span data-ttu-id="57e53-1130">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1130">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1131">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1132">1.8</span><span class="sxs-lookup"><span data-stu-id="57e53-1132">1.8</span></span>|
|[<span data-ttu-id="57e53-1133">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1133">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1134">ReadItem</span></span>|
|[<span data-ttu-id="57e53-1135">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1135">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1136">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-1136">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="57e53-1137">戻り値:</span><span class="sxs-lookup"><span data-stu-id="57e53-1137">Returns:</span></span>

<span data-ttu-id="57e53-1138">型: [Attachmentcontent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-1138">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span></span>

##### <a name="example"></a><span data-ttu-id="57e53-1139">例</span><span class="sxs-lookup"><span data-stu-id="57e53-1139">Example</span></span>

```js
var item = Office.context.mailbox.item;
var listOfAttachments = [];
var options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
      // Handle attachment formats that are not supported.
  }
}
```

<br>

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-18"></a><span data-ttu-id="57e53-1140">getAttachmentsAsync ([オプション], [callback]) > Array. <[Attachmentdetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="57e53-1140">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

<span data-ttu-id="57e53-1141">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1141">Gets the item's attachments as an array.</span></span> <span data-ttu-id="57e53-1142">新規作成モードのみです。</span><span class="sxs-lookup"><span data-stu-id="57e53-1142">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-1143">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-1143">Parameters</span></span>

|<span data-ttu-id="57e53-1144">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-1144">Name</span></span>|<span data-ttu-id="57e53-1145">型</span><span class="sxs-lookup"><span data-stu-id="57e53-1145">Type</span></span>|<span data-ttu-id="57e53-1146">属性</span><span class="sxs-lookup"><span data-stu-id="57e53-1146">Attributes</span></span>|<span data-ttu-id="57e53-1147">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-1147">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="57e53-1148">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1148">Object</span></span>|<span data-ttu-id="57e53-1149">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1149">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1150">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="57e53-1150">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57e53-1151">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1151">Object</span></span>|<span data-ttu-id="57e53-1152">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1152">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1153">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1153">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="57e53-1154">function</span><span class="sxs-lookup"><span data-stu-id="57e53-1154">function</span></span>|<span data-ttu-id="57e53-1155">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1155">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1156">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1156">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-1157">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-1157">Requirements</span></span>

|<span data-ttu-id="57e53-1158">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1158">Requirement</span></span>|<span data-ttu-id="57e53-1159">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1159">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1160">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1161">1.8</span><span class="sxs-lookup"><span data-stu-id="57e53-1161">1.8</span></span>|
|[<span data-ttu-id="57e53-1162">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1162">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1163">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1163">ReadItem</span></span>|
|[<span data-ttu-id="57e53-1164">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1164">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1165">作成</span><span class="sxs-lookup"><span data-stu-id="57e53-1165">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="57e53-1166">戻り値:</span><span class="sxs-lookup"><span data-stu-id="57e53-1166">Returns:</span></span>

<span data-ttu-id="57e53-1167">型: Array. <[attachmentdetails 詳細](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="57e53-1167">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

##### <a name="example"></a><span data-ttu-id="57e53-1168">例</span><span class="sxs-lookup"><span data-stu-id="57e53-1168">Example</span></span>

<span data-ttu-id="57e53-1169">次の例では、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1169">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```js
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var attachment = result.value [i];
      outputString += "<BR>" + i + ". Name: ";
      outputString += attachment.name;
      outputString += "<BR>ID: " + attachment.id;
      outputString += "<BR>contentType: " + attachment.contentType;
      outputString += "<BR>size: " + attachment.size;
      outputString += "<BR>attachmentType: " + attachment.attachmentType;
      outputString += "<BR>isInline: " + attachment.isInline;
    }
  }
}
```

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-18"></a><span data-ttu-id="57e53-1170">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span><span class="sxs-lookup"><span data-stu-id="57e53-1170">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span></span>

<span data-ttu-id="57e53-1171">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1171">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-1172">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="57e53-1172">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-1173">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1173">Requirements</span></span>

|<span data-ttu-id="57e53-1174">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1174">Requirement</span></span>|<span data-ttu-id="57e53-1175">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1175">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1176">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1177">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-1177">1.0</span></span>|
|[<span data-ttu-id="57e53-1178">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1178">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1179">ReadItem</span></span>|
|[<span data-ttu-id="57e53-1180">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1181">読み取り</span><span class="sxs-lookup"><span data-stu-id="57e53-1181">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="57e53-1182">戻り値:</span><span class="sxs-lookup"><span data-stu-id="57e53-1182">Returns:</span></span>

<span data-ttu-id="57e53-1183">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-1183">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span></span>

##### <a name="example"></a><span data-ttu-id="57e53-1184">例</span><span class="sxs-lookup"><span data-stu-id="57e53-1184">Example</span></span>

<span data-ttu-id="57e53-1185">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="57e53-1185">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-18meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-18phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-18tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-18"></a><span data-ttu-id="57e53-1186">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span><span class="sxs-lookup"><span data-stu-id="57e53-1186">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span></span>

<span data-ttu-id="57e53-1187">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1187">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-1188">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="57e53-1188">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-1189">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-1189">Parameters</span></span>

|<span data-ttu-id="57e53-1190">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-1190">Name</span></span>|<span data-ttu-id="57e53-1191">種類</span><span class="sxs-lookup"><span data-stu-id="57e53-1191">Type</span></span>|<span data-ttu-id="57e53-1192">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-1192">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="57e53-1193">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="57e53-1193">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.8)|<span data-ttu-id="57e53-1194">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="57e53-1194">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-1195">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-1195">Requirements</span></span>

|<span data-ttu-id="57e53-1196">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1196">Requirement</span></span>|<span data-ttu-id="57e53-1197">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1197">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1198">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1199">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-1199">1.0</span></span>|
|[<span data-ttu-id="57e53-1200">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1201">制限あり</span><span class="sxs-lookup"><span data-stu-id="57e53-1201">Restricted</span></span>|
|[<span data-ttu-id="57e53-1202">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1203">読み取り</span><span class="sxs-lookup"><span data-stu-id="57e53-1203">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="57e53-1204">戻り値:</span><span class="sxs-lookup"><span data-stu-id="57e53-1204">Returns:</span></span>

<span data-ttu-id="57e53-1205">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1205">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="57e53-1206">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1206">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="57e53-1207">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="57e53-1207">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="57e53-1208">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="57e53-1208">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="57e53-1209">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="57e53-1209">Value of `entityType`</span></span>|<span data-ttu-id="57e53-1210">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="57e53-1210">Type of objects in returned array</span></span>|<span data-ttu-id="57e53-1211">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1211">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="57e53-1212">文字列</span><span class="sxs-lookup"><span data-stu-id="57e53-1212">String</span></span>|<span data-ttu-id="57e53-1213">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="57e53-1213">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="57e53-1214">連絡先</span><span class="sxs-lookup"><span data-stu-id="57e53-1214">Contact</span></span>|<span data-ttu-id="57e53-1215">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="57e53-1215">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="57e53-1216">文字列</span><span class="sxs-lookup"><span data-stu-id="57e53-1216">String</span></span>|<span data-ttu-id="57e53-1217">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="57e53-1217">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="57e53-1218">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="57e53-1218">MeetingSuggestion</span></span>|<span data-ttu-id="57e53-1219">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="57e53-1219">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="57e53-1220">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="57e53-1220">PhoneNumber</span></span>|<span data-ttu-id="57e53-1221">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="57e53-1221">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="57e53-1222">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="57e53-1222">TaskSuggestion</span></span>|<span data-ttu-id="57e53-1223">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="57e53-1223">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="57e53-1224">文字列</span><span class="sxs-lookup"><span data-stu-id="57e53-1224">String</span></span>|<span data-ttu-id="57e53-1225">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="57e53-1225">**Restricted**</span></span>|

<span data-ttu-id="57e53-1226">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span><span class="sxs-lookup"><span data-stu-id="57e53-1226">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span></span>

##### <a name="example"></a><span data-ttu-id="57e53-1227">例</span><span class="sxs-lookup"><span data-stu-id="57e53-1227">Example</span></span>

<span data-ttu-id="57e53-1228">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1228">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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
};
```

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-18meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-18phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-18tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-18"></a><span data-ttu-id="57e53-1229">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span><span class="sxs-lookup"><span data-stu-id="57e53-1229">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span></span>

<span data-ttu-id="57e53-1230">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1230">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-1231">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="57e53-1231">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="57e53-1232">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1232">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-1233">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-1233">Parameters</span></span>

|<span data-ttu-id="57e53-1234">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-1234">Name</span></span>|<span data-ttu-id="57e53-1235">種類</span><span class="sxs-lookup"><span data-stu-id="57e53-1235">Type</span></span>|<span data-ttu-id="57e53-1236">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-1236">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="57e53-1237">String</span><span class="sxs-lookup"><span data-stu-id="57e53-1237">String</span></span>|<span data-ttu-id="57e53-1238">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="57e53-1238">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-1239">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-1239">Requirements</span></span>

|<span data-ttu-id="57e53-1240">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1240">Requirement</span></span>|<span data-ttu-id="57e53-1241">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1241">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1242">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1243">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-1243">1.0</span></span>|
|[<span data-ttu-id="57e53-1244">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1244">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1245">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1245">ReadItem</span></span>|
|[<span data-ttu-id="57e53-1246">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1246">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1247">読み取り</span><span class="sxs-lookup"><span data-stu-id="57e53-1247">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="57e53-1248">戻り値:</span><span class="sxs-lookup"><span data-stu-id="57e53-1248">Returns:</span></span>

<span data-ttu-id="57e53-p174">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-p174">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="57e53-1251">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span><span class="sxs-lookup"><span data-stu-id="57e53-1251">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span></span>

<br>

---
---

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="57e53-1252">getItemIdAsync ([オプション], callback)</span><span class="sxs-lookup"><span data-stu-id="57e53-1252">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="57e53-1253">保存されたアイテムの ID を非同期に取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1253">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="57e53-1254">新規作成モードのみです。</span><span class="sxs-lookup"><span data-stu-id="57e53-1254">Compose mode only.</span></span>

<span data-ttu-id="57e53-1255">このメソッドを呼び出すと、コールバックメソッドによってアイテム ID が返されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1255">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-1256">アドインが新規作成モードの`getItemIdAsync`アイテムに対して呼び出しを行う場合 ( `itemId` EWS または REST API を使用するため)、Outlook がキャッシュモードの場合は、アイテムがサーバーに同期されるまでしばらく時間がかかる場合があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="57e53-1256">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="57e53-1257">アイテムが同期されるまで、 `itemId`は認識されず、を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1257">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-1258">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-1258">Parameters</span></span>

|<span data-ttu-id="57e53-1259">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-1259">Name</span></span>|<span data-ttu-id="57e53-1260">型</span><span class="sxs-lookup"><span data-stu-id="57e53-1260">Type</span></span>|<span data-ttu-id="57e53-1261">属性</span><span class="sxs-lookup"><span data-stu-id="57e53-1261">Attributes</span></span>|<span data-ttu-id="57e53-1262">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-1262">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="57e53-1263">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1263">Object</span></span>|<span data-ttu-id="57e53-1264">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1264">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1265">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="57e53-1265">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57e53-1266">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1266">Object</span></span>|<span data-ttu-id="57e53-1267">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1268">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1268">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="57e53-1269">function</span><span class="sxs-lookup"><span data-stu-id="57e53-1269">function</span></span>||<span data-ttu-id="57e53-1270">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1270">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="57e53-1271">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1271">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="57e53-1272">エラー</span><span class="sxs-lookup"><span data-stu-id="57e53-1272">Errors</span></span>

|<span data-ttu-id="57e53-1273">エラー コード</span><span class="sxs-lookup"><span data-stu-id="57e53-1273">Error code</span></span>|<span data-ttu-id="57e53-1274">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-1274">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="57e53-1275">この id は、アイテムが保存されるまでは取得できません。</span><span class="sxs-lookup"><span data-stu-id="57e53-1275">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-1276">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-1276">Requirements</span></span>

|<span data-ttu-id="57e53-1277">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1277">Requirement</span></span>|<span data-ttu-id="57e53-1278">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1278">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1279">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1280">1.8</span><span class="sxs-lookup"><span data-stu-id="57e53-1280">1.8</span></span>|
|[<span data-ttu-id="57e53-1281">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1281">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1282">ReadItem</span></span>|
|[<span data-ttu-id="57e53-1283">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1283">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1284">作成</span><span class="sxs-lookup"><span data-stu-id="57e53-1284">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="57e53-1285">例</span><span class="sxs-lookup"><span data-stu-id="57e53-1285">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="57e53-1286">次の例は、コールバック関数`result`に渡されるパラメーターの構造を示しています。</span><span class="sxs-lookup"><span data-stu-id="57e53-1286">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="57e53-1287">プロパティ`value`には、アイテムの ID が含まれています。</span><span class="sxs-lookup"><span data-stu-id="57e53-1287">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="57e53-1288">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="57e53-1288">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="57e53-1289">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1289">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-1290">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="57e53-1290">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="57e53-p178">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="57e53-p178">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="57e53-1294">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="57e53-1294">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="57e53-1295">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="57e53-1295">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="57e53-p179">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-1299">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1299">Requirements</span></span>

|<span data-ttu-id="57e53-1300">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1300">Requirement</span></span>|<span data-ttu-id="57e53-1301">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1301">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1302">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1303">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-1303">1.0</span></span>|
|[<span data-ttu-id="57e53-1304">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1305">ReadItem</span></span>|
|[<span data-ttu-id="57e53-1306">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1307">読み取り</span><span class="sxs-lookup"><span data-stu-id="57e53-1307">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="57e53-1308">戻り値:</span><span class="sxs-lookup"><span data-stu-id="57e53-1308">Returns:</span></span>

<span data-ttu-id="57e53-p180">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="57e53-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="57e53-1311">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="57e53-1311">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="57e53-1312">Object</span><span class="sxs-lookup"><span data-stu-id="57e53-1312">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="57e53-1313">例</span><span class="sxs-lookup"><span data-stu-id="57e53-1313">Example</span></span>

<span data-ttu-id="57e53-1314">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="57e53-1314">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="57e53-1315">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="57e53-1315">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="57e53-1316">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1316">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-1317">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="57e53-1317">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="57e53-1318">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="57e53-1318">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="57e53-p181">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="57e53-p181">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-1321">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-1321">Parameters</span></span>

|<span data-ttu-id="57e53-1322">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-1322">Name</span></span>|<span data-ttu-id="57e53-1323">種類</span><span class="sxs-lookup"><span data-stu-id="57e53-1323">Type</span></span>|<span data-ttu-id="57e53-1324">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-1324">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="57e53-1325">String</span><span class="sxs-lookup"><span data-stu-id="57e53-1325">String</span></span>|<span data-ttu-id="57e53-1326">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="57e53-1326">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-1327">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-1327">Requirements</span></span>

|<span data-ttu-id="57e53-1328">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1328">Requirement</span></span>|<span data-ttu-id="57e53-1329">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1329">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1330">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1330">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1331">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-1331">1.0</span></span>|
|[<span data-ttu-id="57e53-1332">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1332">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1333">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1333">ReadItem</span></span>|
|[<span data-ttu-id="57e53-1334">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1334">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1335">読み取り</span><span class="sxs-lookup"><span data-stu-id="57e53-1335">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="57e53-1336">戻り値:</span><span class="sxs-lookup"><span data-stu-id="57e53-1336">Returns:</span></span>

<span data-ttu-id="57e53-1337">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="57e53-1337">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="57e53-1338">型: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="57e53-1338">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="57e53-1339">例</span><span class="sxs-lookup"><span data-stu-id="57e53-1339">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="57e53-1340">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="57e53-1340">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="57e53-1341">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1341">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="57e53-p182">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-p182">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-1344">Web 上の Outlook では、テキストが選択されておらず、カーソルが本文にある場合、このメソッドは文字列 "null" を返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1344">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="57e53-1345">このような状況を確認するには、次のようなコードを含めます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1345">To check for this situation, include code similar to the following:</span></span>
>
> `var selectedText = (asyncResult.value.endPosition === asyncResult.value.startPosition) ? "" : asyncResult.value.data;`

##### <a name="parameters"></a><span data-ttu-id="57e53-1346">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-1346">Parameters</span></span>

|<span data-ttu-id="57e53-1347">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-1347">Name</span></span>|<span data-ttu-id="57e53-1348">型</span><span class="sxs-lookup"><span data-stu-id="57e53-1348">Type</span></span>|<span data-ttu-id="57e53-1349">属性</span><span class="sxs-lookup"><span data-stu-id="57e53-1349">Attributes</span></span>|<span data-ttu-id="57e53-1350">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-1350">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="57e53-1351">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="57e53-1351">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="57e53-p184">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-p184">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="57e53-1355">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1355">Object</span></span>|<span data-ttu-id="57e53-1356">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1356">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1357">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="57e53-1357">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57e53-1358">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1358">Object</span></span>|<span data-ttu-id="57e53-1359">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1359">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1360">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1360">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="57e53-1361">function</span><span class="sxs-lookup"><span data-stu-id="57e53-1361">function</span></span>||<span data-ttu-id="57e53-1362">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1362">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="57e53-1363">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1363">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="57e53-1364">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="57e53-1364">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-1365">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-1365">Requirements</span></span>

|<span data-ttu-id="57e53-1366">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1366">Requirement</span></span>|<span data-ttu-id="57e53-1367">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1367">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1368">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1368">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1369">1.2</span><span class="sxs-lookup"><span data-stu-id="57e53-1369">1.2</span></span>|
|[<span data-ttu-id="57e53-1370">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1370">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1371">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1371">ReadItem</span></span>|
|[<span data-ttu-id="57e53-1372">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1372">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1373">作成</span><span class="sxs-lookup"><span data-stu-id="57e53-1373">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="57e53-1374">戻り値:</span><span class="sxs-lookup"><span data-stu-id="57e53-1374">Returns:</span></span>

<span data-ttu-id="57e53-1375">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="57e53-1375">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="57e53-1376">型:String</span><span class="sxs-lookup"><span data-stu-id="57e53-1376">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="57e53-1377">例</span><span class="sxs-lookup"><span data-stu-id="57e53-1377">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-18"></a><span data-ttu-id="57e53-1378">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span><span class="sxs-lookup"><span data-stu-id="57e53-1378">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span></span>

<span data-ttu-id="57e53-1379">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1379">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="57e53-1380">強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1380">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-1381">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="57e53-1381">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-1382">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1382">Requirements</span></span>

|<span data-ttu-id="57e53-1383">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1383">Requirement</span></span>|<span data-ttu-id="57e53-1384">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1384">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1385">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1385">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1386">1.6</span><span class="sxs-lookup"><span data-stu-id="57e53-1386">1.6</span></span>|
|[<span data-ttu-id="57e53-1387">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1387">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1388">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1388">ReadItem</span></span>|
|[<span data-ttu-id="57e53-1389">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1389">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1390">読み取り</span><span class="sxs-lookup"><span data-stu-id="57e53-1390">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="57e53-1391">戻り値:</span><span class="sxs-lookup"><span data-stu-id="57e53-1391">Returns:</span></span>

<span data-ttu-id="57e53-1392">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="57e53-1392">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span></span>

##### <a name="example"></a><span data-ttu-id="57e53-1393">例</span><span class="sxs-lookup"><span data-stu-id="57e53-1393">Example</span></span>

<span data-ttu-id="57e53-1394">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="57e53-1394">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="57e53-1395">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="57e53-1395">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="57e53-p187">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-p187">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-1398">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="57e53-1398">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="57e53-p188">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="57e53-p188">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="57e53-1402">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="57e53-1402">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="57e53-1403">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="57e53-1403">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="57e53-p189">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-p189">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="57e53-1407">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-1407">Requirements</span></span>

|<span data-ttu-id="57e53-1408">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1408">Requirement</span></span>|<span data-ttu-id="57e53-1409">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1409">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1410">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1411">1.6</span><span class="sxs-lookup"><span data-stu-id="57e53-1411">1.6</span></span>|
|[<span data-ttu-id="57e53-1412">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1412">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1413">ReadItem</span></span>|
|[<span data-ttu-id="57e53-1414">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1414">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1415">読み取り</span><span class="sxs-lookup"><span data-stu-id="57e53-1415">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="57e53-1416">戻り値:</span><span class="sxs-lookup"><span data-stu-id="57e53-1416">Returns:</span></span>

<span data-ttu-id="57e53-p190">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="57e53-p190">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="57e53-1419">例</span><span class="sxs-lookup"><span data-stu-id="57e53-1419">Example</span></span>

<span data-ttu-id="57e53-1420">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="57e53-1420">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="57e53-1421">getSharedPropertiesAsync ([options], callback)</span><span class="sxs-lookup"><span data-stu-id="57e53-1421">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="57e53-1422">共有フォルダー、予定表、またはメールボックス内の選択した予定またはメッセージのプロパティを取得します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1422">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-1423">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-1423">Parameters</span></span>

|<span data-ttu-id="57e53-1424">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-1424">Name</span></span>|<span data-ttu-id="57e53-1425">型</span><span class="sxs-lookup"><span data-stu-id="57e53-1425">Type</span></span>|<span data-ttu-id="57e53-1426">属性</span><span class="sxs-lookup"><span data-stu-id="57e53-1426">Attributes</span></span>|<span data-ttu-id="57e53-1427">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-1427">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="57e53-1428">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1428">Object</span></span>|<span data-ttu-id="57e53-1429">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1429">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1430">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="57e53-1430">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57e53-1431">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1431">Object</span></span>|<span data-ttu-id="57e53-1432">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1432">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1433">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1433">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="57e53-1434">function</span><span class="sxs-lookup"><span data-stu-id="57e53-1434">function</span></span>||<span data-ttu-id="57e53-1435">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1435">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="57e53-1436">共有プロパティは、 [`SharedProperties`](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8) `asyncResult.value`プロパティのオブジェクトとして提供されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1436">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="57e53-1437">このオブジェクトは、アイテムの共有プロパティを取得するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1437">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-1438">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-1438">Requirements</span></span>

|<span data-ttu-id="57e53-1439">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1439">Requirement</span></span>|<span data-ttu-id="57e53-1440">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1440">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1441">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1441">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1442">1.8</span><span class="sxs-lookup"><span data-stu-id="57e53-1442">1.8</span></span>|
|[<span data-ttu-id="57e53-1443">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1443">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1444">ReadItem</span></span>|
|[<span data-ttu-id="57e53-1445">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1445">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1446">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-1446">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-1447">例</span><span class="sxs-lookup"><span data-stu-id="57e53-1447">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="57e53-1448">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="57e53-1448">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="57e53-1449">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1449">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="57e53-p192">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="57e53-p192">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-1453">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-1453">Parameters</span></span>

|<span data-ttu-id="57e53-1454">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-1454">Name</span></span>|<span data-ttu-id="57e53-1455">型</span><span class="sxs-lookup"><span data-stu-id="57e53-1455">Type</span></span>|<span data-ttu-id="57e53-1456">属性</span><span class="sxs-lookup"><span data-stu-id="57e53-1456">Attributes</span></span>|<span data-ttu-id="57e53-1457">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-1457">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="57e53-1458">function</span><span class="sxs-lookup"><span data-stu-id="57e53-1458">function</span></span>||<span data-ttu-id="57e53-1459">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1459">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="57e53-1460">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.8) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1460">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.8) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="57e53-1461">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1461">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="57e53-1462">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1462">Object</span></span>|<span data-ttu-id="57e53-1463">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1463">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1464">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1464">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="57e53-1465">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1465">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-1466">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-1466">Requirements</span></span>

|<span data-ttu-id="57e53-1467">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1467">Requirement</span></span>|<span data-ttu-id="57e53-1468">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1468">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1469">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1470">1.0</span><span class="sxs-lookup"><span data-stu-id="57e53-1470">1.0</span></span>|
|[<span data-ttu-id="57e53-1471">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1472">ReadItem</span></span>|
|[<span data-ttu-id="57e53-1473">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1474">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-1474">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-1475">例</span><span class="sxs-lookup"><span data-stu-id="57e53-1475">Example</span></span>

<span data-ttu-id="57e53-p195">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="57e53-p195">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
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

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="57e53-1479">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="57e53-1479">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="57e53-1480">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1480">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="57e53-1481">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1481">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="57e53-1482">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="57e53-1482">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="57e53-1483">Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="57e53-1483">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="57e53-1484">ユーザーがアプリを閉じたとき、またはインラインフォームの作成が開始されたときに、別のウィンドウで続行するためにフォームをポップアウトした後、セッションが終了します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1484">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-1485">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-1485">Parameters</span></span>

|<span data-ttu-id="57e53-1486">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-1486">Name</span></span>|<span data-ttu-id="57e53-1487">型</span><span class="sxs-lookup"><span data-stu-id="57e53-1487">Type</span></span>|<span data-ttu-id="57e53-1488">属性</span><span class="sxs-lookup"><span data-stu-id="57e53-1488">Attributes</span></span>|<span data-ttu-id="57e53-1489">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-1489">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="57e53-1490">String</span><span class="sxs-lookup"><span data-stu-id="57e53-1490">String</span></span>||<span data-ttu-id="57e53-1491">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="57e53-1491">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="57e53-1492">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1492">Object</span></span>|<span data-ttu-id="57e53-1493">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1493">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1494">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="57e53-1494">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57e53-1495">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1495">Object</span></span>|<span data-ttu-id="57e53-1496">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1496">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1497">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1497">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="57e53-1498">function</span><span class="sxs-lookup"><span data-stu-id="57e53-1498">function</span></span>|<span data-ttu-id="57e53-1499">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1499">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1500">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1500">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="57e53-1501">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1501">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="57e53-1502">エラー</span><span class="sxs-lookup"><span data-stu-id="57e53-1502">Errors</span></span>

|<span data-ttu-id="57e53-1503">エラー コード</span><span class="sxs-lookup"><span data-stu-id="57e53-1503">Error code</span></span>|<span data-ttu-id="57e53-1504">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-1504">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="57e53-1505">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="57e53-1505">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-1506">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-1506">Requirements</span></span>

|<span data-ttu-id="57e53-1507">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1507">Requirement</span></span>|<span data-ttu-id="57e53-1508">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1508">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1509">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1510">1.1</span><span class="sxs-lookup"><span data-stu-id="57e53-1510">1.1</span></span>|
|[<span data-ttu-id="57e53-1511">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1512">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1512">ReadWriteItem</span></span>|
|[<span data-ttu-id="57e53-1513">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1514">作成</span><span class="sxs-lookup"><span data-stu-id="57e53-1514">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-1515">例</span><span class="sxs-lookup"><span data-stu-id="57e53-1515">Example</span></span>

<span data-ttu-id="57e53-1516">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1516">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="57e53-1517">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="57e53-1517">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="57e53-1518">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1518">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="57e53-1519">現在、サポートされて`Office.EventType.AttachmentsChanged`いる`Office.EventType.AppointmentTimeChanged`イベント`Office.EventType.EnhancedLocationsChanged`の`Office.EventType.RecipientsChanged`種類は`Office.EventType.RecurrenceChanged`、、、、、です。</span><span class="sxs-lookup"><span data-stu-id="57e53-1519">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-1520">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-1520">Parameters</span></span>

| <span data-ttu-id="57e53-1521">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-1521">Name</span></span> | <span data-ttu-id="57e53-1522">型</span><span class="sxs-lookup"><span data-stu-id="57e53-1522">Type</span></span> | <span data-ttu-id="57e53-1523">属性</span><span class="sxs-lookup"><span data-stu-id="57e53-1523">Attributes</span></span> | <span data-ttu-id="57e53-1524">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-1524">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="57e53-1525">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="57e53-1525">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="57e53-1526">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="57e53-1526">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="57e53-1527">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1527">Object</span></span> | <span data-ttu-id="57e53-1528">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1528">&lt;optional&gt;</span></span> | <span data-ttu-id="57e53-1529">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="57e53-1529">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="57e53-1530">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1530">Object</span></span> | <span data-ttu-id="57e53-1531">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1531">&lt;optional&gt;</span></span> | <span data-ttu-id="57e53-1532">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1532">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="57e53-1533">function</span><span class="sxs-lookup"><span data-stu-id="57e53-1533">function</span></span>| <span data-ttu-id="57e53-1534">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1534">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1535">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1535">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-1536">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1536">Requirements</span></span>

|<span data-ttu-id="57e53-1537">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1537">Requirement</span></span>| <span data-ttu-id="57e53-1538">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1538">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1539">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="57e53-1540">1.7</span><span class="sxs-lookup"><span data-stu-id="57e53-1540">1.7</span></span> |
|[<span data-ttu-id="57e53-1541">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="57e53-1542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1542">ReadItem</span></span> |
|[<span data-ttu-id="57e53-1543">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57e53-1544">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57e53-1544">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="57e53-1545">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="57e53-1545">saveAsync([options], callback)</span></span>

<span data-ttu-id="57e53-1546">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1546">Asynchronously saves an item.</span></span>

<span data-ttu-id="57e53-1547">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1547">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="57e53-1548">Outlook on the web またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1548">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="57e53-1549">キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1549">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-1550">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="57e53-1550">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="57e53-1551">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1551">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="57e53-p199">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-p199">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="57e53-1555">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="57e53-1555">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="57e53-1556">Outlook on Mac では、会議の保存はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="57e53-1556">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="57e53-1557">`saveAsync` メソッドは、作成モードの会議から呼び出されると失敗します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1557">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="57e53-1558">回避策については、「[Office JS API を使用して Outlook for Mac で会議を下書きとして保存できない](https://support.microsoft.com/help/4505745)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57e53-1558">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="57e53-1559">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1559">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-1560">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-1560">Parameters</span></span>

|<span data-ttu-id="57e53-1561">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-1561">Name</span></span>|<span data-ttu-id="57e53-1562">型</span><span class="sxs-lookup"><span data-stu-id="57e53-1562">Type</span></span>|<span data-ttu-id="57e53-1563">属性</span><span class="sxs-lookup"><span data-stu-id="57e53-1563">Attributes</span></span>|<span data-ttu-id="57e53-1564">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-1564">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="57e53-1565">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1565">Object</span></span>|<span data-ttu-id="57e53-1566">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1566">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1567">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="57e53-1567">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57e53-1568">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1568">Object</span></span>|<span data-ttu-id="57e53-1569">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1569">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1570">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1570">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="57e53-1571">function</span><span class="sxs-lookup"><span data-stu-id="57e53-1571">function</span></span>||<span data-ttu-id="57e53-1572">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1572">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="57e53-1573">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1573">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-1574">Requirements</span><span class="sxs-lookup"><span data-stu-id="57e53-1574">Requirements</span></span>

|<span data-ttu-id="57e53-1575">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1575">Requirement</span></span>|<span data-ttu-id="57e53-1576">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1576">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1577">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1577">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1578">1.3</span><span class="sxs-lookup"><span data-stu-id="57e53-1578">1.3</span></span>|
|[<span data-ttu-id="57e53-1579">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1579">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1580">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1580">ReadWriteItem</span></span>|
|[<span data-ttu-id="57e53-1581">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1581">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1582">作成</span><span class="sxs-lookup"><span data-stu-id="57e53-1582">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="57e53-1583">例</span><span class="sxs-lookup"><span data-stu-id="57e53-1583">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="57e53-p201">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="57e53-p201">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="57e53-1586">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="57e53-1586">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="57e53-1587">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="57e53-1587">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="57e53-p202">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="57e53-p202">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57e53-1591">パラメーター</span><span class="sxs-lookup"><span data-stu-id="57e53-1591">Parameters</span></span>

|<span data-ttu-id="57e53-1592">名前</span><span class="sxs-lookup"><span data-stu-id="57e53-1592">Name</span></span>|<span data-ttu-id="57e53-1593">型</span><span class="sxs-lookup"><span data-stu-id="57e53-1593">Type</span></span>|<span data-ttu-id="57e53-1594">属性</span><span class="sxs-lookup"><span data-stu-id="57e53-1594">Attributes</span></span>|<span data-ttu-id="57e53-1595">説明</span><span class="sxs-lookup"><span data-stu-id="57e53-1595">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="57e53-1596">String</span><span class="sxs-lookup"><span data-stu-id="57e53-1596">String</span></span>||<span data-ttu-id="57e53-p203">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="57e53-p203">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="57e53-1600">Object</span><span class="sxs-lookup"><span data-stu-id="57e53-1600">Object</span></span>|<span data-ttu-id="57e53-1601">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1601">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1602">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="57e53-1602">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57e53-1603">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="57e53-1603">Object</span></span>|<span data-ttu-id="57e53-1604">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1604">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1605">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1605">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="57e53-1606">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="57e53-1606">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="57e53-1607">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="57e53-1607">&lt;optional&gt;</span></span>|<span data-ttu-id="57e53-1608">`text` の場合、Outlook on the web とデスクトップ クライアントでは現在のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1608">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="57e53-1609">フィールドが HTML エディターの場合、データが HTML の場合でもテキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1609">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="57e53-1610">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Outlook on the web では現在のスタイルが適用され、Outlook デスクトップ クライアントでは既定のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1610">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="57e53-1611">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1611">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="57e53-1612">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1612">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="57e53-1613">function</span><span class="sxs-lookup"><span data-stu-id="57e53-1613">function</span></span>||<span data-ttu-id="57e53-1614">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="57e53-1614">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57e53-1615">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1615">Requirements</span></span>

|<span data-ttu-id="57e53-1616">要件</span><span class="sxs-lookup"><span data-stu-id="57e53-1616">Requirement</span></span>|<span data-ttu-id="57e53-1617">値</span><span class="sxs-lookup"><span data-stu-id="57e53-1617">Value</span></span>|
|---|---|
|[<span data-ttu-id="57e53-1618">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57e53-1618">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57e53-1619">1.2</span><span class="sxs-lookup"><span data-stu-id="57e53-1619">1.2</span></span>|
|[<span data-ttu-id="57e53-1620">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57e53-1620">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57e53-1621">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="57e53-1621">ReadWriteItem</span></span>|
|[<span data-ttu-id="57e53-1622">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57e53-1622">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57e53-1623">作成</span><span class="sxs-lookup"><span data-stu-id="57e53-1623">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="57e53-1624">例</span><span class="sxs-lookup"><span data-stu-id="57e53-1624">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
