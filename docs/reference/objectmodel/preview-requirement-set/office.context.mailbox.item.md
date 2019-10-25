---
title: Office. アイテム-プレビュー要件セット
description: ''
ms.date: 10/23/2019
localization_priority: Normal
ms.openlocfilehash: 7a72e6fbbec6dbf9cee07d85237baef93ca7ecd8
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/24/2019
ms.locfileid: "37682663"
---
# <a name="item"></a><span data-ttu-id="ccd50-102">item</span><span class="sxs-lookup"><span data-stu-id="ccd50-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="ccd50-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="ccd50-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="ccd50-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-106">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-106">Requirements</span></span>

|<span data-ttu-id="ccd50-107">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-107">Requirement</span></span>|<span data-ttu-id="ccd50-108">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-110">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-110">1.0</span></span>|
|[<span data-ttu-id="ccd50-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="ccd50-112">Restricted</span></span>|
|[<span data-ttu-id="ccd50-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ccd50-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="ccd50-115">Members and methods</span></span>

| <span data-ttu-id="ccd50-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-116">Member</span></span> | <span data-ttu-id="ccd50-117">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ccd50-118">attachments</span><span class="sxs-lookup"><span data-stu-id="ccd50-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="ccd50-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-119">Member</span></span> |
| [<span data-ttu-id="ccd50-120">bcc</span><span class="sxs-lookup"><span data-stu-id="ccd50-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="ccd50-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-121">Member</span></span> |
| [<span data-ttu-id="ccd50-122">body</span><span class="sxs-lookup"><span data-stu-id="ccd50-122">body</span></span>](#body-body) | <span data-ttu-id="ccd50-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-123">Member</span></span> |
| [<span data-ttu-id="ccd50-124">categories</span><span class="sxs-lookup"><span data-stu-id="ccd50-124">categories</span></span>](#categories-categories) | <span data-ttu-id="ccd50-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-125">Member</span></span> |
| [<span data-ttu-id="ccd50-126">cc</span><span class="sxs-lookup"><span data-stu-id="ccd50-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ccd50-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-127">Member</span></span> |
| [<span data-ttu-id="ccd50-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="ccd50-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="ccd50-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-129">Member</span></span> |
| [<span data-ttu-id="ccd50-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="ccd50-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="ccd50-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-131">Member</span></span> |
| [<span data-ttu-id="ccd50-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="ccd50-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="ccd50-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-133">Member</span></span> |
| [<span data-ttu-id="ccd50-134">end</span><span class="sxs-lookup"><span data-stu-id="ccd50-134">end</span></span>](#end-datetime) | <span data-ttu-id="ccd50-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-135">Member</span></span> |
| [<span data-ttu-id="ccd50-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="ccd50-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="ccd50-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-137">Member</span></span> |
| [<span data-ttu-id="ccd50-138">from</span><span class="sxs-lookup"><span data-stu-id="ccd50-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="ccd50-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-139">Member</span></span> |
| [<span data-ttu-id="ccd50-140">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="ccd50-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="ccd50-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-141">Member</span></span> |
| [<span data-ttu-id="ccd50-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="ccd50-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="ccd50-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-143">Member</span></span> |
| [<span data-ttu-id="ccd50-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="ccd50-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="ccd50-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-145">Member</span></span> |
| [<span data-ttu-id="ccd50-146">itemId</span><span class="sxs-lookup"><span data-stu-id="ccd50-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="ccd50-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-147">Member</span></span> |
| [<span data-ttu-id="ccd50-148">itemType</span><span class="sxs-lookup"><span data-stu-id="ccd50-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="ccd50-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-149">Member</span></span> |
| [<span data-ttu-id="ccd50-150">location</span><span class="sxs-lookup"><span data-stu-id="ccd50-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="ccd50-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-151">Member</span></span> |
| [<span data-ttu-id="ccd50-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="ccd50-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="ccd50-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-153">Member</span></span> |
| [<span data-ttu-id="ccd50-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="ccd50-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="ccd50-155">Member</span><span class="sxs-lookup"><span data-stu-id="ccd50-155">Member</span></span> |
| [<span data-ttu-id="ccd50-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="ccd50-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ccd50-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-157">Member</span></span> |
| [<span data-ttu-id="ccd50-158">organizer</span><span class="sxs-lookup"><span data-stu-id="ccd50-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="ccd50-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-159">Member</span></span> |
| [<span data-ttu-id="ccd50-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="ccd50-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="ccd50-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-161">Member</span></span> |
| [<span data-ttu-id="ccd50-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="ccd50-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ccd50-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-163">Member</span></span> |
| [<span data-ttu-id="ccd50-164">sender</span><span class="sxs-lookup"><span data-stu-id="ccd50-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="ccd50-165">Member</span><span class="sxs-lookup"><span data-stu-id="ccd50-165">Member</span></span> |
| [<span data-ttu-id="ccd50-166">系列 Id</span><span class="sxs-lookup"><span data-stu-id="ccd50-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="ccd50-167">Member</span><span class="sxs-lookup"><span data-stu-id="ccd50-167">Member</span></span> |
| [<span data-ttu-id="ccd50-168">start</span><span class="sxs-lookup"><span data-stu-id="ccd50-168">start</span></span>](#start-datetime) | <span data-ttu-id="ccd50-169">Member</span><span class="sxs-lookup"><span data-stu-id="ccd50-169">Member</span></span> |
| [<span data-ttu-id="ccd50-170">subject</span><span class="sxs-lookup"><span data-stu-id="ccd50-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="ccd50-171">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-171">Member</span></span> |
| [<span data-ttu-id="ccd50-172">to</span><span class="sxs-lookup"><span data-stu-id="ccd50-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ccd50-173">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccd50-173">Member</span></span> |
| [<span data-ttu-id="ccd50-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ccd50-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="ccd50-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccd50-175">Method</span></span> |
| [<span data-ttu-id="ccd50-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="ccd50-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="ccd50-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccd50-177">Method</span></span> |
| [<span data-ttu-id="ccd50-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="ccd50-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="ccd50-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccd50-179">Method</span></span> |
| [<span data-ttu-id="ccd50-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ccd50-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="ccd50-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccd50-181">Method</span></span> |
| [<span data-ttu-id="ccd50-182">close</span><span class="sxs-lookup"><span data-stu-id="ccd50-182">close</span></span>](#close) | <span data-ttu-id="ccd50-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccd50-183">Method</span></span> |
| [<span data-ttu-id="ccd50-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="ccd50-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="ccd50-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccd50-185">Method</span></span> |
| [<span data-ttu-id="ccd50-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="ccd50-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="ccd50-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccd50-187">Method</span></span> |
| [<span data-ttu-id="ccd50-188">getAllInternetHeadersAsync</span><span class="sxs-lookup"><span data-stu-id="ccd50-188">getAllInternetHeadersAsync</span></span>](#getallinternetheadersasyncoptions-callback) | <span data-ttu-id="ccd50-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccd50-189">Method</span></span> |
| [<span data-ttu-id="ccd50-190">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="ccd50-190">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="ccd50-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccd50-191">Method</span></span> |
| [<span data-ttu-id="ccd50-192">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="ccd50-192">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="ccd50-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccd50-193">Method</span></span> |
| [<span data-ttu-id="ccd50-194">getEntities</span><span class="sxs-lookup"><span data-stu-id="ccd50-194">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="ccd50-195">Method</span><span class="sxs-lookup"><span data-stu-id="ccd50-195">Method</span></span> |
| [<span data-ttu-id="ccd50-196">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="ccd50-196">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="ccd50-197">Method</span><span class="sxs-lookup"><span data-stu-id="ccd50-197">Method</span></span> |
| [<span data-ttu-id="ccd50-198">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="ccd50-198">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="ccd50-199">Method</span><span class="sxs-lookup"><span data-stu-id="ccd50-199">Method</span></span> |
| [<span data-ttu-id="ccd50-200">、Office.context.mailbox.item.getinitializationcontextasync</span><span class="sxs-lookup"><span data-stu-id="ccd50-200">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="ccd50-201">Method</span><span class="sxs-lookup"><span data-stu-id="ccd50-201">Method</span></span> |
| [<span data-ttu-id="ccd50-202">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="ccd50-202">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="ccd50-203">Method</span><span class="sxs-lookup"><span data-stu-id="ccd50-203">Method</span></span> |
| [<span data-ttu-id="ccd50-204">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="ccd50-204">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="ccd50-205">Method</span><span class="sxs-lookup"><span data-stu-id="ccd50-205">Method</span></span> |
| [<span data-ttu-id="ccd50-206">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="ccd50-206">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="ccd50-207">Method</span><span class="sxs-lookup"><span data-stu-id="ccd50-207">Method</span></span> |
| [<span data-ttu-id="ccd50-208">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="ccd50-208">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="ccd50-209">Method</span><span class="sxs-lookup"><span data-stu-id="ccd50-209">Method</span></span> |
| [<span data-ttu-id="ccd50-210">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="ccd50-210">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="ccd50-211">Method</span><span class="sxs-lookup"><span data-stu-id="ccd50-211">Method</span></span> |
| [<span data-ttu-id="ccd50-212">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="ccd50-212">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="ccd50-213">Method</span><span class="sxs-lookup"><span data-stu-id="ccd50-213">Method</span></span> |
| [<span data-ttu-id="ccd50-214">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="ccd50-214">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="ccd50-215">Method</span><span class="sxs-lookup"><span data-stu-id="ccd50-215">Method</span></span> |
| [<span data-ttu-id="ccd50-216">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="ccd50-216">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="ccd50-217">Method</span><span class="sxs-lookup"><span data-stu-id="ccd50-217">Method</span></span> |
| [<span data-ttu-id="ccd50-218">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ccd50-218">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="ccd50-219">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccd50-219">Method</span></span> |
| [<span data-ttu-id="ccd50-220">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="ccd50-220">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="ccd50-221">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccd50-221">Method</span></span> |
| [<span data-ttu-id="ccd50-222">saveAsync</span><span class="sxs-lookup"><span data-stu-id="ccd50-222">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="ccd50-223">Method</span><span class="sxs-lookup"><span data-stu-id="ccd50-223">Method</span></span> |
| [<span data-ttu-id="ccd50-224">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="ccd50-224">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="ccd50-225">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccd50-225">Method</span></span> |

### <a name="example"></a><span data-ttu-id="ccd50-226">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-226">Example</span></span>

<span data-ttu-id="ccd50-227">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-227">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="ccd50-228">Members</span><span class="sxs-lookup"><span data-stu-id="ccd50-228">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="ccd50-229">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="ccd50-229">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="ccd50-230">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-230">Gets the item's attachments as an array.</span></span> <span data-ttu-id="ccd50-231">Read mode only.</span><span class="sxs-lookup"><span data-stu-id="ccd50-231">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-232">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-232">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="ccd50-233">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ccd50-233">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="ccd50-234">型</span><span class="sxs-lookup"><span data-stu-id="ccd50-234">Type</span></span>

*   <span data-ttu-id="ccd50-235">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="ccd50-235">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-236">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-236">Requirements</span></span>

|<span data-ttu-id="ccd50-237">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-237">Requirement</span></span>|<span data-ttu-id="ccd50-238">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-239">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-240">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-240">1.0</span></span>|
|[<span data-ttu-id="ccd50-241">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-242">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-243">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-244">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-244">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-245">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-245">Example</span></span>

<span data-ttu-id="ccd50-246">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-246">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="ccd50-247">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ccd50-247">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="ccd50-248">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-248">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="ccd50-249">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ccd50-249">Compose mode only.</span></span>

<span data-ttu-id="ccd50-250">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ccd50-251">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-251">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ccd50-252">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-252">Get 500 members maximum.</span></span>
- <span data-ttu-id="ccd50-253">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="ccd50-253">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="ccd50-254">Type</span><span class="sxs-lookup"><span data-stu-id="ccd50-254">Type</span></span>

*   [<span data-ttu-id="ccd50-255">受信者</span><span class="sxs-lookup"><span data-stu-id="ccd50-255">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="ccd50-256">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-256">Requirements</span></span>

|<span data-ttu-id="ccd50-257">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-257">Requirement</span></span>|<span data-ttu-id="ccd50-258">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-259">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-260">1.1</span><span class="sxs-lookup"><span data-stu-id="ccd50-260">1.1</span></span>|
|[<span data-ttu-id="ccd50-261">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-262">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-264">作成</span><span class="sxs-lookup"><span data-stu-id="ccd50-264">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-265">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-265">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="ccd50-266">body: [Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="ccd50-266">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="ccd50-267">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-267">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="ccd50-268">Type</span><span class="sxs-lookup"><span data-stu-id="ccd50-268">Type</span></span>

*   [<span data-ttu-id="ccd50-269">Body</span><span class="sxs-lookup"><span data-stu-id="ccd50-269">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="ccd50-270">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-270">Requirements</span></span>

|<span data-ttu-id="ccd50-271">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-271">Requirement</span></span>|<span data-ttu-id="ccd50-272">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-273">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-274">1.1</span><span class="sxs-lookup"><span data-stu-id="ccd50-274">1.1</span></span>|
|[<span data-ttu-id="ccd50-275">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-275">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-276">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-277">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-277">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-278">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-278">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-279">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-279">Example</span></span>

<span data-ttu-id="ccd50-280">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-280">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="ccd50-281">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-281">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="ccd50-282">カテゴリ:[カテゴリ](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="ccd50-282">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="ccd50-283">アイテムのカテゴリを管理するためのメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-283">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-284">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-284">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="ccd50-285">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-285">Type</span></span>

*   [<span data-ttu-id="ccd50-286">Categories</span><span class="sxs-lookup"><span data-stu-id="ccd50-286">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="ccd50-287">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-287">Requirements</span></span>

|<span data-ttu-id="ccd50-288">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-288">Requirement</span></span>|<span data-ttu-id="ccd50-289">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-290">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-291">プレビュー</span><span class="sxs-lookup"><span data-stu-id="ccd50-291">Preview</span></span>|
|[<span data-ttu-id="ccd50-292">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-293">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-293">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-294">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-295">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-295">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-296">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-296">Example</span></span>

<span data-ttu-id="ccd50-297">この例では、アイテムのカテゴリを取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-297">This example gets the item's categories.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="ccd50-298">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ccd50-298">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="ccd50-299">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-299">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="ccd50-300">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-300">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ccd50-301">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-301">Read mode</span></span>

<span data-ttu-id="ccd50-302">`cc` プロパティは、メッセージの `EmailAddressDetails` 行にある各受信者について、\*\*\*\* オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-302">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="ccd50-303">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-303">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ccd50-304">ただし、Windows および Mac では、最大500メンバーを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-304">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="ccd50-305">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-305">Compose mode</span></span>

<span data-ttu-id="ccd50-306">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-306">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="ccd50-307">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-307">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ccd50-308">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-308">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ccd50-309">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-309">Get 500 members maximum.</span></span>
- <span data-ttu-id="ccd50-310">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="ccd50-310">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ccd50-311">Type</span><span class="sxs-lookup"><span data-stu-id="ccd50-311">Type</span></span>

*   <span data-ttu-id="ccd50-312">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ccd50-312">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-313">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-313">Requirements</span></span>

|<span data-ttu-id="ccd50-314">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-314">Requirement</span></span>|<span data-ttu-id="ccd50-315">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-316">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-317">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-317">1.0</span></span>|
|[<span data-ttu-id="ccd50-318">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-319">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-320">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-321">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-321">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="ccd50-322">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="ccd50-322">(nullable) conversationId: String</span></span>

<span data-ttu-id="ccd50-323">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-323">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="ccd50-p109">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="ccd50-p110">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="ccd50-328">Type</span><span class="sxs-lookup"><span data-stu-id="ccd50-328">Type</span></span>

*   <span data-ttu-id="ccd50-329">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-329">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-330">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-330">Requirements</span></span>

|<span data-ttu-id="ccd50-331">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-331">Requirement</span></span>|<span data-ttu-id="ccd50-332">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-333">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-334">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-334">1.0</span></span>|
|[<span data-ttu-id="ccd50-335">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-336">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-337">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-338">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-338">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-339">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-339">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="ccd50-340">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="ccd50-340">dateTimeCreated: Date</span></span>

<span data-ttu-id="ccd50-p111">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ccd50-343">型</span><span class="sxs-lookup"><span data-stu-id="ccd50-343">Type</span></span>

*   <span data-ttu-id="ccd50-344">日付</span><span class="sxs-lookup"><span data-stu-id="ccd50-344">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-345">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-345">Requirements</span></span>

|<span data-ttu-id="ccd50-346">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-346">Requirement</span></span>|<span data-ttu-id="ccd50-347">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-347">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-348">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-349">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-349">1.0</span></span>|
|[<span data-ttu-id="ccd50-350">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-350">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-351">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-352">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-352">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-353">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-353">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-354">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-354">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="ccd50-355">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="ccd50-355">dateTimeModified: Date</span></span>

<span data-ttu-id="ccd50-p112">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-358">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-358">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="ccd50-359">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-359">Type</span></span>

*   <span data-ttu-id="ccd50-360">日付</span><span class="sxs-lookup"><span data-stu-id="ccd50-360">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-361">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-361">Requirements</span></span>

|<span data-ttu-id="ccd50-362">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-362">Requirement</span></span>|<span data-ttu-id="ccd50-363">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-363">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-364">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-364">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-365">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-365">1.0</span></span>|
|[<span data-ttu-id="ccd50-366">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-366">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-367">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-367">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-368">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-368">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-369">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-369">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-370">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-370">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="ccd50-371">end: Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="ccd50-371">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="ccd50-372">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-372">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="ccd50-p113">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ccd50-375">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-375">Read mode</span></span>

<span data-ttu-id="ccd50-376">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-376">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="ccd50-377">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-377">Compose mode</span></span>

<span data-ttu-id="ccd50-378">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-378">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="ccd50-379">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-379">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="ccd50-380">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-380">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="ccd50-381">型</span><span class="sxs-lookup"><span data-stu-id="ccd50-381">Type</span></span>

*   <span data-ttu-id="ccd50-382">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="ccd50-382">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-383">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-383">Requirements</span></span>

|<span data-ttu-id="ccd50-384">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-384">Requirement</span></span>|<span data-ttu-id="ccd50-385">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-386">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-387">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-387">1.0</span></span>|
|[<span data-ttu-id="ccd50-388">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-389">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-390">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-391">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-391">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="ccd50-392">enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="ccd50-392">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="ccd50-393">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-393">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ccd50-394">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-394">Read mode</span></span>

<span data-ttu-id="ccd50-395">この`enhancedLocation`プロパティは、予定に関連付けられている場所 ( [locationdetails](/javascript/api/outlook/office.locationdetails)オブジェクトで表される) のセットを取得できる[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-395">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ccd50-396">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-396">Compose mode</span></span>

<span data-ttu-id="ccd50-397">この`enhancedLocation`プロパティは、予定の場所を取得、削除、または追加するためのメソッドを提供する[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-397">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="ccd50-398">Type</span><span class="sxs-lookup"><span data-stu-id="ccd50-398">Type</span></span>

*   [<span data-ttu-id="ccd50-399">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="ccd50-399">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="ccd50-400">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-400">Requirements</span></span>

|<span data-ttu-id="ccd50-401">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-401">Requirement</span></span>|<span data-ttu-id="ccd50-402">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-403">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-403">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-404">プレビュー</span><span class="sxs-lookup"><span data-stu-id="ccd50-404">Preview</span></span>|
|[<span data-ttu-id="ccd50-405">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-405">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-406">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-407">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-407">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-408">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-408">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-409">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-409">Example</span></span>

<span data-ttu-id="ccd50-410">次の例では、予定に関連付けられている現在の場所を取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-410">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="ccd50-411">from: [emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)|[from](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="ccd50-411">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="ccd50-412">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-412">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="ccd50-p114">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-415">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-415">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ccd50-416">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-416">Read mode</span></span>

<span data-ttu-id="ccd50-417">プロパティ`from`は`EmailAddressDetails`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-417">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="ccd50-418">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-418">Compose mode</span></span>

<span data-ttu-id="ccd50-419">プロパティ`from`は、from `From`値を取得するメソッドを提供するオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-419">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ccd50-420">Type</span><span class="sxs-lookup"><span data-stu-id="ccd50-420">Type</span></span>

*   <span data-ttu-id="ccd50-421">[電子メールアドレス](/javascript/api/outlook/office.emailaddressdetails) | [の](/javascript/api/outlook/office.from)詳細</span><span class="sxs-lookup"><span data-stu-id="ccd50-421">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-422">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-422">Requirements</span></span>

|<span data-ttu-id="ccd50-423">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-423">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="ccd50-424">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-425">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-425">1.0</span></span>|<span data-ttu-id="ccd50-426">1.7</span><span class="sxs-lookup"><span data-stu-id="ccd50-426">1.7</span></span>|
|[<span data-ttu-id="ccd50-427">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-427">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-428">ReadItem</span></span>|<span data-ttu-id="ccd50-429">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-429">ReadWriteItem</span></span>|
|[<span data-ttu-id="ccd50-430">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-431">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-431">Read</span></span>|<span data-ttu-id="ccd50-432">作成</span><span class="sxs-lookup"><span data-stu-id="ccd50-432">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="ccd50-433">internetHeaders: [internetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="ccd50-433">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="ccd50-434">メッセージのカスタムインターネットヘッダーを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-434">Gets or sets custom internet headers on a message.</span></span> <span data-ttu-id="ccd50-435">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ccd50-435">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ccd50-436">Type</span><span class="sxs-lookup"><span data-stu-id="ccd50-436">Type</span></span>

*   [<span data-ttu-id="ccd50-437">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="ccd50-437">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="ccd50-438">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-438">Requirements</span></span>

|<span data-ttu-id="ccd50-439">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-439">Requirement</span></span>|<span data-ttu-id="ccd50-440">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-440">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-441">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-441">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-442">プレビュー</span><span class="sxs-lookup"><span data-stu-id="ccd50-442">Preview</span></span>|
|[<span data-ttu-id="ccd50-443">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-443">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-444">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-445">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-445">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-446">作成</span><span class="sxs-lookup"><span data-stu-id="ccd50-446">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-447">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-447">Example</span></span>

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

#### <a name="internetmessageid-string"></a><span data-ttu-id="ccd50-448">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="ccd50-448">internetMessageId: String</span></span>

<span data-ttu-id="ccd50-p116">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ccd50-451">Type</span><span class="sxs-lookup"><span data-stu-id="ccd50-451">Type</span></span>

*   <span data-ttu-id="ccd50-452">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-452">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-453">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-453">Requirements</span></span>

|<span data-ttu-id="ccd50-454">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-454">Requirement</span></span>|<span data-ttu-id="ccd50-455">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-456">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-457">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-457">1.0</span></span>|
|[<span data-ttu-id="ccd50-458">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-458">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-459">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-460">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-460">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-461">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-461">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-462">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-462">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="ccd50-463">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="ccd50-463">itemClass: String</span></span>

<span data-ttu-id="ccd50-p117">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="ccd50-p118">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="ccd50-468">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-468">Type</span></span>|<span data-ttu-id="ccd50-469">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-469">Description</span></span>|<span data-ttu-id="ccd50-470">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="ccd50-470">item class</span></span>|
|---|---|---|
|<span data-ttu-id="ccd50-471">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="ccd50-471">Appointment items</span></span>|<span data-ttu-id="ccd50-472">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="ccd50-472">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="ccd50-473">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="ccd50-473">Message items</span></span>|<span data-ttu-id="ccd50-474">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-474">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="ccd50-475">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-475">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="ccd50-476">Type</span><span class="sxs-lookup"><span data-stu-id="ccd50-476">Type</span></span>

*   <span data-ttu-id="ccd50-477">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-477">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-478">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-478">Requirements</span></span>

|<span data-ttu-id="ccd50-479">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-479">Requirement</span></span>|<span data-ttu-id="ccd50-480">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-481">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-482">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-482">1.0</span></span>|
|[<span data-ttu-id="ccd50-483">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-484">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-485">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-486">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-486">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-487">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-487">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="ccd50-488">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="ccd50-488">(nullable) itemId: String</span></span>

<span data-ttu-id="ccd50-p119">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p119">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-491">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="ccd50-491">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="ccd50-492">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-492">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="ccd50-493">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-493">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="ccd50-494">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ccd50-494">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="ccd50-p121">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="ccd50-497">Type</span><span class="sxs-lookup"><span data-stu-id="ccd50-497">Type</span></span>

*   <span data-ttu-id="ccd50-498">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-498">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-499">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-499">Requirements</span></span>

|<span data-ttu-id="ccd50-500">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-500">Requirement</span></span>|<span data-ttu-id="ccd50-501">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-502">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-503">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-503">1.0</span></span>|
|[<span data-ttu-id="ccd50-504">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-504">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-505">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-506">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-506">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-507">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-507">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-508">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-508">Example</span></span>

<span data-ttu-id="ccd50-p122">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="ccd50-511">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="ccd50-511">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="ccd50-512">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-512">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="ccd50-513">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="ccd50-513">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="ccd50-514">Type</span><span class="sxs-lookup"><span data-stu-id="ccd50-514">Type</span></span>

*   [<span data-ttu-id="ccd50-515">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="ccd50-515">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="ccd50-516">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-516">Requirements</span></span>

|<span data-ttu-id="ccd50-517">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-517">Requirement</span></span>|<span data-ttu-id="ccd50-518">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-518">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-519">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-519">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-520">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-520">1.0</span></span>|
|[<span data-ttu-id="ccd50-521">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-521">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-522">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-522">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-523">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-523">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-524">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-524">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-525">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-525">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="ccd50-526">location: String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="ccd50-526">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="ccd50-527">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-527">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ccd50-528">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-528">Read mode</span></span>

<span data-ttu-id="ccd50-529">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-529">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="ccd50-530">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-530">Compose mode</span></span>

<span data-ttu-id="ccd50-531">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-531">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ccd50-532">型</span><span class="sxs-lookup"><span data-stu-id="ccd50-532">Type</span></span>

*   <span data-ttu-id="ccd50-533">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="ccd50-533">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-534">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-534">Requirements</span></span>

|<span data-ttu-id="ccd50-535">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-535">Requirement</span></span>|<span data-ttu-id="ccd50-536">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-537">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-537">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-538">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-538">1.0</span></span>|
|[<span data-ttu-id="ccd50-539">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-539">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-540">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-541">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-541">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-542">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-542">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="ccd50-543">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="ccd50-543">normalizedSubject: String</span></span>

<span data-ttu-id="ccd50-p123">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="ccd50-p124">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="ccd50-548">型</span><span class="sxs-lookup"><span data-stu-id="ccd50-548">Type</span></span>

*   <span data-ttu-id="ccd50-549">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-549">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-550">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-550">Requirements</span></span>

|<span data-ttu-id="ccd50-551">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-551">Requirement</span></span>|<span data-ttu-id="ccd50-552">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-552">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-553">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-553">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-554">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-554">1.0</span></span>|
|[<span data-ttu-id="ccd50-555">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-555">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-556">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-556">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-557">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-557">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-558">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-558">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-559">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-559">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="ccd50-560">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="ccd50-560">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="ccd50-561">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-561">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="ccd50-562">型</span><span class="sxs-lookup"><span data-stu-id="ccd50-562">Type</span></span>

*   [<span data-ttu-id="ccd50-563">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="ccd50-563">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="ccd50-564">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-564">Requirements</span></span>

|<span data-ttu-id="ccd50-565">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-565">Requirement</span></span>|<span data-ttu-id="ccd50-566">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-567">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-568">1.3</span><span class="sxs-lookup"><span data-stu-id="ccd50-568">1.3</span></span>|
|[<span data-ttu-id="ccd50-569">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-570">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-571">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-572">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-572">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-573">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-573">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="ccd50-574">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ccd50-574">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="ccd50-575">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-575">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="ccd50-576">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-576">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ccd50-577">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-577">Read mode</span></span>

<span data-ttu-id="ccd50-578">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-578">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="ccd50-579">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-579">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ccd50-580">ただし、Windows および Mac では、最大500メンバーを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-580">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="ccd50-581">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-581">Compose mode</span></span>

<span data-ttu-id="ccd50-582">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-582">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="ccd50-583">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-583">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ccd50-584">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-584">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ccd50-585">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-585">Get 500 members maximum.</span></span>
- <span data-ttu-id="ccd50-586">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="ccd50-586">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ccd50-587">Type</span><span class="sxs-lookup"><span data-stu-id="ccd50-587">Type</span></span>

*   <span data-ttu-id="ccd50-588">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ccd50-588">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-589">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-589">Requirements</span></span>

|<span data-ttu-id="ccd50-590">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-590">Requirement</span></span>|<span data-ttu-id="ccd50-591">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-592">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-593">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-593">1.0</span></span>|
|[<span data-ttu-id="ccd50-594">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-594">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-595">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-595">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-596">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-597">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-597">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="ccd50-598">開催者: [emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)|[開催者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="ccd50-598">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="ccd50-599">指定した会議の開催者の電子メールアドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-599">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ccd50-600">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-600">Read mode</span></span>

<span data-ttu-id="ccd50-601">プロパティ`organizer`は、会議の開催者を表す[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-601">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="ccd50-602">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-602">Compose mode</span></span>

<span data-ttu-id="ccd50-603">プロパティ`organizer`は、開催者の値を取得するためのメソッドを提供する[オーガナイザー](/javascript/api/outlook/office.organizer)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-603">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="ccd50-604">Type</span><span class="sxs-lookup"><span data-stu-id="ccd50-604">Type</span></span>

*   <span data-ttu-id="ccd50-605">[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails) | [開催者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="ccd50-605">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-606">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-606">Requirements</span></span>

|<span data-ttu-id="ccd50-607">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-607">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="ccd50-608">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-609">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-609">1.0</span></span>|<span data-ttu-id="ccd50-610">1.7</span><span class="sxs-lookup"><span data-stu-id="ccd50-610">1.7</span></span>|
|[<span data-ttu-id="ccd50-611">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-611">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-612">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-612">ReadItem</span></span>|<span data-ttu-id="ccd50-613">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-613">ReadWriteItem</span></span>|
|[<span data-ttu-id="ccd50-614">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-614">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-615">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-615">Read</span></span>|<span data-ttu-id="ccd50-616">作成</span><span class="sxs-lookup"><span data-stu-id="ccd50-616">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="ccd50-617">(nullable) 定期的なスケジュール:[定期的](/javascript/api/outlook/office.recurrence)なアイテム</span><span class="sxs-lookup"><span data-stu-id="ccd50-617">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="ccd50-618">予定の定期的なパターンを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-618">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="ccd50-619">会議出席依頼の定期的なパターンを取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-619">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="ccd50-620">予定アイテムの読み取りおよび作成モード。</span><span class="sxs-lookup"><span data-stu-id="ccd50-620">Read and compose modes for appointment items.</span></span> <span data-ttu-id="ccd50-621">会議出席依頼アイテムの閲覧モード。</span><span class="sxs-lookup"><span data-stu-id="ccd50-621">Read mode for meeting request items.</span></span>

<span data-ttu-id="ccd50-622">この`recurrence`プロパティは、アイテムが series または series 内のインスタンスの場合、定期的な予定または会議出席依頼に対して[定期的](/javascript/api/outlook/office.recurrence)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-622">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="ccd50-623">`null`は、単一の予定および1つの予定の会議出席依頼に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-623">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="ccd50-624">`undefined`は、会議出席依頼ではないメッセージに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-624">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="ccd50-625">注: 会議出席依頼に`itemClass`は、IPM という値があります。出席依頼。</span><span class="sxs-lookup"><span data-stu-id="ccd50-625">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="ccd50-626">注: 定期的なオブジェクトが`null`の場合は、そのオブジェクトが単一の予定または1つの予定の会議出席依頼であり、データ系列の一部ではないことを示します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-626">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ccd50-627">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-627">Read mode</span></span>

<span data-ttu-id="ccd50-628">この`recurrence`プロパティは、定期的な予定を表す[定期的](/javascript/api/outlook/office.recurrence)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-628">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="ccd50-629">これは、予定および会議出席依頼に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-629">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="ccd50-630">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-630">Compose mode</span></span>

<span data-ttu-id="ccd50-631">この`recurrence`プロパティは、予定の繰り返しを管理するためのメソッドを提供する[定期的](/javascript/api/outlook/office.recurrence)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-631">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="ccd50-632">これは予定に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-632">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="ccd50-633">Type</span><span class="sxs-lookup"><span data-stu-id="ccd50-633">Type</span></span>

* [<span data-ttu-id="ccd50-634">繰り返さ</span><span class="sxs-lookup"><span data-stu-id="ccd50-634">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="ccd50-635">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-635">Requirement</span></span>|<span data-ttu-id="ccd50-636">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-636">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-637">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-638">1.7</span><span class="sxs-lookup"><span data-stu-id="ccd50-638">1.7</span></span>|
|[<span data-ttu-id="ccd50-639">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-639">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-640">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-640">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-641">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-641">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-642">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-642">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="ccd50-643">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ccd50-643">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="ccd50-644">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-644">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="ccd50-645">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-645">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ccd50-646">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-646">Read mode</span></span>

<span data-ttu-id="ccd50-647">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-647">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="ccd50-648">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-648">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ccd50-649">ただし、Windows および Mac では、最大500メンバーを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-649">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="ccd50-650">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-650">Compose mode</span></span>

<span data-ttu-id="ccd50-651">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-651">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="ccd50-652">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-652">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ccd50-653">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-653">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ccd50-654">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-654">Get 500 members maximum.</span></span>
- <span data-ttu-id="ccd50-655">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="ccd50-655">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="ccd50-656">Type</span><span class="sxs-lookup"><span data-stu-id="ccd50-656">Type</span></span>

*   <span data-ttu-id="ccd50-657">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ccd50-657">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-658">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-658">Requirements</span></span>

|<span data-ttu-id="ccd50-659">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-659">Requirement</span></span>|<span data-ttu-id="ccd50-660">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-661">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-662">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-662">1.0</span></span>|
|[<span data-ttu-id="ccd50-663">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-664">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-664">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-665">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-666">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-666">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="ccd50-667">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="ccd50-667">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="ccd50-p135">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p135">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="ccd50-p136">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsfrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p136">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-672">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-672">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="ccd50-673">型</span><span class="sxs-lookup"><span data-stu-id="ccd50-673">Type</span></span>

*   [<span data-ttu-id="ccd50-674">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ccd50-674">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="ccd50-675">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-675">Requirements</span></span>

|<span data-ttu-id="ccd50-676">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-676">Requirement</span></span>|<span data-ttu-id="ccd50-677">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-677">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-678">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-678">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-679">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-679">1.0</span></span>|
|[<span data-ttu-id="ccd50-680">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-680">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-681">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-681">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-682">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-682">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-683">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-683">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-684">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-684">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="ccd50-685">(nullable) 系列 Id: String</span><span class="sxs-lookup"><span data-stu-id="ccd50-685">(nullable) seriesId: String</span></span>

<span data-ttu-id="ccd50-686">インスタンスが属する系列の id を取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-686">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="ccd50-687">Web 上の Outlook およびデスクトップクライアントでは、 `seriesId`は、このアイテムが属する親 (シリーズ) アイテムの Exchange web サービス (EWS) ID を返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-687">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="ccd50-688">ただし、iOS と Android では、 `seriesId`は親アイテムの REST ID を返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-688">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-689">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="ccd50-689">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="ccd50-690">`seriesId`プロパティが OUTLOOK REST API で使用される outlook id と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-690">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="ccd50-691">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-691">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="ccd50-692">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ccd50-692">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="ccd50-693">この`seriesId`プロパティは`null` 、単一の予定、系列のアイテム、会議出席依頼などの親アイテムを持たないアイテムに`undefined`対して、会議出席依頼以外のアイテムに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-693">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="ccd50-694">Type</span><span class="sxs-lookup"><span data-stu-id="ccd50-694">Type</span></span>

* <span data-ttu-id="ccd50-695">文字列</span><span class="sxs-lookup"><span data-stu-id="ccd50-695">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-696">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-696">Requirements</span></span>

|<span data-ttu-id="ccd50-697">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-697">Requirement</span></span>|<span data-ttu-id="ccd50-698">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-698">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-699">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-699">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-700">1.7</span><span class="sxs-lookup"><span data-stu-id="ccd50-700">1.7</span></span>|
|[<span data-ttu-id="ccd50-701">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-701">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-702">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-702">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-703">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-703">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-704">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-704">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-705">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-705">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="ccd50-706">start: Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="ccd50-706">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="ccd50-707">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-707">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="ccd50-p139">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p139">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ccd50-710">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-710">Read mode</span></span>

<span data-ttu-id="ccd50-711">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-711">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="ccd50-712">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-712">Compose mode</span></span>

<span data-ttu-id="ccd50-713">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-713">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="ccd50-714">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-714">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="ccd50-715">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-715">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="ccd50-716">型</span><span class="sxs-lookup"><span data-stu-id="ccd50-716">Type</span></span>

*   <span data-ttu-id="ccd50-717">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="ccd50-717">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-718">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-718">Requirements</span></span>

|<span data-ttu-id="ccd50-719">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-719">Requirement</span></span>|<span data-ttu-id="ccd50-720">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-720">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-721">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-721">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-722">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-722">1.0</span></span>|
|[<span data-ttu-id="ccd50-723">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-723">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-724">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-724">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-725">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-725">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-726">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-726">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="ccd50-727">subject: String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="ccd50-727">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="ccd50-728">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-728">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="ccd50-729">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-729">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ccd50-730">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-730">Read mode</span></span>

<span data-ttu-id="ccd50-p140">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p140">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="ccd50-733">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-733">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="ccd50-734">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-734">Compose mode</span></span>
<span data-ttu-id="ccd50-735">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-735">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="ccd50-736">型</span><span class="sxs-lookup"><span data-stu-id="ccd50-736">Type</span></span>

*   <span data-ttu-id="ccd50-737">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="ccd50-737">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-738">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-738">Requirements</span></span>

|<span data-ttu-id="ccd50-739">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-739">Requirement</span></span>|<span data-ttu-id="ccd50-740">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-740">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-741">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-741">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-742">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-742">1.0</span></span>|
|[<span data-ttu-id="ccd50-743">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-743">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-744">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-744">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-745">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-745">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-746">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-746">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="ccd50-747">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ccd50-747">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="ccd50-748">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-748">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="ccd50-749">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-749">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ccd50-750">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-750">Read mode</span></span>

<span data-ttu-id="ccd50-751">`to` プロパティは、メッセージの `EmailAddressDetails` 行にある各受信者について、\*\*\*\* オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-751">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="ccd50-752">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-752">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ccd50-753">ただし、Windows および Mac では、最大500メンバーを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-753">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="ccd50-754">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ccd50-754">Compose mode</span></span>

<span data-ttu-id="ccd50-755">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-755">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="ccd50-756">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-756">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ccd50-757">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-757">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ccd50-758">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-758">Get 500 members maximum.</span></span>
- <span data-ttu-id="ccd50-759">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="ccd50-759">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ccd50-760">Type</span><span class="sxs-lookup"><span data-stu-id="ccd50-760">Type</span></span>

*   <span data-ttu-id="ccd50-761">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ccd50-761">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-762">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-762">Requirements</span></span>

|<span data-ttu-id="ccd50-763">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-763">Requirement</span></span>|<span data-ttu-id="ccd50-764">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-764">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-765">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-765">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-766">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-766">1.0</span></span>|
|[<span data-ttu-id="ccd50-767">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-767">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-768">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-768">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-769">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-769">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-770">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-770">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="ccd50-771">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccd50-771">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="ccd50-772">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ccd50-772">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ccd50-773">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-773">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="ccd50-774">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-774">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="ccd50-775">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-775">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-776">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-776">Parameters</span></span>
|<span data-ttu-id="ccd50-777">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-777">Name</span></span>|<span data-ttu-id="ccd50-778">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-778">Type</span></span>|<span data-ttu-id="ccd50-779">属性</span><span class="sxs-lookup"><span data-stu-id="ccd50-779">Attributes</span></span>|<span data-ttu-id="ccd50-780">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-780">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="ccd50-781">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-781">String</span></span>||<span data-ttu-id="ccd50-p144">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p144">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="ccd50-784">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-784">String</span></span>||<span data-ttu-id="ccd50-p145">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p145">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="ccd50-787">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-787">Object</span></span>|<span data-ttu-id="ccd50-788">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-788">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-789">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ccd50-789">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ccd50-790">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-790">Object</span></span>|<span data-ttu-id="ccd50-791">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-791">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-792">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-792">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="ccd50-793">Boolean</span><span class="sxs-lookup"><span data-stu-id="ccd50-793">Boolean</span></span>|<span data-ttu-id="ccd50-794">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-794">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-795">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-795">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="ccd50-796">function</span><span class="sxs-lookup"><span data-stu-id="ccd50-796">function</span></span>|<span data-ttu-id="ccd50-797">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-797">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-798">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-798">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ccd50-799">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-799">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ccd50-800">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-800">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ccd50-801">エラー</span><span class="sxs-lookup"><span data-stu-id="ccd50-801">Errors</span></span>

|<span data-ttu-id="ccd50-802">エラー コード</span><span class="sxs-lookup"><span data-stu-id="ccd50-802">Error code</span></span>|<span data-ttu-id="ccd50-803">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-803">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="ccd50-804">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-804">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="ccd50-805">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="ccd50-805">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="ccd50-806">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-806">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-807">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-807">Requirements</span></span>

|<span data-ttu-id="ccd50-808">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-808">Requirement</span></span>|<span data-ttu-id="ccd50-809">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-809">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-810">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-810">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-811">1.1</span><span class="sxs-lookup"><span data-stu-id="ccd50-811">1.1</span></span>|
|[<span data-ttu-id="ccd50-812">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-812">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-813">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-813">ReadWriteItem</span></span>|
|[<span data-ttu-id="ccd50-814">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-814">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-815">作成</span><span class="sxs-lookup"><span data-stu-id="ccd50-815">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="ccd50-816">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-816">Examples</span></span>

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

<span data-ttu-id="ccd50-817">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-817">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="ccd50-818">addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ccd50-818">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ccd50-819">Base64 エンコードのファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-819">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="ccd50-820">この`addFileAttachmentFromBase64Async`メソッドは、base64 エンコードからファイルをアップロードし、新規作成フォームのアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-820">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="ccd50-821">このメソッドは、AsyncResult オブジェクトの添付ファイル識別子を返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-821">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="ccd50-822">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-822">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-823">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-823">Parameters</span></span>

|<span data-ttu-id="ccd50-824">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-824">Name</span></span>|<span data-ttu-id="ccd50-825">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-825">Type</span></span>|<span data-ttu-id="ccd50-826">属性</span><span class="sxs-lookup"><span data-stu-id="ccd50-826">Attributes</span></span>|<span data-ttu-id="ccd50-827">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-827">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="ccd50-828">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-828">String</span></span>||<span data-ttu-id="ccd50-829">電子メールまたはイベントに追加する画像またはファイルの、base64 でエンコードされたコンテンツ。</span><span class="sxs-lookup"><span data-stu-id="ccd50-829">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="ccd50-830">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-830">String</span></span>||<span data-ttu-id="ccd50-p147">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p147">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="ccd50-833">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-833">Object</span></span>|<span data-ttu-id="ccd50-834">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-834">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-835">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ccd50-835">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ccd50-836">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccd50-836">Object</span></span>|<span data-ttu-id="ccd50-837">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-837">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-838">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-838">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="ccd50-839">Boolean</span><span class="sxs-lookup"><span data-stu-id="ccd50-839">Boolean</span></span>|<span data-ttu-id="ccd50-840">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-840">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-841">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-841">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="ccd50-842">function</span><span class="sxs-lookup"><span data-stu-id="ccd50-842">function</span></span>|<span data-ttu-id="ccd50-843">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-843">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-844">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-844">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ccd50-845">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-845">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ccd50-846">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-846">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ccd50-847">エラー</span><span class="sxs-lookup"><span data-stu-id="ccd50-847">Errors</span></span>

|<span data-ttu-id="ccd50-848">エラー コード</span><span class="sxs-lookup"><span data-stu-id="ccd50-848">Error code</span></span>|<span data-ttu-id="ccd50-849">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-849">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="ccd50-850">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-850">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="ccd50-851">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="ccd50-851">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="ccd50-852">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-852">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-853">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-853">Requirements</span></span>

|<span data-ttu-id="ccd50-854">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-854">Requirement</span></span>|<span data-ttu-id="ccd50-855">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-856">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-856">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-857">プレビュー</span><span class="sxs-lookup"><span data-stu-id="ccd50-857">Preview</span></span>|
|[<span data-ttu-id="ccd50-858">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-858">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-859">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-859">ReadWriteItem</span></span>|
|[<span data-ttu-id="ccd50-860">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-860">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-861">作成</span><span class="sxs-lookup"><span data-stu-id="ccd50-861">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="ccd50-862">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-862">Examples</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="ccd50-863">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ccd50-863">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="ccd50-864">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-864">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="ccd50-865">現在、サポートされて`Office.EventType.AttachmentsChanged`いる`Office.EventType.AppointmentTimeChanged`イベント`Office.EventType.EnhancedLocationsChanged`の`Office.EventType.RecipientsChanged`種類は`Office.EventType.RecurrenceChanged`、、、、、です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-865">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-866">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-866">Parameters</span></span>

| <span data-ttu-id="ccd50-867">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-867">Name</span></span> | <span data-ttu-id="ccd50-868">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-868">Type</span></span> | <span data-ttu-id="ccd50-869">属性</span><span class="sxs-lookup"><span data-stu-id="ccd50-869">Attributes</span></span> | <span data-ttu-id="ccd50-870">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-870">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="ccd50-871">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="ccd50-871">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="ccd50-872">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="ccd50-872">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="ccd50-873">Function</span><span class="sxs-lookup"><span data-stu-id="ccd50-873">Function</span></span> || <span data-ttu-id="ccd50-p148">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p148">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="ccd50-877">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-877">Object</span></span> | <span data-ttu-id="ccd50-878">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-878">&lt;optional&gt;</span></span> | <span data-ttu-id="ccd50-879">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ccd50-879">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="ccd50-880">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccd50-880">Object</span></span> | <span data-ttu-id="ccd50-881">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-881">&lt;optional&gt;</span></span> | <span data-ttu-id="ccd50-882">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-882">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="ccd50-883">function</span><span class="sxs-lookup"><span data-stu-id="ccd50-883">function</span></span>| <span data-ttu-id="ccd50-884">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-884">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-885">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-885">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-886">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-886">Requirements</span></span>

|<span data-ttu-id="ccd50-887">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-887">Requirement</span></span>| <span data-ttu-id="ccd50-888">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-888">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-889">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-889">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccd50-890">1.7</span><span class="sxs-lookup"><span data-stu-id="ccd50-890">1.7</span></span> |
|[<span data-ttu-id="ccd50-891">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-891">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccd50-892">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-892">ReadItem</span></span> |
|[<span data-ttu-id="ccd50-893">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-893">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccd50-894">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-894">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="ccd50-895">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-895">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="ccd50-896">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ccd50-896">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ccd50-897">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-897">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="ccd50-p149">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p149">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="ccd50-901">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-901">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="ccd50-902">Office アドインを Outlook on the web で実行している場合、編集中のアイテム以外のアイテムに `addItemAttachmentAsync` メソッドでアイテムを添付できます。ただし、これはサポートされていないため、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-902">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-903">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-903">Parameters</span></span>

|<span data-ttu-id="ccd50-904">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-904">Name</span></span>|<span data-ttu-id="ccd50-905">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-905">Type</span></span>|<span data-ttu-id="ccd50-906">属性</span><span class="sxs-lookup"><span data-stu-id="ccd50-906">Attributes</span></span>|<span data-ttu-id="ccd50-907">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-907">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="ccd50-908">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-908">String</span></span>||<span data-ttu-id="ccd50-p150">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p150">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="ccd50-911">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-911">String</span></span>||<span data-ttu-id="ccd50-912">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="ccd50-912">The subject of the item to be attached.</span></span> <span data-ttu-id="ccd50-913">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-913">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="ccd50-914">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-914">Object</span></span>|<span data-ttu-id="ccd50-915">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-915">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-916">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ccd50-916">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ccd50-917">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccd50-917">Object</span></span>|<span data-ttu-id="ccd50-918">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-918">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-919">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-919">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ccd50-920">関数</span><span class="sxs-lookup"><span data-stu-id="ccd50-920">function</span></span>|<span data-ttu-id="ccd50-921">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-921">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-922">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-922">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ccd50-923">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-923">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ccd50-924">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-924">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ccd50-925">エラー</span><span class="sxs-lookup"><span data-stu-id="ccd50-925">Errors</span></span>

|<span data-ttu-id="ccd50-926">エラー コード</span><span class="sxs-lookup"><span data-stu-id="ccd50-926">Error code</span></span>|<span data-ttu-id="ccd50-927">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-927">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="ccd50-928">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-928">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-929">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-929">Requirements</span></span>

|<span data-ttu-id="ccd50-930">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-930">Requirement</span></span>|<span data-ttu-id="ccd50-931">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-931">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-932">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-932">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-933">1.1</span><span class="sxs-lookup"><span data-stu-id="ccd50-933">1.1</span></span>|
|[<span data-ttu-id="ccd50-934">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-934">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-935">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-935">ReadWriteItem</span></span>|
|[<span data-ttu-id="ccd50-936">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-936">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-937">作成</span><span class="sxs-lookup"><span data-stu-id="ccd50-937">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-938">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-938">Example</span></span>

<span data-ttu-id="ccd50-939">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-939">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="ccd50-940">close()</span><span class="sxs-lookup"><span data-stu-id="ccd50-940">close()</span></span>

<span data-ttu-id="ccd50-941">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-941">Closes the current item that is being composed.</span></span>

<span data-ttu-id="ccd50-p152">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p152">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-944">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-944">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="ccd50-945">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-945">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-946">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-946">Requirements</span></span>

|<span data-ttu-id="ccd50-947">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-947">Requirement</span></span>|<span data-ttu-id="ccd50-948">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-948">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-949">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-949">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-950">1.3</span><span class="sxs-lookup"><span data-stu-id="ccd50-950">1.3</span></span>|
|[<span data-ttu-id="ccd50-951">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-951">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-952">制限あり</span><span class="sxs-lookup"><span data-stu-id="ccd50-952">Restricted</span></span>|
|[<span data-ttu-id="ccd50-953">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-953">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-954">新規作成</span><span class="sxs-lookup"><span data-stu-id="ccd50-954">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="ccd50-955">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="ccd50-955">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="ccd50-956">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-956">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-957">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-957">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ccd50-958">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-958">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ccd50-959">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="ccd50-959">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="ccd50-p153">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-963">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-963">Parameters</span></span>

|<span data-ttu-id="ccd50-964">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-964">Name</span></span>|<span data-ttu-id="ccd50-965">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-965">Type</span></span>|<span data-ttu-id="ccd50-966">属性</span><span class="sxs-lookup"><span data-stu-id="ccd50-966">Attributes</span></span>|<span data-ttu-id="ccd50-967">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-967">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="ccd50-968">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-968">String &#124; Object</span></span>||<span data-ttu-id="ccd50-p154">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ccd50-971">**または**</span><span class="sxs-lookup"><span data-stu-id="ccd50-971">**OR**</span></span><br/><span data-ttu-id="ccd50-p155">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="ccd50-974">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-974">String</span></span>|<span data-ttu-id="ccd50-975">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-975">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-p156">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="ccd50-978">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-978">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="ccd50-979">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-979">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-980">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="ccd50-980">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="ccd50-981">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-981">String</span></span>||<span data-ttu-id="ccd50-p157">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="ccd50-984">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-984">String</span></span>||<span data-ttu-id="ccd50-985">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-985">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="ccd50-986">文字列</span><span class="sxs-lookup"><span data-stu-id="ccd50-986">String</span></span>||<span data-ttu-id="ccd50-p158">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="ccd50-989">ブール値</span><span class="sxs-lookup"><span data-stu-id="ccd50-989">Boolean</span></span>||<span data-ttu-id="ccd50-p159">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="ccd50-992">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-992">String</span></span>||<span data-ttu-id="ccd50-p160">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="ccd50-996">function</span><span class="sxs-lookup"><span data-stu-id="ccd50-996">function</span></span>|<span data-ttu-id="ccd50-997">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-997">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-998">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-998">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-999">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-999">Requirements</span></span>

|<span data-ttu-id="ccd50-1000">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1000">Requirement</span></span>|<span data-ttu-id="ccd50-1001">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1001">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1002">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1002">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1003">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-1003">1.0</span></span>|
|[<span data-ttu-id="ccd50-1004">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1004">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1005">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1005">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-1006">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1006">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1007">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-1007">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ccd50-1008">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1008">Examples</span></span>

<span data-ttu-id="ccd50-1009">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1009">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="ccd50-1010">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1010">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="ccd50-1011">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1011">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ccd50-1012">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1012">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="ccd50-1013">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1013">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="ccd50-1014">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1014">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="ccd50-1015">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="ccd50-1015">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="ccd50-1016">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1016">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-1017">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1017">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ccd50-1018">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1018">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ccd50-1019">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1019">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="ccd50-p161">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p161">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-1023">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-1023">Parameters</span></span>

|<span data-ttu-id="ccd50-1024">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-1024">Name</span></span>|<span data-ttu-id="ccd50-1025">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-1025">Type</span></span>|<span data-ttu-id="ccd50-1026">属性</span><span class="sxs-lookup"><span data-stu-id="ccd50-1026">Attributes</span></span>|<span data-ttu-id="ccd50-1027">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-1027">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="ccd50-1028">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1028">String &#124; Object</span></span>||<span data-ttu-id="ccd50-p162">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p162">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ccd50-1031">**または**</span><span class="sxs-lookup"><span data-stu-id="ccd50-1031">**OR**</span></span><br/><span data-ttu-id="ccd50-p163">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p163">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="ccd50-1034">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-1034">String</span></span>|<span data-ttu-id="ccd50-1035">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1035">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-p164">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p164">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="ccd50-1038">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1038">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="ccd50-1039">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1040">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1040">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="ccd50-1041">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-1041">String</span></span>||<span data-ttu-id="ccd50-p165">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p165">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="ccd50-1044">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-1044">String</span></span>||<span data-ttu-id="ccd50-1045">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1045">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="ccd50-1046">文字列</span><span class="sxs-lookup"><span data-stu-id="ccd50-1046">String</span></span>||<span data-ttu-id="ccd50-p166">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p166">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="ccd50-1049">ブール値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1049">Boolean</span></span>||<span data-ttu-id="ccd50-p167">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p167">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="ccd50-1052">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-1052">String</span></span>||<span data-ttu-id="ccd50-p168">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p168">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="ccd50-1056">function</span><span class="sxs-lookup"><span data-stu-id="ccd50-1056">function</span></span>|<span data-ttu-id="ccd50-1057">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1058">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-1059">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1059">Requirements</span></span>

|<span data-ttu-id="ccd50-1060">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1060">Requirement</span></span>|<span data-ttu-id="ccd50-1061">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1061">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1062">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1062">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1063">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-1063">1.0</span></span>|
|[<span data-ttu-id="ccd50-1064">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1064">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1065">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1065">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-1066">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1066">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1067">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-1067">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ccd50-1068">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1068">Examples</span></span>

<span data-ttu-id="ccd50-1069">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1069">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="ccd50-1070">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1070">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="ccd50-1071">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1071">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ccd50-1072">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1072">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="ccd50-1073">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1073">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="ccd50-1074">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1074">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getallinternetheadersasyncoptions-callback"></a><span data-ttu-id="ccd50-1075">getAllInternetHeadersAsync ([オプション], [callback])</span><span class="sxs-lookup"><span data-stu-id="ccd50-1075">getAllInternetHeadersAsync([options], [callback])</span></span>

<span data-ttu-id="ccd50-1076">メッセージのすべてのインターネットヘッダーを文字列として取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1076">Gets all the internet headers for the message as a string.</span></span> <span data-ttu-id="ccd50-1077">Read mode only.</span><span class="sxs-lookup"><span data-stu-id="ccd50-1077">Read mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-1078">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-1078">Parameters</span></span>

|<span data-ttu-id="ccd50-1079">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-1079">Name</span></span>|<span data-ttu-id="ccd50-1080">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-1080">Type</span></span>|<span data-ttu-id="ccd50-1081">属性</span><span class="sxs-lookup"><span data-stu-id="ccd50-1081">Attributes</span></span>|<span data-ttu-id="ccd50-1082">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-1082">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="ccd50-1083">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1083">Object</span></span>|<span data-ttu-id="ccd50-1084">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1084">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1085">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1085">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ccd50-1086">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1086">Object</span></span>|<span data-ttu-id="ccd50-1087">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1087">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1088">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1088">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ccd50-1089">関数</span><span class="sxs-lookup"><span data-stu-id="ccd50-1089">function</span></span>|<span data-ttu-id="ccd50-1090">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1090">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1091">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1091">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="ccd50-1092">成功した場合、インターネットヘッダーデータは、文字列として asyncResult プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1092">On success, the internet headers data is provided in the asyncResult.value property as a string.</span></span> <span data-ttu-id="ccd50-1093">返される文字列値の書式情報については、 [RFC 2183](https://tools.ietf.org/html/rfc2183)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1093">Refer to [RFC 2183](https://tools.ietf.org/html/rfc2183) for the formatting information of the returned string value.</span></span> <span data-ttu-id="ccd50-1094">呼び出しが失敗した場合、asyncResult. error プロパティには、エラーの理由と共にエラーコードが含まれます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1094">If the call fails, the asyncResult.error property will contain an error code with the reason for the failure.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-1095">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-1095">Requirements</span></span>

|<span data-ttu-id="ccd50-1096">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1096">Requirement</span></span>|<span data-ttu-id="ccd50-1097">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1097">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1098">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1098">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1099">プレビュー</span><span class="sxs-lookup"><span data-stu-id="ccd50-1099">Preview</span></span>|
|[<span data-ttu-id="ccd50-1100">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1100">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1101">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1101">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-1102">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1102">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1103">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-1103">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ccd50-1104">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ccd50-1104">Returns:</span></span>

<span data-ttu-id="ccd50-1105">[RFC 2183](https://tools.ietf.org/html/rfc2183)に従って書式設定された文字列としてのインターネットヘッダーデータ。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1105">The internet headers data as a string formatted according to [RFC 2183](https://tools.ietf.org/html/rfc2183).</span></span>

<span data-ttu-id="ccd50-1106">型:String</span><span class="sxs-lookup"><span data-stu-id="ccd50-1106">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="ccd50-1107">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1107">Example</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="ccd50-1108">getAttachmentContentAsync (attachmentId, [options], [callback]) > [Attachmentcontent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="ccd50-1108">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="ccd50-1109">メッセージまたは予定から指定された添付ファイルを取得し`AttachmentContent` 、それをオブジェクトとして返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1109">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="ccd50-1110">メソッド`getAttachmentContentAsync`は、指定された id の添付ファイルをアイテムから取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1110">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="ccd50-1111">ベストプラクティスとして、識別子を使用して、または`getAttachmentsAsync` `item.attachments`の呼び出しで attachmentIds を取得したのと同じセッションの添付ファイルを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1111">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="ccd50-1112">Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1112">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="ccd50-1113">ユーザーがアプリを閉じたとき、またはインラインフォームの作成が開始されたときに、別のウィンドウで続行するためにフォームをポップアウトした後、セッションが終了します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1113">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-1114">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-1114">Parameters</span></span>

|<span data-ttu-id="ccd50-1115">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-1115">Name</span></span>|<span data-ttu-id="ccd50-1116">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-1116">Type</span></span>|<span data-ttu-id="ccd50-1117">属性</span><span class="sxs-lookup"><span data-stu-id="ccd50-1117">Attributes</span></span>|<span data-ttu-id="ccd50-1118">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-1118">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="ccd50-1119">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-1119">String</span></span>||<span data-ttu-id="ccd50-1120">取得する添付ファイルの識別子を指定します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1120">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="ccd50-1121">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1121">Object</span></span>|<span data-ttu-id="ccd50-1122">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1122">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1123">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1123">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ccd50-1124">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1124">Object</span></span>|<span data-ttu-id="ccd50-1125">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1125">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1126">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1126">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ccd50-1127">関数</span><span class="sxs-lookup"><span data-stu-id="ccd50-1127">function</span></span>|<span data-ttu-id="ccd50-1128">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1128">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1129">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1129">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-1130">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1130">Requirements</span></span>

|<span data-ttu-id="ccd50-1131">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1131">Requirement</span></span>|<span data-ttu-id="ccd50-1132">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1132">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1133">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1134">プレビュー</span><span class="sxs-lookup"><span data-stu-id="ccd50-1134">Preview</span></span>|
|[<span data-ttu-id="ccd50-1135">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1136">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-1137">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1138">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-1138">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ccd50-1139">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ccd50-1139">Returns:</span></span>

<span data-ttu-id="ccd50-1140">型: [Attachmentcontent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="ccd50-1140">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="ccd50-1141">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1141">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="ccd50-1142">getAttachmentsAsync ([オプション], [callback]) > Array. <[Attachmentdetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="ccd50-1142">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="ccd50-1143">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1143">Gets the item's attachments as an array.</span></span> <span data-ttu-id="ccd50-1144">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1144">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-1145">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-1145">Parameters</span></span>

|<span data-ttu-id="ccd50-1146">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-1146">Name</span></span>|<span data-ttu-id="ccd50-1147">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-1147">Type</span></span>|<span data-ttu-id="ccd50-1148">属性</span><span class="sxs-lookup"><span data-stu-id="ccd50-1148">Attributes</span></span>|<span data-ttu-id="ccd50-1149">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-1149">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="ccd50-1150">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1150">Object</span></span>|<span data-ttu-id="ccd50-1151">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1151">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1152">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1152">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ccd50-1153">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1153">Object</span></span>|<span data-ttu-id="ccd50-1154">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1154">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1155">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1155">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ccd50-1156">function</span><span class="sxs-lookup"><span data-stu-id="ccd50-1156">function</span></span>|<span data-ttu-id="ccd50-1157">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1158">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1158">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-1159">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-1159">Requirements</span></span>

|<span data-ttu-id="ccd50-1160">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1160">Requirement</span></span>|<span data-ttu-id="ccd50-1161">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1161">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1162">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1162">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1163">プレビュー</span><span class="sxs-lookup"><span data-stu-id="ccd50-1163">Preview</span></span>|
|[<span data-ttu-id="ccd50-1164">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1164">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1165">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1165">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-1166">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1166">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1167">作成</span><span class="sxs-lookup"><span data-stu-id="ccd50-1167">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="ccd50-1168">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ccd50-1168">Returns:</span></span>

<span data-ttu-id="ccd50-1169">型: Array. <[attachmentdetails 詳細](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="ccd50-1169">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="ccd50-1170">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1170">Example</span></span>

<span data-ttu-id="ccd50-1171">次の例では、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1171">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="ccd50-1172">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="ccd50-1172">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="ccd50-1173">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1173">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-1174">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1174">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-1175">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1175">Requirements</span></span>

|<span data-ttu-id="ccd50-1176">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1176">Requirement</span></span>|<span data-ttu-id="ccd50-1177">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1177">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1178">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1179">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-1179">1.0</span></span>|
|[<span data-ttu-id="ccd50-1180">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1180">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1181">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-1182">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1182">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1183">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-1183">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ccd50-1184">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ccd50-1184">Returns:</span></span>

<span data-ttu-id="ccd50-1185">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="ccd50-1185">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="ccd50-1186">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1186">Example</span></span>

<span data-ttu-id="ccd50-1187">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1187">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="ccd50-1188">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="ccd50-1188">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="ccd50-1189">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1189">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-1190">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1190">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-1191">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-1191">Parameters</span></span>

|<span data-ttu-id="ccd50-1192">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-1192">Name</span></span>|<span data-ttu-id="ccd50-1193">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-1193">Type</span></span>|<span data-ttu-id="ccd50-1194">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-1194">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="ccd50-1195">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="ccd50-1195">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="ccd50-1196">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1196">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-1197">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-1197">Requirements</span></span>

|<span data-ttu-id="ccd50-1198">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1198">Requirement</span></span>|<span data-ttu-id="ccd50-1199">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1199">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1200">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1201">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-1201">1.0</span></span>|
|[<span data-ttu-id="ccd50-1202">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1202">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1203">制限あり</span><span class="sxs-lookup"><span data-stu-id="ccd50-1203">Restricted</span></span>|
|[<span data-ttu-id="ccd50-1204">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1204">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1205">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-1205">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ccd50-1206">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ccd50-1206">Returns:</span></span>

<span data-ttu-id="ccd50-1207">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1207">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="ccd50-1208">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1208">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="ccd50-1209">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1209">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="ccd50-1210">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1210">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="ccd50-1211">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1211">Value of `entityType`</span></span>|<span data-ttu-id="ccd50-1212">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="ccd50-1212">Type of objects in returned array</span></span>|<span data-ttu-id="ccd50-1213">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1213">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="ccd50-1214">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-1214">String</span></span>|<span data-ttu-id="ccd50-1215">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="ccd50-1215">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="ccd50-1216">連絡先</span><span class="sxs-lookup"><span data-stu-id="ccd50-1216">Contact</span></span>|<span data-ttu-id="ccd50-1217">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ccd50-1217">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="ccd50-1218">文字列</span><span class="sxs-lookup"><span data-stu-id="ccd50-1218">String</span></span>|<span data-ttu-id="ccd50-1219">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ccd50-1219">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="ccd50-1220">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="ccd50-1220">MeetingSuggestion</span></span>|<span data-ttu-id="ccd50-1221">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ccd50-1221">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="ccd50-1222">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="ccd50-1222">PhoneNumber</span></span>|<span data-ttu-id="ccd50-1223">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="ccd50-1223">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="ccd50-1224">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="ccd50-1224">TaskSuggestion</span></span>|<span data-ttu-id="ccd50-1225">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ccd50-1225">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="ccd50-1226">文字列</span><span class="sxs-lookup"><span data-stu-id="ccd50-1226">String</span></span>|<span data-ttu-id="ccd50-1227">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="ccd50-1227">**Restricted**</span></span>|

<span data-ttu-id="ccd50-1228">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="ccd50-1228">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="ccd50-1229">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1229">Example</span></span>

<span data-ttu-id="ccd50-1230">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1230">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="ccd50-1231">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="ccd50-1231">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="ccd50-1232">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1232">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-1233">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1233">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ccd50-1234">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1234">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-1235">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-1235">Parameters</span></span>

|<span data-ttu-id="ccd50-1236">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-1236">Name</span></span>|<span data-ttu-id="ccd50-1237">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-1237">Type</span></span>|<span data-ttu-id="ccd50-1238">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-1238">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="ccd50-1239">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-1239">String</span></span>|<span data-ttu-id="ccd50-1240">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1240">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-1241">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-1241">Requirements</span></span>

|<span data-ttu-id="ccd50-1242">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1242">Requirement</span></span>|<span data-ttu-id="ccd50-1243">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1243">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1244">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1244">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1245">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-1245">1.0</span></span>|
|[<span data-ttu-id="ccd50-1246">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1246">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1247">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1247">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-1248">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1248">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1249">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-1249">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ccd50-1250">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ccd50-1250">Returns:</span></span>

<span data-ttu-id="ccd50-p174">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p174">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="ccd50-1253">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="ccd50-1253">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

<br>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="ccd50-1254">、Office.context.mailbox.item.getinitializationcontextasync ([オプション], [callback])</span><span class="sxs-lookup"><span data-stu-id="ccd50-1254">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="ccd50-1255">[アクション可能なメッセージによってアドインがアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されたときに渡される初期化データを取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1255">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-1256">このメソッドは、Outlook 2016 以降の Windows (16.0.8413.1000 より後のバージョン) および Outlook on the Office 365 でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1256">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-1257">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-1257">Parameters</span></span>

|<span data-ttu-id="ccd50-1258">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-1258">Name</span></span>|<span data-ttu-id="ccd50-1259">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-1259">Type</span></span>|<span data-ttu-id="ccd50-1260">属性</span><span class="sxs-lookup"><span data-stu-id="ccd50-1260">Attributes</span></span>|<span data-ttu-id="ccd50-1261">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-1261">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="ccd50-1262">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccd50-1262">Object</span></span>|<span data-ttu-id="ccd50-1263">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1263">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1264">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1264">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ccd50-1265">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccd50-1265">Object</span></span>|<span data-ttu-id="ccd50-1266">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1266">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1267">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1267">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ccd50-1268">function</span><span class="sxs-lookup"><span data-stu-id="ccd50-1268">function</span></span>|<span data-ttu-id="ccd50-1269">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1269">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1270">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1270">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ccd50-1271">成功すると、初期化データが文字列とし`asyncResult.value`てプロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1271">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="ccd50-1272">初期化コンテキストがない場合、 `asyncResult`オブジェクトには、 `Error` `code`プロパティがに`9020`設定されたオブジェクトと`name`プロパティがに`GenericResponseError`設定されたオブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1272">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-1273">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-1273">Requirements</span></span>

|<span data-ttu-id="ccd50-1274">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1274">Requirement</span></span>|<span data-ttu-id="ccd50-1275">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1275">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1276">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1277">プレビュー</span><span class="sxs-lookup"><span data-stu-id="ccd50-1277">Preview</span></span>|
|[<span data-ttu-id="ccd50-1278">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1279">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-1280">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1281">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-1281">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-1282">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1282">Example</span></span>

```js
// Get the initialization context (if present).
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object.
        var context = JSON.parse(asyncResult.value);
        // Do something with context.
      } else {
        // Empty context, treat as no context.
      }
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

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="ccd50-1283">getItemIdAsync ([オプション], callback)</span><span class="sxs-lookup"><span data-stu-id="ccd50-1283">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="ccd50-1284">保存されたアイテムの ID を非同期に取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1284">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="ccd50-1285">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1285">Compose mode only.</span></span>

<span data-ttu-id="ccd50-1286">このメソッドを呼び出すと、コールバックメソッドによってアイテム ID が返されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1286">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-1287">アドインが新規作成モードの`getItemIdAsync`アイテムに対して呼び出しを行う場合 ( `itemId` EWS または REST API を使用するため)、Outlook がキャッシュモードの場合は、アイテムがサーバーに同期されるまでしばらく時間がかかる場合があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1287">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="ccd50-1288">アイテムが同期されるまで、 `itemId`は認識されず、を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1288">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-1289">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-1289">Parameters</span></span>

|<span data-ttu-id="ccd50-1290">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-1290">Name</span></span>|<span data-ttu-id="ccd50-1291">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-1291">Type</span></span>|<span data-ttu-id="ccd50-1292">属性</span><span class="sxs-lookup"><span data-stu-id="ccd50-1292">Attributes</span></span>|<span data-ttu-id="ccd50-1293">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-1293">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="ccd50-1294">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccd50-1294">Object</span></span>|<span data-ttu-id="ccd50-1295">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1295">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1296">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1296">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ccd50-1297">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccd50-1297">Object</span></span>|<span data-ttu-id="ccd50-1298">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1298">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1299">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1299">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ccd50-1300">function</span><span class="sxs-lookup"><span data-stu-id="ccd50-1300">function</span></span>||<span data-ttu-id="ccd50-1301">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1301">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ccd50-1302">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1302">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ccd50-1303">エラー</span><span class="sxs-lookup"><span data-stu-id="ccd50-1303">Errors</span></span>

|<span data-ttu-id="ccd50-1304">エラー コード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1304">Error code</span></span>|<span data-ttu-id="ccd50-1305">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-1305">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="ccd50-1306">この id は、アイテムが保存されるまでは取得できません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1306">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-1307">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-1307">Requirements</span></span>

|<span data-ttu-id="ccd50-1308">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1308">Requirement</span></span>|<span data-ttu-id="ccd50-1309">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1309">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1310">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1310">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1311">プレビュー</span><span class="sxs-lookup"><span data-stu-id="ccd50-1311">Preview</span></span>|
|[<span data-ttu-id="ccd50-1312">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1312">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1313">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1313">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-1314">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1314">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1315">作成</span><span class="sxs-lookup"><span data-stu-id="ccd50-1315">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="ccd50-1316">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1316">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="ccd50-1317">次の例は、コールバック関数`result`に渡されるパラメーターの構造を示しています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1317">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="ccd50-1318">プロパティ`value`には、アイテムの ID が含まれています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1318">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="ccd50-1319">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="ccd50-1319">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="ccd50-1320">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1320">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-1321">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1321">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ccd50-p178">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p178">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="ccd50-1325">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1325">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="ccd50-1326">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1326">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="ccd50-p179">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-1330">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1330">Requirements</span></span>

|<span data-ttu-id="ccd50-1331">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1331">Requirement</span></span>|<span data-ttu-id="ccd50-1332">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1332">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1333">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1334">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-1334">1.0</span></span>|
|[<span data-ttu-id="ccd50-1335">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1336">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-1337">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1338">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-1338">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ccd50-1339">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ccd50-1339">Returns:</span></span>

<span data-ttu-id="ccd50-p180">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="ccd50-1342">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="ccd50-1342">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ccd50-1343">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1343">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="ccd50-1344">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1344">Example</span></span>

<span data-ttu-id="ccd50-1345">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1345">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="ccd50-1346">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="ccd50-1346">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="ccd50-1347">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1347">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-1348">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1348">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ccd50-1349">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1349">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="ccd50-p181">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p181">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-1352">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-1352">Parameters</span></span>

|<span data-ttu-id="ccd50-1353">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-1353">Name</span></span>|<span data-ttu-id="ccd50-1354">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-1354">Type</span></span>|<span data-ttu-id="ccd50-1355">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-1355">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="ccd50-1356">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-1356">String</span></span>|<span data-ttu-id="ccd50-1357">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1357">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-1358">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-1358">Requirements</span></span>

|<span data-ttu-id="ccd50-1359">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1359">Requirement</span></span>|<span data-ttu-id="ccd50-1360">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1360">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1361">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1362">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-1362">1.0</span></span>|
|[<span data-ttu-id="ccd50-1363">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1364">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-1365">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1366">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-1366">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ccd50-1367">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ccd50-1367">Returns:</span></span>

<span data-ttu-id="ccd50-1368">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1368">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="ccd50-1369">型: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="ccd50-1369">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="ccd50-1370">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1370">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="ccd50-1371">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="ccd50-1371">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="ccd50-1372">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1372">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="ccd50-p182">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p182">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-1375">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-1375">Parameters</span></span>

|<span data-ttu-id="ccd50-1376">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-1376">Name</span></span>|<span data-ttu-id="ccd50-1377">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-1377">Type</span></span>|<span data-ttu-id="ccd50-1378">属性</span><span class="sxs-lookup"><span data-stu-id="ccd50-1378">Attributes</span></span>|<span data-ttu-id="ccd50-1379">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-1379">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="ccd50-1380">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="ccd50-1380">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="ccd50-p183">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p183">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="ccd50-1384">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1384">Object</span></span>|<span data-ttu-id="ccd50-1385">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1385">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1386">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1386">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ccd50-1387">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1387">Object</span></span>|<span data-ttu-id="ccd50-1388">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1388">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1389">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1389">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ccd50-1390">function</span><span class="sxs-lookup"><span data-stu-id="ccd50-1390">function</span></span>||<span data-ttu-id="ccd50-1391">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1391">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ccd50-1392">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1392">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="ccd50-1393">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1393">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-1394">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-1394">Requirements</span></span>

|<span data-ttu-id="ccd50-1395">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1395">Requirement</span></span>|<span data-ttu-id="ccd50-1396">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1396">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1397">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1398">1.2</span><span class="sxs-lookup"><span data-stu-id="ccd50-1398">1.2</span></span>|
|[<span data-ttu-id="ccd50-1399">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1399">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1400">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-1401">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1401">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1402">作成</span><span class="sxs-lookup"><span data-stu-id="ccd50-1402">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="ccd50-1403">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ccd50-1403">Returns:</span></span>

<span data-ttu-id="ccd50-1404">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1404">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="ccd50-1405">型:String</span><span class="sxs-lookup"><span data-stu-id="ccd50-1405">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="ccd50-1406">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1406">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="ccd50-1407">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="ccd50-1407">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="ccd50-1408">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1408">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="ccd50-1409">強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1409">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-1410">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1410">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-1411">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1411">Requirements</span></span>

|<span data-ttu-id="ccd50-1412">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1412">Requirement</span></span>|<span data-ttu-id="ccd50-1413">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1413">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1414">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1414">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1415">1.6</span><span class="sxs-lookup"><span data-stu-id="ccd50-1415">1.6</span></span>|
|[<span data-ttu-id="ccd50-1416">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1416">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1417">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1417">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-1418">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1418">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1419">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-1419">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ccd50-1420">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ccd50-1420">Returns:</span></span>

<span data-ttu-id="ccd50-1421">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="ccd50-1421">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="ccd50-1422">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1422">Example</span></span>

<span data-ttu-id="ccd50-1423">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1423">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="ccd50-1424">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="ccd50-1424">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="ccd50-p186">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p186">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-1427">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1427">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ccd50-p187">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p187">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="ccd50-1431">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1431">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="ccd50-1432">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1432">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="ccd50-p188">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p188">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccd50-1436">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-1436">Requirements</span></span>

|<span data-ttu-id="ccd50-1437">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1437">Requirement</span></span>|<span data-ttu-id="ccd50-1438">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1438">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1439">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1439">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1440">1.6</span><span class="sxs-lookup"><span data-stu-id="ccd50-1440">1.6</span></span>|
|[<span data-ttu-id="ccd50-1441">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1441">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1442">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1442">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-1443">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1443">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1444">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccd50-1444">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ccd50-1445">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ccd50-1445">Returns:</span></span>

<span data-ttu-id="ccd50-p189">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p189">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="ccd50-1448">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1448">Example</span></span>

<span data-ttu-id="ccd50-1449">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1449">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="ccd50-1450">getSharedPropertiesAsync ([options], callback)</span><span class="sxs-lookup"><span data-stu-id="ccd50-1450">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="ccd50-1451">共有フォルダー、予定表、またはメールボックス内の選択した予定またはメッセージのプロパティを取得します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1451">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-1452">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-1452">Parameters</span></span>

|<span data-ttu-id="ccd50-1453">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-1453">Name</span></span>|<span data-ttu-id="ccd50-1454">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-1454">Type</span></span>|<span data-ttu-id="ccd50-1455">属性</span><span class="sxs-lookup"><span data-stu-id="ccd50-1455">Attributes</span></span>|<span data-ttu-id="ccd50-1456">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-1456">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="ccd50-1457">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1457">Object</span></span>|<span data-ttu-id="ccd50-1458">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1458">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1459">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1459">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ccd50-1460">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1460">Object</span></span>|<span data-ttu-id="ccd50-1461">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1461">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1462">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1462">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ccd50-1463">function</span><span class="sxs-lookup"><span data-stu-id="ccd50-1463">function</span></span>||<span data-ttu-id="ccd50-1464">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1464">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ccd50-1465">共有プロパティは、 [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) `asyncResult.value`プロパティのオブジェクトとして提供されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1465">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="ccd50-1466">このオブジェクトは、アイテムの共有プロパティを取得するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1466">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-1467">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-1467">Requirements</span></span>

|<span data-ttu-id="ccd50-1468">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1468">Requirement</span></span>|<span data-ttu-id="ccd50-1469">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1469">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1470">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1471">プレビュー</span><span class="sxs-lookup"><span data-stu-id="ccd50-1471">Preview</span></span>|
|[<span data-ttu-id="ccd50-1472">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1473">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-1474">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1475">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-1475">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-1476">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1476">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="ccd50-1477">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ccd50-1477">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="ccd50-1478">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1478">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="ccd50-p191">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p191">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-1482">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-1482">Parameters</span></span>

|<span data-ttu-id="ccd50-1483">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-1483">Name</span></span>|<span data-ttu-id="ccd50-1484">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-1484">Type</span></span>|<span data-ttu-id="ccd50-1485">属性</span><span class="sxs-lookup"><span data-stu-id="ccd50-1485">Attributes</span></span>|<span data-ttu-id="ccd50-1486">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-1486">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="ccd50-1487">function</span><span class="sxs-lookup"><span data-stu-id="ccd50-1487">function</span></span>||<span data-ttu-id="ccd50-1488">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1488">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ccd50-1489">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1489">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="ccd50-1490">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1490">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="ccd50-1491">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1491">Object</span></span>|<span data-ttu-id="ccd50-1492">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1492">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1493">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1493">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="ccd50-1494">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1494">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-1495">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-1495">Requirements</span></span>

|<span data-ttu-id="ccd50-1496">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1496">Requirement</span></span>|<span data-ttu-id="ccd50-1497">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1497">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1498">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1498">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1499">1.0</span><span class="sxs-lookup"><span data-stu-id="ccd50-1499">1.0</span></span>|
|[<span data-ttu-id="ccd50-1500">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1500">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1501">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1501">ReadItem</span></span>|
|[<span data-ttu-id="ccd50-1502">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1502">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1503">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-1503">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-1504">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1504">Example</span></span>

<span data-ttu-id="ccd50-p194">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p194">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="ccd50-1508">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ccd50-1508">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="ccd50-1509">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1509">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="ccd50-1510">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1510">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="ccd50-1511">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1511">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="ccd50-1512">Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1512">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="ccd50-1513">ユーザーがアプリを閉じたとき、またはインラインフォームの作成が開始されたときに、別のウィンドウで続行するためにフォームをポップアウトした後、セッションが終了します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1513">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-1514">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-1514">Parameters</span></span>

|<span data-ttu-id="ccd50-1515">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-1515">Name</span></span>|<span data-ttu-id="ccd50-1516">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-1516">Type</span></span>|<span data-ttu-id="ccd50-1517">属性</span><span class="sxs-lookup"><span data-stu-id="ccd50-1517">Attributes</span></span>|<span data-ttu-id="ccd50-1518">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-1518">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="ccd50-1519">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-1519">String</span></span>||<span data-ttu-id="ccd50-1520">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1520">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="ccd50-1521">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1521">Object</span></span>|<span data-ttu-id="ccd50-1522">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1522">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1523">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1523">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ccd50-1524">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1524">Object</span></span>|<span data-ttu-id="ccd50-1525">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1525">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1526">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1526">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ccd50-1527">function</span><span class="sxs-lookup"><span data-stu-id="ccd50-1527">function</span></span>|<span data-ttu-id="ccd50-1528">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1528">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1529">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1529">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ccd50-1530">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1530">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ccd50-1531">エラー</span><span class="sxs-lookup"><span data-stu-id="ccd50-1531">Errors</span></span>

|<span data-ttu-id="ccd50-1532">エラー コード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1532">Error code</span></span>|<span data-ttu-id="ccd50-1533">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-1533">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="ccd50-1534">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1534">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-1535">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-1535">Requirements</span></span>

|<span data-ttu-id="ccd50-1536">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1536">Requirement</span></span>|<span data-ttu-id="ccd50-1537">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1537">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1538">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1539">1.1</span><span class="sxs-lookup"><span data-stu-id="ccd50-1539">1.1</span></span>|
|[<span data-ttu-id="ccd50-1540">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1540">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1541">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1541">ReadWriteItem</span></span>|
|[<span data-ttu-id="ccd50-1542">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1542">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1543">作成</span><span class="sxs-lookup"><span data-stu-id="ccd50-1543">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-1544">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1544">Example</span></span>

<span data-ttu-id="ccd50-1545">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1545">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="ccd50-1546">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ccd50-1546">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="ccd50-1547">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1547">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="ccd50-1548">現在、サポートされて`Office.EventType.AttachmentsChanged`いる`Office.EventType.AppointmentTimeChanged`イベント`Office.EventType.EnhancedLocationsChanged`の`Office.EventType.RecipientsChanged`種類は`Office.EventType.RecurrenceChanged`、、、、、です。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1548">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-1549">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-1549">Parameters</span></span>

| <span data-ttu-id="ccd50-1550">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-1550">Name</span></span> | <span data-ttu-id="ccd50-1551">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-1551">Type</span></span> | <span data-ttu-id="ccd50-1552">属性</span><span class="sxs-lookup"><span data-stu-id="ccd50-1552">Attributes</span></span> | <span data-ttu-id="ccd50-1553">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-1553">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="ccd50-1554">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="ccd50-1554">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="ccd50-1555">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1555">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="ccd50-1556">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccd50-1556">Object</span></span> | <span data-ttu-id="ccd50-1557">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1557">&lt;optional&gt;</span></span> | <span data-ttu-id="ccd50-1558">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1558">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="ccd50-1559">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1559">Object</span></span> | <span data-ttu-id="ccd50-1560">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1560">&lt;optional&gt;</span></span> | <span data-ttu-id="ccd50-1561">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1561">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="ccd50-1562">function</span><span class="sxs-lookup"><span data-stu-id="ccd50-1562">function</span></span>| <span data-ttu-id="ccd50-1563">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1563">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1564">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1564">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-1565">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1565">Requirements</span></span>

|<span data-ttu-id="ccd50-1566">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1566">Requirement</span></span>| <span data-ttu-id="ccd50-1567">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1567">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1568">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1568">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccd50-1569">1.7</span><span class="sxs-lookup"><span data-stu-id="ccd50-1569">1.7</span></span> |
|[<span data-ttu-id="ccd50-1570">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1570">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccd50-1571">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1571">ReadItem</span></span> |
|[<span data-ttu-id="ccd50-1572">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1572">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccd50-1573">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccd50-1573">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="ccd50-1574">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="ccd50-1574">saveAsync([options], callback)</span></span>

<span data-ttu-id="ccd50-1575">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1575">Asynchronously saves an item.</span></span>

<span data-ttu-id="ccd50-1576">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1576">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="ccd50-1577">Outlook on the web またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1577">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="ccd50-1578">キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1578">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-1579">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1579">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="ccd50-1580">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1580">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="ccd50-p198">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p198">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="ccd50-1584">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1584">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="ccd50-1585">Outlook on Mac では、会議の保存はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1585">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="ccd50-1586">`saveAsync` メソッドは、作成モードの会議から呼び出されると失敗します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1586">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="ccd50-1587">回避策については、「[Office JS API を使用して Outlook for Mac で会議を下書きとして保存できない](https://support.microsoft.com/help/4505745)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1587">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="ccd50-1588">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1588">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-1589">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-1589">Parameters</span></span>

|<span data-ttu-id="ccd50-1590">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-1590">Name</span></span>|<span data-ttu-id="ccd50-1591">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-1591">Type</span></span>|<span data-ttu-id="ccd50-1592">属性</span><span class="sxs-lookup"><span data-stu-id="ccd50-1592">Attributes</span></span>|<span data-ttu-id="ccd50-1593">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-1593">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="ccd50-1594">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccd50-1594">Object</span></span>|<span data-ttu-id="ccd50-1595">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1595">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1596">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1596">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ccd50-1597">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1597">Object</span></span>|<span data-ttu-id="ccd50-1598">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1598">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1599">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1599">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ccd50-1600">関数</span><span class="sxs-lookup"><span data-stu-id="ccd50-1600">function</span></span>||<span data-ttu-id="ccd50-1601">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1601">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ccd50-1602">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1602">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-1603">Requirements</span><span class="sxs-lookup"><span data-stu-id="ccd50-1603">Requirements</span></span>

|<span data-ttu-id="ccd50-1604">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1604">Requirement</span></span>|<span data-ttu-id="ccd50-1605">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1605">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1606">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1607">1.3</span><span class="sxs-lookup"><span data-stu-id="ccd50-1607">1.3</span></span>|
|[<span data-ttu-id="ccd50-1608">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1609">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1609">ReadWriteItem</span></span>|
|[<span data-ttu-id="ccd50-1610">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1611">作成</span><span class="sxs-lookup"><span data-stu-id="ccd50-1611">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="ccd50-1612">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1612">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="ccd50-p200">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p200">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="ccd50-1615">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="ccd50-1615">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="ccd50-1616">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1616">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="ccd50-p201">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p201">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccd50-1620">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccd50-1620">Parameters</span></span>

|<span data-ttu-id="ccd50-1621">名前</span><span class="sxs-lookup"><span data-stu-id="ccd50-1621">Name</span></span>|<span data-ttu-id="ccd50-1622">種類</span><span class="sxs-lookup"><span data-stu-id="ccd50-1622">Type</span></span>|<span data-ttu-id="ccd50-1623">属性</span><span class="sxs-lookup"><span data-stu-id="ccd50-1623">Attributes</span></span>|<span data-ttu-id="ccd50-1624">説明</span><span class="sxs-lookup"><span data-stu-id="ccd50-1624">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="ccd50-1625">String</span><span class="sxs-lookup"><span data-stu-id="ccd50-1625">String</span></span>||<span data-ttu-id="ccd50-p202">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-p202">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="ccd50-1629">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1629">Object</span></span>|<span data-ttu-id="ccd50-1630">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1630">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1631">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1631">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ccd50-1632">Object</span><span class="sxs-lookup"><span data-stu-id="ccd50-1632">Object</span></span>|<span data-ttu-id="ccd50-1633">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1633">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1634">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1634">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="ccd50-1635">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="ccd50-1635">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="ccd50-1636">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccd50-1636">&lt;optional&gt;</span></span>|<span data-ttu-id="ccd50-1637">`text` の場合、Outlook on the web とデスクトップ クライアントでは現在のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1637">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="ccd50-1638">フィールドが HTML エディターの場合、データが HTML の場合でもテキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1638">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="ccd50-1639">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Outlook on the web では現在のスタイルが適用され、Outlook デスクトップ クライアントでは既定のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1639">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="ccd50-1640">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1640">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="ccd50-1641">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1641">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="ccd50-1642">function</span><span class="sxs-lookup"><span data-stu-id="ccd50-1642">function</span></span>||<span data-ttu-id="ccd50-1643">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccd50-1643">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccd50-1644">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1644">Requirements</span></span>

|<span data-ttu-id="ccd50-1645">要件</span><span class="sxs-lookup"><span data-stu-id="ccd50-1645">Requirement</span></span>|<span data-ttu-id="ccd50-1646">値</span><span class="sxs-lookup"><span data-stu-id="ccd50-1646">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccd50-1647">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccd50-1647">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ccd50-1648">1.2</span><span class="sxs-lookup"><span data-stu-id="ccd50-1648">1.2</span></span>|
|[<span data-ttu-id="ccd50-1649">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccd50-1649">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ccd50-1650">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ccd50-1650">ReadWriteItem</span></span>|
|[<span data-ttu-id="ccd50-1651">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccd50-1651">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ccd50-1652">作成</span><span class="sxs-lookup"><span data-stu-id="ccd50-1652">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ccd50-1653">例</span><span class="sxs-lookup"><span data-stu-id="ccd50-1653">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
