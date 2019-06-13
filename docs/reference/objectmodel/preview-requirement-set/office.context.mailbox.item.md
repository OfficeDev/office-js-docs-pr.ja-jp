---
title: Office. アイテム-プレビュー要件セット
description: ''
ms.date: 06/11/2019
localization_priority: Normal
ms.openlocfilehash: 9844fdeba274271a226501d6c0a8694e164f6ac7
ms.sourcegitcommit: 3f84b2caa73d7fe1eb0d15e32ea4dec459e2ff53
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/12/2019
ms.locfileid: "34910274"
---
# <a name="item"></a><span data-ttu-id="2357b-102">item</span><span class="sxs-lookup"><span data-stu-id="2357b-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="2357b-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="2357b-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="2357b-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-106">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-106">Requirements</span></span>

|<span data-ttu-id="2357b-107">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-107">Requirement</span></span>|<span data-ttu-id="2357b-108">値</span><span class="sxs-lookup"><span data-stu-id="2357b-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-110">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-110">1.0</span></span>|
|[<span data-ttu-id="2357b-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="2357b-112">Restricted</span></span>|
|[<span data-ttu-id="2357b-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2357b-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-115">Members and methods</span></span>

| <span data-ttu-id="2357b-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-116">Member</span></span> | <span data-ttu-id="2357b-117">種類</span><span class="sxs-lookup"><span data-stu-id="2357b-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2357b-118">attachments</span><span class="sxs-lookup"><span data-stu-id="2357b-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="2357b-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-119">Member</span></span> |
| [<span data-ttu-id="2357b-120">bcc</span><span class="sxs-lookup"><span data-stu-id="2357b-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="2357b-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-121">Member</span></span> |
| [<span data-ttu-id="2357b-122">body</span><span class="sxs-lookup"><span data-stu-id="2357b-122">body</span></span>](#body-body) | <span data-ttu-id="2357b-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-123">Member</span></span> |
| [<span data-ttu-id="2357b-124">categories</span><span class="sxs-lookup"><span data-stu-id="2357b-124">categories</span></span>](#categories-categories) | <span data-ttu-id="2357b-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-125">Member</span></span> |
| [<span data-ttu-id="2357b-126">cc</span><span class="sxs-lookup"><span data-stu-id="2357b-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2357b-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-127">Member</span></span> |
| [<span data-ttu-id="2357b-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="2357b-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="2357b-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-129">Member</span></span> |
| [<span data-ttu-id="2357b-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="2357b-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="2357b-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-131">Member</span></span> |
| [<span data-ttu-id="2357b-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="2357b-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="2357b-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-133">Member</span></span> |
| [<span data-ttu-id="2357b-134">end</span><span class="sxs-lookup"><span data-stu-id="2357b-134">end</span></span>](#end-datetime) | <span data-ttu-id="2357b-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-135">Member</span></span> |
| [<span data-ttu-id="2357b-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="2357b-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="2357b-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-137">Member</span></span> |
| [<span data-ttu-id="2357b-138">from</span><span class="sxs-lookup"><span data-stu-id="2357b-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="2357b-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-139">Member</span></span> |
| [<span data-ttu-id="2357b-140">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="2357b-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="2357b-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-141">Member</span></span> |
| [<span data-ttu-id="2357b-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="2357b-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="2357b-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-143">Member</span></span> |
| [<span data-ttu-id="2357b-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="2357b-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="2357b-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-145">Member</span></span> |
| [<span data-ttu-id="2357b-146">itemId</span><span class="sxs-lookup"><span data-stu-id="2357b-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="2357b-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-147">Member</span></span> |
| [<span data-ttu-id="2357b-148">itemType</span><span class="sxs-lookup"><span data-stu-id="2357b-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="2357b-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-149">Member</span></span> |
| [<span data-ttu-id="2357b-150">location</span><span class="sxs-lookup"><span data-stu-id="2357b-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="2357b-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-151">Member</span></span> |
| [<span data-ttu-id="2357b-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="2357b-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="2357b-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-153">Member</span></span> |
| [<span data-ttu-id="2357b-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="2357b-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="2357b-155">Member</span><span class="sxs-lookup"><span data-stu-id="2357b-155">Member</span></span> |
| [<span data-ttu-id="2357b-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="2357b-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2357b-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-157">Member</span></span> |
| [<span data-ttu-id="2357b-158">organizer</span><span class="sxs-lookup"><span data-stu-id="2357b-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="2357b-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-159">Member</span></span> |
| [<span data-ttu-id="2357b-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="2357b-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="2357b-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-161">Member</span></span> |
| [<span data-ttu-id="2357b-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="2357b-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2357b-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-163">Member</span></span> |
| [<span data-ttu-id="2357b-164">sender</span><span class="sxs-lookup"><span data-stu-id="2357b-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="2357b-165">Member</span><span class="sxs-lookup"><span data-stu-id="2357b-165">Member</span></span> |
| [<span data-ttu-id="2357b-166">系列 Id</span><span class="sxs-lookup"><span data-stu-id="2357b-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="2357b-167">Member</span><span class="sxs-lookup"><span data-stu-id="2357b-167">Member</span></span> |
| [<span data-ttu-id="2357b-168">start</span><span class="sxs-lookup"><span data-stu-id="2357b-168">start</span></span>](#start-datetime) | <span data-ttu-id="2357b-169">Member</span><span class="sxs-lookup"><span data-stu-id="2357b-169">Member</span></span> |
| [<span data-ttu-id="2357b-170">subject</span><span class="sxs-lookup"><span data-stu-id="2357b-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="2357b-171">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-171">Member</span></span> |
| [<span data-ttu-id="2357b-172">to</span><span class="sxs-lookup"><span data-stu-id="2357b-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2357b-173">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-173">Member</span></span> |
| [<span data-ttu-id="2357b-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2357b-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="2357b-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-175">Method</span></span> |
| [<span data-ttu-id="2357b-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="2357b-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="2357b-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-177">Method</span></span> |
| [<span data-ttu-id="2357b-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="2357b-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="2357b-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-179">Method</span></span> |
| [<span data-ttu-id="2357b-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2357b-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="2357b-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-181">Method</span></span> |
| [<span data-ttu-id="2357b-182">close</span><span class="sxs-lookup"><span data-stu-id="2357b-182">close</span></span>](#close) | <span data-ttu-id="2357b-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-183">Method</span></span> |
| [<span data-ttu-id="2357b-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="2357b-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="2357b-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-185">Method</span></span> |
| [<span data-ttu-id="2357b-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="2357b-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="2357b-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-187">Method</span></span> |
| [<span data-ttu-id="2357b-188">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="2357b-188">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="2357b-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-189">Method</span></span> |
| [<span data-ttu-id="2357b-190">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="2357b-190">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="2357b-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-191">Method</span></span> |
| [<span data-ttu-id="2357b-192">getEntities</span><span class="sxs-lookup"><span data-stu-id="2357b-192">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="2357b-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-193">Method</span></span> |
| [<span data-ttu-id="2357b-194">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="2357b-194">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="2357b-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-195">Method</span></span> |
| [<span data-ttu-id="2357b-196">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="2357b-196">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="2357b-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-197">Method</span></span> |
| [<span data-ttu-id="2357b-198">、Office.context.mailbox.item.getinitializationcontextasync</span><span class="sxs-lookup"><span data-stu-id="2357b-198">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="2357b-199">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-199">Method</span></span> |
| [<span data-ttu-id="2357b-200">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="2357b-200">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="2357b-201">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-201">Method</span></span> |
| [<span data-ttu-id="2357b-202">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="2357b-202">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="2357b-203">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-203">Method</span></span> |
| [<span data-ttu-id="2357b-204">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="2357b-204">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="2357b-205">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-205">Method</span></span> |
| [<span data-ttu-id="2357b-206">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="2357b-206">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="2357b-207">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-207">Method</span></span> |
| [<span data-ttu-id="2357b-208">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="2357b-208">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="2357b-209">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-209">Method</span></span> |
| [<span data-ttu-id="2357b-210">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="2357b-210">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="2357b-211">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-211">Method</span></span> |
| [<span data-ttu-id="2357b-212">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="2357b-212">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="2357b-213">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-213">Method</span></span> |
| [<span data-ttu-id="2357b-214">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="2357b-214">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="2357b-215">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-215">Method</span></span> |
| [<span data-ttu-id="2357b-216">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2357b-216">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="2357b-217">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-217">Method</span></span> |
| [<span data-ttu-id="2357b-218">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="2357b-218">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="2357b-219">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-219">Method</span></span> |
| [<span data-ttu-id="2357b-220">saveAsync</span><span class="sxs-lookup"><span data-stu-id="2357b-220">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="2357b-221">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-221">Method</span></span> |
| [<span data-ttu-id="2357b-222">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="2357b-222">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="2357b-223">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-223">Method</span></span> |

### <a name="example"></a><span data-ttu-id="2357b-224">例</span><span class="sxs-lookup"><span data-stu-id="2357b-224">Example</span></span>

<span data-ttu-id="2357b-225">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="2357b-225">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="2357b-226">メンバー</span><span class="sxs-lookup"><span data-stu-id="2357b-226">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="2357b-227">添付ファイル: <[Attachmentdetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="2357b-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="2357b-228">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-228">Gets the item's attachments as an array.</span></span> <span data-ttu-id="2357b-229">Read mode only.</span><span class="sxs-lookup"><span data-stu-id="2357b-229">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-230">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="2357b-230">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="2357b-231">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2357b-231">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="2357b-232">型</span><span class="sxs-lookup"><span data-stu-id="2357b-232">Type</span></span>

*   <span data-ttu-id="2357b-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="2357b-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-234">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-234">Requirements</span></span>

|<span data-ttu-id="2357b-235">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-235">Requirement</span></span>|<span data-ttu-id="2357b-236">値</span><span class="sxs-lookup"><span data-stu-id="2357b-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-237">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-238">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-238">1.0</span></span>|
|[<span data-ttu-id="2357b-239">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-240">ReadItem</span></span>|
|[<span data-ttu-id="2357b-241">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-242">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-242">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-243">例</span><span class="sxs-lookup"><span data-stu-id="2357b-243">Example</span></span>

<span data-ttu-id="2357b-244">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="2357b-244">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

---
---

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="2357b-245">bcc:[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2357b-245">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="2357b-246">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-246">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="2357b-247">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="2357b-247">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2357b-248">型</span><span class="sxs-lookup"><span data-stu-id="2357b-248">Type</span></span>

*   [<span data-ttu-id="2357b-249">受信者</span><span class="sxs-lookup"><span data-stu-id="2357b-249">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="2357b-250">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-250">Requirements</span></span>

|<span data-ttu-id="2357b-251">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-251">Requirement</span></span>|<span data-ttu-id="2357b-252">値</span><span class="sxs-lookup"><span data-stu-id="2357b-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-253">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-254">1.1</span><span class="sxs-lookup"><span data-stu-id="2357b-254">1.1</span></span>|
|[<span data-ttu-id="2357b-255">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-256">ReadItem</span></span>|
|[<span data-ttu-id="2357b-257">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-258">作成</span><span class="sxs-lookup"><span data-stu-id="2357b-258">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-259">例</span><span class="sxs-lookup"><span data-stu-id="2357b-259">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

---
---

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="2357b-260">本文:[本文](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="2357b-260">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="2357b-261">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-261">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="2357b-262">型</span><span class="sxs-lookup"><span data-stu-id="2357b-262">Type</span></span>

*   [<span data-ttu-id="2357b-263">Body</span><span class="sxs-lookup"><span data-stu-id="2357b-263">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="2357b-264">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-264">Requirements</span></span>

|<span data-ttu-id="2357b-265">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-265">Requirement</span></span>|<span data-ttu-id="2357b-266">値</span><span class="sxs-lookup"><span data-stu-id="2357b-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-267">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-268">1.1</span><span class="sxs-lookup"><span data-stu-id="2357b-268">1.1</span></span>|
|[<span data-ttu-id="2357b-269">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-270">ReadItem</span></span>|
|[<span data-ttu-id="2357b-271">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-272">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-272">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-273">例</span><span class="sxs-lookup"><span data-stu-id="2357b-273">Example</span></span>

<span data-ttu-id="2357b-274">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-274">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="2357b-275">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="2357b-275">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

---
---

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="2357b-276">カテゴリ:[カテゴリ](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="2357b-276">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="2357b-277">アイテムのカテゴリを管理するためのメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-277">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-278">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2357b-278">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="2357b-279">型</span><span class="sxs-lookup"><span data-stu-id="2357b-279">Type</span></span>

*   [<span data-ttu-id="2357b-280">カテゴリ</span><span class="sxs-lookup"><span data-stu-id="2357b-280">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="2357b-281">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-281">Requirements</span></span>

|<span data-ttu-id="2357b-282">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-282">Requirement</span></span>|<span data-ttu-id="2357b-283">値</span><span class="sxs-lookup"><span data-stu-id="2357b-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-284">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-285">プレビュー</span><span class="sxs-lookup"><span data-stu-id="2357b-285">Preview</span></span>|
|[<span data-ttu-id="2357b-286">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-287">ReadItem</span></span>|
|[<span data-ttu-id="2357b-288">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-289">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-289">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-290">例</span><span class="sxs-lookup"><span data-stu-id="2357b-290">Example</span></span>

<span data-ttu-id="2357b-291">この例では、アイテムのカテゴリを取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-291">This example gets the item's categories.</span></span>

```javascript
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="2357b-292">cc: <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)>|[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2357b-292">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="2357b-293">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="2357b-293">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="2357b-294">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="2357b-294">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2357b-295">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="2357b-295">Read mode</span></span>

<span data-ttu-id="2357b-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="2357b-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="2357b-298">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="2357b-298">Compose mode</span></span>

<span data-ttu-id="2357b-299">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-299">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2357b-300">型</span><span class="sxs-lookup"><span data-stu-id="2357b-300">Type</span></span>

*   <span data-ttu-id="2357b-301">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2357b-301">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-302">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-302">Requirements</span></span>

|<span data-ttu-id="2357b-303">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-303">Requirement</span></span>|<span data-ttu-id="2357b-304">値</span><span class="sxs-lookup"><span data-stu-id="2357b-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-305">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-306">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-306">1.0</span></span>|
|[<span data-ttu-id="2357b-307">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-307">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-308">ReadItem</span></span>|
|[<span data-ttu-id="2357b-309">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-309">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-310">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-310">Compose or Read</span></span>|

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="2357b-311">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="2357b-311">(nullable) conversationId: String</span></span>

<span data-ttu-id="2357b-312">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-312">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="2357b-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="2357b-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="2357b-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="2357b-317">Type</span><span class="sxs-lookup"><span data-stu-id="2357b-317">Type</span></span>

*   <span data-ttu-id="2357b-318">String</span><span class="sxs-lookup"><span data-stu-id="2357b-318">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-319">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-319">Requirements</span></span>

|<span data-ttu-id="2357b-320">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-320">Requirement</span></span>|<span data-ttu-id="2357b-321">値</span><span class="sxs-lookup"><span data-stu-id="2357b-321">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-322">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-322">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-323">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-323">1.0</span></span>|
|[<span data-ttu-id="2357b-324">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-324">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-325">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-325">ReadItem</span></span>|
|[<span data-ttu-id="2357b-326">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-326">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-327">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-327">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-328">例</span><span class="sxs-lookup"><span data-stu-id="2357b-328">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="2357b-329">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="2357b-329">dateTimeCreated: Date</span></span>

<span data-ttu-id="2357b-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="2357b-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2357b-332">型</span><span class="sxs-lookup"><span data-stu-id="2357b-332">Type</span></span>

*   <span data-ttu-id="2357b-333">日付</span><span class="sxs-lookup"><span data-stu-id="2357b-333">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-334">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-334">Requirements</span></span>

|<span data-ttu-id="2357b-335">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-335">Requirement</span></span>|<span data-ttu-id="2357b-336">値</span><span class="sxs-lookup"><span data-stu-id="2357b-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-337">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-338">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-338">1.0</span></span>|
|[<span data-ttu-id="2357b-339">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-339">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-340">ReadItem</span></span>|
|[<span data-ttu-id="2357b-341">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-341">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-342">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-342">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-343">例</span><span class="sxs-lookup"><span data-stu-id="2357b-343">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="2357b-344">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="2357b-344">dateTimeModified: Date</span></span>

<span data-ttu-id="2357b-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="2357b-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-347">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2357b-347">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="2357b-348">型</span><span class="sxs-lookup"><span data-stu-id="2357b-348">Type</span></span>

*   <span data-ttu-id="2357b-349">日付</span><span class="sxs-lookup"><span data-stu-id="2357b-349">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-350">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-350">Requirements</span></span>

|<span data-ttu-id="2357b-351">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-351">Requirement</span></span>|<span data-ttu-id="2357b-352">値</span><span class="sxs-lookup"><span data-stu-id="2357b-352">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-353">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-353">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-354">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-354">1.0</span></span>|
|[<span data-ttu-id="2357b-355">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-355">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-356">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-356">ReadItem</span></span>|
|[<span data-ttu-id="2357b-357">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-357">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-358">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-358">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-359">例</span><span class="sxs-lookup"><span data-stu-id="2357b-359">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="2357b-360">終了: 日付 |[時間](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="2357b-360">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="2357b-361">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="2357b-361">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="2357b-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="2357b-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2357b-364">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="2357b-364">Read mode</span></span>

<span data-ttu-id="2357b-365">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-365">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="2357b-366">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="2357b-366">Compose mode</span></span>

<span data-ttu-id="2357b-367">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-367">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="2357b-368">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2357b-368">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="2357b-369">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="2357b-369">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="2357b-370">型</span><span class="sxs-lookup"><span data-stu-id="2357b-370">Type</span></span>

*   <span data-ttu-id="2357b-371">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="2357b-371">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-372">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-372">Requirements</span></span>

|<span data-ttu-id="2357b-373">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-373">Requirement</span></span>|<span data-ttu-id="2357b-374">値</span><span class="sxs-lookup"><span data-stu-id="2357b-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-375">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-376">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-376">1.0</span></span>|
|[<span data-ttu-id="2357b-377">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-377">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-378">ReadItem</span></span>|
|[<span data-ttu-id="2357b-379">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-379">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-380">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-380">Compose or Read</span></span>|

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="2357b-381">enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="2357b-381">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="2357b-382">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="2357b-382">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2357b-383">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="2357b-383">Read mode</span></span>

<span data-ttu-id="2357b-384">この`enhancedLocation`プロパティは、予定に関連付けられている場所 ( [locationdetails](/javascript/api/outlook/office.locationdetails)オブジェクトで表される) のセットを取得できる[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-384">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="2357b-385">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="2357b-385">Compose mode</span></span>

<span data-ttu-id="2357b-386">この`enhancedLocation`プロパティは、予定の場所を取得、削除、または追加するためのメソッドを提供する[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-386">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="2357b-387">型</span><span class="sxs-lookup"><span data-stu-id="2357b-387">Type</span></span>

*   [<span data-ttu-id="2357b-388">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="2357b-388">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="2357b-389">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-389">Requirements</span></span>

|<span data-ttu-id="2357b-390">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-390">Requirement</span></span>|<span data-ttu-id="2357b-391">値</span><span class="sxs-lookup"><span data-stu-id="2357b-391">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-392">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-392">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-393">プレビュー</span><span class="sxs-lookup"><span data-stu-id="2357b-393">Preview</span></span>|
|[<span data-ttu-id="2357b-394">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-394">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-395">ReadItem</span></span>|
|[<span data-ttu-id="2357b-396">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-396">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-397">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-397">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-398">例</span><span class="sxs-lookup"><span data-stu-id="2357b-398">Example</span></span>

<span data-ttu-id="2357b-399">次の例では、予定に関連付けられている現在の場所を取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-399">The following example gets the current locations associated with the appointment.</span></span>

```javascript
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

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="2357b-400">from: [emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)|[from](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="2357b-400">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="2357b-401">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-401">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="2357b-p112">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="2357b-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-404">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="2357b-404">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2357b-405">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="2357b-405">Read mode</span></span>

<span data-ttu-id="2357b-406">プロパティ`from`は`EmailAddressDetails`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-406">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="2357b-407">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="2357b-407">Compose mode</span></span>

<span data-ttu-id="2357b-408">プロパティ`from`は、from `From`値を取得するメソッドを提供するオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-408">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2357b-409">型</span><span class="sxs-lookup"><span data-stu-id="2357b-409">Type</span></span>

*   <span data-ttu-id="2357b-410">[電子メールアドレス](/javascript/api/outlook/office.emailaddressdetails) | [の](/javascript/api/outlook/office.from)詳細</span><span class="sxs-lookup"><span data-stu-id="2357b-410">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-411">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-411">Requirements</span></span>

|<span data-ttu-id="2357b-412">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-412">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="2357b-413">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-414">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-414">1.0</span></span>|<span data-ttu-id="2357b-415">1.7</span><span class="sxs-lookup"><span data-stu-id="2357b-415">1.7</span></span>|
|[<span data-ttu-id="2357b-416">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-416">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-417">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-417">ReadItem</span></span>|<span data-ttu-id="2357b-418">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2357b-418">ReadWriteItem</span></span>|
|[<span data-ttu-id="2357b-419">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-419">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-420">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-420">Read</span></span>|<span data-ttu-id="2357b-421">作成</span><span class="sxs-lookup"><span data-stu-id="2357b-421">Compose</span></span>|

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="2357b-422">internetHeaders: [internetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="2357b-422">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="2357b-423">メッセージのインターネットヘッダーを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="2357b-423">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="2357b-424">型</span><span class="sxs-lookup"><span data-stu-id="2357b-424">Type</span></span>

*   [<span data-ttu-id="2357b-425">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="2357b-425">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="2357b-426">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-426">Requirements</span></span>

|<span data-ttu-id="2357b-427">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-427">Requirement</span></span>|<span data-ttu-id="2357b-428">値</span><span class="sxs-lookup"><span data-stu-id="2357b-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-429">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-430">プレビュー</span><span class="sxs-lookup"><span data-stu-id="2357b-430">Preview</span></span>|
|[<span data-ttu-id="2357b-431">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-431">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-432">ReadItem</span></span>|
|[<span data-ttu-id="2357b-433">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-433">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-434">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-434">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-435">例</span><span class="sxs-lookup"><span data-stu-id="2357b-435">Example</span></span>

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="2357b-436">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="2357b-436">internetMessageId: String</span></span>

<span data-ttu-id="2357b-p113">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="2357b-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2357b-439">Type</span><span class="sxs-lookup"><span data-stu-id="2357b-439">Type</span></span>

*   <span data-ttu-id="2357b-440">String</span><span class="sxs-lookup"><span data-stu-id="2357b-440">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-441">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-441">Requirements</span></span>

|<span data-ttu-id="2357b-442">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-442">Requirement</span></span>|<span data-ttu-id="2357b-443">値</span><span class="sxs-lookup"><span data-stu-id="2357b-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-444">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-445">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-445">1.0</span></span>|
|[<span data-ttu-id="2357b-446">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-447">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-447">ReadItem</span></span>|
|[<span data-ttu-id="2357b-448">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-449">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-449">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-450">例</span><span class="sxs-lookup"><span data-stu-id="2357b-450">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="2357b-451">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="2357b-451">itemClass: String</span></span>

<span data-ttu-id="2357b-p114">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="2357b-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="2357b-p115">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="2357b-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="2357b-456">型</span><span class="sxs-lookup"><span data-stu-id="2357b-456">Type</span></span>|<span data-ttu-id="2357b-457">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-457">Description</span></span>|<span data-ttu-id="2357b-458">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="2357b-458">item class</span></span>|
|---|---|---|
|<span data-ttu-id="2357b-459">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="2357b-459">Appointment items</span></span>|<span data-ttu-id="2357b-460">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="2357b-460">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="2357b-461">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="2357b-461">Message items</span></span>|<span data-ttu-id="2357b-462">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="2357b-462">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="2357b-463">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-463">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="2357b-464">型</span><span class="sxs-lookup"><span data-stu-id="2357b-464">Type</span></span>

*   <span data-ttu-id="2357b-465">String</span><span class="sxs-lookup"><span data-stu-id="2357b-465">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-466">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-466">Requirements</span></span>

|<span data-ttu-id="2357b-467">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-467">Requirement</span></span>|<span data-ttu-id="2357b-468">値</span><span class="sxs-lookup"><span data-stu-id="2357b-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-469">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-470">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-470">1.0</span></span>|
|[<span data-ttu-id="2357b-471">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-472">ReadItem</span></span>|
|[<span data-ttu-id="2357b-473">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-474">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-474">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-475">例</span><span class="sxs-lookup"><span data-stu-id="2357b-475">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="2357b-476">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="2357b-476">(nullable) itemId: String</span></span>

<span data-ttu-id="2357b-p116">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="2357b-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-479">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="2357b-479">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="2357b-480">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="2357b-480">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="2357b-481">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2357b-481">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="2357b-482">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2357b-482">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="2357b-p118">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="2357b-485">Type</span><span class="sxs-lookup"><span data-stu-id="2357b-485">Type</span></span>

*   <span data-ttu-id="2357b-486">String</span><span class="sxs-lookup"><span data-stu-id="2357b-486">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-487">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-487">Requirements</span></span>

|<span data-ttu-id="2357b-488">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-488">Requirement</span></span>|<span data-ttu-id="2357b-489">値</span><span class="sxs-lookup"><span data-stu-id="2357b-489">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-490">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-490">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-491">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-491">1.0</span></span>|
|[<span data-ttu-id="2357b-492">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-492">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-493">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-493">ReadItem</span></span>|
|[<span data-ttu-id="2357b-494">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-494">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-495">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-495">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-496">例</span><span class="sxs-lookup"><span data-stu-id="2357b-496">Example</span></span>

<span data-ttu-id="2357b-p119">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

---
---

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="2357b-499">itemType: [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="2357b-499">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="2357b-500">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-500">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="2357b-501">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="2357b-501">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="2357b-502">型</span><span class="sxs-lookup"><span data-stu-id="2357b-502">Type</span></span>

*   [<span data-ttu-id="2357b-503">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="2357b-503">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="2357b-504">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-504">Requirements</span></span>

|<span data-ttu-id="2357b-505">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-505">Requirement</span></span>|<span data-ttu-id="2357b-506">値</span><span class="sxs-lookup"><span data-stu-id="2357b-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-507">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-508">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-508">1.0</span></span>|
|[<span data-ttu-id="2357b-509">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-510">ReadItem</span></span>|
|[<span data-ttu-id="2357b-511">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-512">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-512">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-513">例</span><span class="sxs-lookup"><span data-stu-id="2357b-513">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="2357b-514">場所: String |[場所](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="2357b-514">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="2357b-515">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="2357b-515">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2357b-516">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="2357b-516">Read mode</span></span>

<span data-ttu-id="2357b-517">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-517">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="2357b-518">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="2357b-518">Compose mode</span></span>

<span data-ttu-id="2357b-519">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-519">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2357b-520">型</span><span class="sxs-lookup"><span data-stu-id="2357b-520">Type</span></span>

*   <span data-ttu-id="2357b-521">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="2357b-521">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-522">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-522">Requirements</span></span>

|<span data-ttu-id="2357b-523">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-523">Requirement</span></span>|<span data-ttu-id="2357b-524">値</span><span class="sxs-lookup"><span data-stu-id="2357b-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-525">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-525">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-526">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-526">1.0</span></span>|
|[<span data-ttu-id="2357b-527">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-527">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-528">ReadItem</span></span>|
|[<span data-ttu-id="2357b-529">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-529">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-530">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-530">Compose or Read</span></span>|

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="2357b-531">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="2357b-531">normalizedSubject: String</span></span>

<span data-ttu-id="2357b-p120">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="2357b-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="2357b-p121">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="2357b-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="2357b-536">型</span><span class="sxs-lookup"><span data-stu-id="2357b-536">Type</span></span>

*   <span data-ttu-id="2357b-537">String</span><span class="sxs-lookup"><span data-stu-id="2357b-537">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-538">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-538">Requirements</span></span>

|<span data-ttu-id="2357b-539">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-539">Requirement</span></span>|<span data-ttu-id="2357b-540">値</span><span class="sxs-lookup"><span data-stu-id="2357b-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-541">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-542">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-542">1.0</span></span>|
|[<span data-ttu-id="2357b-543">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-544">ReadItem</span></span>|
|[<span data-ttu-id="2357b-545">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-546">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-547">例</span><span class="sxs-lookup"><span data-stu-id="2357b-547">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="2357b-548">notificationMessages: [Notificationmessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="2357b-548">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="2357b-549">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-549">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="2357b-550">型</span><span class="sxs-lookup"><span data-stu-id="2357b-550">Type</span></span>

*   [<span data-ttu-id="2357b-551">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="2357b-551">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="2357b-552">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-552">Requirements</span></span>

|<span data-ttu-id="2357b-553">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-553">Requirement</span></span>|<span data-ttu-id="2357b-554">値</span><span class="sxs-lookup"><span data-stu-id="2357b-554">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-555">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-555">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-556">1.3</span><span class="sxs-lookup"><span data-stu-id="2357b-556">1.3</span></span>|
|[<span data-ttu-id="2357b-557">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-557">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-558">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-558">ReadItem</span></span>|
|[<span data-ttu-id="2357b-559">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-559">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-560">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-560">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-561">例</span><span class="sxs-lookup"><span data-stu-id="2357b-561">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="2357b-562">任意出席者: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)>|[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2357b-562">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="2357b-563">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="2357b-563">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="2357b-564">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="2357b-564">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2357b-565">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="2357b-565">Read mode</span></span>

<span data-ttu-id="2357b-566">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-566">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="2357b-567">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="2357b-567">Compose mode</span></span>

<span data-ttu-id="2357b-568">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-568">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2357b-569">型</span><span class="sxs-lookup"><span data-stu-id="2357b-569">Type</span></span>

*   <span data-ttu-id="2357b-570">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2357b-570">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-571">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-571">Requirements</span></span>

|<span data-ttu-id="2357b-572">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-572">Requirement</span></span>|<span data-ttu-id="2357b-573">値</span><span class="sxs-lookup"><span data-stu-id="2357b-573">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-574">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-574">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-575">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-575">1.0</span></span>|
|[<span data-ttu-id="2357b-576">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-576">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-577">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-577">ReadItem</span></span>|
|[<span data-ttu-id="2357b-578">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-578">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-579">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-579">Compose or Read</span></span>|

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="2357b-580">開催者: [emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)|[開催者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="2357b-580">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="2357b-581">指定した会議の開催者の電子メールアドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-581">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2357b-582">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="2357b-582">Read mode</span></span>

<span data-ttu-id="2357b-583">プロパティ`organizer`は、会議の開催者を表す[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-583">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="2357b-584">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="2357b-584">Compose mode</span></span>

<span data-ttu-id="2357b-585">プロパティ`organizer`は、開催者の値を取得するためのメソッドを提供する[オーガナイザー](/javascript/api/outlook/office.organizer)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-585">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="2357b-586">型</span><span class="sxs-lookup"><span data-stu-id="2357b-586">Type</span></span>

*   <span data-ttu-id="2357b-587">[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails) | [開催者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="2357b-587">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-588">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-588">Requirements</span></span>

|<span data-ttu-id="2357b-589">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-589">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="2357b-590">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-590">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-591">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-591">1.0</span></span>|<span data-ttu-id="2357b-592">1.7</span><span class="sxs-lookup"><span data-stu-id="2357b-592">1.7</span></span>|
|[<span data-ttu-id="2357b-593">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-593">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-594">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-594">ReadItem</span></span>|<span data-ttu-id="2357b-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2357b-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="2357b-596">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-597">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-597">Read</span></span>|<span data-ttu-id="2357b-598">作成</span><span class="sxs-lookup"><span data-stu-id="2357b-598">Compose</span></span>|

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="2357b-599">(nullable) 定期的なスケジュール:[定期的](/javascript/api/outlook/office.recurrence)なアイテム</span><span class="sxs-lookup"><span data-stu-id="2357b-599">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="2357b-600">予定の定期的なパターンを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="2357b-600">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="2357b-601">会議出席依頼の定期的なパターンを取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-601">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="2357b-602">予定アイテムの読み取りおよび作成モード。</span><span class="sxs-lookup"><span data-stu-id="2357b-602">Read and compose modes for appointment items.</span></span> <span data-ttu-id="2357b-603">会議出席依頼アイテムの閲覧モード。</span><span class="sxs-lookup"><span data-stu-id="2357b-603">Read mode for meeting request items.</span></span>

<span data-ttu-id="2357b-604">この`recurrence`プロパティは、アイテムが series または series 内のインスタンスの場合、定期的な予定または会議出席依頼に対して[定期的](/javascript/api/outlook/office.recurrence)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-604">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="2357b-605">`null`は、単一の予定および1つの予定の会議出席依頼に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-605">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="2357b-606">`undefined`は、会議出席依頼ではないメッセージに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-606">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="2357b-607">注: 会議出席依頼に`itemClass`は、IPM という値があります。出席依頼。</span><span class="sxs-lookup"><span data-stu-id="2357b-607">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="2357b-608">注: 定期的なオブジェクトが`null`の場合は、そのオブジェクトが単一の予定または1つの予定の会議出席依頼であり、データ系列の一部ではないことを示します。</span><span class="sxs-lookup"><span data-stu-id="2357b-608">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2357b-609">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="2357b-609">Read mode</span></span>

<span data-ttu-id="2357b-610">この`recurrence`プロパティは、定期的な予定を表す[定期的](/javascript/api/outlook/office.recurrence)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-610">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="2357b-611">これは、予定および会議出席依頼に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-611">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="2357b-612">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="2357b-612">Compose mode</span></span>

<span data-ttu-id="2357b-613">この`recurrence`プロパティは、予定の繰り返しを管理するためのメソッドを提供する[定期的](/javascript/api/outlook/office.recurrence)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-613">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="2357b-614">これは予定に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-614">This is available for appointments.</span></span>

```javascript
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

##### <a name="type"></a><span data-ttu-id="2357b-615">型</span><span class="sxs-lookup"><span data-stu-id="2357b-615">Type</span></span>

* [<span data-ttu-id="2357b-616">繰り返さ</span><span class="sxs-lookup"><span data-stu-id="2357b-616">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="2357b-617">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-617">Requirement</span></span>|<span data-ttu-id="2357b-618">値</span><span class="sxs-lookup"><span data-stu-id="2357b-618">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-619">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-619">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-620">1.7</span><span class="sxs-lookup"><span data-stu-id="2357b-620">1.7</span></span>|
|[<span data-ttu-id="2357b-621">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-621">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-622">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-622">ReadItem</span></span>|
|[<span data-ttu-id="2357b-623">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-623">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-624">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-624">Compose or Read</span></span>|

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="2357b-625">requiredatて dees: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)>|[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2357b-625">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="2357b-626">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="2357b-626">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="2357b-627">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="2357b-627">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2357b-628">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="2357b-628">Read mode</span></span>

<span data-ttu-id="2357b-629">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-629">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="2357b-630">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="2357b-630">Compose mode</span></span>

<span data-ttu-id="2357b-631">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-631">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="2357b-632">型</span><span class="sxs-lookup"><span data-stu-id="2357b-632">Type</span></span>

*   <span data-ttu-id="2357b-633">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2357b-633">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-634">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-634">Requirements</span></span>

|<span data-ttu-id="2357b-635">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-635">Requirement</span></span>|<span data-ttu-id="2357b-636">値</span><span class="sxs-lookup"><span data-stu-id="2357b-636">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-637">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-638">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-638">1.0</span></span>|
|[<span data-ttu-id="2357b-639">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-639">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-640">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-640">ReadItem</span></span>|
|[<span data-ttu-id="2357b-641">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-641">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-642">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-642">Compose or Read</span></span>|

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="2357b-643">sender: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="2357b-643">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="2357b-p128">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="2357b-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="2357b-p129">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsfrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="2357b-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-648">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="2357b-648">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="2357b-649">型</span><span class="sxs-lookup"><span data-stu-id="2357b-649">Type</span></span>

*   [<span data-ttu-id="2357b-650">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2357b-650">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="2357b-651">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-651">Requirements</span></span>

|<span data-ttu-id="2357b-652">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-652">Requirement</span></span>|<span data-ttu-id="2357b-653">値</span><span class="sxs-lookup"><span data-stu-id="2357b-653">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-654">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-654">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-655">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-655">1.0</span></span>|
|[<span data-ttu-id="2357b-656">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-656">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-657">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-657">ReadItem</span></span>|
|[<span data-ttu-id="2357b-658">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-658">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-659">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-659">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-660">例</span><span class="sxs-lookup"><span data-stu-id="2357b-660">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="2357b-661">(nullable) 系列 Id: String</span><span class="sxs-lookup"><span data-stu-id="2357b-661">(nullable) seriesId: String</span></span>

<span data-ttu-id="2357b-662">インスタンスが属する系列の id を取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-662">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="2357b-663">OWA および Outlook で、は`seriesId` 、このアイテムが属する親 (シリーズ) アイテムの Exchange Web サービス (EWS) ID を返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-663">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="2357b-664">ただし、iOS と Android では、 `seriesId`は親アイテムの REST ID を返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-664">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-665">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="2357b-665">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="2357b-666">`seriesId`プロパティが OUTLOOK REST API で使用される outlook id と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="2357b-666">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="2357b-667">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2357b-667">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="2357b-668">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2357b-668">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="2357b-669">この`seriesId`プロパティは`null` 、単一の予定、系列のアイテム、会議出席依頼などの親アイテムを持たないアイテムに`undefined`対して、会議出席依頼以外のアイテムに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-669">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="2357b-670">Type</span><span class="sxs-lookup"><span data-stu-id="2357b-670">Type</span></span>

* <span data-ttu-id="2357b-671">String</span><span class="sxs-lookup"><span data-stu-id="2357b-671">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-672">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-672">Requirements</span></span>

|<span data-ttu-id="2357b-673">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-673">Requirement</span></span>|<span data-ttu-id="2357b-674">値</span><span class="sxs-lookup"><span data-stu-id="2357b-674">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-675">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-675">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-676">1.7</span><span class="sxs-lookup"><span data-stu-id="2357b-676">1.7</span></span>|
|[<span data-ttu-id="2357b-677">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-677">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-678">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-678">ReadItem</span></span>|
|[<span data-ttu-id="2357b-679">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-679">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-680">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-680">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-681">例</span><span class="sxs-lookup"><span data-stu-id="2357b-681">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

---
---

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="2357b-682">開始: 日付 |[時間](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="2357b-682">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="2357b-683">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="2357b-683">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="2357b-p132">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="2357b-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2357b-686">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="2357b-686">Read mode</span></span>

<span data-ttu-id="2357b-687">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-687">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="2357b-688">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="2357b-688">Compose mode</span></span>

<span data-ttu-id="2357b-689">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-689">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="2357b-690">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2357b-690">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="2357b-691">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="2357b-691">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="2357b-692">型</span><span class="sxs-lookup"><span data-stu-id="2357b-692">Type</span></span>

*   <span data-ttu-id="2357b-693">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="2357b-693">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-694">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-694">Requirements</span></span>

|<span data-ttu-id="2357b-695">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-695">Requirement</span></span>|<span data-ttu-id="2357b-696">値</span><span class="sxs-lookup"><span data-stu-id="2357b-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-697">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-698">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-698">1.0</span></span>|
|[<span data-ttu-id="2357b-699">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-699">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-700">ReadItem</span></span>|
|[<span data-ttu-id="2357b-701">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-701">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-702">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-702">Compose or Read</span></span>|

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="2357b-703">subject: String |[件名](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="2357b-703">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="2357b-704">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="2357b-704">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="2357b-705">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="2357b-705">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2357b-706">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="2357b-706">Read mode</span></span>

<span data-ttu-id="2357b-p133">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="2357b-709">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="2357b-709">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="2357b-710">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="2357b-710">Compose mode</span></span>
<span data-ttu-id="2357b-711">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-711">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="2357b-712">型</span><span class="sxs-lookup"><span data-stu-id="2357b-712">Type</span></span>

*   <span data-ttu-id="2357b-713">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="2357b-713">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-714">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-714">Requirements</span></span>

|<span data-ttu-id="2357b-715">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-715">Requirement</span></span>|<span data-ttu-id="2357b-716">値</span><span class="sxs-lookup"><span data-stu-id="2357b-716">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-717">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-717">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-718">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-718">1.0</span></span>|
|[<span data-ttu-id="2357b-719">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-719">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-720">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-720">ReadItem</span></span>|
|[<span data-ttu-id="2357b-721">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-721">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-722">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-722">Compose or Read</span></span>|

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="2357b-723">宛先: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)>|[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2357b-723">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="2357b-724">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="2357b-724">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="2357b-725">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="2357b-725">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2357b-726">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="2357b-726">Read mode</span></span>

<span data-ttu-id="2357b-p135">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="2357b-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="2357b-729">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="2357b-729">Compose mode</span></span>

<span data-ttu-id="2357b-730">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-730">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2357b-731">型</span><span class="sxs-lookup"><span data-stu-id="2357b-731">Type</span></span>

*   <span data-ttu-id="2357b-732">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2357b-732">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-733">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-733">Requirements</span></span>

|<span data-ttu-id="2357b-734">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-734">Requirement</span></span>|<span data-ttu-id="2357b-735">値</span><span class="sxs-lookup"><span data-stu-id="2357b-735">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-736">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-736">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-737">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-737">1.0</span></span>|
|[<span data-ttu-id="2357b-738">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-738">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-739">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-739">ReadItem</span></span>|
|[<span data-ttu-id="2357b-740">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-740">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-741">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-741">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="2357b-742">メソッド</span><span class="sxs-lookup"><span data-stu-id="2357b-742">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="2357b-743">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2357b-743">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="2357b-744">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="2357b-744">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="2357b-745">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="2357b-745">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="2357b-746">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-746">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-747">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-747">Parameters</span></span>
|<span data-ttu-id="2357b-748">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-748">Name</span></span>|<span data-ttu-id="2357b-749">種類</span><span class="sxs-lookup"><span data-stu-id="2357b-749">Type</span></span>|<span data-ttu-id="2357b-750">属性</span><span class="sxs-lookup"><span data-stu-id="2357b-750">Attributes</span></span>|<span data-ttu-id="2357b-751">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-751">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="2357b-752">String</span><span class="sxs-lookup"><span data-stu-id="2357b-752">String</span></span>||<span data-ttu-id="2357b-p136">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="2357b-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="2357b-755">String</span><span class="sxs-lookup"><span data-stu-id="2357b-755">String</span></span>||<span data-ttu-id="2357b-p137">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="2357b-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="2357b-758">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-758">Object</span></span>|<span data-ttu-id="2357b-759">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-759">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-760">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="2357b-760">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2357b-761">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-761">Object</span></span>|<span data-ttu-id="2357b-762">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-762">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-763">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-763">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="2357b-764">Boolean</span><span class="sxs-lookup"><span data-stu-id="2357b-764">Boolean</span></span>|<span data-ttu-id="2357b-765">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-765">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-766">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="2357b-766">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="2357b-767">function</span><span class="sxs-lookup"><span data-stu-id="2357b-767">function</span></span>|<span data-ttu-id="2357b-768">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-768">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-769">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-769">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2357b-770">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-770">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="2357b-771">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="2357b-771">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2357b-772">エラー</span><span class="sxs-lookup"><span data-stu-id="2357b-772">Errors</span></span>

|<span data-ttu-id="2357b-773">エラー コード</span><span class="sxs-lookup"><span data-stu-id="2357b-773">Error code</span></span>|<span data-ttu-id="2357b-774">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-774">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="2357b-775">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="2357b-775">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="2357b-776">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="2357b-776">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="2357b-777">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="2357b-777">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-778">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-778">Requirements</span></span>

|<span data-ttu-id="2357b-779">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-779">Requirement</span></span>|<span data-ttu-id="2357b-780">値</span><span class="sxs-lookup"><span data-stu-id="2357b-780">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-781">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-782">1.1</span><span class="sxs-lookup"><span data-stu-id="2357b-782">1.1</span></span>|
|[<span data-ttu-id="2357b-783">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-783">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-784">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2357b-784">ReadWriteItem</span></span>|
|[<span data-ttu-id="2357b-785">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-785">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-786">作成</span><span class="sxs-lookup"><span data-stu-id="2357b-786">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="2357b-787">例</span><span class="sxs-lookup"><span data-stu-id="2357b-787">Examples</span></span>

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

<span data-ttu-id="2357b-788">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="2357b-788">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

---
---

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="2357b-789">addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2357b-789">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="2357b-790">Base64 エンコードのファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="2357b-790">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="2357b-791">この`addFileAttachmentFromBase64Async`メソッドは、base64 エンコードからファイルをアップロードし、新規作成フォームのアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="2357b-791">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="2357b-792">このメソッドは、AsyncResult オブジェクトの添付ファイル識別子を返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-792">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="2357b-793">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-793">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-794">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-794">Parameters</span></span>

|<span data-ttu-id="2357b-795">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-795">Name</span></span>|<span data-ttu-id="2357b-796">種類</span><span class="sxs-lookup"><span data-stu-id="2357b-796">Type</span></span>|<span data-ttu-id="2357b-797">属性</span><span class="sxs-lookup"><span data-stu-id="2357b-797">Attributes</span></span>|<span data-ttu-id="2357b-798">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-798">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="2357b-799">String</span><span class="sxs-lookup"><span data-stu-id="2357b-799">String</span></span>||<span data-ttu-id="2357b-800">電子メールまたはイベントに追加する画像またはファイルの、base64 でエンコードされたコンテンツ。</span><span class="sxs-lookup"><span data-stu-id="2357b-800">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="2357b-801">String</span><span class="sxs-lookup"><span data-stu-id="2357b-801">String</span></span>||<span data-ttu-id="2357b-p139">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="2357b-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="2357b-804">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-804">Object</span></span>|<span data-ttu-id="2357b-805">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-805">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-806">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="2357b-806">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2357b-807">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-807">Object</span></span>|<span data-ttu-id="2357b-808">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-808">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-809">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-809">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="2357b-810">Boolean</span><span class="sxs-lookup"><span data-stu-id="2357b-810">Boolean</span></span>|<span data-ttu-id="2357b-811">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-811">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-812">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="2357b-812">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="2357b-813">function</span><span class="sxs-lookup"><span data-stu-id="2357b-813">function</span></span>|<span data-ttu-id="2357b-814">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-814">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-815">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-815">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2357b-816">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-816">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="2357b-817">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="2357b-817">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2357b-818">エラー</span><span class="sxs-lookup"><span data-stu-id="2357b-818">Errors</span></span>

|<span data-ttu-id="2357b-819">エラー コード</span><span class="sxs-lookup"><span data-stu-id="2357b-819">Error code</span></span>|<span data-ttu-id="2357b-820">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-820">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="2357b-821">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="2357b-821">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="2357b-822">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="2357b-822">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="2357b-823">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="2357b-823">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-824">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-824">Requirements</span></span>

|<span data-ttu-id="2357b-825">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-825">Requirement</span></span>|<span data-ttu-id="2357b-826">値</span><span class="sxs-lookup"><span data-stu-id="2357b-826">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-827">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-827">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-828">プレビュー</span><span class="sxs-lookup"><span data-stu-id="2357b-828">Preview</span></span>|
|[<span data-ttu-id="2357b-829">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-829">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-830">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2357b-830">ReadWriteItem</span></span>|
|[<span data-ttu-id="2357b-831">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-831">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-832">作成</span><span class="sxs-lookup"><span data-stu-id="2357b-832">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="2357b-833">例</span><span class="sxs-lookup"><span data-stu-id="2357b-833">Examples</span></span>

```javascript
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

---
---

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="2357b-834">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2357b-834">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="2357b-835">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="2357b-835">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="2357b-836">現在、サポートされて`Office.EventType.AttachmentsChanged`いる`Office.EventType.AppointmentTimeChanged`イベント`Office.EventType.EnhancedLocationsChanged`の`Office.EventType.RecipientsChanged`種類は`Office.EventType.RecurrenceChanged`、、、、、です。</span><span class="sxs-lookup"><span data-stu-id="2357b-836">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-837">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-837">Parameters</span></span>

| <span data-ttu-id="2357b-838">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-838">Name</span></span> | <span data-ttu-id="2357b-839">型</span><span class="sxs-lookup"><span data-stu-id="2357b-839">Type</span></span> | <span data-ttu-id="2357b-840">属性</span><span class="sxs-lookup"><span data-stu-id="2357b-840">Attributes</span></span> | <span data-ttu-id="2357b-841">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-841">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="2357b-842">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="2357b-842">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="2357b-843">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="2357b-843">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="2357b-844">Function</span><span class="sxs-lookup"><span data-stu-id="2357b-844">Function</span></span> || <span data-ttu-id="2357b-p140">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="2357b-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="2357b-848">Object</span><span class="sxs-lookup"><span data-stu-id="2357b-848">Object</span></span> | <span data-ttu-id="2357b-849">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-849">&lt;optional&gt;</span></span> | <span data-ttu-id="2357b-850">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="2357b-850">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="2357b-851">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-851">Object</span></span> | <span data-ttu-id="2357b-852">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-852">&lt;optional&gt;</span></span> | <span data-ttu-id="2357b-853">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-853">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="2357b-854">function</span><span class="sxs-lookup"><span data-stu-id="2357b-854">function</span></span>| <span data-ttu-id="2357b-855">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-855">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-856">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-856">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-857">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-857">Requirements</span></span>

|<span data-ttu-id="2357b-858">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-858">Requirement</span></span>| <span data-ttu-id="2357b-859">値</span><span class="sxs-lookup"><span data-stu-id="2357b-859">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-860">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-860">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2357b-861">1.7</span><span class="sxs-lookup"><span data-stu-id="2357b-861">1.7</span></span> |
|[<span data-ttu-id="2357b-862">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-862">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2357b-863">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-863">ReadItem</span></span> |
|[<span data-ttu-id="2357b-864">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-864">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2357b-865">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-865">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="2357b-866">例</span><span class="sxs-lookup"><span data-stu-id="2357b-866">Example</span></span>

```javascript
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="2357b-867">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2357b-867">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="2357b-868">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="2357b-868">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="2357b-p141">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="2357b-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="2357b-872">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-872">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="2357b-873">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="2357b-873">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-874">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-874">Parameters</span></span>

|<span data-ttu-id="2357b-875">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-875">Name</span></span>|<span data-ttu-id="2357b-876">型</span><span class="sxs-lookup"><span data-stu-id="2357b-876">Type</span></span>|<span data-ttu-id="2357b-877">属性</span><span class="sxs-lookup"><span data-stu-id="2357b-877">Attributes</span></span>|<span data-ttu-id="2357b-878">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-878">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="2357b-879">String</span><span class="sxs-lookup"><span data-stu-id="2357b-879">String</span></span>||<span data-ttu-id="2357b-p142">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="2357b-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="2357b-882">String</span><span class="sxs-lookup"><span data-stu-id="2357b-882">String</span></span>||<span data-ttu-id="2357b-883">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="2357b-883">The subject of the item to be attached.</span></span> <span data-ttu-id="2357b-884">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="2357b-884">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="2357b-885">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-885">Object</span></span>|<span data-ttu-id="2357b-886">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-886">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-887">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="2357b-887">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2357b-888">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-888">Object</span></span>|<span data-ttu-id="2357b-889">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-889">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-890">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-890">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2357b-891">function</span><span class="sxs-lookup"><span data-stu-id="2357b-891">function</span></span>|<span data-ttu-id="2357b-892">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-892">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-893">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-893">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2357b-894">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-894">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="2357b-895">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="2357b-895">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2357b-896">エラー</span><span class="sxs-lookup"><span data-stu-id="2357b-896">Errors</span></span>

|<span data-ttu-id="2357b-897">エラー コード</span><span class="sxs-lookup"><span data-stu-id="2357b-897">Error code</span></span>|<span data-ttu-id="2357b-898">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-898">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="2357b-899">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="2357b-899">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-900">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-900">Requirements</span></span>

|<span data-ttu-id="2357b-901">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-901">Requirement</span></span>|<span data-ttu-id="2357b-902">値</span><span class="sxs-lookup"><span data-stu-id="2357b-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-903">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-904">1.1</span><span class="sxs-lookup"><span data-stu-id="2357b-904">1.1</span></span>|
|[<span data-ttu-id="2357b-905">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-905">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2357b-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="2357b-907">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-907">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-908">作成</span><span class="sxs-lookup"><span data-stu-id="2357b-908">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-909">例</span><span class="sxs-lookup"><span data-stu-id="2357b-909">Example</span></span>

<span data-ttu-id="2357b-910">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-910">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

---
---

#### <a name="close"></a><span data-ttu-id="2357b-911">close()</span><span class="sxs-lookup"><span data-stu-id="2357b-911">close()</span></span>

<span data-ttu-id="2357b-912">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="2357b-912">Closes the current item that is being composed.</span></span>

<span data-ttu-id="2357b-p144">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="2357b-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-915">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-915">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="2357b-916">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="2357b-916">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-917">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-917">Requirements</span></span>

|<span data-ttu-id="2357b-918">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-918">Requirement</span></span>|<span data-ttu-id="2357b-919">値</span><span class="sxs-lookup"><span data-stu-id="2357b-919">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-920">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-920">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-921">1.3</span><span class="sxs-lookup"><span data-stu-id="2357b-921">1.3</span></span>|
|[<span data-ttu-id="2357b-922">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-922">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-923">制限あり</span><span class="sxs-lookup"><span data-stu-id="2357b-923">Restricted</span></span>|
|[<span data-ttu-id="2357b-924">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-924">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-925">新規作成</span><span class="sxs-lookup"><span data-stu-id="2357b-925">Compose</span></span>|

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="2357b-926">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="2357b-926">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="2357b-927">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-927">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-928">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2357b-928">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2357b-929">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-929">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="2357b-930">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="2357b-930">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="2357b-p145">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="2357b-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-934">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-934">Parameters</span></span>

|<span data-ttu-id="2357b-935">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-935">Name</span></span>|<span data-ttu-id="2357b-936">型</span><span class="sxs-lookup"><span data-stu-id="2357b-936">Type</span></span>|<span data-ttu-id="2357b-937">属性</span><span class="sxs-lookup"><span data-stu-id="2357b-937">Attributes</span></span>|<span data-ttu-id="2357b-938">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-938">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="2357b-939">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="2357b-939">String &#124; Object</span></span>||<span data-ttu-id="2357b-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="2357b-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="2357b-942">**または**</span><span class="sxs-lookup"><span data-stu-id="2357b-942">**OR**</span></span><br/><span data-ttu-id="2357b-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="2357b-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="2357b-945">String</span><span class="sxs-lookup"><span data-stu-id="2357b-945">String</span></span>|<span data-ttu-id="2357b-946">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-946">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="2357b-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="2357b-949">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-949">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="2357b-950">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-950">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-951">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="2357b-951">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="2357b-952">String</span><span class="sxs-lookup"><span data-stu-id="2357b-952">String</span></span>||<span data-ttu-id="2357b-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="2357b-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="2357b-955">String</span><span class="sxs-lookup"><span data-stu-id="2357b-955">String</span></span>||<span data-ttu-id="2357b-956">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="2357b-956">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="2357b-957">文字列</span><span class="sxs-lookup"><span data-stu-id="2357b-957">String</span></span>||<span data-ttu-id="2357b-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="2357b-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="2357b-960">ブール値</span><span class="sxs-lookup"><span data-stu-id="2357b-960">Boolean</span></span>||<span data-ttu-id="2357b-p151">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="2357b-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="2357b-963">String</span><span class="sxs-lookup"><span data-stu-id="2357b-963">String</span></span>||<span data-ttu-id="2357b-p152">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="2357b-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="2357b-967">function</span><span class="sxs-lookup"><span data-stu-id="2357b-967">function</span></span>|<span data-ttu-id="2357b-968">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-968">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-969">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-969">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-970">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-970">Requirements</span></span>

|<span data-ttu-id="2357b-971">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-971">Requirement</span></span>|<span data-ttu-id="2357b-972">値</span><span class="sxs-lookup"><span data-stu-id="2357b-972">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-973">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-973">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-974">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-974">1.0</span></span>|
|[<span data-ttu-id="2357b-975">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-975">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-976">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-976">ReadItem</span></span>|
|[<span data-ttu-id="2357b-977">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-977">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-978">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-978">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="2357b-979">例</span><span class="sxs-lookup"><span data-stu-id="2357b-979">Examples</span></span>

<span data-ttu-id="2357b-980">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="2357b-980">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="2357b-981">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="2357b-981">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="2357b-982">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="2357b-982">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="2357b-983">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="2357b-983">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="2357b-984">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="2357b-984">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="2357b-985">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="2357b-985">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="2357b-986">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="2357b-986">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="2357b-987">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-987">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-988">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2357b-988">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2357b-989">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-989">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="2357b-990">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="2357b-990">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="2357b-p153">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="2357b-p153">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-994">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-994">Parameters</span></span>

|<span data-ttu-id="2357b-995">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-995">Name</span></span>|<span data-ttu-id="2357b-996">型</span><span class="sxs-lookup"><span data-stu-id="2357b-996">Type</span></span>|<span data-ttu-id="2357b-997">属性</span><span class="sxs-lookup"><span data-stu-id="2357b-997">Attributes</span></span>|<span data-ttu-id="2357b-998">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-998">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="2357b-999">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="2357b-999">String &#124; Object</span></span>||<span data-ttu-id="2357b-p154">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="2357b-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="2357b-1002">**または**</span><span class="sxs-lookup"><span data-stu-id="2357b-1002">**OR**</span></span><br/><span data-ttu-id="2357b-p155">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="2357b-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="2357b-1005">String</span><span class="sxs-lookup"><span data-stu-id="2357b-1005">String</span></span>|<span data-ttu-id="2357b-1006">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1006">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-p156">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="2357b-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="2357b-1009">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1009">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="2357b-1010">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1010">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1011">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="2357b-1011">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="2357b-1012">String</span><span class="sxs-lookup"><span data-stu-id="2357b-1012">String</span></span>||<span data-ttu-id="2357b-p157">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="2357b-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="2357b-1015">String</span><span class="sxs-lookup"><span data-stu-id="2357b-1015">String</span></span>||<span data-ttu-id="2357b-1016">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="2357b-1016">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="2357b-1017">文字列</span><span class="sxs-lookup"><span data-stu-id="2357b-1017">String</span></span>||<span data-ttu-id="2357b-p158">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="2357b-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="2357b-1020">ブール値</span><span class="sxs-lookup"><span data-stu-id="2357b-1020">Boolean</span></span>||<span data-ttu-id="2357b-p159">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="2357b-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="2357b-1023">String</span><span class="sxs-lookup"><span data-stu-id="2357b-1023">String</span></span>||<span data-ttu-id="2357b-p160">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="2357b-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="2357b-1027">function</span><span class="sxs-lookup"><span data-stu-id="2357b-1027">function</span></span>|<span data-ttu-id="2357b-1028">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1028">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1029">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1029">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-1030">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1030">Requirements</span></span>

|<span data-ttu-id="2357b-1031">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1031">Requirement</span></span>|<span data-ttu-id="2357b-1032">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1032">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1033">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1033">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-1034">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-1034">1.0</span></span>|
|[<span data-ttu-id="2357b-1035">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1035">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-1036">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-1036">ReadItem</span></span>|
|[<span data-ttu-id="2357b-1037">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1037">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-1038">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-1038">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="2357b-1039">例</span><span class="sxs-lookup"><span data-stu-id="2357b-1039">Examples</span></span>

<span data-ttu-id="2357b-1040">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1040">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="2357b-1041">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1041">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="2357b-1042">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1042">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="2357b-1043">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1043">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="2357b-1044">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1044">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="2357b-1045">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1045">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

---
---

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="2357b-1046">getAttachmentContentAsync (attachmentId, [options], [callback]) > [Attachmentcontent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="2357b-1046">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="2357b-1047">メッセージまたは予定から指定された添付ファイルを取得し`AttachmentContent` 、それをオブジェクトとして返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1047">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="2357b-1048">メソッド`getAttachmentContentAsync`は、指定された id の添付ファイルをアイテムから取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1048">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="2357b-1049">ベストプラクティスとして、識別子を使用して、または`getAttachmentsAsync` `item.attachments`の呼び出しで attachmentIds を取得したのと同じセッションの添付ファイルを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2357b-1049">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="2357b-1050">Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="2357b-1050">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="2357b-1051">ユーザーがアプリを閉じたとき、またはインラインフォームの作成が開始されたときに、別のウィンドウで続行するためにフォームをポップアウトした後、セッションが終了します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1051">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-1052">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-1052">Parameters</span></span>

|<span data-ttu-id="2357b-1053">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-1053">Name</span></span>|<span data-ttu-id="2357b-1054">型</span><span class="sxs-lookup"><span data-stu-id="2357b-1054">Type</span></span>|<span data-ttu-id="2357b-1055">属性</span><span class="sxs-lookup"><span data-stu-id="2357b-1055">Attributes</span></span>|<span data-ttu-id="2357b-1056">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-1056">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="2357b-1057">String</span><span class="sxs-lookup"><span data-stu-id="2357b-1057">String</span></span>||<span data-ttu-id="2357b-1058">取得する添付ファイルの識別子を指定します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1058">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="2357b-1059">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1059">Object</span></span>|<span data-ttu-id="2357b-1060">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1060">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1061">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="2357b-1061">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2357b-1062">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1062">Object</span></span>|<span data-ttu-id="2357b-1063">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1064">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1064">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2357b-1065">function</span><span class="sxs-lookup"><span data-stu-id="2357b-1065">function</span></span>|<span data-ttu-id="2357b-1066">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1067">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1067">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-1068">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1068">Requirements</span></span>

|<span data-ttu-id="2357b-1069">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1069">Requirement</span></span>|<span data-ttu-id="2357b-1070">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1070">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1071">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1071">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-1072">プレビュー</span><span class="sxs-lookup"><span data-stu-id="2357b-1072">Preview</span></span>|
|[<span data-ttu-id="2357b-1073">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1073">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-1074">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-1074">ReadItem</span></span>|
|[<span data-ttu-id="2357b-1075">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1075">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-1076">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-1076">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2357b-1077">戻り値:</span><span class="sxs-lookup"><span data-stu-id="2357b-1077">Returns:</span></span>

<span data-ttu-id="2357b-1078">型: [Attachmentcontent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="2357b-1078">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="2357b-1079">例</span><span class="sxs-lookup"><span data-stu-id="2357b-1079">Example</span></span>

```javascript
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
  if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
    // Handle file attachment.
  } else if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Eml) {
    // Handle email item attachment.
  } else if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
    // Handle .icalender attachment.
  } else if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Url) {
    // Handle cloud attachment.
  } else {
    // Handle attachment formats that are not supported.
  }
}
```

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="2357b-1080">getAttachmentsAsync ([オプション], [callback]) > Array. <[Attachmentdetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="2357b-1080">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="2357b-1081">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1081">Gets the item's attachments as an array.</span></span> <span data-ttu-id="2357b-1082">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="2357b-1082">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-1083">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-1083">Parameters</span></span>

|<span data-ttu-id="2357b-1084">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-1084">Name</span></span>|<span data-ttu-id="2357b-1085">型</span><span class="sxs-lookup"><span data-stu-id="2357b-1085">Type</span></span>|<span data-ttu-id="2357b-1086">属性</span><span class="sxs-lookup"><span data-stu-id="2357b-1086">Attributes</span></span>|<span data-ttu-id="2357b-1087">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-1087">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="2357b-1088">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1088">Object</span></span>|<span data-ttu-id="2357b-1089">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1089">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1090">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="2357b-1090">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2357b-1091">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1091">Object</span></span>|<span data-ttu-id="2357b-1092">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1092">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1093">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1093">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2357b-1094">function</span><span class="sxs-lookup"><span data-stu-id="2357b-1094">function</span></span>|<span data-ttu-id="2357b-1095">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1096">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1096">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-1097">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1097">Requirements</span></span>

|<span data-ttu-id="2357b-1098">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1098">Requirement</span></span>|<span data-ttu-id="2357b-1099">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1099">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1100">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1100">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-1101">プレビュー</span><span class="sxs-lookup"><span data-stu-id="2357b-1101">Preview</span></span>|
|[<span data-ttu-id="2357b-1102">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1102">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-1103">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-1103">ReadItem</span></span>|
|[<span data-ttu-id="2357b-1104">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1104">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-1105">作成</span><span class="sxs-lookup"><span data-stu-id="2357b-1105">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="2357b-1106">戻り値:</span><span class="sxs-lookup"><span data-stu-id="2357b-1106">Returns:</span></span>

<span data-ttu-id="2357b-1107">型: Array. <[attachmentdetails 詳細](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="2357b-1107">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="2357b-1108">例</span><span class="sxs-lookup"><span data-stu-id="2357b-1108">Example</span></span>

<span data-ttu-id="2357b-1109">次の例では、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1109">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
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

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="2357b-1110">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="2357b-1110">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="2357b-1111">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1111">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-1112">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2357b-1112">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-1113">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1113">Requirements</span></span>

|<span data-ttu-id="2357b-1114">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1114">Requirement</span></span>|<span data-ttu-id="2357b-1115">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1115">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1116">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1116">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-1117">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-1117">1.0</span></span>|
|[<span data-ttu-id="2357b-1118">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1118">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-1119">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-1119">ReadItem</span></span>|
|[<span data-ttu-id="2357b-1120">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1120">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-1121">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-1121">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2357b-1122">戻り値:</span><span class="sxs-lookup"><span data-stu-id="2357b-1122">Returns:</span></span>

<span data-ttu-id="2357b-1123">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="2357b-1123">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="2357b-1124">例</span><span class="sxs-lookup"><span data-stu-id="2357b-1124">Example</span></span>

<span data-ttu-id="2357b-1125">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="2357b-1125">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="2357b-1126">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="2357b-1126">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="2357b-1127">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1127">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-1128">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2357b-1128">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-1129">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-1129">Parameters</span></span>

|<span data-ttu-id="2357b-1130">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-1130">Name</span></span>|<span data-ttu-id="2357b-1131">型</span><span class="sxs-lookup"><span data-stu-id="2357b-1131">Type</span></span>|<span data-ttu-id="2357b-1132">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-1132">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="2357b-1133">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="2357b-1133">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="2357b-1134">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="2357b-1134">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-1135">Requirements</span><span class="sxs-lookup"><span data-stu-id="2357b-1135">Requirements</span></span>

|<span data-ttu-id="2357b-1136">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1136">Requirement</span></span>|<span data-ttu-id="2357b-1137">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1138">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-1139">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-1139">1.0</span></span>|
|[<span data-ttu-id="2357b-1140">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-1141">制限あり</span><span class="sxs-lookup"><span data-stu-id="2357b-1141">Restricted</span></span>|
|[<span data-ttu-id="2357b-1142">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-1143">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2357b-1144">戻り値:</span><span class="sxs-lookup"><span data-stu-id="2357b-1144">Returns:</span></span>

<span data-ttu-id="2357b-1145">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1145">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="2357b-1146">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1146">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="2357b-1147">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="2357b-1147">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="2357b-1148">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="2357b-1148">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="2357b-1149">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="2357b-1149">Value of `entityType`</span></span>|<span data-ttu-id="2357b-1150">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="2357b-1150">Type of objects in returned array</span></span>|<span data-ttu-id="2357b-1151">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1151">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="2357b-1152">String</span><span class="sxs-lookup"><span data-stu-id="2357b-1152">String</span></span>|<span data-ttu-id="2357b-1153">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="2357b-1153">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="2357b-1154">連絡先</span><span class="sxs-lookup"><span data-stu-id="2357b-1154">Contact</span></span>|<span data-ttu-id="2357b-1155">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2357b-1155">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="2357b-1156">文字列</span><span class="sxs-lookup"><span data-stu-id="2357b-1156">String</span></span>|<span data-ttu-id="2357b-1157">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2357b-1157">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="2357b-1158">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="2357b-1158">MeetingSuggestion</span></span>|<span data-ttu-id="2357b-1159">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2357b-1159">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="2357b-1160">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="2357b-1160">PhoneNumber</span></span>|<span data-ttu-id="2357b-1161">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="2357b-1161">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="2357b-1162">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="2357b-1162">TaskSuggestion</span></span>|<span data-ttu-id="2357b-1163">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2357b-1163">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="2357b-1164">文字列</span><span class="sxs-lookup"><span data-stu-id="2357b-1164">String</span></span>|<span data-ttu-id="2357b-1165">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="2357b-1165">**Restricted**</span></span>|

<span data-ttu-id="2357b-1166">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="2357b-1166">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="2357b-1167">例</span><span class="sxs-lookup"><span data-stu-id="2357b-1167">Example</span></span>

<span data-ttu-id="2357b-1168">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1168">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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
};
```

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="2357b-1169">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="2357b-1169">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="2357b-1170">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1170">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-1171">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2357b-1171">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2357b-1172">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1172">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-1173">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-1173">Parameters</span></span>

|<span data-ttu-id="2357b-1174">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-1174">Name</span></span>|<span data-ttu-id="2357b-1175">型</span><span class="sxs-lookup"><span data-stu-id="2357b-1175">Type</span></span>|<span data-ttu-id="2357b-1176">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-1176">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="2357b-1177">String</span><span class="sxs-lookup"><span data-stu-id="2357b-1177">String</span></span>|<span data-ttu-id="2357b-1178">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="2357b-1178">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-1179">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1179">Requirements</span></span>

|<span data-ttu-id="2357b-1180">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1180">Requirement</span></span>|<span data-ttu-id="2357b-1181">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1181">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-1183">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-1183">1.0</span></span>|
|[<span data-ttu-id="2357b-1184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-1185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-1185">ReadItem</span></span>|
|[<span data-ttu-id="2357b-1186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-1187">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-1187">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2357b-1188">戻り値:</span><span class="sxs-lookup"><span data-stu-id="2357b-1188">Returns:</span></span>

<span data-ttu-id="2357b-p164">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="2357b-1191">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="2357b-1191">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="2357b-1192">、Office.context.mailbox.item.getinitializationcontextasync ([オプション], [callback])</span><span class="sxs-lookup"><span data-stu-id="2357b-1192">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="2357b-1193">[アクション可能なメッセージによってアドインがアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されたときに渡される初期化データを取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1193">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-1194">このメソッドは、Outlook 2016 以降の Windows (16.0.8413.1000 より後のバージョン) および Outlook on the Office 365 でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="2357b-1194">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-1195">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-1195">Parameters</span></span>

|<span data-ttu-id="2357b-1196">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-1196">Name</span></span>|<span data-ttu-id="2357b-1197">型</span><span class="sxs-lookup"><span data-stu-id="2357b-1197">Type</span></span>|<span data-ttu-id="2357b-1198">属性</span><span class="sxs-lookup"><span data-stu-id="2357b-1198">Attributes</span></span>|<span data-ttu-id="2357b-1199">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-1199">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="2357b-1200">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1200">Object</span></span>|<span data-ttu-id="2357b-1201">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1201">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1202">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="2357b-1202">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2357b-1203">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1203">Object</span></span>|<span data-ttu-id="2357b-1204">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1204">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1205">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1205">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2357b-1206">function</span><span class="sxs-lookup"><span data-stu-id="2357b-1206">function</span></span>|<span data-ttu-id="2357b-1207">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1207">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1208">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1208">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2357b-1209">成功すると、初期化データが文字列とし`asyncResult.value`てプロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1209">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="2357b-1210">初期化コンテキストがない場合、 `asyncResult`オブジェクトには、 `Error` `code`プロパティがに`9020`設定されたオブジェクトと`name`プロパティがに`GenericResponseError`設定されたオブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1210">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-1211">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1211">Requirements</span></span>

|<span data-ttu-id="2357b-1212">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1212">Requirement</span></span>|<span data-ttu-id="2357b-1213">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1213">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1214">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-1215">プレビュー</span><span class="sxs-lookup"><span data-stu-id="2357b-1215">Preview</span></span>|
|[<span data-ttu-id="2357b-1216">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1216">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-1217">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-1217">ReadItem</span></span>|
|[<span data-ttu-id="2357b-1218">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1218">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-1219">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-1219">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-1220">例</span><span class="sxs-lookup"><span data-stu-id="2357b-1220">Example</span></span>

```javascript
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

---
---

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="2357b-1221">getItemIdAsync ([オプション], callback)</span><span class="sxs-lookup"><span data-stu-id="2357b-1221">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="2357b-1222">保存されたアイテムの ID を非同期に取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1222">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="2357b-1223">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="2357b-1223">Compose mode only.</span></span>

<span data-ttu-id="2357b-1224">このメソッドを呼び出すと、コールバックメソッドによってアイテム ID が返されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1224">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-1225">アドインが新規作成モードの`getItemIdAsync`アイテムに対して呼び出しを行う場合 ( `itemId` EWS または REST API を使用するため)、Outlook がキャッシュモードの場合は、アイテムがサーバーに同期されるまでしばらく時間がかかる場合があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="2357b-1225">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="2357b-1226">アイテムが同期されるまで、 `itemId`は認識されず、を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1226">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-1227">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-1227">Parameters</span></span>

|<span data-ttu-id="2357b-1228">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-1228">Name</span></span>|<span data-ttu-id="2357b-1229">型</span><span class="sxs-lookup"><span data-stu-id="2357b-1229">Type</span></span>|<span data-ttu-id="2357b-1230">属性</span><span class="sxs-lookup"><span data-stu-id="2357b-1230">Attributes</span></span>|<span data-ttu-id="2357b-1231">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-1231">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="2357b-1232">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1232">Object</span></span>|<span data-ttu-id="2357b-1233">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1233">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1234">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="2357b-1234">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2357b-1235">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1235">Object</span></span>|<span data-ttu-id="2357b-1236">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1236">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1237">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1237">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2357b-1238">function</span><span class="sxs-lookup"><span data-stu-id="2357b-1238">function</span></span>||<span data-ttu-id="2357b-1239">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1239">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2357b-1240">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1240">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2357b-1241">エラー</span><span class="sxs-lookup"><span data-stu-id="2357b-1241">Errors</span></span>

|<span data-ttu-id="2357b-1242">エラー コード</span><span class="sxs-lookup"><span data-stu-id="2357b-1242">Error code</span></span>|<span data-ttu-id="2357b-1243">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-1243">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="2357b-1244">この id は、アイテムが保存されるまでは取得できません。</span><span class="sxs-lookup"><span data-stu-id="2357b-1244">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-1245">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1245">Requirements</span></span>

|<span data-ttu-id="2357b-1246">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1246">Requirement</span></span>|<span data-ttu-id="2357b-1247">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1247">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1248">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1248">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-1249">プレビュー</span><span class="sxs-lookup"><span data-stu-id="2357b-1249">Preview</span></span>|
|[<span data-ttu-id="2357b-1250">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1250">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-1251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-1251">ReadItem</span></span>|
|[<span data-ttu-id="2357b-1252">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1252">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-1253">作成</span><span class="sxs-lookup"><span data-stu-id="2357b-1253">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="2357b-1254">例</span><span class="sxs-lookup"><span data-stu-id="2357b-1254">Examples</span></span>

```javascript
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="2357b-1255">次の例は、コールバック関数`result`に渡されるパラメーターの構造を示しています。</span><span class="sxs-lookup"><span data-stu-id="2357b-1255">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="2357b-1256">プロパティ`value`には、アイテムの ID が含まれています。</span><span class="sxs-lookup"><span data-stu-id="2357b-1256">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="2357b-1257">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="2357b-1257">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="2357b-1258">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1258">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-1259">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2357b-1259">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2357b-p168">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="2357b-p168">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="2357b-1263">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="2357b-1263">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="2357b-1264">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="2357b-1264">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="2357b-p169">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-1268">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1268">Requirements</span></span>

|<span data-ttu-id="2357b-1269">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1269">Requirement</span></span>|<span data-ttu-id="2357b-1270">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1271">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-1272">1.0</span></span>|
|[<span data-ttu-id="2357b-1273">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-1274">ReadItem</span></span>|
|[<span data-ttu-id="2357b-1275">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-1276">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2357b-1277">戻り値:</span><span class="sxs-lookup"><span data-stu-id="2357b-1277">Returns:</span></span>

<span data-ttu-id="2357b-p170">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="2357b-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="2357b-1280">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="2357b-1280">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2357b-1281">Object</span><span class="sxs-lookup"><span data-stu-id="2357b-1281">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="2357b-1282">例</span><span class="sxs-lookup"><span data-stu-id="2357b-1282">Example</span></span>

<span data-ttu-id="2357b-1283">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="2357b-1283">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="2357b-1284">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="2357b-1284">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="2357b-1285">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1285">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-1286">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2357b-1286">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2357b-1287">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="2357b-1287">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="2357b-p171">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="2357b-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-1290">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-1290">Parameters</span></span>

|<span data-ttu-id="2357b-1291">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-1291">Name</span></span>|<span data-ttu-id="2357b-1292">型</span><span class="sxs-lookup"><span data-stu-id="2357b-1292">Type</span></span>|<span data-ttu-id="2357b-1293">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-1293">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="2357b-1294">String</span><span class="sxs-lookup"><span data-stu-id="2357b-1294">String</span></span>|<span data-ttu-id="2357b-1295">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="2357b-1295">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-1296">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1296">Requirements</span></span>

|<span data-ttu-id="2357b-1297">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1297">Requirement</span></span>|<span data-ttu-id="2357b-1298">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1298">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1299">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1299">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-1300">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-1300">1.0</span></span>|
|[<span data-ttu-id="2357b-1301">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1301">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-1302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-1302">ReadItem</span></span>|
|[<span data-ttu-id="2357b-1303">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1303">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-1304">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-1304">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2357b-1305">戻り値:</span><span class="sxs-lookup"><span data-stu-id="2357b-1305">Returns:</span></span>

<span data-ttu-id="2357b-1306">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="2357b-1306">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="2357b-1307">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="2357b-1307">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2357b-1308">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="2357b-1308">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="2357b-1309">例</span><span class="sxs-lookup"><span data-stu-id="2357b-1309">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="2357b-1310">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="2357b-1310">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="2357b-1311">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1311">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="2357b-p172">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-p172">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-1314">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-1314">Parameters</span></span>

|<span data-ttu-id="2357b-1315">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-1315">Name</span></span>|<span data-ttu-id="2357b-1316">型</span><span class="sxs-lookup"><span data-stu-id="2357b-1316">Type</span></span>|<span data-ttu-id="2357b-1317">属性</span><span class="sxs-lookup"><span data-stu-id="2357b-1317">Attributes</span></span>|<span data-ttu-id="2357b-1318">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-1318">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="2357b-1319">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="2357b-1319">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="2357b-p173">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="2357b-p173">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="2357b-1323">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1323">Object</span></span>|<span data-ttu-id="2357b-1324">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1324">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1325">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="2357b-1325">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2357b-1326">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1326">Object</span></span>|<span data-ttu-id="2357b-1327">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1327">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1328">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1328">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2357b-1329">function</span><span class="sxs-lookup"><span data-stu-id="2357b-1329">function</span></span>||<span data-ttu-id="2357b-1330">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1330">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2357b-1331">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1331">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="2357b-1332">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="2357b-1332">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-1333">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1333">Requirements</span></span>

|<span data-ttu-id="2357b-1334">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1334">Requirement</span></span>|<span data-ttu-id="2357b-1335">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1335">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1336">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-1337">1.2</span><span class="sxs-lookup"><span data-stu-id="2357b-1337">1.2</span></span>|
|[<span data-ttu-id="2357b-1338">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1338">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-1339">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2357b-1339">ReadWriteItem</span></span>|
|[<span data-ttu-id="2357b-1340">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1340">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-1341">作成</span><span class="sxs-lookup"><span data-stu-id="2357b-1341">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="2357b-1342">戻り値:</span><span class="sxs-lookup"><span data-stu-id="2357b-1342">Returns:</span></span>

<span data-ttu-id="2357b-1343">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="2357b-1343">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="2357b-1344">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="2357b-1344">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2357b-1345">String</span><span class="sxs-lookup"><span data-stu-id="2357b-1345">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="2357b-1346">例</span><span class="sxs-lookup"><span data-stu-id="2357b-1346">Example</span></span>

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

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="2357b-1347">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="2357b-1347">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="2357b-1348">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1348">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="2357b-1349">強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1349">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-1350">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2357b-1350">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-1351">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1351">Requirements</span></span>

|<span data-ttu-id="2357b-1352">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1352">Requirement</span></span>|<span data-ttu-id="2357b-1353">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1353">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1354">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1354">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-1355">1.6</span><span class="sxs-lookup"><span data-stu-id="2357b-1355">1.6</span></span>|
|[<span data-ttu-id="2357b-1356">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1356">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-1357">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-1357">ReadItem</span></span>|
|[<span data-ttu-id="2357b-1358">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1358">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-1359">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-1359">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2357b-1360">戻り値:</span><span class="sxs-lookup"><span data-stu-id="2357b-1360">Returns:</span></span>

<span data-ttu-id="2357b-1361">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="2357b-1361">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="2357b-1362">例</span><span class="sxs-lookup"><span data-stu-id="2357b-1362">Example</span></span>

<span data-ttu-id="2357b-1363">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="2357b-1363">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="2357b-1364">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="2357b-1364">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="2357b-p176">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-p176">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-1367">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2357b-1367">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2357b-p177">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="2357b-p177">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="2357b-1371">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="2357b-1371">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="2357b-1372">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="2357b-1372">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="2357b-p178">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-p178">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2357b-1376">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1376">Requirements</span></span>

|<span data-ttu-id="2357b-1377">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1377">Requirement</span></span>|<span data-ttu-id="2357b-1378">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1378">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1379">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1379">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-1380">1.6</span><span class="sxs-lookup"><span data-stu-id="2357b-1380">1.6</span></span>|
|[<span data-ttu-id="2357b-1381">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1381">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-1382">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-1382">ReadItem</span></span>|
|[<span data-ttu-id="2357b-1383">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1383">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-1384">読み取り</span><span class="sxs-lookup"><span data-stu-id="2357b-1384">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2357b-1385">戻り値:</span><span class="sxs-lookup"><span data-stu-id="2357b-1385">Returns:</span></span>

<span data-ttu-id="2357b-p179">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="2357b-p179">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="2357b-1388">例</span><span class="sxs-lookup"><span data-stu-id="2357b-1388">Example</span></span>

<span data-ttu-id="2357b-1389">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="2357b-1389">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="2357b-1390">getSharedPropertiesAsync ([options], callback)</span><span class="sxs-lookup"><span data-stu-id="2357b-1390">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="2357b-1391">共有フォルダー、予定表、またはメールボックス内の選択した予定またはメッセージのプロパティを取得します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1391">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-1392">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-1392">Parameters</span></span>

|<span data-ttu-id="2357b-1393">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-1393">Name</span></span>|<span data-ttu-id="2357b-1394">型</span><span class="sxs-lookup"><span data-stu-id="2357b-1394">Type</span></span>|<span data-ttu-id="2357b-1395">属性</span><span class="sxs-lookup"><span data-stu-id="2357b-1395">Attributes</span></span>|<span data-ttu-id="2357b-1396">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-1396">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="2357b-1397">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1397">Object</span></span>|<span data-ttu-id="2357b-1398">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1398">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1399">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="2357b-1399">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2357b-1400">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1400">Object</span></span>|<span data-ttu-id="2357b-1401">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1401">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1402">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1402">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2357b-1403">function</span><span class="sxs-lookup"><span data-stu-id="2357b-1403">function</span></span>||<span data-ttu-id="2357b-1404">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1404">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2357b-1405">共有プロパティは、 [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) `asyncResult.value`プロパティのオブジェクトとして提供されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1405">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="2357b-1406">このオブジェクトは、アイテムの共有プロパティを取得するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1406">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-1407">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1407">Requirements</span></span>

|<span data-ttu-id="2357b-1408">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1408">Requirement</span></span>|<span data-ttu-id="2357b-1409">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1409">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1410">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-1411">プレビュー</span><span class="sxs-lookup"><span data-stu-id="2357b-1411">Preview</span></span>|
|[<span data-ttu-id="2357b-1412">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1412">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-1413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-1413">ReadItem</span></span>|
|[<span data-ttu-id="2357b-1414">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1414">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-1415">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-1415">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-1416">例</span><span class="sxs-lookup"><span data-stu-id="2357b-1416">Example</span></span>

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="2357b-1417">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2357b-1417">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="2357b-1418">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1418">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="2357b-p181">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="2357b-p181">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-1422">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-1422">Parameters</span></span>

|<span data-ttu-id="2357b-1423">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-1423">Name</span></span>|<span data-ttu-id="2357b-1424">型</span><span class="sxs-lookup"><span data-stu-id="2357b-1424">Type</span></span>|<span data-ttu-id="2357b-1425">属性</span><span class="sxs-lookup"><span data-stu-id="2357b-1425">Attributes</span></span>|<span data-ttu-id="2357b-1426">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-1426">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="2357b-1427">function</span><span class="sxs-lookup"><span data-stu-id="2357b-1427">function</span></span>||<span data-ttu-id="2357b-1428">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1428">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2357b-1429">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1429">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="2357b-1430">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1430">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="2357b-1431">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1431">Object</span></span>|<span data-ttu-id="2357b-1432">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1432">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1433">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1433">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="2357b-1434">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1434">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-1435">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1435">Requirements</span></span>

|<span data-ttu-id="2357b-1436">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1436">Requirement</span></span>|<span data-ttu-id="2357b-1437">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1437">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1438">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1438">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-1439">1.0</span><span class="sxs-lookup"><span data-stu-id="2357b-1439">1.0</span></span>|
|[<span data-ttu-id="2357b-1440">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1440">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-1441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-1441">ReadItem</span></span>|
|[<span data-ttu-id="2357b-1442">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1442">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-1443">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-1443">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-1444">例</span><span class="sxs-lookup"><span data-stu-id="2357b-1444">Example</span></span>

<span data-ttu-id="2357b-p184">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="2357b-p184">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="2357b-1448">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2357b-1448">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="2357b-1449">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1449">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="2357b-1450">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1450">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="2357b-1451">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="2357b-1451">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="2357b-1452">Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="2357b-1452">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="2357b-1453">ユーザーがアプリを閉じたとき、またはインラインフォームの作成が開始されたときに、別のウィンドウで続行するためにフォームをポップアウトした後、セッションが終了します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1453">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-1454">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-1454">Parameters</span></span>

|<span data-ttu-id="2357b-1455">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-1455">Name</span></span>|<span data-ttu-id="2357b-1456">型</span><span class="sxs-lookup"><span data-stu-id="2357b-1456">Type</span></span>|<span data-ttu-id="2357b-1457">属性</span><span class="sxs-lookup"><span data-stu-id="2357b-1457">Attributes</span></span>|<span data-ttu-id="2357b-1458">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-1458">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="2357b-1459">String</span><span class="sxs-lookup"><span data-stu-id="2357b-1459">String</span></span>||<span data-ttu-id="2357b-1460">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="2357b-1460">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="2357b-1461">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1461">Object</span></span>|<span data-ttu-id="2357b-1462">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1462">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1463">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="2357b-1463">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2357b-1464">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1464">Object</span></span>|<span data-ttu-id="2357b-1465">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1465">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1466">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1466">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2357b-1467">function</span><span class="sxs-lookup"><span data-stu-id="2357b-1467">function</span></span>|<span data-ttu-id="2357b-1468">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1468">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1469">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1469">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2357b-1470">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1470">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2357b-1471">エラー</span><span class="sxs-lookup"><span data-stu-id="2357b-1471">Errors</span></span>

|<span data-ttu-id="2357b-1472">エラー コード</span><span class="sxs-lookup"><span data-stu-id="2357b-1472">Error code</span></span>|<span data-ttu-id="2357b-1473">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-1473">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="2357b-1474">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="2357b-1474">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-1475">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1475">Requirements</span></span>

|<span data-ttu-id="2357b-1476">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1476">Requirement</span></span>|<span data-ttu-id="2357b-1477">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1477">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1478">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1478">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-1479">1.1</span><span class="sxs-lookup"><span data-stu-id="2357b-1479">1.1</span></span>|
|[<span data-ttu-id="2357b-1480">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1480">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-1481">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2357b-1481">ReadWriteItem</span></span>|
|[<span data-ttu-id="2357b-1482">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1482">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-1483">作成</span><span class="sxs-lookup"><span data-stu-id="2357b-1483">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-1484">例</span><span class="sxs-lookup"><span data-stu-id="2357b-1484">Example</span></span>

<span data-ttu-id="2357b-1485">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1485">The following code removes an attachment with an identifier of '0'.</span></span>

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

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="2357b-1486">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2357b-1486">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="2357b-1487">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1487">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="2357b-1488">現在、サポートされて`Office.EventType.AttachmentsChanged`いる`Office.EventType.AppointmentTimeChanged`イベント`Office.EventType.EnhancedLocationsChanged`の`Office.EventType.RecipientsChanged`種類は`Office.EventType.RecurrenceChanged`、、、、、です。</span><span class="sxs-lookup"><span data-stu-id="2357b-1488">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-1489">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-1489">Parameters</span></span>

| <span data-ttu-id="2357b-1490">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-1490">Name</span></span> | <span data-ttu-id="2357b-1491">型</span><span class="sxs-lookup"><span data-stu-id="2357b-1491">Type</span></span> | <span data-ttu-id="2357b-1492">属性</span><span class="sxs-lookup"><span data-stu-id="2357b-1492">Attributes</span></span> | <span data-ttu-id="2357b-1493">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-1493">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="2357b-1494">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="2357b-1494">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="2357b-1495">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="2357b-1495">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="2357b-1496">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1496">Object</span></span> | <span data-ttu-id="2357b-1497">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1497">&lt;optional&gt;</span></span> | <span data-ttu-id="2357b-1498">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="2357b-1498">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="2357b-1499">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1499">Object</span></span> | <span data-ttu-id="2357b-1500">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1500">&lt;optional&gt;</span></span> | <span data-ttu-id="2357b-1501">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1501">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="2357b-1502">関数</span><span class="sxs-lookup"><span data-stu-id="2357b-1502">function</span></span>| <span data-ttu-id="2357b-1503">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1503">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1504">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1504">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-1505">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1505">Requirements</span></span>

|<span data-ttu-id="2357b-1506">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1506">Requirement</span></span>| <span data-ttu-id="2357b-1507">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1507">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1508">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2357b-1509">1.7</span><span class="sxs-lookup"><span data-stu-id="2357b-1509">1.7</span></span> |
|[<span data-ttu-id="2357b-1510">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2357b-1511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2357b-1511">ReadItem</span></span> |
|[<span data-ttu-id="2357b-1512">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2357b-1513">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2357b-1513">Compose or Read</span></span> |

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="2357b-1514">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="2357b-1514">saveAsync([options], callback)</span></span>

<span data-ttu-id="2357b-1515">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1515">Asynchronously saves an item.</span></span>

<span data-ttu-id="2357b-p186">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-p186">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-1519">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="2357b-1519">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="2357b-1520">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1520">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="2357b-p188">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-p188">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="2357b-1524">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="2357b-1524">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="2357b-1525">Outlook for Mac では、会議の保存はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2357b-1525">Outlook for Mac does not support saving a meeting.</span></span> <span data-ttu-id="2357b-1526">新規`saveAsync`作成モードで会議から呼び出された場合、メソッドは失敗します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1526">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="2357b-1527">回避策については[、「OFFICE JS API を使用して Outlook For Mac で会議を下書きとして保存できません](https://support.microsoft.com/help/4505745)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2357b-1527">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="2357b-1528">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1528">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-1529">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-1529">Parameters</span></span>

|<span data-ttu-id="2357b-1530">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-1530">Name</span></span>|<span data-ttu-id="2357b-1531">型</span><span class="sxs-lookup"><span data-stu-id="2357b-1531">Type</span></span>|<span data-ttu-id="2357b-1532">属性</span><span class="sxs-lookup"><span data-stu-id="2357b-1532">Attributes</span></span>|<span data-ttu-id="2357b-1533">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-1533">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="2357b-1534">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1534">Object</span></span>|<span data-ttu-id="2357b-1535">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1535">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1536">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="2357b-1536">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2357b-1537">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1537">Object</span></span>|<span data-ttu-id="2357b-1538">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1538">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1539">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1539">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2357b-1540">関数</span><span class="sxs-lookup"><span data-stu-id="2357b-1540">function</span></span>||<span data-ttu-id="2357b-1541">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1541">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2357b-1542">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1542">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-1543">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1543">Requirements</span></span>

|<span data-ttu-id="2357b-1544">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1544">Requirement</span></span>|<span data-ttu-id="2357b-1545">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1545">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1546">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-1547">1.3</span><span class="sxs-lookup"><span data-stu-id="2357b-1547">1.3</span></span>|
|[<span data-ttu-id="2357b-1548">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-1549">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2357b-1549">ReadWriteItem</span></span>|
|[<span data-ttu-id="2357b-1550">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-1551">作成</span><span class="sxs-lookup"><span data-stu-id="2357b-1551">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="2357b-1552">例</span><span class="sxs-lookup"><span data-stu-id="2357b-1552">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="2357b-p190">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="2357b-p190">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="2357b-1555">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="2357b-1555">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="2357b-1556">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="2357b-1556">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="2357b-p191">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="2357b-p191">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2357b-1560">パラメーター</span><span class="sxs-lookup"><span data-stu-id="2357b-1560">Parameters</span></span>

|<span data-ttu-id="2357b-1561">名前</span><span class="sxs-lookup"><span data-stu-id="2357b-1561">Name</span></span>|<span data-ttu-id="2357b-1562">型</span><span class="sxs-lookup"><span data-stu-id="2357b-1562">Type</span></span>|<span data-ttu-id="2357b-1563">属性</span><span class="sxs-lookup"><span data-stu-id="2357b-1563">Attributes</span></span>|<span data-ttu-id="2357b-1564">説明</span><span class="sxs-lookup"><span data-stu-id="2357b-1564">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="2357b-1565">String</span><span class="sxs-lookup"><span data-stu-id="2357b-1565">String</span></span>||<span data-ttu-id="2357b-p192">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="2357b-p192">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="2357b-1569">Object</span><span class="sxs-lookup"><span data-stu-id="2357b-1569">Object</span></span>|<span data-ttu-id="2357b-1570">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1570">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1571">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="2357b-1571">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2357b-1572">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2357b-1572">Object</span></span>|<span data-ttu-id="2357b-1573">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1573">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-1574">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1574">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="2357b-1575">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="2357b-1575">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="2357b-1576">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2357b-1576">&lt;optional&gt;</span></span>|<span data-ttu-id="2357b-p193">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-p193">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="2357b-p194">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-p194">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="2357b-1581">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="2357b-1581">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="2357b-1582">function</span><span class="sxs-lookup"><span data-stu-id="2357b-1582">function</span></span>||<span data-ttu-id="2357b-1583">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="2357b-1583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2357b-1584">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1584">Requirements</span></span>

|<span data-ttu-id="2357b-1585">要件</span><span class="sxs-lookup"><span data-stu-id="2357b-1585">Requirement</span></span>|<span data-ttu-id="2357b-1586">値</span><span class="sxs-lookup"><span data-stu-id="2357b-1586">Value</span></span>|
|---|---|
|[<span data-ttu-id="2357b-1587">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2357b-1587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2357b-1588">1.2</span><span class="sxs-lookup"><span data-stu-id="2357b-1588">1.2</span></span>|
|[<span data-ttu-id="2357b-1589">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2357b-1589">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2357b-1590">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2357b-1590">ReadWriteItem</span></span>|
|[<span data-ttu-id="2357b-1591">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2357b-1591">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2357b-1592">作成</span><span class="sxs-lookup"><span data-stu-id="2357b-1592">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2357b-1593">例</span><span class="sxs-lookup"><span data-stu-id="2357b-1593">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
