---
title: Office. アイテム-プレビュー要件セット
description: ''
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: a529dff046f48eff65b70813617bbb9d216dba5e
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001629"
---
# <a name="item"></a><span data-ttu-id="d92fd-102">item</span><span class="sxs-lookup"><span data-stu-id="d92fd-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="d92fd-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="d92fd-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="d92fd-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-106">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-106">Requirements</span></span>

|<span data-ttu-id="d92fd-107">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-107">Requirement</span></span>|<span data-ttu-id="d92fd-108">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-110">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-110">1.0</span></span>|
|[<span data-ttu-id="d92fd-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="d92fd-112">Restricted</span></span>|
|[<span data-ttu-id="d92fd-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d92fd-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-115">Members and methods</span></span>

| <span data-ttu-id="d92fd-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-116">Member</span></span> | <span data-ttu-id="d92fd-117">種類</span><span class="sxs-lookup"><span data-stu-id="d92fd-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d92fd-118">attachments</span><span class="sxs-lookup"><span data-stu-id="d92fd-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="d92fd-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-119">Member</span></span> |
| [<span data-ttu-id="d92fd-120">bcc</span><span class="sxs-lookup"><span data-stu-id="d92fd-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="d92fd-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-121">Member</span></span> |
| [<span data-ttu-id="d92fd-122">body</span><span class="sxs-lookup"><span data-stu-id="d92fd-122">body</span></span>](#body-body) | <span data-ttu-id="d92fd-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-123">Member</span></span> |
| [<span data-ttu-id="d92fd-124">categories</span><span class="sxs-lookup"><span data-stu-id="d92fd-124">categories</span></span>](#categories-categories) | <span data-ttu-id="d92fd-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-125">Member</span></span> |
| [<span data-ttu-id="d92fd-126">cc</span><span class="sxs-lookup"><span data-stu-id="d92fd-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d92fd-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-127">Member</span></span> |
| [<span data-ttu-id="d92fd-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="d92fd-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="d92fd-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-129">Member</span></span> |
| [<span data-ttu-id="d92fd-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="d92fd-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="d92fd-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-131">Member</span></span> |
| [<span data-ttu-id="d92fd-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="d92fd-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="d92fd-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-133">Member</span></span> |
| [<span data-ttu-id="d92fd-134">end</span><span class="sxs-lookup"><span data-stu-id="d92fd-134">end</span></span>](#end-datetime) | <span data-ttu-id="d92fd-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-135">Member</span></span> |
| [<span data-ttu-id="d92fd-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="d92fd-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="d92fd-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-137">Member</span></span> |
| [<span data-ttu-id="d92fd-138">from</span><span class="sxs-lookup"><span data-stu-id="d92fd-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="d92fd-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-139">Member</span></span> |
| [<span data-ttu-id="d92fd-140">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="d92fd-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="d92fd-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-141">Member</span></span> |
| [<span data-ttu-id="d92fd-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="d92fd-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="d92fd-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-143">Member</span></span> |
| [<span data-ttu-id="d92fd-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="d92fd-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="d92fd-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-145">Member</span></span> |
| [<span data-ttu-id="d92fd-146">itemId</span><span class="sxs-lookup"><span data-stu-id="d92fd-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="d92fd-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-147">Member</span></span> |
| [<span data-ttu-id="d92fd-148">itemType</span><span class="sxs-lookup"><span data-stu-id="d92fd-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="d92fd-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-149">Member</span></span> |
| [<span data-ttu-id="d92fd-150">location</span><span class="sxs-lookup"><span data-stu-id="d92fd-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="d92fd-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-151">Member</span></span> |
| [<span data-ttu-id="d92fd-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="d92fd-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="d92fd-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-153">Member</span></span> |
| [<span data-ttu-id="d92fd-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="d92fd-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="d92fd-155">Member</span><span class="sxs-lookup"><span data-stu-id="d92fd-155">Member</span></span> |
| [<span data-ttu-id="d92fd-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="d92fd-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d92fd-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-157">Member</span></span> |
| [<span data-ttu-id="d92fd-158">organizer</span><span class="sxs-lookup"><span data-stu-id="d92fd-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="d92fd-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-159">Member</span></span> |
| [<span data-ttu-id="d92fd-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="d92fd-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="d92fd-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-161">Member</span></span> |
| [<span data-ttu-id="d92fd-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="d92fd-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d92fd-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-163">Member</span></span> |
| [<span data-ttu-id="d92fd-164">sender</span><span class="sxs-lookup"><span data-stu-id="d92fd-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="d92fd-165">Member</span><span class="sxs-lookup"><span data-stu-id="d92fd-165">Member</span></span> |
| [<span data-ttu-id="d92fd-166">系列 Id</span><span class="sxs-lookup"><span data-stu-id="d92fd-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="d92fd-167">Member</span><span class="sxs-lookup"><span data-stu-id="d92fd-167">Member</span></span> |
| [<span data-ttu-id="d92fd-168">start</span><span class="sxs-lookup"><span data-stu-id="d92fd-168">start</span></span>](#start-datetime) | <span data-ttu-id="d92fd-169">Member</span><span class="sxs-lookup"><span data-stu-id="d92fd-169">Member</span></span> |
| [<span data-ttu-id="d92fd-170">subject</span><span class="sxs-lookup"><span data-stu-id="d92fd-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="d92fd-171">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-171">Member</span></span> |
| [<span data-ttu-id="d92fd-172">to</span><span class="sxs-lookup"><span data-stu-id="d92fd-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d92fd-173">メンバー</span><span class="sxs-lookup"><span data-stu-id="d92fd-173">Member</span></span> |
| [<span data-ttu-id="d92fd-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d92fd-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="d92fd-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-175">Method</span></span> |
| [<span data-ttu-id="d92fd-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="d92fd-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="d92fd-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-177">Method</span></span> |
| [<span data-ttu-id="d92fd-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d92fd-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="d92fd-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-179">Method</span></span> |
| [<span data-ttu-id="d92fd-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d92fd-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="d92fd-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-181">Method</span></span> |
| [<span data-ttu-id="d92fd-182">close</span><span class="sxs-lookup"><span data-stu-id="d92fd-182">close</span></span>](#close) | <span data-ttu-id="d92fd-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-183">Method</span></span> |
| [<span data-ttu-id="d92fd-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="d92fd-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="d92fd-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-185">Method</span></span> |
| [<span data-ttu-id="d92fd-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="d92fd-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="d92fd-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-187">Method</span></span> |
| [<span data-ttu-id="d92fd-188">getAllInternetHeadersAsync</span><span class="sxs-lookup"><span data-stu-id="d92fd-188">getAllInternetHeadersAsync</span></span>](#getallinternetheadersasyncoptions-callback) | <span data-ttu-id="d92fd-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-189">Method</span></span> |
| [<span data-ttu-id="d92fd-190">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="d92fd-190">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="d92fd-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-191">Method</span></span> |
| [<span data-ttu-id="d92fd-192">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="d92fd-192">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="d92fd-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-193">Method</span></span> |
| [<span data-ttu-id="d92fd-194">getEntities</span><span class="sxs-lookup"><span data-stu-id="d92fd-194">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="d92fd-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-195">Method</span></span> |
| [<span data-ttu-id="d92fd-196">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="d92fd-196">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="d92fd-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-197">Method</span></span> |
| [<span data-ttu-id="d92fd-198">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="d92fd-198">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="d92fd-199">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-199">Method</span></span> |
| [<span data-ttu-id="d92fd-200">、Office.context.mailbox.item.getinitializationcontextasync</span><span class="sxs-lookup"><span data-stu-id="d92fd-200">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="d92fd-201">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-201">Method</span></span> |
| [<span data-ttu-id="d92fd-202">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="d92fd-202">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="d92fd-203">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-203">Method</span></span> |
| [<span data-ttu-id="d92fd-204">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="d92fd-204">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="d92fd-205">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-205">Method</span></span> |
| [<span data-ttu-id="d92fd-206">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="d92fd-206">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="d92fd-207">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-207">Method</span></span> |
| [<span data-ttu-id="d92fd-208">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="d92fd-208">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="d92fd-209">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-209">Method</span></span> |
| [<span data-ttu-id="d92fd-210">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="d92fd-210">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="d92fd-211">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-211">Method</span></span> |
| [<span data-ttu-id="d92fd-212">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="d92fd-212">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="d92fd-213">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-213">Method</span></span> |
| [<span data-ttu-id="d92fd-214">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="d92fd-214">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="d92fd-215">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-215">Method</span></span> |
| [<span data-ttu-id="d92fd-216">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="d92fd-216">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="d92fd-217">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-217">Method</span></span> |
| [<span data-ttu-id="d92fd-218">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d92fd-218">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="d92fd-219">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-219">Method</span></span> |
| [<span data-ttu-id="d92fd-220">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d92fd-220">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="d92fd-221">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-221">Method</span></span> |
| [<span data-ttu-id="d92fd-222">saveAsync</span><span class="sxs-lookup"><span data-stu-id="d92fd-222">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="d92fd-223">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-223">Method</span></span> |
| [<span data-ttu-id="d92fd-224">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="d92fd-224">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="d92fd-225">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-225">Method</span></span> |

### <a name="example"></a><span data-ttu-id="d92fd-226">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-226">Example</span></span>

<span data-ttu-id="d92fd-227">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-227">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="d92fd-228">Members</span><span class="sxs-lookup"><span data-stu-id="d92fd-228">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="d92fd-229">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d92fd-229">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="d92fd-230">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-230">Gets the item's attachments as an array.</span></span> <span data-ttu-id="d92fd-231">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d92fd-231">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-232">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-232">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="d92fd-233">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d92fd-233">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="d92fd-234">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-234">Type</span></span>

*   <span data-ttu-id="d92fd-235">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d92fd-235">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-236">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-236">Requirements</span></span>

|<span data-ttu-id="d92fd-237">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-237">Requirement</span></span>|<span data-ttu-id="d92fd-238">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-239">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-240">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-240">1.0</span></span>|
|[<span data-ttu-id="d92fd-241">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-242">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-243">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-244">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-244">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-245">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-245">Example</span></span>

<span data-ttu-id="d92fd-246">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-246">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="d92fd-247">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d92fd-247">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="d92fd-248">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-248">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="d92fd-249">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d92fd-249">Compose mode only.</span></span>

<span data-ttu-id="d92fd-250">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d92fd-251">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-251">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d92fd-252">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-252">Get 500 members maximum.</span></span>
- <span data-ttu-id="d92fd-253">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-253">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="d92fd-254">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-254">Type</span></span>

*   [<span data-ttu-id="d92fd-255">受信者</span><span class="sxs-lookup"><span data-stu-id="d92fd-255">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="d92fd-256">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-256">Requirements</span></span>

|<span data-ttu-id="d92fd-257">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-257">Requirement</span></span>|<span data-ttu-id="d92fd-258">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-259">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-260">1.1</span><span class="sxs-lookup"><span data-stu-id="d92fd-260">1.1</span></span>|
|[<span data-ttu-id="d92fd-261">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-262">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-264">作成</span><span class="sxs-lookup"><span data-stu-id="d92fd-264">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-265">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-265">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="d92fd-266">body: [Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="d92fd-266">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="d92fd-267">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-267">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d92fd-268">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-268">Type</span></span>

*   [<span data-ttu-id="d92fd-269">Body</span><span class="sxs-lookup"><span data-stu-id="d92fd-269">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="d92fd-270">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-270">Requirements</span></span>

|<span data-ttu-id="d92fd-271">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-271">Requirement</span></span>|<span data-ttu-id="d92fd-272">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-273">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-274">1.1</span><span class="sxs-lookup"><span data-stu-id="d92fd-274">1.1</span></span>|
|[<span data-ttu-id="d92fd-275">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-275">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-276">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-277">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-277">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-278">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-278">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-279">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-279">Example</span></span>

<span data-ttu-id="d92fd-280">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-280">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="d92fd-281">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-281">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="d92fd-282">カテゴリ:[カテゴリ](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="d92fd-282">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="d92fd-283">アイテムのカテゴリを管理するためのメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-283">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-284">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-284">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="d92fd-285">種類</span><span class="sxs-lookup"><span data-stu-id="d92fd-285">Type</span></span>

*   [<span data-ttu-id="d92fd-286">Categories</span><span class="sxs-lookup"><span data-stu-id="d92fd-286">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="d92fd-287">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-287">Requirements</span></span>

|<span data-ttu-id="d92fd-288">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-288">Requirement</span></span>|<span data-ttu-id="d92fd-289">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-290">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-291">1.8</span><span class="sxs-lookup"><span data-stu-id="d92fd-291">1.8</span></span>|
|[<span data-ttu-id="d92fd-292">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-293">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-293">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-294">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-295">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-295">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-296">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-296">Example</span></span>

<span data-ttu-id="d92fd-297">この例では、アイテムのカテゴリを取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-297">This example gets the item's categories.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="d92fd-298">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d92fd-298">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="d92fd-299">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-299">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="d92fd-300">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-300">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d92fd-301">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-301">Read mode</span></span>

<span data-ttu-id="d92fd-302">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-302">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="d92fd-303">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-303">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d92fd-304">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-304">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="d92fd-305">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-305">Compose mode</span></span>

<span data-ttu-id="d92fd-306">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-306">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="d92fd-307">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-307">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d92fd-308">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-308">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d92fd-309">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-309">Get 500 members maximum.</span></span>
- <span data-ttu-id="d92fd-310">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-310">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d92fd-311">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-311">Type</span></span>

*   <span data-ttu-id="d92fd-312">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d92fd-312">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-313">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-313">Requirements</span></span>

|<span data-ttu-id="d92fd-314">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-314">Requirement</span></span>|<span data-ttu-id="d92fd-315">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-316">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-317">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-317">1.0</span></span>|
|[<span data-ttu-id="d92fd-318">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-319">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-320">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-321">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-321">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="d92fd-322">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="d92fd-322">(nullable) conversationId: String</span></span>

<span data-ttu-id="d92fd-323">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-323">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="d92fd-p109">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="d92fd-p110">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="d92fd-328">Type</span><span class="sxs-lookup"><span data-stu-id="d92fd-328">Type</span></span>

*   <span data-ttu-id="d92fd-329">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-329">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-330">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-330">Requirements</span></span>

|<span data-ttu-id="d92fd-331">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-331">Requirement</span></span>|<span data-ttu-id="d92fd-332">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-333">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-334">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-334">1.0</span></span>|
|[<span data-ttu-id="d92fd-335">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-336">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-337">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-338">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-338">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-339">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-339">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="d92fd-340">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="d92fd-340">dateTimeCreated: Date</span></span>

<span data-ttu-id="d92fd-p111">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d92fd-343">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-343">Type</span></span>

*   <span data-ttu-id="d92fd-344">日付</span><span class="sxs-lookup"><span data-stu-id="d92fd-344">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-345">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-345">Requirements</span></span>

|<span data-ttu-id="d92fd-346">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-346">Requirement</span></span>|<span data-ttu-id="d92fd-347">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-347">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-348">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-349">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-349">1.0</span></span>|
|[<span data-ttu-id="d92fd-350">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-350">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-351">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-352">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-352">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-353">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-353">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-354">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-354">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="d92fd-355">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="d92fd-355">dateTimeModified: Date</span></span>

<span data-ttu-id="d92fd-p112">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-358">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-358">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="d92fd-359">種類</span><span class="sxs-lookup"><span data-stu-id="d92fd-359">Type</span></span>

*   <span data-ttu-id="d92fd-360">日付</span><span class="sxs-lookup"><span data-stu-id="d92fd-360">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-361">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-361">Requirements</span></span>

|<span data-ttu-id="d92fd-362">必要条件</span><span class="sxs-lookup"><span data-stu-id="d92fd-362">Requirement</span></span>|<span data-ttu-id="d92fd-363">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-363">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-364">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-364">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-365">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-365">1.0</span></span>|
|[<span data-ttu-id="d92fd-366">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-366">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-367">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-367">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-368">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-368">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-369">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-369">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-370">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-370">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="d92fd-371">end: Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="d92fd-371">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="d92fd-372">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-372">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="d92fd-p113">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d92fd-375">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-375">Read mode</span></span>

<span data-ttu-id="d92fd-376">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-376">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="d92fd-377">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-377">Compose mode</span></span>

<span data-ttu-id="d92fd-378">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-378">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="d92fd-379">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-379">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="d92fd-380">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-380">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="d92fd-381">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-381">Type</span></span>

*   <span data-ttu-id="d92fd-382">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="d92fd-382">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-383">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-383">Requirements</span></span>

|<span data-ttu-id="d92fd-384">必要条件</span><span class="sxs-lookup"><span data-stu-id="d92fd-384">Requirement</span></span>|<span data-ttu-id="d92fd-385">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-386">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-387">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-387">1.0</span></span>|
|[<span data-ttu-id="d92fd-388">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-389">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-390">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-391">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-391">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="d92fd-392">enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="d92fd-392">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="d92fd-393">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-393">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d92fd-394">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-394">Read mode</span></span>

<span data-ttu-id="d92fd-395">この`enhancedLocation`プロパティは、予定に関連付けられている場所 ( [locationdetails](/javascript/api/outlook/office.locationdetails)オブジェクトで表される) のセットを取得できる[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-395">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d92fd-396">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-396">Compose mode</span></span>

<span data-ttu-id="d92fd-397">この`enhancedLocation`プロパティは、予定の場所を取得、削除、または追加するためのメソッドを提供する[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-397">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d92fd-398">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-398">Type</span></span>

*   [<span data-ttu-id="d92fd-399">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="d92fd-399">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="d92fd-400">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-400">Requirements</span></span>

|<span data-ttu-id="d92fd-401">必要条件</span><span class="sxs-lookup"><span data-stu-id="d92fd-401">Requirement</span></span>|<span data-ttu-id="d92fd-402">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-403">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-403">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-404">1.8</span><span class="sxs-lookup"><span data-stu-id="d92fd-404">1.8</span></span>|
|[<span data-ttu-id="d92fd-405">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-405">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-406">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-407">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-407">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-408">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-408">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-409">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-409">Example</span></span>

<span data-ttu-id="d92fd-410">次の例では、予定に関連付けられている現在の場所を取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-410">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="d92fd-411">from: [emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)|[from](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="d92fd-411">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="d92fd-412">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-412">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="d92fd-p114">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-415">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-415">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d92fd-416">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-416">Read mode</span></span>

<span data-ttu-id="d92fd-417">プロパティ`from`は`EmailAddressDetails`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-417">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="d92fd-418">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-418">Compose mode</span></span>

<span data-ttu-id="d92fd-419">プロパティ`from`は、from `From`値を取得するメソッドを提供するオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-419">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d92fd-420">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-420">Type</span></span>

*   <span data-ttu-id="d92fd-421">[電子メールアドレス](/javascript/api/outlook/office.emailaddressdetails) | [の](/javascript/api/outlook/office.from)詳細</span><span class="sxs-lookup"><span data-stu-id="d92fd-421">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-422">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-422">Requirements</span></span>

|<span data-ttu-id="d92fd-423">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-423">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="d92fd-424">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-425">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-425">1.0</span></span>|<span data-ttu-id="d92fd-426">1.7</span><span class="sxs-lookup"><span data-stu-id="d92fd-426">1.7</span></span>|
|[<span data-ttu-id="d92fd-427">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-427">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-428">ReadItem</span></span>|<span data-ttu-id="d92fd-429">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-429">ReadWriteItem</span></span>|
|[<span data-ttu-id="d92fd-430">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-431">Read</span><span class="sxs-lookup"><span data-stu-id="d92fd-431">Read</span></span>|<span data-ttu-id="d92fd-432">Compose</span><span class="sxs-lookup"><span data-stu-id="d92fd-432">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="d92fd-433">internetHeaders: [internetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="d92fd-433">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="d92fd-434">メッセージのカスタムインターネットヘッダーを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-434">Gets or sets custom internet headers on a message.</span></span> <span data-ttu-id="d92fd-435">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d92fd-435">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d92fd-436">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-436">Type</span></span>

*   [<span data-ttu-id="d92fd-437">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="d92fd-437">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="d92fd-438">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-438">Requirements</span></span>

|<span data-ttu-id="d92fd-439">必要条件</span><span class="sxs-lookup"><span data-stu-id="d92fd-439">Requirement</span></span>|<span data-ttu-id="d92fd-440">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-440">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-441">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-441">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-442">1.8</span><span class="sxs-lookup"><span data-stu-id="d92fd-442">1.8</span></span>|
|[<span data-ttu-id="d92fd-443">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-443">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-444">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-445">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-445">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-446">作成</span><span class="sxs-lookup"><span data-stu-id="d92fd-446">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-447">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-447">Example</span></span>

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

#### <a name="internetmessageid-string"></a><span data-ttu-id="d92fd-448">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="d92fd-448">internetMessageId: String</span></span>

<span data-ttu-id="d92fd-p116">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d92fd-451">Type</span><span class="sxs-lookup"><span data-stu-id="d92fd-451">Type</span></span>

*   <span data-ttu-id="d92fd-452">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-452">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-453">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-453">Requirements</span></span>

|<span data-ttu-id="d92fd-454">必要条件</span><span class="sxs-lookup"><span data-stu-id="d92fd-454">Requirement</span></span>|<span data-ttu-id="d92fd-455">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-456">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-457">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-457">1.0</span></span>|
|[<span data-ttu-id="d92fd-458">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-458">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-459">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-460">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-460">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-461">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-461">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-462">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-462">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="d92fd-463">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="d92fd-463">itemClass: String</span></span>

<span data-ttu-id="d92fd-p117">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="d92fd-p118">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="d92fd-468">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-468">Type</span></span>|<span data-ttu-id="d92fd-469">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-469">Description</span></span>|<span data-ttu-id="d92fd-470">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="d92fd-470">item class</span></span>|
|---|---|---|
|<span data-ttu-id="d92fd-471">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="d92fd-471">Appointment items</span></span>|<span data-ttu-id="d92fd-472">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="d92fd-472">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="d92fd-473">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="d92fd-473">Message items</span></span>|<span data-ttu-id="d92fd-474">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-474">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="d92fd-475">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-475">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="d92fd-476">Type</span><span class="sxs-lookup"><span data-stu-id="d92fd-476">Type</span></span>

*   <span data-ttu-id="d92fd-477">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-477">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-478">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-478">Requirements</span></span>

|<span data-ttu-id="d92fd-479">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-479">Requirement</span></span>|<span data-ttu-id="d92fd-480">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-481">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-482">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-482">1.0</span></span>|
|[<span data-ttu-id="d92fd-483">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-484">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-485">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-486">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-486">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-487">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-487">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="d92fd-488">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="d92fd-488">(nullable) itemId: String</span></span>

<span data-ttu-id="d92fd-489">現在のアイテムの[Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)を取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-489">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item.</span></span> <span data-ttu-id="d92fd-490">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d92fd-490">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-491">`itemId`プロパティによって返される識別子は、 [Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)と同じです。</span><span class="sxs-lookup"><span data-stu-id="d92fd-491">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="d92fd-492">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-492">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="d92fd-493">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-493">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="d92fd-494">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d92fd-494">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="d92fd-p121">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="d92fd-497">Type</span><span class="sxs-lookup"><span data-stu-id="d92fd-497">Type</span></span>

*   <span data-ttu-id="d92fd-498">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-498">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-499">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-499">Requirements</span></span>

|<span data-ttu-id="d92fd-500">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-500">Requirement</span></span>|<span data-ttu-id="d92fd-501">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-502">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-503">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-503">1.0</span></span>|
|[<span data-ttu-id="d92fd-504">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-504">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-505">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-506">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-506">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-507">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-507">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-508">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-508">Example</span></span>

<span data-ttu-id="d92fd-p122">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="d92fd-511">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="d92fd-511">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="d92fd-512">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-512">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="d92fd-513">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="d92fd-513">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d92fd-514">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-514">Type</span></span>

*   [<span data-ttu-id="d92fd-515">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="d92fd-515">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="d92fd-516">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-516">Requirements</span></span>

|<span data-ttu-id="d92fd-517">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-517">Requirement</span></span>|<span data-ttu-id="d92fd-518">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-518">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-519">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-519">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-520">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-520">1.0</span></span>|
|[<span data-ttu-id="d92fd-521">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-521">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-522">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-522">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-523">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-523">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-524">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-524">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-525">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-525">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="d92fd-526">location: String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="d92fd-526">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="d92fd-527">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-527">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d92fd-528">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-528">Read mode</span></span>

<span data-ttu-id="d92fd-529">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-529">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="d92fd-530">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-530">Compose mode</span></span>

<span data-ttu-id="d92fd-531">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-531">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d92fd-532">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-532">Type</span></span>

*   <span data-ttu-id="d92fd-533">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="d92fd-533">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-534">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-534">Requirements</span></span>

|<span data-ttu-id="d92fd-535">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-535">Requirement</span></span>|<span data-ttu-id="d92fd-536">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-537">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-537">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-538">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-538">1.0</span></span>|
|[<span data-ttu-id="d92fd-539">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-539">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-540">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-541">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-541">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-542">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-542">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="d92fd-543">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="d92fd-543">normalizedSubject: String</span></span>

<span data-ttu-id="d92fd-p123">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="d92fd-p124">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="d92fd-548">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-548">Type</span></span>

*   <span data-ttu-id="d92fd-549">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-549">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-550">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-550">Requirements</span></span>

|<span data-ttu-id="d92fd-551">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-551">Requirement</span></span>|<span data-ttu-id="d92fd-552">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-552">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-553">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-553">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-554">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-554">1.0</span></span>|
|[<span data-ttu-id="d92fd-555">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-555">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-556">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-556">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-557">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-557">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-558">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-558">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-559">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-559">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="d92fd-560">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="d92fd-560">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="d92fd-561">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-561">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d92fd-562">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-562">Type</span></span>

*   [<span data-ttu-id="d92fd-563">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="d92fd-563">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="d92fd-564">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-564">Requirements</span></span>

|<span data-ttu-id="d92fd-565">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-565">Requirement</span></span>|<span data-ttu-id="d92fd-566">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-567">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-568">1.3</span><span class="sxs-lookup"><span data-stu-id="d92fd-568">1.3</span></span>|
|[<span data-ttu-id="d92fd-569">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-570">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-571">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-572">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-572">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-573">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-573">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="d92fd-574">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d92fd-574">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="d92fd-575">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-575">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="d92fd-576">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-576">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d92fd-577">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-577">Read mode</span></span>

<span data-ttu-id="d92fd-578">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-578">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="d92fd-579">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-579">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d92fd-580">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-580">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="d92fd-581">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-581">Compose mode</span></span>

<span data-ttu-id="d92fd-582">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-582">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="d92fd-583">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-583">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d92fd-584">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-584">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d92fd-585">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-585">Get 500 members maximum.</span></span>
- <span data-ttu-id="d92fd-586">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-586">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d92fd-587">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-587">Type</span></span>

*   <span data-ttu-id="d92fd-588">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d92fd-588">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-589">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-589">Requirements</span></span>

|<span data-ttu-id="d92fd-590">必要条件</span><span class="sxs-lookup"><span data-stu-id="d92fd-590">Requirement</span></span>|<span data-ttu-id="d92fd-591">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-592">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-593">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-593">1.0</span></span>|
|[<span data-ttu-id="d92fd-594">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-594">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-595">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-595">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-596">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-597">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-597">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="d92fd-598">開催者: [emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)|[開催者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="d92fd-598">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="d92fd-599">指定した会議の開催者の電子メールアドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-599">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d92fd-600">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-600">Read mode</span></span>

<span data-ttu-id="d92fd-601">プロパティ`organizer`は、会議の開催者を表す[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-601">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="d92fd-602">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-602">Compose mode</span></span>

<span data-ttu-id="d92fd-603">プロパティ`organizer`は、開催者の値を取得するためのメソッドを提供する[オーガナイザー](/javascript/api/outlook/office.organizer)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-603">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="d92fd-604">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-604">Type</span></span>

*   <span data-ttu-id="d92fd-605">[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails) | [開催者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="d92fd-605">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-606">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-606">Requirements</span></span>

|<span data-ttu-id="d92fd-607">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-607">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="d92fd-608">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-609">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-609">1.0</span></span>|<span data-ttu-id="d92fd-610">1.7</span><span class="sxs-lookup"><span data-stu-id="d92fd-610">1.7</span></span>|
|[<span data-ttu-id="d92fd-611">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-611">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-612">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-612">ReadItem</span></span>|<span data-ttu-id="d92fd-613">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-613">ReadWriteItem</span></span>|
|[<span data-ttu-id="d92fd-614">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-614">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-615">Read</span><span class="sxs-lookup"><span data-stu-id="d92fd-615">Read</span></span>|<span data-ttu-id="d92fd-616">Compose</span><span class="sxs-lookup"><span data-stu-id="d92fd-616">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="d92fd-617">(nullable) 定期的なスケジュール:[定期的](/javascript/api/outlook/office.recurrence)なアイテム</span><span class="sxs-lookup"><span data-stu-id="d92fd-617">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="d92fd-618">予定の定期的なパターンを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-618">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="d92fd-619">会議出席依頼の定期的なパターンを取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-619">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="d92fd-620">予定アイテムの読み取りおよび作成モード。</span><span class="sxs-lookup"><span data-stu-id="d92fd-620">Read and compose modes for appointment items.</span></span> <span data-ttu-id="d92fd-621">会議出席依頼アイテムの閲覧モード。</span><span class="sxs-lookup"><span data-stu-id="d92fd-621">Read mode for meeting request items.</span></span>

<span data-ttu-id="d92fd-622">この`recurrence`プロパティは、アイテムが series または series 内のインスタンスの場合、定期的な予定または会議出席依頼に対して[定期的](/javascript/api/outlook/office.recurrence)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-622">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="d92fd-623">`null`は、単一の予定および1つの予定の会議出席依頼に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-623">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="d92fd-624">`undefined`は、会議出席依頼ではないメッセージに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-624">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="d92fd-625">注: 会議出席依頼に`itemClass`は、IPM という値があります。出席依頼。</span><span class="sxs-lookup"><span data-stu-id="d92fd-625">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="d92fd-626">注: 定期的なオブジェクトが`null`の場合は、そのオブジェクトが単一の予定または1つの予定の会議出席依頼であり、データ系列の一部ではないことを示します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-626">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d92fd-627">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-627">Read mode</span></span>

<span data-ttu-id="d92fd-628">この`recurrence`プロパティは、定期的な予定を表す[定期的](/javascript/api/outlook/office.recurrence)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-628">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="d92fd-629">これは、予定および会議出席依頼に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-629">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="d92fd-630">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-630">Compose mode</span></span>

<span data-ttu-id="d92fd-631">この`recurrence`プロパティは、予定の繰り返しを管理するためのメソッドを提供する[定期的](/javascript/api/outlook/office.recurrence)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-631">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="d92fd-632">これは予定に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-632">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="d92fd-633">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-633">Type</span></span>

* [<span data-ttu-id="d92fd-634">繰り返さ</span><span class="sxs-lookup"><span data-stu-id="d92fd-634">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="d92fd-635">必要条件</span><span class="sxs-lookup"><span data-stu-id="d92fd-635">Requirement</span></span>|<span data-ttu-id="d92fd-636">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-636">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-637">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-638">1.7</span><span class="sxs-lookup"><span data-stu-id="d92fd-638">1.7</span></span>|
|[<span data-ttu-id="d92fd-639">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-639">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-640">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-640">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-641">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-641">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-642">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-642">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="d92fd-643">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d92fd-643">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="d92fd-644">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-644">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="d92fd-645">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-645">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d92fd-646">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-646">Read mode</span></span>

<span data-ttu-id="d92fd-647">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-647">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="d92fd-648">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-648">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d92fd-649">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-649">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="d92fd-650">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-650">Compose mode</span></span>

<span data-ttu-id="d92fd-651">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-651">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="d92fd-652">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-652">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d92fd-653">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-653">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d92fd-654">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-654">Get 500 members maximum.</span></span>
- <span data-ttu-id="d92fd-655">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-655">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="d92fd-656">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-656">Type</span></span>

*   <span data-ttu-id="d92fd-657">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d92fd-657">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-658">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-658">Requirements</span></span>

|<span data-ttu-id="d92fd-659">必要条件</span><span class="sxs-lookup"><span data-stu-id="d92fd-659">Requirement</span></span>|<span data-ttu-id="d92fd-660">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-661">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-662">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-662">1.0</span></span>|
|[<span data-ttu-id="d92fd-663">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-664">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-664">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-665">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-666">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-666">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="d92fd-667">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="d92fd-667">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="d92fd-p135">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p135">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="d92fd-p136">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsfrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p136">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-672">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-672">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="d92fd-673">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-673">Type</span></span>

*   [<span data-ttu-id="d92fd-674">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d92fd-674">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="d92fd-675">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-675">Requirements</span></span>

|<span data-ttu-id="d92fd-676">必要条件</span><span class="sxs-lookup"><span data-stu-id="d92fd-676">Requirement</span></span>|<span data-ttu-id="d92fd-677">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-677">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-678">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-678">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-679">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-679">1.0</span></span>|
|[<span data-ttu-id="d92fd-680">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-680">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-681">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-681">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-682">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-682">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-683">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-683">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-684">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-684">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="d92fd-685">(nullable) 系列 Id: String</span><span class="sxs-lookup"><span data-stu-id="d92fd-685">(nullable) seriesId: String</span></span>

<span data-ttu-id="d92fd-686">インスタンスが属する系列の id を取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-686">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="d92fd-687">Web 上の Outlook およびデスクトップクライアントでは、 `seriesId`は、このアイテムが属する親 (シリーズ) アイテムの Exchange web サービス (EWS) ID を返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-687">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="d92fd-688">ただし、iOS と Android では、 `seriesId`は親アイテムの REST ID を返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-688">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-689">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="d92fd-689">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="d92fd-690">`seriesId`プロパティが OUTLOOK REST API で使用される outlook id と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-690">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="d92fd-691">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-691">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="d92fd-692">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d92fd-692">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="d92fd-693">この`seriesId`プロパティは`null` 、単一の予定、系列のアイテム、会議出席依頼などの親アイテムを持たないアイテムに`undefined`対して、会議出席依頼以外のアイテムに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-693">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="d92fd-694">Type</span><span class="sxs-lookup"><span data-stu-id="d92fd-694">Type</span></span>

* <span data-ttu-id="d92fd-695">文字列</span><span class="sxs-lookup"><span data-stu-id="d92fd-695">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-696">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-696">Requirements</span></span>

|<span data-ttu-id="d92fd-697">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-697">Requirement</span></span>|<span data-ttu-id="d92fd-698">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-698">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-699">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-699">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-700">1.7</span><span class="sxs-lookup"><span data-stu-id="d92fd-700">1.7</span></span>|
|[<span data-ttu-id="d92fd-701">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-701">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-702">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-702">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-703">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-703">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-704">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-704">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-705">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-705">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="d92fd-706">start: Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="d92fd-706">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="d92fd-707">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-707">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="d92fd-p139">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p139">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d92fd-710">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-710">Read mode</span></span>

<span data-ttu-id="d92fd-711">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-711">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="d92fd-712">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-712">Compose mode</span></span>

<span data-ttu-id="d92fd-713">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-713">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="d92fd-714">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-714">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="d92fd-715">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-715">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="d92fd-716">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-716">Type</span></span>

*   <span data-ttu-id="d92fd-717">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="d92fd-717">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-718">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-718">Requirements</span></span>

|<span data-ttu-id="d92fd-719">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-719">Requirement</span></span>|<span data-ttu-id="d92fd-720">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-720">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-721">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-721">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-722">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-722">1.0</span></span>|
|[<span data-ttu-id="d92fd-723">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-723">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-724">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-724">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-725">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-725">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-726">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-726">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="d92fd-727">subject: String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="d92fd-727">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="d92fd-728">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-728">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="d92fd-729">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-729">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d92fd-730">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-730">Read mode</span></span>

<span data-ttu-id="d92fd-p140">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p140">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="d92fd-733">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-733">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="d92fd-734">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-734">Compose mode</span></span>
<span data-ttu-id="d92fd-735">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-735">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="d92fd-736">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-736">Type</span></span>

*   <span data-ttu-id="d92fd-737">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="d92fd-737">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-738">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-738">Requirements</span></span>

|<span data-ttu-id="d92fd-739">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-739">Requirement</span></span>|<span data-ttu-id="d92fd-740">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-740">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-741">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-741">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-742">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-742">1.0</span></span>|
|[<span data-ttu-id="d92fd-743">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-743">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-744">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-744">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-745">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-745">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-746">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-746">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="d92fd-747">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d92fd-747">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="d92fd-748">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-748">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="d92fd-749">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-749">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d92fd-750">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-750">Read mode</span></span>

<span data-ttu-id="d92fd-751">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-751">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="d92fd-752">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-752">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d92fd-753">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-753">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="d92fd-754">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d92fd-754">Compose mode</span></span>

<span data-ttu-id="d92fd-755">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-755">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="d92fd-756">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-756">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d92fd-757">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-757">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d92fd-758">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-758">Get 500 members maximum.</span></span>
- <span data-ttu-id="d92fd-759">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-759">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d92fd-760">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-760">Type</span></span>

*   <span data-ttu-id="d92fd-761">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d92fd-761">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-762">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-762">Requirements</span></span>

|<span data-ttu-id="d92fd-763">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-763">Requirement</span></span>|<span data-ttu-id="d92fd-764">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-764">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-765">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-765">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-766">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-766">1.0</span></span>|
|[<span data-ttu-id="d92fd-767">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-767">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-768">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-768">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-769">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-769">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-770">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-770">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="d92fd-771">メソッド</span><span class="sxs-lookup"><span data-stu-id="d92fd-771">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="d92fd-772">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d92fd-772">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d92fd-773">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-773">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="d92fd-774">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-774">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="d92fd-775">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-775">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-776">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-776">Parameters</span></span>
|<span data-ttu-id="d92fd-777">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-777">Name</span></span>|<span data-ttu-id="d92fd-778">種類</span><span class="sxs-lookup"><span data-stu-id="d92fd-778">Type</span></span>|<span data-ttu-id="d92fd-779">属性</span><span class="sxs-lookup"><span data-stu-id="d92fd-779">Attributes</span></span>|<span data-ttu-id="d92fd-780">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-780">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="d92fd-781">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-781">String</span></span>||<span data-ttu-id="d92fd-p144">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p144">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="d92fd-784">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-784">String</span></span>||<span data-ttu-id="d92fd-p145">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p145">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="d92fd-787">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-787">Object</span></span>|<span data-ttu-id="d92fd-788">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-788">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-789">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d92fd-789">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d92fd-790">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-790">Object</span></span>|<span data-ttu-id="d92fd-791">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-791">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-792">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-792">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="d92fd-793">Boolean</span><span class="sxs-lookup"><span data-stu-id="d92fd-793">Boolean</span></span>|<span data-ttu-id="d92fd-794">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-794">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-795">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-795">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="d92fd-796">function</span><span class="sxs-lookup"><span data-stu-id="d92fd-796">function</span></span>|<span data-ttu-id="d92fd-797">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-797">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-798">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-798">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d92fd-799">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-799">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d92fd-800">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-800">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d92fd-801">エラー</span><span class="sxs-lookup"><span data-stu-id="d92fd-801">Errors</span></span>

|<span data-ttu-id="d92fd-802">エラー コード</span><span class="sxs-lookup"><span data-stu-id="d92fd-802">Error code</span></span>|<span data-ttu-id="d92fd-803">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-803">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="d92fd-804">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-804">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="d92fd-805">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="d92fd-805">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="d92fd-806">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-806">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-807">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-807">Requirements</span></span>

|<span data-ttu-id="d92fd-808">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-808">Requirement</span></span>|<span data-ttu-id="d92fd-809">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-809">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-810">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-810">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-811">1.1</span><span class="sxs-lookup"><span data-stu-id="d92fd-811">1.1</span></span>|
|[<span data-ttu-id="d92fd-812">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-812">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-813">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-813">ReadWriteItem</span></span>|
|[<span data-ttu-id="d92fd-814">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-814">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-815">作成</span><span class="sxs-lookup"><span data-stu-id="d92fd-815">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d92fd-816">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-816">Examples</span></span>

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

<span data-ttu-id="d92fd-817">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-817">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="d92fd-818">addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d92fd-818">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d92fd-819">Base64 エンコードのファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-819">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="d92fd-820">この`addFileAttachmentFromBase64Async`メソッドは、base64 エンコードからファイルをアップロードし、新規作成フォームのアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-820">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="d92fd-821">このメソッドは、AsyncResult オブジェクトの添付ファイル識別子を返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-821">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="d92fd-822">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-822">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-823">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-823">Parameters</span></span>

|<span data-ttu-id="d92fd-824">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-824">Name</span></span>|<span data-ttu-id="d92fd-825">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-825">Type</span></span>|<span data-ttu-id="d92fd-826">属性</span><span class="sxs-lookup"><span data-stu-id="d92fd-826">Attributes</span></span>|<span data-ttu-id="d92fd-827">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-827">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="d92fd-828">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-828">String</span></span>||<span data-ttu-id="d92fd-829">電子メールまたはイベントに追加する画像またはファイルの、base64 でエンコードされたコンテンツ。</span><span class="sxs-lookup"><span data-stu-id="d92fd-829">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="d92fd-830">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-830">String</span></span>||<span data-ttu-id="d92fd-p147">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p147">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="d92fd-833">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-833">Object</span></span>|<span data-ttu-id="d92fd-834">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-834">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-835">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d92fd-835">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d92fd-836">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-836">Object</span></span>|<span data-ttu-id="d92fd-837">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-837">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-838">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-838">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="d92fd-839">Boolean</span><span class="sxs-lookup"><span data-stu-id="d92fd-839">Boolean</span></span>|<span data-ttu-id="d92fd-840">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-840">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-841">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-841">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="d92fd-842">function</span><span class="sxs-lookup"><span data-stu-id="d92fd-842">function</span></span>|<span data-ttu-id="d92fd-843">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-843">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-844">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-844">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d92fd-845">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-845">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d92fd-846">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-846">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d92fd-847">エラー</span><span class="sxs-lookup"><span data-stu-id="d92fd-847">Errors</span></span>

|<span data-ttu-id="d92fd-848">エラー コード</span><span class="sxs-lookup"><span data-stu-id="d92fd-848">Error code</span></span>|<span data-ttu-id="d92fd-849">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-849">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="d92fd-850">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-850">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="d92fd-851">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="d92fd-851">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="d92fd-852">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-852">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-853">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-853">Requirements</span></span>

|<span data-ttu-id="d92fd-854">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-854">Requirement</span></span>|<span data-ttu-id="d92fd-855">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-856">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-856">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-857">1.8</span><span class="sxs-lookup"><span data-stu-id="d92fd-857">1.8</span></span>|
|[<span data-ttu-id="d92fd-858">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-858">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-859">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-859">ReadWriteItem</span></span>|
|[<span data-ttu-id="d92fd-860">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-860">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-861">作成</span><span class="sxs-lookup"><span data-stu-id="d92fd-861">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d92fd-862">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-862">Examples</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="d92fd-863">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d92fd-863">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="d92fd-864">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-864">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="d92fd-865">現在、サポートされて`Office.EventType.AttachmentsChanged`いる`Office.EventType.AppointmentTimeChanged`イベント`Office.EventType.EnhancedLocationsChanged`の`Office.EventType.RecipientsChanged`種類は`Office.EventType.RecurrenceChanged`、、、、、です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-865">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-866">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-866">Parameters</span></span>

| <span data-ttu-id="d92fd-867">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-867">Name</span></span> | <span data-ttu-id="d92fd-868">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-868">Type</span></span> | <span data-ttu-id="d92fd-869">属性</span><span class="sxs-lookup"><span data-stu-id="d92fd-869">Attributes</span></span> | <span data-ttu-id="d92fd-870">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-870">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d92fd-871">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d92fd-871">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d92fd-872">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="d92fd-872">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="d92fd-873">Function</span><span class="sxs-lookup"><span data-stu-id="d92fd-873">Function</span></span> || <span data-ttu-id="d92fd-p148">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p148">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="d92fd-877">Object</span><span class="sxs-lookup"><span data-stu-id="d92fd-877">Object</span></span> | <span data-ttu-id="d92fd-878">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-878">&lt;optional&gt;</span></span> | <span data-ttu-id="d92fd-879">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d92fd-879">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d92fd-880">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-880">Object</span></span> | <span data-ttu-id="d92fd-881">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-881">&lt;optional&gt;</span></span> | <span data-ttu-id="d92fd-882">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-882">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d92fd-883">function</span><span class="sxs-lookup"><span data-stu-id="d92fd-883">function</span></span>| <span data-ttu-id="d92fd-884">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-884">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-885">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-885">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-886">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-886">Requirements</span></span>

|<span data-ttu-id="d92fd-887">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-887">Requirement</span></span>| <span data-ttu-id="d92fd-888">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-888">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-889">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-889">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d92fd-890">1.7</span><span class="sxs-lookup"><span data-stu-id="d92fd-890">1.7</span></span> |
|[<span data-ttu-id="d92fd-891">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-891">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d92fd-892">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-892">ReadItem</span></span> |
|[<span data-ttu-id="d92fd-893">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-893">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d92fd-894">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-894">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="d92fd-895">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-895">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="d92fd-896">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d92fd-896">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d92fd-897">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-897">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="d92fd-p149">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p149">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="d92fd-901">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-901">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="d92fd-902">Office アドインを Outlook on the web で実行している場合、編集中のアイテム以外のアイテムに `addItemAttachmentAsync` メソッドでアイテムを添付できます。ただし、これはサポートされていないため、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-902">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-903">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-903">Parameters</span></span>

|<span data-ttu-id="d92fd-904">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-904">Name</span></span>|<span data-ttu-id="d92fd-905">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-905">Type</span></span>|<span data-ttu-id="d92fd-906">属性</span><span class="sxs-lookup"><span data-stu-id="d92fd-906">Attributes</span></span>|<span data-ttu-id="d92fd-907">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-907">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="d92fd-908">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-908">String</span></span>||<span data-ttu-id="d92fd-p150">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p150">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="d92fd-911">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-911">String</span></span>||<span data-ttu-id="d92fd-912">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="d92fd-912">The subject of the item to be attached.</span></span> <span data-ttu-id="d92fd-913">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-913">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="d92fd-914">Object</span><span class="sxs-lookup"><span data-stu-id="d92fd-914">Object</span></span>|<span data-ttu-id="d92fd-915">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-915">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-916">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d92fd-916">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d92fd-917">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-917">Object</span></span>|<span data-ttu-id="d92fd-918">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-918">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-919">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-919">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d92fd-920">関数</span><span class="sxs-lookup"><span data-stu-id="d92fd-920">function</span></span>|<span data-ttu-id="d92fd-921">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-921">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-922">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-922">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d92fd-923">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-923">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d92fd-924">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-924">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d92fd-925">エラー</span><span class="sxs-lookup"><span data-stu-id="d92fd-925">Errors</span></span>

|<span data-ttu-id="d92fd-926">エラー コード</span><span class="sxs-lookup"><span data-stu-id="d92fd-926">Error code</span></span>|<span data-ttu-id="d92fd-927">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-927">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="d92fd-928">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-928">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-929">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-929">Requirements</span></span>

|<span data-ttu-id="d92fd-930">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-930">Requirement</span></span>|<span data-ttu-id="d92fd-931">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-931">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-932">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-932">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-933">1.1</span><span class="sxs-lookup"><span data-stu-id="d92fd-933">1.1</span></span>|
|[<span data-ttu-id="d92fd-934">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-934">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-935">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-935">ReadWriteItem</span></span>|
|[<span data-ttu-id="d92fd-936">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-936">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-937">作成</span><span class="sxs-lookup"><span data-stu-id="d92fd-937">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-938">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-938">Example</span></span>

<span data-ttu-id="d92fd-939">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-939">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="d92fd-940">close()</span><span class="sxs-lookup"><span data-stu-id="d92fd-940">close()</span></span>

<span data-ttu-id="d92fd-941">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-941">Closes the current item that is being composed.</span></span>

<span data-ttu-id="d92fd-p152">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p152">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-944">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-944">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="d92fd-945">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-945">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-946">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-946">Requirements</span></span>

|<span data-ttu-id="d92fd-947">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-947">Requirement</span></span>|<span data-ttu-id="d92fd-948">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-948">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-949">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-949">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-950">1.3</span><span class="sxs-lookup"><span data-stu-id="d92fd-950">1.3</span></span>|
|[<span data-ttu-id="d92fd-951">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-951">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-952">制限あり</span><span class="sxs-lookup"><span data-stu-id="d92fd-952">Restricted</span></span>|
|[<span data-ttu-id="d92fd-953">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-953">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-954">新規作成</span><span class="sxs-lookup"><span data-stu-id="d92fd-954">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="d92fd-955">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="d92fd-955">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="d92fd-956">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-956">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-957">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-957">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d92fd-958">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-958">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d92fd-959">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="d92fd-959">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="d92fd-p153">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-963">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-963">Parameters</span></span>

|<span data-ttu-id="d92fd-964">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-964">Name</span></span>|<span data-ttu-id="d92fd-965">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-965">Type</span></span>|<span data-ttu-id="d92fd-966">属性</span><span class="sxs-lookup"><span data-stu-id="d92fd-966">Attributes</span></span>|<span data-ttu-id="d92fd-967">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-967">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="d92fd-968">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="d92fd-968">String &#124; Object</span></span>||<span data-ttu-id="d92fd-p154">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d92fd-971">**または**</span><span class="sxs-lookup"><span data-stu-id="d92fd-971">**OR**</span></span><br/><span data-ttu-id="d92fd-p155">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="d92fd-974">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-974">String</span></span>|<span data-ttu-id="d92fd-975">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-975">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-p156">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="d92fd-978">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-978">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="d92fd-979">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-979">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-980">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="d92fd-980">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="d92fd-981">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-981">String</span></span>||<span data-ttu-id="d92fd-p157">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="d92fd-984">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-984">String</span></span>||<span data-ttu-id="d92fd-985">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-985">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="d92fd-986">文字列</span><span class="sxs-lookup"><span data-stu-id="d92fd-986">String</span></span>||<span data-ttu-id="d92fd-p158">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="d92fd-989">ブール値</span><span class="sxs-lookup"><span data-stu-id="d92fd-989">Boolean</span></span>||<span data-ttu-id="d92fd-p159">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="d92fd-992">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-992">String</span></span>||<span data-ttu-id="d92fd-p160">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="d92fd-996">function</span><span class="sxs-lookup"><span data-stu-id="d92fd-996">function</span></span>|<span data-ttu-id="d92fd-997">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-997">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-998">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-998">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-999">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-999">Requirements</span></span>

|<span data-ttu-id="d92fd-1000">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1000">Requirement</span></span>|<span data-ttu-id="d92fd-1001">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1001">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1002">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1002">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1003">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-1003">1.0</span></span>|
|[<span data-ttu-id="d92fd-1004">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1004">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1005">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1005">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-1006">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1006">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1007">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-1007">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d92fd-1008">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1008">Examples</span></span>

<span data-ttu-id="d92fd-1009">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1009">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="d92fd-1010">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1010">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="d92fd-1011">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1011">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d92fd-1012">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1012">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d92fd-1013">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1013">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d92fd-1014">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1014">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="d92fd-1015">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="d92fd-1015">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="d92fd-1016">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1016">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-1017">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1017">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d92fd-1018">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1018">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d92fd-1019">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1019">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="d92fd-p161">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p161">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-1023">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-1023">Parameters</span></span>

|<span data-ttu-id="d92fd-1024">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-1024">Name</span></span>|<span data-ttu-id="d92fd-1025">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-1025">Type</span></span>|<span data-ttu-id="d92fd-1026">属性</span><span class="sxs-lookup"><span data-stu-id="d92fd-1026">Attributes</span></span>|<span data-ttu-id="d92fd-1027">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-1027">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="d92fd-1028">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="d92fd-1028">String &#124; Object</span></span>||<span data-ttu-id="d92fd-p162">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p162">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d92fd-1031">**または**</span><span class="sxs-lookup"><span data-stu-id="d92fd-1031">**OR**</span></span><br/><span data-ttu-id="d92fd-p163">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p163">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="d92fd-1034">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-1034">String</span></span>|<span data-ttu-id="d92fd-1035">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1035">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-p164">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p164">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="d92fd-1038">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1038">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="d92fd-1039">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1040">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1040">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="d92fd-1041">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-1041">String</span></span>||<span data-ttu-id="d92fd-p165">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p165">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="d92fd-1044">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-1044">String</span></span>||<span data-ttu-id="d92fd-1045">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1045">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="d92fd-1046">文字列</span><span class="sxs-lookup"><span data-stu-id="d92fd-1046">String</span></span>||<span data-ttu-id="d92fd-p166">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p166">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="d92fd-1049">ブール値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1049">Boolean</span></span>||<span data-ttu-id="d92fd-p167">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p167">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="d92fd-1052">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-1052">String</span></span>||<span data-ttu-id="d92fd-p168">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p168">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="d92fd-1056">function</span><span class="sxs-lookup"><span data-stu-id="d92fd-1056">function</span></span>|<span data-ttu-id="d92fd-1057">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1058">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-1059">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1059">Requirements</span></span>

|<span data-ttu-id="d92fd-1060">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1060">Requirement</span></span>|<span data-ttu-id="d92fd-1061">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1061">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1062">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1062">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1063">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-1063">1.0</span></span>|
|[<span data-ttu-id="d92fd-1064">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1064">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1065">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1065">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-1066">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1066">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1067">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-1067">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d92fd-1068">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1068">Examples</span></span>

<span data-ttu-id="d92fd-1069">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1069">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="d92fd-1070">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1070">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="d92fd-1071">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1071">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d92fd-1072">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1072">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d92fd-1073">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1073">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d92fd-1074">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1074">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getallinternetheadersasyncoptions-callback"></a><span data-ttu-id="d92fd-1075">getAllInternetHeadersAsync ([オプション], [callback])</span><span class="sxs-lookup"><span data-stu-id="d92fd-1075">getAllInternetHeadersAsync([options], [callback])</span></span>

<span data-ttu-id="d92fd-1076">メッセージのすべてのインターネットヘッダーを文字列として取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1076">Gets all the internet headers for the message as a string.</span></span> <span data-ttu-id="d92fd-1077">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1077">Read mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-1078">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-1078">Parameters</span></span>

|<span data-ttu-id="d92fd-1079">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-1079">Name</span></span>|<span data-ttu-id="d92fd-1080">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-1080">Type</span></span>|<span data-ttu-id="d92fd-1081">属性</span><span class="sxs-lookup"><span data-stu-id="d92fd-1081">Attributes</span></span>|<span data-ttu-id="d92fd-1082">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-1082">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="d92fd-1083">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1083">Object</span></span>|<span data-ttu-id="d92fd-1084">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1084">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1085">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1085">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d92fd-1086">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1086">Object</span></span>|<span data-ttu-id="d92fd-1087">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1087">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1088">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1088">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d92fd-1089">関数</span><span class="sxs-lookup"><span data-stu-id="d92fd-1089">function</span></span>|<span data-ttu-id="d92fd-1090">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1090">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1091">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1091">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="d92fd-1092">成功した場合、インターネットヘッダーデータは、文字列として asyncResult プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1092">On success, the internet headers data is provided in the asyncResult.value property as a string.</span></span> <span data-ttu-id="d92fd-1093">返される文字列値の書式情報については、 [RFC 2183](https://tools.ietf.org/html/rfc2183)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1093">Refer to [RFC 2183](https://tools.ietf.org/html/rfc2183) for the formatting information of the returned string value.</span></span> <span data-ttu-id="d92fd-1094">呼び出しが失敗した場合、asyncResult. error プロパティには、エラーの理由と共にエラーコードが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1094">If the call fails, the asyncResult.error property will contain an error code with the reason for the failure.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-1095">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-1095">Requirements</span></span>

|<span data-ttu-id="d92fd-1096">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1096">Requirement</span></span>|<span data-ttu-id="d92fd-1097">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1097">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1098">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1098">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1099">1.8</span><span class="sxs-lookup"><span data-stu-id="d92fd-1099">1.8</span></span>|
|[<span data-ttu-id="d92fd-1100">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1100">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1101">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1101">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-1102">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1102">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1103">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-1103">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d92fd-1104">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d92fd-1104">Returns:</span></span>

<span data-ttu-id="d92fd-1105">[RFC 2183](https://tools.ietf.org/html/rfc2183)に従って書式設定された文字列としてのインターネットヘッダーデータ。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1105">The internet headers data as a string formatted according to [RFC 2183](https://tools.ietf.org/html/rfc2183).</span></span>

<span data-ttu-id="d92fd-1106">型:String</span><span class="sxs-lookup"><span data-stu-id="d92fd-1106">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d92fd-1107">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1107">Example</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="d92fd-1108">getAttachmentContentAsync (attachmentId, [options], [callback]) > [Attachmentcontent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="d92fd-1108">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="d92fd-1109">メッセージまたは予定から指定された添付ファイルを取得し`AttachmentContent` 、それをオブジェクトとして返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1109">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="d92fd-1110">メソッド`getAttachmentContentAsync`は、指定された id の添付ファイルをアイテムから取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1110">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="d92fd-1111">ベストプラクティスとして、識別子を使用して、または`getAttachmentsAsync` `item.attachments`の呼び出しで attachmentIds を取得したのと同じセッションの添付ファイルを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1111">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="d92fd-1112">Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1112">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="d92fd-1113">ユーザーがアプリを閉じたとき、またはインラインフォームの作成が開始されたときに、別のウィンドウで続行するためにフォームをポップアウトした後、セッションが終了します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1113">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-1114">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-1114">Parameters</span></span>

|<span data-ttu-id="d92fd-1115">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-1115">Name</span></span>|<span data-ttu-id="d92fd-1116">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-1116">Type</span></span>|<span data-ttu-id="d92fd-1117">属性</span><span class="sxs-lookup"><span data-stu-id="d92fd-1117">Attributes</span></span>|<span data-ttu-id="d92fd-1118">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-1118">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="d92fd-1119">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-1119">String</span></span>||<span data-ttu-id="d92fd-1120">取得する添付ファイルの識別子を指定します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1120">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="d92fd-1121">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1121">Object</span></span>|<span data-ttu-id="d92fd-1122">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1122">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1123">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1123">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d92fd-1124">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1124">Object</span></span>|<span data-ttu-id="d92fd-1125">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1125">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1126">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1126">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d92fd-1127">関数</span><span class="sxs-lookup"><span data-stu-id="d92fd-1127">function</span></span>|<span data-ttu-id="d92fd-1128">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1128">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1129">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1129">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-1130">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1130">Requirements</span></span>

|<span data-ttu-id="d92fd-1131">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1131">Requirement</span></span>|<span data-ttu-id="d92fd-1132">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1132">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1133">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1134">1.8</span><span class="sxs-lookup"><span data-stu-id="d92fd-1134">1.8</span></span>|
|[<span data-ttu-id="d92fd-1135">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1136">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-1137">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1138">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-1138">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d92fd-1139">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d92fd-1139">Returns:</span></span>

<span data-ttu-id="d92fd-1140">型: [Attachmentcontent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="d92fd-1140">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="d92fd-1141">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1141">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="d92fd-1142">getAttachmentsAsync ([オプション], [callback]) > Array. <[Attachmentdetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d92fd-1142">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="d92fd-1143">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1143">Gets the item's attachments as an array.</span></span> <span data-ttu-id="d92fd-1144">新規作成モードのみです。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1144">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-1145">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-1145">Parameters</span></span>

|<span data-ttu-id="d92fd-1146">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-1146">Name</span></span>|<span data-ttu-id="d92fd-1147">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-1147">Type</span></span>|<span data-ttu-id="d92fd-1148">属性</span><span class="sxs-lookup"><span data-stu-id="d92fd-1148">Attributes</span></span>|<span data-ttu-id="d92fd-1149">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-1149">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="d92fd-1150">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1150">Object</span></span>|<span data-ttu-id="d92fd-1151">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1151">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1152">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1152">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d92fd-1153">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1153">Object</span></span>|<span data-ttu-id="d92fd-1154">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1154">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1155">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1155">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d92fd-1156">function</span><span class="sxs-lookup"><span data-stu-id="d92fd-1156">function</span></span>|<span data-ttu-id="d92fd-1157">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1158">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1158">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-1159">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-1159">Requirements</span></span>

|<span data-ttu-id="d92fd-1160">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1160">Requirement</span></span>|<span data-ttu-id="d92fd-1161">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1161">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1162">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1162">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1163">1.8</span><span class="sxs-lookup"><span data-stu-id="d92fd-1163">1.8</span></span>|
|[<span data-ttu-id="d92fd-1164">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1164">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1165">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1165">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-1166">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1166">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1167">作成</span><span class="sxs-lookup"><span data-stu-id="d92fd-1167">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="d92fd-1168">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d92fd-1168">Returns:</span></span>

<span data-ttu-id="d92fd-1169">型: Array. <[attachmentdetails 詳細](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d92fd-1169">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="d92fd-1170">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1170">Example</span></span>

<span data-ttu-id="d92fd-1171">次の例では、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1171">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="d92fd-1172">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="d92fd-1172">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="d92fd-1173">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1173">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-1174">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1174">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-1175">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1175">Requirements</span></span>

|<span data-ttu-id="d92fd-1176">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1176">Requirement</span></span>|<span data-ttu-id="d92fd-1177">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1177">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1178">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1179">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-1179">1.0</span></span>|
|[<span data-ttu-id="d92fd-1180">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1180">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1181">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-1182">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1182">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1183">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-1183">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d92fd-1184">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d92fd-1184">Returns:</span></span>

<span data-ttu-id="d92fd-1185">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="d92fd-1185">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="d92fd-1186">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1186">Example</span></span>

<span data-ttu-id="d92fd-1187">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1187">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="d92fd-1188">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="d92fd-1188">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="d92fd-1189">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1189">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-1190">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1190">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-1191">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-1191">Parameters</span></span>

|<span data-ttu-id="d92fd-1192">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-1192">Name</span></span>|<span data-ttu-id="d92fd-1193">種類</span><span class="sxs-lookup"><span data-stu-id="d92fd-1193">Type</span></span>|<span data-ttu-id="d92fd-1194">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-1194">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="d92fd-1195">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="d92fd-1195">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="d92fd-1196">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1196">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-1197">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-1197">Requirements</span></span>

|<span data-ttu-id="d92fd-1198">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1198">Requirement</span></span>|<span data-ttu-id="d92fd-1199">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1199">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1200">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1201">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-1201">1.0</span></span>|
|[<span data-ttu-id="d92fd-1202">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1202">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1203">制限あり</span><span class="sxs-lookup"><span data-stu-id="d92fd-1203">Restricted</span></span>|
|[<span data-ttu-id="d92fd-1204">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1204">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1205">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-1205">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d92fd-1206">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d92fd-1206">Returns:</span></span>

<span data-ttu-id="d92fd-1207">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1207">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="d92fd-1208">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1208">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="d92fd-1209">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1209">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="d92fd-1210">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1210">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="d92fd-1211">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1211">Value of `entityType`</span></span>|<span data-ttu-id="d92fd-1212">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="d92fd-1212">Type of objects in returned array</span></span>|<span data-ttu-id="d92fd-1213">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1213">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="d92fd-1214">文字列</span><span class="sxs-lookup"><span data-stu-id="d92fd-1214">String</span></span>|<span data-ttu-id="d92fd-1215">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="d92fd-1215">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="d92fd-1216">連絡先</span><span class="sxs-lookup"><span data-stu-id="d92fd-1216">Contact</span></span>|<span data-ttu-id="d92fd-1217">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d92fd-1217">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="d92fd-1218">文字列</span><span class="sxs-lookup"><span data-stu-id="d92fd-1218">String</span></span>|<span data-ttu-id="d92fd-1219">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d92fd-1219">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="d92fd-1220">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="d92fd-1220">MeetingSuggestion</span></span>|<span data-ttu-id="d92fd-1221">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d92fd-1221">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="d92fd-1222">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="d92fd-1222">PhoneNumber</span></span>|<span data-ttu-id="d92fd-1223">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="d92fd-1223">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="d92fd-1224">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="d92fd-1224">TaskSuggestion</span></span>|<span data-ttu-id="d92fd-1225">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d92fd-1225">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="d92fd-1226">文字列</span><span class="sxs-lookup"><span data-stu-id="d92fd-1226">String</span></span>|<span data-ttu-id="d92fd-1227">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="d92fd-1227">**Restricted**</span></span>|

<span data-ttu-id="d92fd-1228">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="d92fd-1228">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="d92fd-1229">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1229">Example</span></span>

<span data-ttu-id="d92fd-1230">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1230">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="d92fd-1231">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="d92fd-1231">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="d92fd-1232">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1232">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-1233">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1233">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d92fd-1234">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1234">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-1235">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-1235">Parameters</span></span>

|<span data-ttu-id="d92fd-1236">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-1236">Name</span></span>|<span data-ttu-id="d92fd-1237">種類</span><span class="sxs-lookup"><span data-stu-id="d92fd-1237">Type</span></span>|<span data-ttu-id="d92fd-1238">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-1238">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="d92fd-1239">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-1239">String</span></span>|<span data-ttu-id="d92fd-1240">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1240">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-1241">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-1241">Requirements</span></span>

|<span data-ttu-id="d92fd-1242">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1242">Requirement</span></span>|<span data-ttu-id="d92fd-1243">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1243">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1244">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1244">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1245">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-1245">1.0</span></span>|
|[<span data-ttu-id="d92fd-1246">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1246">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1247">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1247">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-1248">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1248">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1249">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-1249">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d92fd-1250">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d92fd-1250">Returns:</span></span>

<span data-ttu-id="d92fd-p174">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p174">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="d92fd-1253">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="d92fd-1253">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

<br>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="d92fd-1254">、Office.context.mailbox.item.getinitializationcontextasync ([オプション], [callback])</span><span class="sxs-lookup"><span data-stu-id="d92fd-1254">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="d92fd-1255">[アクション可能なメッセージによってアドインがアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されたときに渡される初期化データを取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1255">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-1256">このメソッドは、Outlook 2016 以降の Windows (16.0.8413.1000 より後のバージョン) および Outlook on the Office 365 でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1256">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-1257">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-1257">Parameters</span></span>

|<span data-ttu-id="d92fd-1258">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-1258">Name</span></span>|<span data-ttu-id="d92fd-1259">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-1259">Type</span></span>|<span data-ttu-id="d92fd-1260">属性</span><span class="sxs-lookup"><span data-stu-id="d92fd-1260">Attributes</span></span>|<span data-ttu-id="d92fd-1261">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-1261">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="d92fd-1262">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1262">Object</span></span>|<span data-ttu-id="d92fd-1263">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1263">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1264">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1264">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d92fd-1265">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1265">Object</span></span>|<span data-ttu-id="d92fd-1266">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1266">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1267">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1267">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d92fd-1268">function</span><span class="sxs-lookup"><span data-stu-id="d92fd-1268">function</span></span>|<span data-ttu-id="d92fd-1269">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1269">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1270">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1270">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d92fd-1271">成功すると、初期化データが文字列とし`asyncResult.value`てプロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1271">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="d92fd-1272">初期化コンテキストがない場合、 `asyncResult`オブジェクトには、 `Error` `code`プロパティがに`9020`設定されたオブジェクトと`name`プロパティがに`GenericResponseError`設定されたオブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1272">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-1273">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-1273">Requirements</span></span>

|<span data-ttu-id="d92fd-1274">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1274">Requirement</span></span>|<span data-ttu-id="d92fd-1275">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1275">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1276">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1277">プレビュー</span><span class="sxs-lookup"><span data-stu-id="d92fd-1277">Preview</span></span>|
|[<span data-ttu-id="d92fd-1278">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1279">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-1280">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1281">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-1281">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-1282">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1282">Example</span></span>

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

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="d92fd-1283">getItemIdAsync ([オプション], callback)</span><span class="sxs-lookup"><span data-stu-id="d92fd-1283">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="d92fd-1284">保存されたアイテムの ID を非同期に取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1284">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="d92fd-1285">新規作成モードのみです。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1285">Compose mode only.</span></span>

<span data-ttu-id="d92fd-1286">このメソッドを呼び出すと、コールバックメソッドによってアイテム ID が返されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1286">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-1287">アドインが新規作成モードの`getItemIdAsync`アイテムに対して呼び出しを行う場合 ( `itemId` EWS または REST API を使用するため)、Outlook がキャッシュモードの場合は、アイテムがサーバーに同期されるまでしばらく時間がかかる場合があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1287">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="d92fd-1288">アイテムが同期されるまで、 `itemId`は認識されず、を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1288">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-1289">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-1289">Parameters</span></span>

|<span data-ttu-id="d92fd-1290">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-1290">Name</span></span>|<span data-ttu-id="d92fd-1291">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-1291">Type</span></span>|<span data-ttu-id="d92fd-1292">属性</span><span class="sxs-lookup"><span data-stu-id="d92fd-1292">Attributes</span></span>|<span data-ttu-id="d92fd-1293">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-1293">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="d92fd-1294">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1294">Object</span></span>|<span data-ttu-id="d92fd-1295">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1295">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1296">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1296">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d92fd-1297">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1297">Object</span></span>|<span data-ttu-id="d92fd-1298">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1298">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1299">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1299">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d92fd-1300">function</span><span class="sxs-lookup"><span data-stu-id="d92fd-1300">function</span></span>||<span data-ttu-id="d92fd-1301">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1301">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d92fd-1302">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1302">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d92fd-1303">エラー</span><span class="sxs-lookup"><span data-stu-id="d92fd-1303">Errors</span></span>

|<span data-ttu-id="d92fd-1304">エラー コード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1304">Error code</span></span>|<span data-ttu-id="d92fd-1305">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-1305">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="d92fd-1306">この id は、アイテムが保存されるまでは取得できません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1306">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-1307">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-1307">Requirements</span></span>

|<span data-ttu-id="d92fd-1308">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1308">Requirement</span></span>|<span data-ttu-id="d92fd-1309">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1309">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1310">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1310">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1311">1.8</span><span class="sxs-lookup"><span data-stu-id="d92fd-1311">1.8</span></span>|
|[<span data-ttu-id="d92fd-1312">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1312">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1313">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1313">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-1314">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1314">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1315">作成</span><span class="sxs-lookup"><span data-stu-id="d92fd-1315">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d92fd-1316">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1316">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="d92fd-1317">次の例は、コールバック関数`result`に渡されるパラメーターの構造を示しています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1317">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="d92fd-1318">プロパティ`value`には、アイテムの ID が含まれています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1318">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="d92fd-1319">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d92fd-1319">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="d92fd-1320">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1320">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-1321">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1321">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d92fd-p178">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p178">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d92fd-1325">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1325">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d92fd-1326">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1326">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="d92fd-p179">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-1330">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1330">Requirements</span></span>

|<span data-ttu-id="d92fd-1331">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1331">Requirement</span></span>|<span data-ttu-id="d92fd-1332">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1332">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1333">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1334">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-1334">1.0</span></span>|
|[<span data-ttu-id="d92fd-1335">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1336">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-1337">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1338">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-1338">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d92fd-1339">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d92fd-1339">Returns:</span></span>

<span data-ttu-id="d92fd-p180">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="d92fd-1342">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="d92fd-1342">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d92fd-1343">Object</span><span class="sxs-lookup"><span data-stu-id="d92fd-1343">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d92fd-1344">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1344">Example</span></span>

<span data-ttu-id="d92fd-1345">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1345">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="d92fd-1346">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="d92fd-1346">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="d92fd-1347">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1347">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-1348">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1348">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d92fd-1349">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1349">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="d92fd-p181">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p181">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-1352">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-1352">Parameters</span></span>

|<span data-ttu-id="d92fd-1353">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-1353">Name</span></span>|<span data-ttu-id="d92fd-1354">種類</span><span class="sxs-lookup"><span data-stu-id="d92fd-1354">Type</span></span>|<span data-ttu-id="d92fd-1355">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-1355">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="d92fd-1356">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-1356">String</span></span>|<span data-ttu-id="d92fd-1357">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1357">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-1358">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-1358">Requirements</span></span>

|<span data-ttu-id="d92fd-1359">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1359">Requirement</span></span>|<span data-ttu-id="d92fd-1360">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1360">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1361">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1362">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-1362">1.0</span></span>|
|[<span data-ttu-id="d92fd-1363">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1364">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-1365">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1366">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-1366">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d92fd-1367">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d92fd-1367">Returns:</span></span>

<span data-ttu-id="d92fd-1368">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1368">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="d92fd-1369">型: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="d92fd-1369">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="d92fd-1370">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1370">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="d92fd-1371">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="d92fd-1371">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="d92fd-1372">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1372">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="d92fd-p182">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p182">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-1375">Web 上の Outlook では、テキストが選択されておらず、カーソルが本文にある場合、このメソッドは文字列 "null" を返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1375">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="d92fd-1376">このような状況を確認するには、次のようなコードを含めます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1376">To check for this situation, include code similar to the following:</span></span>
>
> `var selectedText = (asyncResult.value.endPosition === asyncResult.value.startPosition) ? "" : asyncResult.value.data;`

##### <a name="parameters"></a><span data-ttu-id="d92fd-1377">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-1377">Parameters</span></span>

|<span data-ttu-id="d92fd-1378">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-1378">Name</span></span>|<span data-ttu-id="d92fd-1379">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-1379">Type</span></span>|<span data-ttu-id="d92fd-1380">属性</span><span class="sxs-lookup"><span data-stu-id="d92fd-1380">Attributes</span></span>|<span data-ttu-id="d92fd-1381">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-1381">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="d92fd-1382">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d92fd-1382">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="d92fd-p184">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p184">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="d92fd-1386">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1386">Object</span></span>|<span data-ttu-id="d92fd-1387">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1387">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1388">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1388">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d92fd-1389">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1389">Object</span></span>|<span data-ttu-id="d92fd-1390">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1390">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1391">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1391">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d92fd-1392">function</span><span class="sxs-lookup"><span data-stu-id="d92fd-1392">function</span></span>||<span data-ttu-id="d92fd-1393">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1393">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d92fd-1394">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1394">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="d92fd-1395">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1395">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-1396">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-1396">Requirements</span></span>

|<span data-ttu-id="d92fd-1397">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1397">Requirement</span></span>|<span data-ttu-id="d92fd-1398">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1398">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1399">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1399">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1400">1.2</span><span class="sxs-lookup"><span data-stu-id="d92fd-1400">1.2</span></span>|
|[<span data-ttu-id="d92fd-1401">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1401">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1402">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-1403">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1403">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1404">作成</span><span class="sxs-lookup"><span data-stu-id="d92fd-1404">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="d92fd-1405">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d92fd-1405">Returns:</span></span>

<span data-ttu-id="d92fd-1406">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1406">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="d92fd-1407">型:String</span><span class="sxs-lookup"><span data-stu-id="d92fd-1407">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d92fd-1408">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1408">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="d92fd-1409">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="d92fd-1409">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="d92fd-1410">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1410">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="d92fd-1411">強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1411">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-1412">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1412">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-1413">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1413">Requirements</span></span>

|<span data-ttu-id="d92fd-1414">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1414">Requirement</span></span>|<span data-ttu-id="d92fd-1415">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1415">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1416">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1416">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1417">1.6</span><span class="sxs-lookup"><span data-stu-id="d92fd-1417">1.6</span></span>|
|[<span data-ttu-id="d92fd-1418">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1418">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1419">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1419">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-1420">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1420">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1421">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-1421">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d92fd-1422">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d92fd-1422">Returns:</span></span>

<span data-ttu-id="d92fd-1423">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="d92fd-1423">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="d92fd-1424">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1424">Example</span></span>

<span data-ttu-id="d92fd-1425">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1425">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="d92fd-1426">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d92fd-1426">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="d92fd-p187">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p187">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-1429">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1429">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d92fd-p188">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p188">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d92fd-1433">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1433">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d92fd-1434">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1434">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="d92fd-p189">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p189">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d92fd-1438">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-1438">Requirements</span></span>

|<span data-ttu-id="d92fd-1439">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1439">Requirement</span></span>|<span data-ttu-id="d92fd-1440">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1440">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1441">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1441">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1442">1.6</span><span class="sxs-lookup"><span data-stu-id="d92fd-1442">1.6</span></span>|
|[<span data-ttu-id="d92fd-1443">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1443">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1444">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-1445">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1445">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1446">読み取り</span><span class="sxs-lookup"><span data-stu-id="d92fd-1446">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d92fd-1447">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d92fd-1447">Returns:</span></span>

<span data-ttu-id="d92fd-p190">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p190">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="d92fd-1450">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1450">Example</span></span>

<span data-ttu-id="d92fd-1451">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1451">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="d92fd-1452">getSharedPropertiesAsync ([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d92fd-1452">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="d92fd-1453">共有フォルダー、予定表、またはメールボックス内の選択した予定またはメッセージのプロパティを取得します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1453">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-1454">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-1454">Parameters</span></span>

|<span data-ttu-id="d92fd-1455">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-1455">Name</span></span>|<span data-ttu-id="d92fd-1456">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-1456">Type</span></span>|<span data-ttu-id="d92fd-1457">属性</span><span class="sxs-lookup"><span data-stu-id="d92fd-1457">Attributes</span></span>|<span data-ttu-id="d92fd-1458">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-1458">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="d92fd-1459">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1459">Object</span></span>|<span data-ttu-id="d92fd-1460">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1460">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1461">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1461">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d92fd-1462">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1462">Object</span></span>|<span data-ttu-id="d92fd-1463">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1463">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1464">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1464">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d92fd-1465">function</span><span class="sxs-lookup"><span data-stu-id="d92fd-1465">function</span></span>||<span data-ttu-id="d92fd-1466">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1466">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d92fd-1467">共有プロパティは、 [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) `asyncResult.value`プロパティのオブジェクトとして提供されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1467">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="d92fd-1468">このオブジェクトは、アイテムの共有プロパティを取得するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1468">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-1469">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-1469">Requirements</span></span>

|<span data-ttu-id="d92fd-1470">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1470">Requirement</span></span>|<span data-ttu-id="d92fd-1471">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1471">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1472">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1473">1.8</span><span class="sxs-lookup"><span data-stu-id="d92fd-1473">1.8</span></span>|
|[<span data-ttu-id="d92fd-1474">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1475">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-1476">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1477">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-1477">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-1478">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1478">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="d92fd-1479">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d92fd-1479">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="d92fd-1480">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1480">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="d92fd-p192">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p192">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-1484">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-1484">Parameters</span></span>

|<span data-ttu-id="d92fd-1485">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-1485">Name</span></span>|<span data-ttu-id="d92fd-1486">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-1486">Type</span></span>|<span data-ttu-id="d92fd-1487">属性</span><span class="sxs-lookup"><span data-stu-id="d92fd-1487">Attributes</span></span>|<span data-ttu-id="d92fd-1488">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-1488">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="d92fd-1489">function</span><span class="sxs-lookup"><span data-stu-id="d92fd-1489">function</span></span>||<span data-ttu-id="d92fd-1490">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1490">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d92fd-1491">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1491">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="d92fd-1492">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1492">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="d92fd-1493">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1493">Object</span></span>|<span data-ttu-id="d92fd-1494">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1494">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1495">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1495">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="d92fd-1496">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1496">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-1497">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-1497">Requirements</span></span>

|<span data-ttu-id="d92fd-1498">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1498">Requirement</span></span>|<span data-ttu-id="d92fd-1499">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1499">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1500">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1500">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1501">1.0</span><span class="sxs-lookup"><span data-stu-id="d92fd-1501">1.0</span></span>|
|[<span data-ttu-id="d92fd-1502">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1502">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1503">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1503">ReadItem</span></span>|
|[<span data-ttu-id="d92fd-1504">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1504">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1505">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-1505">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-1506">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1506">Example</span></span>

<span data-ttu-id="d92fd-p195">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p195">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="d92fd-1510">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d92fd-1510">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="d92fd-1511">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1511">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="d92fd-1512">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1512">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="d92fd-1513">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1513">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="d92fd-1514">Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1514">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="d92fd-1515">ユーザーがアプリを閉じたとき、またはインラインフォームの作成が開始されたときに、別のウィンドウで続行するためにフォームをポップアウトした後、セッションが終了します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1515">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-1516">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-1516">Parameters</span></span>

|<span data-ttu-id="d92fd-1517">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-1517">Name</span></span>|<span data-ttu-id="d92fd-1518">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-1518">Type</span></span>|<span data-ttu-id="d92fd-1519">属性</span><span class="sxs-lookup"><span data-stu-id="d92fd-1519">Attributes</span></span>|<span data-ttu-id="d92fd-1520">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-1520">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="d92fd-1521">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-1521">String</span></span>||<span data-ttu-id="d92fd-1522">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1522">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="d92fd-1523">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1523">Object</span></span>|<span data-ttu-id="d92fd-1524">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1524">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1525">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1525">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d92fd-1526">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1526">Object</span></span>|<span data-ttu-id="d92fd-1527">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1527">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1528">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1528">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d92fd-1529">function</span><span class="sxs-lookup"><span data-stu-id="d92fd-1529">function</span></span>|<span data-ttu-id="d92fd-1530">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1530">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1531">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1531">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d92fd-1532">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1532">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d92fd-1533">エラー</span><span class="sxs-lookup"><span data-stu-id="d92fd-1533">Errors</span></span>

|<span data-ttu-id="d92fd-1534">エラー コード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1534">Error code</span></span>|<span data-ttu-id="d92fd-1535">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-1535">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="d92fd-1536">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1536">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-1537">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-1537">Requirements</span></span>

|<span data-ttu-id="d92fd-1538">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1538">Requirement</span></span>|<span data-ttu-id="d92fd-1539">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1539">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1540">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1541">1.1</span><span class="sxs-lookup"><span data-stu-id="d92fd-1541">1.1</span></span>|
|[<span data-ttu-id="d92fd-1542">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1542">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1543">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1543">ReadWriteItem</span></span>|
|[<span data-ttu-id="d92fd-1544">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1544">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1545">作成</span><span class="sxs-lookup"><span data-stu-id="d92fd-1545">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-1546">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1546">Example</span></span>

<span data-ttu-id="d92fd-1547">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1547">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="d92fd-1548">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d92fd-1548">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="d92fd-1549">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1549">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="d92fd-1550">現在、サポートされて`Office.EventType.AttachmentsChanged`いる`Office.EventType.AppointmentTimeChanged`イベント`Office.EventType.EnhancedLocationsChanged`の`Office.EventType.RecipientsChanged`種類は`Office.EventType.RecurrenceChanged`、、、、、です。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1550">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-1551">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-1551">Parameters</span></span>

| <span data-ttu-id="d92fd-1552">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-1552">Name</span></span> | <span data-ttu-id="d92fd-1553">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-1553">Type</span></span> | <span data-ttu-id="d92fd-1554">属性</span><span class="sxs-lookup"><span data-stu-id="d92fd-1554">Attributes</span></span> | <span data-ttu-id="d92fd-1555">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-1555">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d92fd-1556">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d92fd-1556">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d92fd-1557">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1557">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="d92fd-1558">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1558">Object</span></span> | <span data-ttu-id="d92fd-1559">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1559">&lt;optional&gt;</span></span> | <span data-ttu-id="d92fd-1560">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1560">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d92fd-1561">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1561">Object</span></span> | <span data-ttu-id="d92fd-1562">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1562">&lt;optional&gt;</span></span> | <span data-ttu-id="d92fd-1563">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1563">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d92fd-1564">function</span><span class="sxs-lookup"><span data-stu-id="d92fd-1564">function</span></span>| <span data-ttu-id="d92fd-1565">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1565">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1566">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1566">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-1567">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1567">Requirements</span></span>

|<span data-ttu-id="d92fd-1568">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1568">Requirement</span></span>| <span data-ttu-id="d92fd-1569">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1569">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1570">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d92fd-1571">1.7</span><span class="sxs-lookup"><span data-stu-id="d92fd-1571">1.7</span></span> |
|[<span data-ttu-id="d92fd-1572">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1572">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d92fd-1573">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1573">ReadItem</span></span> |
|[<span data-ttu-id="d92fd-1574">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1574">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d92fd-1575">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d92fd-1575">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="d92fd-1576">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d92fd-1576">saveAsync([options], callback)</span></span>

<span data-ttu-id="d92fd-1577">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1577">Asynchronously saves an item.</span></span>

<span data-ttu-id="d92fd-1578">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1578">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="d92fd-1579">Outlook on the web またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1579">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="d92fd-1580">キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1580">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-1581">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1581">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="d92fd-1582">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1582">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="d92fd-p199">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p199">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="d92fd-1586">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1586">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="d92fd-1587">Outlook on Mac では、会議の保存はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1587">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="d92fd-1588">`saveAsync` メソッドは、作成モードの会議から呼び出されると失敗します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1588">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="d92fd-1589">回避策については、「[Office JS API を使用して Outlook for Mac で会議を下書きとして保存できない](https://support.microsoft.com/help/4505745)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1589">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="d92fd-1590">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1590">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-1591">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-1591">Parameters</span></span>

|<span data-ttu-id="d92fd-1592">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-1592">Name</span></span>|<span data-ttu-id="d92fd-1593">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-1593">Type</span></span>|<span data-ttu-id="d92fd-1594">属性</span><span class="sxs-lookup"><span data-stu-id="d92fd-1594">Attributes</span></span>|<span data-ttu-id="d92fd-1595">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-1595">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="d92fd-1596">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1596">Object</span></span>|<span data-ttu-id="d92fd-1597">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1597">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1598">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1598">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d92fd-1599">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1599">Object</span></span>|<span data-ttu-id="d92fd-1600">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1600">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1601">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1601">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d92fd-1602">関数</span><span class="sxs-lookup"><span data-stu-id="d92fd-1602">function</span></span>||<span data-ttu-id="d92fd-1603">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1603">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d92fd-1604">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1604">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-1605">Requirements</span><span class="sxs-lookup"><span data-stu-id="d92fd-1605">Requirements</span></span>

|<span data-ttu-id="d92fd-1606">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1606">Requirement</span></span>|<span data-ttu-id="d92fd-1607">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1607">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1608">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1609">1.3</span><span class="sxs-lookup"><span data-stu-id="d92fd-1609">1.3</span></span>|
|[<span data-ttu-id="d92fd-1610">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1610">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1611">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1611">ReadWriteItem</span></span>|
|[<span data-ttu-id="d92fd-1612">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1613">作成</span><span class="sxs-lookup"><span data-stu-id="d92fd-1613">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d92fd-1614">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1614">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="d92fd-p201">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p201">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="d92fd-1617">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="d92fd-1617">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="d92fd-1618">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1618">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="d92fd-p202">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p202">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d92fd-1622">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d92fd-1622">Parameters</span></span>

|<span data-ttu-id="d92fd-1623">名前</span><span class="sxs-lookup"><span data-stu-id="d92fd-1623">Name</span></span>|<span data-ttu-id="d92fd-1624">型</span><span class="sxs-lookup"><span data-stu-id="d92fd-1624">Type</span></span>|<span data-ttu-id="d92fd-1625">属性</span><span class="sxs-lookup"><span data-stu-id="d92fd-1625">Attributes</span></span>|<span data-ttu-id="d92fd-1626">説明</span><span class="sxs-lookup"><span data-stu-id="d92fd-1626">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="d92fd-1627">String</span><span class="sxs-lookup"><span data-stu-id="d92fd-1627">String</span></span>||<span data-ttu-id="d92fd-p203">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-p203">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="d92fd-1631">Object</span><span class="sxs-lookup"><span data-stu-id="d92fd-1631">Object</span></span>|<span data-ttu-id="d92fd-1632">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1632">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1633">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1633">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d92fd-1634">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d92fd-1634">Object</span></span>|<span data-ttu-id="d92fd-1635">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1635">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1636">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1636">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="d92fd-1637">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d92fd-1637">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="d92fd-1638">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d92fd-1638">&lt;optional&gt;</span></span>|<span data-ttu-id="d92fd-1639">`text` の場合、Outlook on the web とデスクトップ クライアントでは現在のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1639">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="d92fd-1640">フィールドが HTML エディターの場合、データが HTML の場合でもテキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1640">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="d92fd-1641">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Outlook on the web では現在のスタイルが適用され、Outlook デスクトップ クライアントでは既定のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1641">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="d92fd-1642">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1642">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="d92fd-1643">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1643">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="d92fd-1644">function</span><span class="sxs-lookup"><span data-stu-id="d92fd-1644">function</span></span>||<span data-ttu-id="d92fd-1645">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d92fd-1645">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d92fd-1646">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1646">Requirements</span></span>

|<span data-ttu-id="d92fd-1647">要件</span><span class="sxs-lookup"><span data-stu-id="d92fd-1647">Requirement</span></span>|<span data-ttu-id="d92fd-1648">値</span><span class="sxs-lookup"><span data-stu-id="d92fd-1648">Value</span></span>|
|---|---|
|[<span data-ttu-id="d92fd-1649">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d92fd-1649">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d92fd-1650">1.2</span><span class="sxs-lookup"><span data-stu-id="d92fd-1650">1.2</span></span>|
|[<span data-ttu-id="d92fd-1651">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d92fd-1651">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d92fd-1652">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d92fd-1652">ReadWriteItem</span></span>|
|[<span data-ttu-id="d92fd-1653">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d92fd-1653">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d92fd-1654">作成</span><span class="sxs-lookup"><span data-stu-id="d92fd-1654">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d92fd-1655">例</span><span class="sxs-lookup"><span data-stu-id="d92fd-1655">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
