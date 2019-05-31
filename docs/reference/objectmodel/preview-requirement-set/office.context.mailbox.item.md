---
title: Office. アイテム-プレビュー要件セット
description: ''
ms.date: 05/30/2019
localization_priority: Normal
ms.openlocfilehash: 12ec5d5558b558c87587e34472c33116478d14b3
ms.sourcegitcommit: b299b8a5dfffb6102cb14b431bdde4861abfb47f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/30/2019
ms.locfileid: "34589203"
---
# <a name="item"></a><span data-ttu-id="d36ad-102">item</span><span class="sxs-lookup"><span data-stu-id="d36ad-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="d36ad-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="d36ad-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="d36ad-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-106">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-106">Requirements</span></span>

|<span data-ttu-id="d36ad-107">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-107">Requirement</span></span>|<span data-ttu-id="d36ad-108">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-110">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-110">1.0</span></span>|
|[<span data-ttu-id="d36ad-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="d36ad-112">Restricted</span></span>|
|[<span data-ttu-id="d36ad-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d36ad-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-115">Members and methods</span></span>

| <span data-ttu-id="d36ad-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-116">Member</span></span> | <span data-ttu-id="d36ad-117">種類</span><span class="sxs-lookup"><span data-stu-id="d36ad-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d36ad-118">attachments</span><span class="sxs-lookup"><span data-stu-id="d36ad-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="d36ad-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-119">Member</span></span> |
| [<span data-ttu-id="d36ad-120">bcc</span><span class="sxs-lookup"><span data-stu-id="d36ad-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="d36ad-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-121">Member</span></span> |
| [<span data-ttu-id="d36ad-122">body</span><span class="sxs-lookup"><span data-stu-id="d36ad-122">body</span></span>](#body-body) | <span data-ttu-id="d36ad-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-123">Member</span></span> |
| [<span data-ttu-id="d36ad-124">categories</span><span class="sxs-lookup"><span data-stu-id="d36ad-124">categories</span></span>](#categories-categories) | <span data-ttu-id="d36ad-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-125">Member</span></span> |
| [<span data-ttu-id="d36ad-126">cc</span><span class="sxs-lookup"><span data-stu-id="d36ad-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d36ad-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-127">Member</span></span> |
| [<span data-ttu-id="d36ad-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="d36ad-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="d36ad-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-129">Member</span></span> |
| [<span data-ttu-id="d36ad-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="d36ad-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="d36ad-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-131">Member</span></span> |
| [<span data-ttu-id="d36ad-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="d36ad-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="d36ad-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-133">Member</span></span> |
| [<span data-ttu-id="d36ad-134">end</span><span class="sxs-lookup"><span data-stu-id="d36ad-134">end</span></span>](#end-datetime) | <span data-ttu-id="d36ad-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-135">Member</span></span> |
| [<span data-ttu-id="d36ad-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="d36ad-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="d36ad-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-137">Member</span></span> |
| [<span data-ttu-id="d36ad-138">from</span><span class="sxs-lookup"><span data-stu-id="d36ad-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="d36ad-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-139">Member</span></span> |
| [<span data-ttu-id="d36ad-140">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="d36ad-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="d36ad-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-141">Member</span></span> |
| [<span data-ttu-id="d36ad-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="d36ad-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="d36ad-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-143">Member</span></span> |
| [<span data-ttu-id="d36ad-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="d36ad-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="d36ad-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-145">Member</span></span> |
| [<span data-ttu-id="d36ad-146">itemId</span><span class="sxs-lookup"><span data-stu-id="d36ad-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="d36ad-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-147">Member</span></span> |
| [<span data-ttu-id="d36ad-148">itemType</span><span class="sxs-lookup"><span data-stu-id="d36ad-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="d36ad-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-149">Member</span></span> |
| [<span data-ttu-id="d36ad-150">location</span><span class="sxs-lookup"><span data-stu-id="d36ad-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="d36ad-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-151">Member</span></span> |
| [<span data-ttu-id="d36ad-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="d36ad-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="d36ad-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-153">Member</span></span> |
| [<span data-ttu-id="d36ad-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="d36ad-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="d36ad-155">Member</span><span class="sxs-lookup"><span data-stu-id="d36ad-155">Member</span></span> |
| [<span data-ttu-id="d36ad-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="d36ad-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d36ad-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-157">Member</span></span> |
| [<span data-ttu-id="d36ad-158">organizer</span><span class="sxs-lookup"><span data-stu-id="d36ad-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="d36ad-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-159">Member</span></span> |
| [<span data-ttu-id="d36ad-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="d36ad-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="d36ad-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-161">Member</span></span> |
| [<span data-ttu-id="d36ad-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="d36ad-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d36ad-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-163">Member</span></span> |
| [<span data-ttu-id="d36ad-164">sender</span><span class="sxs-lookup"><span data-stu-id="d36ad-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="d36ad-165">Member</span><span class="sxs-lookup"><span data-stu-id="d36ad-165">Member</span></span> |
| [<span data-ttu-id="d36ad-166">系列 Id</span><span class="sxs-lookup"><span data-stu-id="d36ad-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="d36ad-167">Member</span><span class="sxs-lookup"><span data-stu-id="d36ad-167">Member</span></span> |
| [<span data-ttu-id="d36ad-168">start</span><span class="sxs-lookup"><span data-stu-id="d36ad-168">start</span></span>](#start-datetime) | <span data-ttu-id="d36ad-169">Member</span><span class="sxs-lookup"><span data-stu-id="d36ad-169">Member</span></span> |
| [<span data-ttu-id="d36ad-170">subject</span><span class="sxs-lookup"><span data-stu-id="d36ad-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="d36ad-171">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-171">Member</span></span> |
| [<span data-ttu-id="d36ad-172">to</span><span class="sxs-lookup"><span data-stu-id="d36ad-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d36ad-173">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-173">Member</span></span> |
| [<span data-ttu-id="d36ad-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d36ad-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="d36ad-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-175">Method</span></span> |
| [<span data-ttu-id="d36ad-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="d36ad-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="d36ad-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-177">Method</span></span> |
| [<span data-ttu-id="d36ad-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d36ad-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="d36ad-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-179">Method</span></span> |
| [<span data-ttu-id="d36ad-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d36ad-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="d36ad-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-181">Method</span></span> |
| [<span data-ttu-id="d36ad-182">close</span><span class="sxs-lookup"><span data-stu-id="d36ad-182">close</span></span>](#close) | <span data-ttu-id="d36ad-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-183">Method</span></span> |
| [<span data-ttu-id="d36ad-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="d36ad-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="d36ad-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-185">Method</span></span> |
| [<span data-ttu-id="d36ad-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="d36ad-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="d36ad-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-187">Method</span></span> |
| [<span data-ttu-id="d36ad-188">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="d36ad-188">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="d36ad-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-189">Method</span></span> |
| [<span data-ttu-id="d36ad-190">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="d36ad-190">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="d36ad-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-191">Method</span></span> |
| [<span data-ttu-id="d36ad-192">getEntities</span><span class="sxs-lookup"><span data-stu-id="d36ad-192">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="d36ad-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-193">Method</span></span> |
| [<span data-ttu-id="d36ad-194">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="d36ad-194">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="d36ad-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-195">Method</span></span> |
| [<span data-ttu-id="d36ad-196">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="d36ad-196">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="d36ad-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-197">Method</span></span> |
| [<span data-ttu-id="d36ad-198">、Office.context.mailbox.item.getinitializationcontextasync</span><span class="sxs-lookup"><span data-stu-id="d36ad-198">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="d36ad-199">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-199">Method</span></span> |
| [<span data-ttu-id="d36ad-200">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="d36ad-200">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="d36ad-201">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-201">Method</span></span> |
| [<span data-ttu-id="d36ad-202">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="d36ad-202">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="d36ad-203">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-203">Method</span></span> |
| [<span data-ttu-id="d36ad-204">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="d36ad-204">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="d36ad-205">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-205">Method</span></span> |
| [<span data-ttu-id="d36ad-206">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="d36ad-206">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="d36ad-207">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-207">Method</span></span> |
| [<span data-ttu-id="d36ad-208">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="d36ad-208">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="d36ad-209">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-209">Method</span></span> |
| [<span data-ttu-id="d36ad-210">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="d36ad-210">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="d36ad-211">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-211">Method</span></span> |
| [<span data-ttu-id="d36ad-212">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="d36ad-212">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="d36ad-213">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-213">Method</span></span> |
| [<span data-ttu-id="d36ad-214">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d36ad-214">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="d36ad-215">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-215">Method</span></span> |
| [<span data-ttu-id="d36ad-216">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d36ad-216">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="d36ad-217">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-217">Method</span></span> |
| [<span data-ttu-id="d36ad-218">saveAsync</span><span class="sxs-lookup"><span data-stu-id="d36ad-218">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="d36ad-219">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-219">Method</span></span> |
| [<span data-ttu-id="d36ad-220">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="d36ad-220">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="d36ad-221">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-221">Method</span></span> |

### <a name="example"></a><span data-ttu-id="d36ad-222">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-222">Example</span></span>

<span data-ttu-id="d36ad-223">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d36ad-223">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="d36ad-224">メンバー</span><span class="sxs-lookup"><span data-stu-id="d36ad-224">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="d36ad-225">添付ファイル: <[Attachmentdetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d36ad-225">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="d36ad-226">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-226">Gets the item's attachments as an array.</span></span> <span data-ttu-id="d36ad-227">Read mode only.</span><span class="sxs-lookup"><span data-stu-id="d36ad-227">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-228">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-228">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="d36ad-229">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d36ad-229">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="d36ad-230">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-230">Type</span></span>

*   <span data-ttu-id="d36ad-231">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d36ad-231">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-232">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-232">Requirements</span></span>

|<span data-ttu-id="d36ad-233">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-233">Requirement</span></span>|<span data-ttu-id="d36ad-234">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-235">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-236">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-236">1.0</span></span>|
|[<span data-ttu-id="d36ad-237">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-237">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-238">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-239">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-239">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-240">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-240">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-241">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-241">Example</span></span>

<span data-ttu-id="d36ad-242">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-242">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="d36ad-243">bcc:[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d36ad-243">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="d36ad-244">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-244">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="d36ad-245">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d36ad-245">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d36ad-246">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-246">Type</span></span>

*   [<span data-ttu-id="d36ad-247">受信者</span><span class="sxs-lookup"><span data-stu-id="d36ad-247">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="d36ad-248">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-248">Requirements</span></span>

|<span data-ttu-id="d36ad-249">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-249">Requirement</span></span>|<span data-ttu-id="d36ad-250">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-251">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-252">1.1</span><span class="sxs-lookup"><span data-stu-id="d36ad-252">1.1</span></span>|
|[<span data-ttu-id="d36ad-253">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-254">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-255">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-256">作成</span><span class="sxs-lookup"><span data-stu-id="d36ad-256">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-257">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-257">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="d36ad-258">本文:[本文](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="d36ad-258">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="d36ad-259">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-259">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d36ad-260">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-260">Type</span></span>

*   [<span data-ttu-id="d36ad-261">Body</span><span class="sxs-lookup"><span data-stu-id="d36ad-261">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="d36ad-262">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-262">Requirements</span></span>

|<span data-ttu-id="d36ad-263">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-263">Requirement</span></span>|<span data-ttu-id="d36ad-264">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-265">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-266">1.1</span><span class="sxs-lookup"><span data-stu-id="d36ad-266">1.1</span></span>|
|[<span data-ttu-id="d36ad-267">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-267">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-268">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-269">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-269">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-270">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-270">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-271">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-271">Example</span></span>

<span data-ttu-id="d36ad-272">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-272">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="d36ad-273">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-273">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

---
---

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="d36ad-274">カテゴリ:[カテゴリ](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="d36ad-274">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="d36ad-275">アイテムのカテゴリを管理するためのメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-275">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-276">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-276">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="d36ad-277">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-277">Type</span></span>

*   [<span data-ttu-id="d36ad-278">カテゴリ</span><span class="sxs-lookup"><span data-stu-id="d36ad-278">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="d36ad-279">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-279">Requirements</span></span>

|<span data-ttu-id="d36ad-280">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-280">Requirement</span></span>|<span data-ttu-id="d36ad-281">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-282">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-283">プレビュー</span><span class="sxs-lookup"><span data-stu-id="d36ad-283">Preview</span></span>|
|[<span data-ttu-id="d36ad-284">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-285">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-286">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-287">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-287">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-288">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-288">Example</span></span>

<span data-ttu-id="d36ad-289">この例では、アイテムのカテゴリを取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-289">This example gets the item's categories.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="d36ad-290">cc: <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)>|[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d36ad-290">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="d36ad-291">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-291">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="d36ad-292">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-292">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d36ad-293">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-293">Read mode</span></span>

<span data-ttu-id="d36ad-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="d36ad-296">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-296">Compose mode</span></span>

<span data-ttu-id="d36ad-297">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-297">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d36ad-298">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-298">Type</span></span>

*   <span data-ttu-id="d36ad-299">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d36ad-299">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-300">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-300">Requirements</span></span>

|<span data-ttu-id="d36ad-301">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-301">Requirement</span></span>|<span data-ttu-id="d36ad-302">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-303">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-304">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-304">1.0</span></span>|
|[<span data-ttu-id="d36ad-305">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-305">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-306">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-307">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-307">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-308">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-308">Compose or Read</span></span>|

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="d36ad-309">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="d36ad-309">(nullable) conversationId: String</span></span>

<span data-ttu-id="d36ad-310">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-310">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="d36ad-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="d36ad-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="d36ad-315">Type</span><span class="sxs-lookup"><span data-stu-id="d36ad-315">Type</span></span>

*   <span data-ttu-id="d36ad-316">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-316">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-317">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-317">Requirements</span></span>

|<span data-ttu-id="d36ad-318">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-318">Requirement</span></span>|<span data-ttu-id="d36ad-319">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-320">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-321">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-321">1.0</span></span>|
|[<span data-ttu-id="d36ad-322">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-323">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-324">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-325">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-325">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-326">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-326">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="d36ad-327">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="d36ad-327">dateTimeCreated: Date</span></span>

<span data-ttu-id="d36ad-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d36ad-330">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-330">Type</span></span>

*   <span data-ttu-id="d36ad-331">日付</span><span class="sxs-lookup"><span data-stu-id="d36ad-331">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-332">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-332">Requirements</span></span>

|<span data-ttu-id="d36ad-333">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-333">Requirement</span></span>|<span data-ttu-id="d36ad-334">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-334">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-335">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-335">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-336">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-336">1.0</span></span>|
|[<span data-ttu-id="d36ad-337">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-337">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-338">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-338">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-339">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-339">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-340">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-340">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-341">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-341">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="d36ad-342">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="d36ad-342">dateTimeModified: Date</span></span>

<span data-ttu-id="d36ad-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-345">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-345">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="d36ad-346">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-346">Type</span></span>

*   <span data-ttu-id="d36ad-347">日付</span><span class="sxs-lookup"><span data-stu-id="d36ad-347">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-348">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-348">Requirements</span></span>

|<span data-ttu-id="d36ad-349">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-349">Requirement</span></span>|<span data-ttu-id="d36ad-350">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-351">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-352">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-352">1.0</span></span>|
|[<span data-ttu-id="d36ad-353">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-354">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-355">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-356">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-356">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-357">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-357">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="d36ad-358">終了: 日付 |[時間](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="d36ad-358">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="d36ad-359">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-359">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="d36ad-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d36ad-362">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-362">Read mode</span></span>

<span data-ttu-id="d36ad-363">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-363">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="d36ad-364">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-364">Compose mode</span></span>

<span data-ttu-id="d36ad-365">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-365">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="d36ad-366">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-366">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="d36ad-367">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-367">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="d36ad-368">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-368">Type</span></span>

*   <span data-ttu-id="d36ad-369">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="d36ad-369">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-370">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-370">Requirements</span></span>

|<span data-ttu-id="d36ad-371">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-371">Requirement</span></span>|<span data-ttu-id="d36ad-372">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-372">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-373">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-373">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-374">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-374">1.0</span></span>|
|[<span data-ttu-id="d36ad-375">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-375">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-376">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-376">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-377">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-377">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-378">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-378">Compose or Read</span></span>|

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="d36ad-379">enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="d36ad-379">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="d36ad-380">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-380">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d36ad-381">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-381">Read mode</span></span>

<span data-ttu-id="d36ad-382">この`enhancedLocation`プロパティは、予定に関連付けられている場所 ( [locationdetails](/javascript/api/outlook/office.locationdetails)オブジェクトで表される) のセットを取得できる[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-382">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d36ad-383">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-383">Compose mode</span></span>

<span data-ttu-id="d36ad-384">この`enhancedLocation`プロパティは、予定の場所を取得、削除、または追加するためのメソッドを提供する[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-384">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d36ad-385">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-385">Type</span></span>

*   [<span data-ttu-id="d36ad-386">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="d36ad-386">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="d36ad-387">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-387">Requirements</span></span>

|<span data-ttu-id="d36ad-388">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-388">Requirement</span></span>|<span data-ttu-id="d36ad-389">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-390">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-391">プレビュー</span><span class="sxs-lookup"><span data-stu-id="d36ad-391">Preview</span></span>|
|[<span data-ttu-id="d36ad-392">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-393">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-394">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-395">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-395">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-396">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-396">Example</span></span>

<span data-ttu-id="d36ad-397">次の例では、予定に関連付けられている現在の場所を取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-397">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="d36ad-398">from: [emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)|[from](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="d36ad-398">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="d36ad-399">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-399">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="d36ad-p112">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-402">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-402">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d36ad-403">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-403">Read mode</span></span>

<span data-ttu-id="d36ad-404">プロパティ`from`は`EmailAddressDetails`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-404">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="d36ad-405">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-405">Compose mode</span></span>

<span data-ttu-id="d36ad-406">プロパティ`from`は、from `From`値を取得するメソッドを提供するオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-406">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d36ad-407">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-407">Type</span></span>

*   <span data-ttu-id="d36ad-408">[電子メールアドレス](/javascript/api/outlook/office.emailaddressdetails) | [の](/javascript/api/outlook/office.from)詳細</span><span class="sxs-lookup"><span data-stu-id="d36ad-408">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-409">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-409">Requirements</span></span>

|<span data-ttu-id="d36ad-410">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-410">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="d36ad-411">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-412">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-412">1.0</span></span>|<span data-ttu-id="d36ad-413">1.7</span><span class="sxs-lookup"><span data-stu-id="d36ad-413">1.7</span></span>|
|[<span data-ttu-id="d36ad-414">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-415">ReadItem</span></span>|<span data-ttu-id="d36ad-416">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-416">ReadWriteItem</span></span>|
|[<span data-ttu-id="d36ad-417">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-418">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-418">Read</span></span>|<span data-ttu-id="d36ad-419">作成</span><span class="sxs-lookup"><span data-stu-id="d36ad-419">Compose</span></span>|

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="d36ad-420">internetHeaders: [internetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="d36ad-420">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="d36ad-421">メッセージのインターネットヘッダーを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-421">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="d36ad-422">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-422">Type</span></span>

*   [<span data-ttu-id="d36ad-423">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="d36ad-423">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="d36ad-424">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-424">Requirements</span></span>

|<span data-ttu-id="d36ad-425">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-425">Requirement</span></span>|<span data-ttu-id="d36ad-426">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-426">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-427">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-427">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-428">プレビュー</span><span class="sxs-lookup"><span data-stu-id="d36ad-428">Preview</span></span>|
|[<span data-ttu-id="d36ad-429">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-429">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-430">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-430">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-431">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-431">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-432">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-432">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-433">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-433">Example</span></span>

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="d36ad-434">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="d36ad-434">internetMessageId: String</span></span>

<span data-ttu-id="d36ad-p113">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d36ad-437">Type</span><span class="sxs-lookup"><span data-stu-id="d36ad-437">Type</span></span>

*   <span data-ttu-id="d36ad-438">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-438">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-439">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-439">Requirements</span></span>

|<span data-ttu-id="d36ad-440">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-440">Requirement</span></span>|<span data-ttu-id="d36ad-441">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-442">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-443">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-443">1.0</span></span>|
|[<span data-ttu-id="d36ad-444">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-445">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-446">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-447">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-447">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-448">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-448">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="d36ad-449">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="d36ad-449">itemClass: String</span></span>

<span data-ttu-id="d36ad-p114">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="d36ad-p115">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="d36ad-454">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-454">Type</span></span>|<span data-ttu-id="d36ad-455">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-455">Description</span></span>|<span data-ttu-id="d36ad-456">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="d36ad-456">item class</span></span>|
|---|---|---|
|<span data-ttu-id="d36ad-457">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="d36ad-457">Appointment items</span></span>|<span data-ttu-id="d36ad-458">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="d36ad-458">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="d36ad-459">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="d36ad-459">Message items</span></span>|<span data-ttu-id="d36ad-460">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-460">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="d36ad-461">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-461">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="d36ad-462">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-462">Type</span></span>

*   <span data-ttu-id="d36ad-463">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-463">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-464">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-464">Requirements</span></span>

|<span data-ttu-id="d36ad-465">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-465">Requirement</span></span>|<span data-ttu-id="d36ad-466">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-467">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-468">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-468">1.0</span></span>|
|[<span data-ttu-id="d36ad-469">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-470">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-471">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-472">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-473">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-473">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="d36ad-474">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="d36ad-474">(nullable) itemId: String</span></span>

<span data-ttu-id="d36ad-p116">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-477">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="d36ad-477">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="d36ad-478">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-478">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="d36ad-479">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-479">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="d36ad-480">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d36ad-480">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="d36ad-p118">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="d36ad-483">Type</span><span class="sxs-lookup"><span data-stu-id="d36ad-483">Type</span></span>

*   <span data-ttu-id="d36ad-484">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-484">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-485">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-485">Requirements</span></span>

|<span data-ttu-id="d36ad-486">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-486">Requirement</span></span>|<span data-ttu-id="d36ad-487">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-488">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-489">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-489">1.0</span></span>|
|[<span data-ttu-id="d36ad-490">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-491">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-492">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-493">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-493">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-494">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-494">Example</span></span>

<span data-ttu-id="d36ad-p119">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="d36ad-497">itemType: [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="d36ad-497">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="d36ad-498">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-498">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="d36ad-499">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="d36ad-499">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d36ad-500">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-500">Type</span></span>

*   [<span data-ttu-id="d36ad-501">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="d36ad-501">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="d36ad-502">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-502">Requirements</span></span>

|<span data-ttu-id="d36ad-503">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-503">Requirement</span></span>|<span data-ttu-id="d36ad-504">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-505">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-506">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-506">1.0</span></span>|
|[<span data-ttu-id="d36ad-507">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-508">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-509">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-510">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-510">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-511">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-511">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="d36ad-512">場所: String |[場所](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="d36ad-512">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="d36ad-513">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-513">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d36ad-514">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-514">Read mode</span></span>

<span data-ttu-id="d36ad-515">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-515">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="d36ad-516">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-516">Compose mode</span></span>

<span data-ttu-id="d36ad-517">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-517">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d36ad-518">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-518">Type</span></span>

*   <span data-ttu-id="d36ad-519">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="d36ad-519">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-520">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-520">Requirements</span></span>

|<span data-ttu-id="d36ad-521">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-521">Requirement</span></span>|<span data-ttu-id="d36ad-522">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-522">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-523">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-524">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-524">1.0</span></span>|
|[<span data-ttu-id="d36ad-525">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-525">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-526">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-526">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-527">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-527">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-528">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-528">Compose or Read</span></span>|

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="d36ad-529">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="d36ad-529">normalizedSubject: String</span></span>

<span data-ttu-id="d36ad-p120">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="d36ad-p121">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="d36ad-534">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-534">Type</span></span>

*   <span data-ttu-id="d36ad-535">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-535">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-536">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-536">Requirements</span></span>

|<span data-ttu-id="d36ad-537">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-537">Requirement</span></span>|<span data-ttu-id="d36ad-538">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-539">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-540">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-540">1.0</span></span>|
|[<span data-ttu-id="d36ad-541">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-542">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-543">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-544">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-544">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-545">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-545">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="d36ad-546">notificationMessages: [Notificationmessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="d36ad-546">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="d36ad-547">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-547">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d36ad-548">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-548">Type</span></span>

*   [<span data-ttu-id="d36ad-549">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="d36ad-549">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="d36ad-550">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-550">Requirements</span></span>

|<span data-ttu-id="d36ad-551">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-551">Requirement</span></span>|<span data-ttu-id="d36ad-552">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-552">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-553">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-553">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-554">1.3</span><span class="sxs-lookup"><span data-stu-id="d36ad-554">1.3</span></span>|
|[<span data-ttu-id="d36ad-555">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-555">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-556">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-556">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-557">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-557">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-558">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-558">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-559">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-559">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="d36ad-560">任意出席者: <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)>|[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d36ad-560">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="d36ad-561">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-561">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="d36ad-562">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-562">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d36ad-563">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-563">Read mode</span></span>

<span data-ttu-id="d36ad-564">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-564">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="d36ad-565">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-565">Compose mode</span></span>

<span data-ttu-id="d36ad-566">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-566">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d36ad-567">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-567">Type</span></span>

*   <span data-ttu-id="d36ad-568">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d36ad-568">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-569">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-569">Requirements</span></span>

|<span data-ttu-id="d36ad-570">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-570">Requirement</span></span>|<span data-ttu-id="d36ad-571">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-571">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-572">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-572">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-573">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-573">1.0</span></span>|
|[<span data-ttu-id="d36ad-574">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-574">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-575">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-575">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-576">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-576">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-577">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-577">Compose or Read</span></span>|

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="d36ad-578">開催者: [emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)|[開催者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="d36ad-578">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="d36ad-579">指定した会議の開催者の電子メールアドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-579">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d36ad-580">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-580">Read mode</span></span>

<span data-ttu-id="d36ad-581">プロパティ`organizer`は、会議の開催者を表す[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-581">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="d36ad-582">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-582">Compose mode</span></span>

<span data-ttu-id="d36ad-583">プロパティ`organizer`は、開催者の値を取得するためのメソッドを提供する[オーガナイザー](/javascript/api/outlook/office.organizer)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-583">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="d36ad-584">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-584">Type</span></span>

*   <span data-ttu-id="d36ad-585">[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails) | [開催者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="d36ad-585">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-586">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-586">Requirements</span></span>

|<span data-ttu-id="d36ad-587">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-587">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="d36ad-588">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-589">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-589">1.0</span></span>|<span data-ttu-id="d36ad-590">1.7</span><span class="sxs-lookup"><span data-stu-id="d36ad-590">1.7</span></span>|
|[<span data-ttu-id="d36ad-591">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-591">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-592">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-592">ReadItem</span></span>|<span data-ttu-id="d36ad-593">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-593">ReadWriteItem</span></span>|
|[<span data-ttu-id="d36ad-594">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-594">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-595">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-595">Read</span></span>|<span data-ttu-id="d36ad-596">作成</span><span class="sxs-lookup"><span data-stu-id="d36ad-596">Compose</span></span>|

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="d36ad-597">(nullable) 定期的なスケジュール:[定期的](/javascript/api/outlook/office.recurrence)なアイテム</span><span class="sxs-lookup"><span data-stu-id="d36ad-597">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="d36ad-598">予定の定期的なパターンを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-598">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="d36ad-599">会議出席依頼の定期的なパターンを取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-599">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="d36ad-600">予定アイテムの読み取りおよび作成モード。</span><span class="sxs-lookup"><span data-stu-id="d36ad-600">Read and compose modes for appointment items.</span></span> <span data-ttu-id="d36ad-601">会議出席依頼アイテムの閲覧モード。</span><span class="sxs-lookup"><span data-stu-id="d36ad-601">Read mode for meeting request items.</span></span>

<span data-ttu-id="d36ad-602">この`recurrence`プロパティは、アイテムが series または series 内のインスタンスの場合、定期的な予定または会議出席依頼に対して[定期的](/javascript/api/outlook/office.recurrence)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-602">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="d36ad-603">`null`は、単一の予定および1つの予定の会議出席依頼に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-603">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="d36ad-604">`undefined`は、会議出席依頼ではないメッセージに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-604">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="d36ad-605">注: 会議出席依頼に`itemClass`は、IPM という値があります。出席依頼。</span><span class="sxs-lookup"><span data-stu-id="d36ad-605">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="d36ad-606">注: 定期的なオブジェクトが`null`の場合は、そのオブジェクトが単一の予定または1つの予定の会議出席依頼であり、データ系列の一部ではないことを示します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-606">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d36ad-607">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-607">Read mode</span></span>

<span data-ttu-id="d36ad-608">この`recurrence`プロパティは、定期的な予定を表す[定期的](/javascript/api/outlook/office.recurrence)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-608">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="d36ad-609">これは、予定および会議出席依頼に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-609">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="d36ad-610">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-610">Compose mode</span></span>

<span data-ttu-id="d36ad-611">この`recurrence`プロパティは、予定の繰り返しを管理するためのメソッドを提供する[定期的](/javascript/api/outlook/office.recurrence)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-611">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="d36ad-612">これは予定に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-612">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="d36ad-613">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-613">Type</span></span>

* [<span data-ttu-id="d36ad-614">繰り返さ</span><span class="sxs-lookup"><span data-stu-id="d36ad-614">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="d36ad-615">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-615">Requirement</span></span>|<span data-ttu-id="d36ad-616">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-616">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-617">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-617">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-618">1.7</span><span class="sxs-lookup"><span data-stu-id="d36ad-618">1.7</span></span>|
|[<span data-ttu-id="d36ad-619">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-619">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-620">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-620">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-621">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-621">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-622">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-622">Compose or Read</span></span>|

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="d36ad-623">requiredat< dees: 配列.[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)>|[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d36ad-623">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="d36ad-624">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-624">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="d36ad-625">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-625">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d36ad-626">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-626">Read mode</span></span>

<span data-ttu-id="d36ad-627">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-627">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="d36ad-628">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-628">Compose mode</span></span>

<span data-ttu-id="d36ad-629">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-629">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="d36ad-630">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-630">Type</span></span>

*   <span data-ttu-id="d36ad-631">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d36ad-631">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-632">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-632">Requirements</span></span>

|<span data-ttu-id="d36ad-633">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-633">Requirement</span></span>|<span data-ttu-id="d36ad-634">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-635">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-636">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-636">1.0</span></span>|
|[<span data-ttu-id="d36ad-637">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-638">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-639">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-640">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-640">Compose or Read</span></span>|

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="d36ad-641">sender: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="d36ad-641">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="d36ad-p128">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="d36ad-p129">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsfrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-646">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-646">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="d36ad-647">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-647">Type</span></span>

*   [<span data-ttu-id="d36ad-648">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d36ad-648">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="d36ad-649">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-649">Requirements</span></span>

|<span data-ttu-id="d36ad-650">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-650">Requirement</span></span>|<span data-ttu-id="d36ad-651">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-651">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-652">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-652">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-653">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-653">1.0</span></span>|
|[<span data-ttu-id="d36ad-654">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-654">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-655">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-655">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-656">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-656">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-657">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-657">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-658">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-658">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="d36ad-659">(nullable) 系列 Id: String</span><span class="sxs-lookup"><span data-stu-id="d36ad-659">(nullable) seriesId: String</span></span>

<span data-ttu-id="d36ad-660">インスタンスが属する系列の id を取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-660">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="d36ad-661">OWA および Outlook で、は`seriesId` 、このアイテムが属する親 (シリーズ) アイテムの Exchange Web サービス (EWS) ID を返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-661">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="d36ad-662">ただし、iOS と Android では、 `seriesId`は親アイテムの REST ID を返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-662">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-663">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="d36ad-663">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="d36ad-664">`seriesId`プロパティが OUTLOOK REST API で使用される outlook id と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-664">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="d36ad-665">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-665">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="d36ad-666">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d36ad-666">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="d36ad-667">この`seriesId`プロパティは`null` 、単一の予定、系列のアイテム、会議出席依頼などの親アイテムを持たないアイテムに`undefined`対して、会議出席依頼以外のアイテムに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-667">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="d36ad-668">Type</span><span class="sxs-lookup"><span data-stu-id="d36ad-668">Type</span></span>

* <span data-ttu-id="d36ad-669">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-669">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-670">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-670">Requirements</span></span>

|<span data-ttu-id="d36ad-671">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-671">Requirement</span></span>|<span data-ttu-id="d36ad-672">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-673">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-674">1.7</span><span class="sxs-lookup"><span data-stu-id="d36ad-674">1.7</span></span>|
|[<span data-ttu-id="d36ad-675">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-675">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-676">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-677">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-677">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-678">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-678">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-679">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-679">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="d36ad-680">開始: 日付 |[時間](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="d36ad-680">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="d36ad-681">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-681">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="d36ad-p132">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d36ad-684">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-684">Read mode</span></span>

<span data-ttu-id="d36ad-685">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-685">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="d36ad-686">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-686">Compose mode</span></span>

<span data-ttu-id="d36ad-687">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-687">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="d36ad-688">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-688">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="d36ad-689">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-689">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="d36ad-690">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-690">Type</span></span>

*   <span data-ttu-id="d36ad-691">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="d36ad-691">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-692">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-692">Requirements</span></span>

|<span data-ttu-id="d36ad-693">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-693">Requirement</span></span>|<span data-ttu-id="d36ad-694">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-695">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-696">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-696">1.0</span></span>|
|[<span data-ttu-id="d36ad-697">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-697">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-698">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-698">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-699">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-699">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-700">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-700">Compose or Read</span></span>|

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="d36ad-701">subject: String |[件名](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="d36ad-701">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="d36ad-702">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-702">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="d36ad-703">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-703">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d36ad-704">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-704">Read mode</span></span>

<span data-ttu-id="d36ad-p133">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="d36ad-707">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d36ad-707">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="d36ad-708">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-708">Compose mode</span></span>
<span data-ttu-id="d36ad-709">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-709">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="d36ad-710">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-710">Type</span></span>

*   <span data-ttu-id="d36ad-711">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="d36ad-711">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-712">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-712">Requirements</span></span>

|<span data-ttu-id="d36ad-713">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-713">Requirement</span></span>|<span data-ttu-id="d36ad-714">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-715">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-715">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-716">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-716">1.0</span></span>|
|[<span data-ttu-id="d36ad-717">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-717">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-718">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-718">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-719">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-719">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-720">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-720">Compose or Read</span></span>|

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="d36ad-721">宛先: <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d36ad-721">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="d36ad-722">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-722">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="d36ad-723">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-723">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d36ad-724">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-724">Read mode</span></span>

<span data-ttu-id="d36ad-p135">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="d36ad-727">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d36ad-727">Compose mode</span></span>

<span data-ttu-id="d36ad-728">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-728">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d36ad-729">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-729">Type</span></span>

*   <span data-ttu-id="d36ad-730">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d36ad-730">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-731">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-731">Requirements</span></span>

|<span data-ttu-id="d36ad-732">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-732">Requirement</span></span>|<span data-ttu-id="d36ad-733">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-734">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-735">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-735">1.0</span></span>|
|[<span data-ttu-id="d36ad-736">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-736">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-737">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-737">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-738">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-738">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-739">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-739">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="d36ad-740">メソッド</span><span class="sxs-lookup"><span data-stu-id="d36ad-740">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="d36ad-741">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d36ad-741">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d36ad-742">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-742">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="d36ad-743">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-743">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="d36ad-744">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-744">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-745">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-745">Parameters</span></span>
|<span data-ttu-id="d36ad-746">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-746">Name</span></span>|<span data-ttu-id="d36ad-747">種類</span><span class="sxs-lookup"><span data-stu-id="d36ad-747">Type</span></span>|<span data-ttu-id="d36ad-748">属性</span><span class="sxs-lookup"><span data-stu-id="d36ad-748">Attributes</span></span>|<span data-ttu-id="d36ad-749">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-749">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="d36ad-750">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-750">String</span></span>||<span data-ttu-id="d36ad-p136">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="d36ad-753">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-753">String</span></span>||<span data-ttu-id="d36ad-p137">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="d36ad-756">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-756">Object</span></span>|<span data-ttu-id="d36ad-757">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-757">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-758">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d36ad-758">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d36ad-759">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-759">Object</span></span>|<span data-ttu-id="d36ad-760">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-760">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-761">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-761">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="d36ad-762">Boolean</span><span class="sxs-lookup"><span data-stu-id="d36ad-762">Boolean</span></span>|<span data-ttu-id="d36ad-763">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-763">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-764">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-764">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="d36ad-765">function</span><span class="sxs-lookup"><span data-stu-id="d36ad-765">function</span></span>|<span data-ttu-id="d36ad-766">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-766">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-767">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-767">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d36ad-768">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-768">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d36ad-769">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-769">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d36ad-770">エラー</span><span class="sxs-lookup"><span data-stu-id="d36ad-770">Errors</span></span>

|<span data-ttu-id="d36ad-771">エラー コード</span><span class="sxs-lookup"><span data-stu-id="d36ad-771">Error code</span></span>|<span data-ttu-id="d36ad-772">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-772">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="d36ad-773">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="d36ad-773">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="d36ad-774">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="d36ad-774">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="d36ad-775">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-775">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-776">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-776">Requirements</span></span>

|<span data-ttu-id="d36ad-777">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-777">Requirement</span></span>|<span data-ttu-id="d36ad-778">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-778">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-779">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-779">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-780">1.1</span><span class="sxs-lookup"><span data-stu-id="d36ad-780">1.1</span></span>|
|[<span data-ttu-id="d36ad-781">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-781">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-782">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-782">ReadWriteItem</span></span>|
|[<span data-ttu-id="d36ad-783">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-783">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-784">作成</span><span class="sxs-lookup"><span data-stu-id="d36ad-784">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d36ad-785">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-785">Examples</span></span>

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

<span data-ttu-id="d36ad-786">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-786">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="d36ad-787">addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d36ad-787">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d36ad-788">Base64 エンコードのファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-788">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="d36ad-789">この`addFileAttachmentFromBase64Async`メソッドは、base64 エンコードからファイルをアップロードし、新規作成フォームのアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-789">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="d36ad-790">このメソッドは、AsyncResult オブジェクトの添付ファイル識別子を返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-790">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="d36ad-791">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-791">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-792">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-792">Parameters</span></span>

|<span data-ttu-id="d36ad-793">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-793">Name</span></span>|<span data-ttu-id="d36ad-794">種類</span><span class="sxs-lookup"><span data-stu-id="d36ad-794">Type</span></span>|<span data-ttu-id="d36ad-795">属性</span><span class="sxs-lookup"><span data-stu-id="d36ad-795">Attributes</span></span>|<span data-ttu-id="d36ad-796">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-796">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="d36ad-797">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-797">String</span></span>||<span data-ttu-id="d36ad-798">電子メールまたはイベントに追加する画像またはファイルの、base64 でエンコードされたコンテンツ。</span><span class="sxs-lookup"><span data-stu-id="d36ad-798">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="d36ad-799">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-799">String</span></span>||<span data-ttu-id="d36ad-p139">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="d36ad-802">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-802">Object</span></span>|<span data-ttu-id="d36ad-803">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-803">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-804">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d36ad-804">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d36ad-805">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-805">Object</span></span>|<span data-ttu-id="d36ad-806">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-806">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-807">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-807">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="d36ad-808">Boolean</span><span class="sxs-lookup"><span data-stu-id="d36ad-808">Boolean</span></span>|<span data-ttu-id="d36ad-809">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-809">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-810">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-810">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="d36ad-811">function</span><span class="sxs-lookup"><span data-stu-id="d36ad-811">function</span></span>|<span data-ttu-id="d36ad-812">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-812">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-813">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-813">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d36ad-814">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-814">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d36ad-815">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-815">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d36ad-816">エラー</span><span class="sxs-lookup"><span data-stu-id="d36ad-816">Errors</span></span>

|<span data-ttu-id="d36ad-817">エラー コード</span><span class="sxs-lookup"><span data-stu-id="d36ad-817">Error code</span></span>|<span data-ttu-id="d36ad-818">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-818">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="d36ad-819">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="d36ad-819">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="d36ad-820">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="d36ad-820">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="d36ad-821">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-821">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-822">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-822">Requirements</span></span>

|<span data-ttu-id="d36ad-823">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-823">Requirement</span></span>|<span data-ttu-id="d36ad-824">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-824">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-825">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-825">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-826">プレビュー</span><span class="sxs-lookup"><span data-stu-id="d36ad-826">Preview</span></span>|
|[<span data-ttu-id="d36ad-827">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-827">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-828">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-828">ReadWriteItem</span></span>|
|[<span data-ttu-id="d36ad-829">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-829">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-830">作成</span><span class="sxs-lookup"><span data-stu-id="d36ad-830">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d36ad-831">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-831">Examples</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="d36ad-832">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d36ad-832">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="d36ad-833">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-833">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="d36ad-834">現在、サポートされて`Office.EventType.AttachmentsChanged`いる`Office.EventType.AppointmentTimeChanged`イベント`Office.EventType.EnhancedLocationsChanged`の`Office.EventType.RecipientsChanged`種類は`Office.EventType.RecurrenceChanged`、、、、、です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-834">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-835">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-835">Parameters</span></span>

| <span data-ttu-id="d36ad-836">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-836">Name</span></span> | <span data-ttu-id="d36ad-837">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-837">Type</span></span> | <span data-ttu-id="d36ad-838">属性</span><span class="sxs-lookup"><span data-stu-id="d36ad-838">Attributes</span></span> | <span data-ttu-id="d36ad-839">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-839">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d36ad-840">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d36ad-840">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d36ad-841">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="d36ad-841">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="d36ad-842">Function</span><span class="sxs-lookup"><span data-stu-id="d36ad-842">Function</span></span> || <span data-ttu-id="d36ad-p140">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="d36ad-846">Object</span><span class="sxs-lookup"><span data-stu-id="d36ad-846">Object</span></span> | <span data-ttu-id="d36ad-847">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-847">&lt;optional&gt;</span></span> | <span data-ttu-id="d36ad-848">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d36ad-848">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d36ad-849">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-849">Object</span></span> | <span data-ttu-id="d36ad-850">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-850">&lt;optional&gt;</span></span> | <span data-ttu-id="d36ad-851">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-851">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d36ad-852">function</span><span class="sxs-lookup"><span data-stu-id="d36ad-852">function</span></span>| <span data-ttu-id="d36ad-853">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-853">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-854">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-854">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-855">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-855">Requirements</span></span>

|<span data-ttu-id="d36ad-856">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-856">Requirement</span></span>| <span data-ttu-id="d36ad-857">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-857">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-858">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-858">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d36ad-859">1.7</span><span class="sxs-lookup"><span data-stu-id="d36ad-859">1.7</span></span> |
|[<span data-ttu-id="d36ad-860">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-860">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d36ad-861">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-861">ReadItem</span></span> |
|[<span data-ttu-id="d36ad-862">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-862">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d36ad-863">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-863">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="d36ad-864">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-864">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="d36ad-865">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d36ad-865">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d36ad-866">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-866">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="d36ad-p141">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="d36ad-870">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-870">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="d36ad-871">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-871">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-872">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-872">Parameters</span></span>

|<span data-ttu-id="d36ad-873">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-873">Name</span></span>|<span data-ttu-id="d36ad-874">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-874">Type</span></span>|<span data-ttu-id="d36ad-875">属性</span><span class="sxs-lookup"><span data-stu-id="d36ad-875">Attributes</span></span>|<span data-ttu-id="d36ad-876">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-876">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="d36ad-877">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-877">String</span></span>||<span data-ttu-id="d36ad-p142">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="d36ad-880">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-880">String</span></span>||<span data-ttu-id="d36ad-881">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="d36ad-881">The subject of the item to be attached.</span></span> <span data-ttu-id="d36ad-882">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-882">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="d36ad-883">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-883">Object</span></span>|<span data-ttu-id="d36ad-884">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-884">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-885">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d36ad-885">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d36ad-886">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-886">Object</span></span>|<span data-ttu-id="d36ad-887">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-887">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-888">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-888">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d36ad-889">function</span><span class="sxs-lookup"><span data-stu-id="d36ad-889">function</span></span>|<span data-ttu-id="d36ad-890">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-890">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-891">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-891">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d36ad-892">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-892">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d36ad-893">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-893">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d36ad-894">エラー</span><span class="sxs-lookup"><span data-stu-id="d36ad-894">Errors</span></span>

|<span data-ttu-id="d36ad-895">エラー コード</span><span class="sxs-lookup"><span data-stu-id="d36ad-895">Error code</span></span>|<span data-ttu-id="d36ad-896">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-896">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="d36ad-897">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-897">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-898">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-898">Requirements</span></span>

|<span data-ttu-id="d36ad-899">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-899">Requirement</span></span>|<span data-ttu-id="d36ad-900">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-900">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-901">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-901">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-902">1.1</span><span class="sxs-lookup"><span data-stu-id="d36ad-902">1.1</span></span>|
|[<span data-ttu-id="d36ad-903">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-903">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-904">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-904">ReadWriteItem</span></span>|
|[<span data-ttu-id="d36ad-905">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-905">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-906">作成</span><span class="sxs-lookup"><span data-stu-id="d36ad-906">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-907">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-907">Example</span></span>

<span data-ttu-id="d36ad-908">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-908">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="d36ad-909">close()</span><span class="sxs-lookup"><span data-stu-id="d36ad-909">close()</span></span>

<span data-ttu-id="d36ad-910">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-910">Closes the current item that is being composed.</span></span>

<span data-ttu-id="d36ad-p144">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-913">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-913">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="d36ad-914">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-914">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-915">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-915">Requirements</span></span>

|<span data-ttu-id="d36ad-916">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-916">Requirement</span></span>|<span data-ttu-id="d36ad-917">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-917">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-918">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-918">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-919">1.3</span><span class="sxs-lookup"><span data-stu-id="d36ad-919">1.3</span></span>|
|[<span data-ttu-id="d36ad-920">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-920">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-921">制限あり</span><span class="sxs-lookup"><span data-stu-id="d36ad-921">Restricted</span></span>|
|[<span data-ttu-id="d36ad-922">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-922">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-923">新規作成</span><span class="sxs-lookup"><span data-stu-id="d36ad-923">Compose</span></span>|

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="d36ad-924">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="d36ad-924">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="d36ad-925">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-925">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-926">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-926">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d36ad-927">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-927">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d36ad-928">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="d36ad-928">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="d36ad-p145">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-932">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-932">Parameters</span></span>

|<span data-ttu-id="d36ad-933">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-933">Name</span></span>|<span data-ttu-id="d36ad-934">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-934">Type</span></span>|<span data-ttu-id="d36ad-935">属性</span><span class="sxs-lookup"><span data-stu-id="d36ad-935">Attributes</span></span>|<span data-ttu-id="d36ad-936">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-936">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="d36ad-937">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="d36ad-937">String &#124; Object</span></span>||<span data-ttu-id="d36ad-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d36ad-940">**または**</span><span class="sxs-lookup"><span data-stu-id="d36ad-940">**OR**</span></span><br/><span data-ttu-id="d36ad-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="d36ad-943">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-943">String</span></span>|<span data-ttu-id="d36ad-944">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-944">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="d36ad-947">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-947">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="d36ad-948">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-948">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-949">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="d36ad-949">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="d36ad-950">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-950">String</span></span>||<span data-ttu-id="d36ad-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="d36ad-953">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-953">String</span></span>||<span data-ttu-id="d36ad-954">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-954">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="d36ad-955">文字列</span><span class="sxs-lookup"><span data-stu-id="d36ad-955">String</span></span>||<span data-ttu-id="d36ad-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="d36ad-958">ブール値</span><span class="sxs-lookup"><span data-stu-id="d36ad-958">Boolean</span></span>||<span data-ttu-id="d36ad-p151">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="d36ad-961">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-961">String</span></span>||<span data-ttu-id="d36ad-p152">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="d36ad-965">function</span><span class="sxs-lookup"><span data-stu-id="d36ad-965">function</span></span>|<span data-ttu-id="d36ad-966">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-966">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-967">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-968">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-968">Requirements</span></span>

|<span data-ttu-id="d36ad-969">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-969">Requirement</span></span>|<span data-ttu-id="d36ad-970">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-970">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-971">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-971">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-972">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-972">1.0</span></span>|
|[<span data-ttu-id="d36ad-973">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-973">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-974">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-974">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-975">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-975">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-976">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-976">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d36ad-977">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-977">Examples</span></span>

<span data-ttu-id="d36ad-978">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-978">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="d36ad-979">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-979">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="d36ad-980">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-980">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d36ad-981">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-981">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d36ad-982">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-982">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d36ad-983">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-983">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="d36ad-984">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="d36ad-984">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="d36ad-985">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-985">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-986">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-986">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d36ad-987">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-987">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d36ad-988">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="d36ad-988">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="d36ad-p153">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p153">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-992">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-992">Parameters</span></span>

|<span data-ttu-id="d36ad-993">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-993">Name</span></span>|<span data-ttu-id="d36ad-994">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-994">Type</span></span>|<span data-ttu-id="d36ad-995">属性</span><span class="sxs-lookup"><span data-stu-id="d36ad-995">Attributes</span></span>|<span data-ttu-id="d36ad-996">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-996">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="d36ad-997">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="d36ad-997">String &#124; Object</span></span>||<span data-ttu-id="d36ad-p154">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d36ad-1000">**または**</span><span class="sxs-lookup"><span data-stu-id="d36ad-1000">**OR**</span></span><br/><span data-ttu-id="d36ad-p155">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="d36ad-1003">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-1003">String</span></span>|<span data-ttu-id="d36ad-1004">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-p156">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="d36ad-1007">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1007">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="d36ad-1008">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1008">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1009">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1009">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="d36ad-1010">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-1010">String</span></span>||<span data-ttu-id="d36ad-p157">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="d36ad-1013">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-1013">String</span></span>||<span data-ttu-id="d36ad-1014">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1014">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="d36ad-1015">文字列</span><span class="sxs-lookup"><span data-stu-id="d36ad-1015">String</span></span>||<span data-ttu-id="d36ad-p158">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="d36ad-1018">ブール値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1018">Boolean</span></span>||<span data-ttu-id="d36ad-p159">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="d36ad-1021">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-1021">String</span></span>||<span data-ttu-id="d36ad-p160">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="d36ad-1025">function</span><span class="sxs-lookup"><span data-stu-id="d36ad-1025">function</span></span>|<span data-ttu-id="d36ad-1026">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1026">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1027">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1027">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-1028">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1028">Requirements</span></span>

|<span data-ttu-id="d36ad-1029">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1029">Requirement</span></span>|<span data-ttu-id="d36ad-1030">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1030">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-1031">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-1031">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-1032">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-1032">1.0</span></span>|
|[<span data-ttu-id="d36ad-1033">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1033">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-1034">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-1034">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-1035">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1035">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-1036">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-1036">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d36ad-1037">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-1037">Examples</span></span>

<span data-ttu-id="d36ad-1038">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1038">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="d36ad-1039">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1039">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="d36ad-1040">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1040">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d36ad-1041">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1041">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d36ad-1042">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1042">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d36ad-1043">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1043">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="d36ad-1044">getAttachmentContentAsync (attachmentId, [options], [callback]) > [Attachmentcontent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="d36ad-1044">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="d36ad-1045">メッセージまたは予定から指定された添付ファイルを取得し`AttachmentContent` 、それをオブジェクトとして返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1045">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="d36ad-1046">メソッド`getAttachmentContentAsync`は、指定された id の添付ファイルをアイテムから取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1046">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="d36ad-1047">ベストプラクティスとして、識別子を使用して、または`getAttachmentsAsync` `item.attachments`の呼び出しで attachmentIds を取得したのと同じセッションの添付ファイルを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1047">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="d36ad-1048">Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1048">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="d36ad-1049">ユーザーがアプリを閉じたとき、またはインラインフォームの作成が開始されたときに、別のウィンドウで続行するためにフォームをポップアウトした後、セッションが終了します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1049">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-1050">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-1050">Parameters</span></span>

|<span data-ttu-id="d36ad-1051">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-1051">Name</span></span>|<span data-ttu-id="d36ad-1052">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-1052">Type</span></span>|<span data-ttu-id="d36ad-1053">属性</span><span class="sxs-lookup"><span data-stu-id="d36ad-1053">Attributes</span></span>|<span data-ttu-id="d36ad-1054">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-1054">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="d36ad-1055">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-1055">String</span></span>||<span data-ttu-id="d36ad-1056">取得する添付ファイルの識別子を指定します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1056">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="d36ad-1057">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-1057">Object</span></span>|<span data-ttu-id="d36ad-1058">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1058">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1059">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1059">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d36ad-1060">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-1060">Object</span></span>|<span data-ttu-id="d36ad-1061">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1061">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1062">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1062">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d36ad-1063">function</span><span class="sxs-lookup"><span data-stu-id="d36ad-1063">function</span></span>|<span data-ttu-id="d36ad-1064">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1064">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1065">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1065">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-1066">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1066">Requirements</span></span>

|<span data-ttu-id="d36ad-1067">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1067">Requirement</span></span>|<span data-ttu-id="d36ad-1068">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-1069">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-1070">プレビュー</span><span class="sxs-lookup"><span data-stu-id="d36ad-1070">Preview</span></span>|
|[<span data-ttu-id="d36ad-1071">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1071">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-1072">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-1072">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-1073">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1073">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-1074">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-1074">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d36ad-1075">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d36ad-1075">Returns:</span></span>

<span data-ttu-id="d36ad-1076">型: [Attachmentcontent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="d36ad-1076">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="d36ad-1077">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-1077">Example</span></span>

```javascript
var item = Office.context.mailbox.item;
var listOfAttachments = [];
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var options = {asyncContext: {type: result.value[i].attachmentType}};
      getAttachmentContentAsync(result.value[i].id, options, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  if (result.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
    // Handle file attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.Eml) {
    // Handle email item attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
    // Handle .icalender attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.Url) {
    // Handle cloud attachment.
  } else {
    // Handle attachment formats that are not supported.
  }
}
```

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="d36ad-1078">getAttachmentsAsync ([オプション], [callback])] > <[Attachmentdetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d36ad-1078">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="d36ad-1079">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1079">Gets the item's attachments as an array.</span></span> <span data-ttu-id="d36ad-1080">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1080">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-1081">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-1081">Parameters</span></span>

|<span data-ttu-id="d36ad-1082">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-1082">Name</span></span>|<span data-ttu-id="d36ad-1083">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-1083">Type</span></span>|<span data-ttu-id="d36ad-1084">属性</span><span class="sxs-lookup"><span data-stu-id="d36ad-1084">Attributes</span></span>|<span data-ttu-id="d36ad-1085">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-1085">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="d36ad-1086">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-1086">Object</span></span>|<span data-ttu-id="d36ad-1087">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1087">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1088">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1088">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d36ad-1089">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-1089">Object</span></span>|<span data-ttu-id="d36ad-1090">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1090">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1091">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1091">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d36ad-1092">function</span><span class="sxs-lookup"><span data-stu-id="d36ad-1092">function</span></span>|<span data-ttu-id="d36ad-1093">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1093">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1094">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1094">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-1095">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1095">Requirements</span></span>

|<span data-ttu-id="d36ad-1096">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1096">Requirement</span></span>|<span data-ttu-id="d36ad-1097">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1097">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-1098">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-1098">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-1099">プレビュー</span><span class="sxs-lookup"><span data-stu-id="d36ad-1099">Preview</span></span>|
|[<span data-ttu-id="d36ad-1100">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1100">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-1101">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-1101">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-1102">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1102">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-1103">作成</span><span class="sxs-lookup"><span data-stu-id="d36ad-1103">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="d36ad-1104">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d36ad-1104">Returns:</span></span>

<span data-ttu-id="d36ad-1105">型: <[Attachmentdetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d36ad-1105">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="d36ad-1106">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-1106">Example</span></span>

<span data-ttu-id="d36ad-1107">次の例では、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1107">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="d36ad-1108">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="d36ad-1108">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="d36ad-1109">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1109">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-1110">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1110">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-1111">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1111">Requirements</span></span>

|<span data-ttu-id="d36ad-1112">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1112">Requirement</span></span>|<span data-ttu-id="d36ad-1113">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-1114">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-1114">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-1115">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-1115">1.0</span></span>|
|[<span data-ttu-id="d36ad-1116">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1116">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-1117">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-1117">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-1118">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1118">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-1119">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-1119">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d36ad-1120">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d36ad-1120">Returns:</span></span>

<span data-ttu-id="d36ad-1121">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="d36ad-1121">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="d36ad-1122">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-1122">Example</span></span>

<span data-ttu-id="d36ad-1123">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1123">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="d36ad-1124">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="d36ad-1124">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="d36ad-1125">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1125">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-1126">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1126">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-1127">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-1127">Parameters</span></span>

|<span data-ttu-id="d36ad-1128">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-1128">Name</span></span>|<span data-ttu-id="d36ad-1129">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-1129">Type</span></span>|<span data-ttu-id="d36ad-1130">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-1130">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="d36ad-1131">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="d36ad-1131">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="d36ad-1132">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1132">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-1133">Requirements</span><span class="sxs-lookup"><span data-stu-id="d36ad-1133">Requirements</span></span>

|<span data-ttu-id="d36ad-1134">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1134">Requirement</span></span>|<span data-ttu-id="d36ad-1135">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1135">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-1136">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-1136">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-1137">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-1137">1.0</span></span>|
|[<span data-ttu-id="d36ad-1138">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1138">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-1139">制限あり</span><span class="sxs-lookup"><span data-stu-id="d36ad-1139">Restricted</span></span>|
|[<span data-ttu-id="d36ad-1140">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1140">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-1141">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-1141">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d36ad-1142">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d36ad-1142">Returns:</span></span>

<span data-ttu-id="d36ad-1143">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1143">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="d36ad-1144">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1144">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="d36ad-1145">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1145">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="d36ad-1146">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1146">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="d36ad-1147">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1147">Value of `entityType`</span></span>|<span data-ttu-id="d36ad-1148">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="d36ad-1148">Type of objects in returned array</span></span>|<span data-ttu-id="d36ad-1149">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1149">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="d36ad-1150">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-1150">String</span></span>|<span data-ttu-id="d36ad-1151">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="d36ad-1151">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="d36ad-1152">連絡先</span><span class="sxs-lookup"><span data-stu-id="d36ad-1152">Contact</span></span>|<span data-ttu-id="d36ad-1153">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d36ad-1153">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="d36ad-1154">文字列</span><span class="sxs-lookup"><span data-stu-id="d36ad-1154">String</span></span>|<span data-ttu-id="d36ad-1155">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d36ad-1155">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="d36ad-1156">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="d36ad-1156">MeetingSuggestion</span></span>|<span data-ttu-id="d36ad-1157">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d36ad-1157">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="d36ad-1158">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="d36ad-1158">PhoneNumber</span></span>|<span data-ttu-id="d36ad-1159">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="d36ad-1159">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="d36ad-1160">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="d36ad-1160">TaskSuggestion</span></span>|<span data-ttu-id="d36ad-1161">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d36ad-1161">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="d36ad-1162">文字列</span><span class="sxs-lookup"><span data-stu-id="d36ad-1162">String</span></span>|<span data-ttu-id="d36ad-1163">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="d36ad-1163">**Restricted**</span></span>|

<span data-ttu-id="d36ad-1164">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="d36ad-1164">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="d36ad-1165">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-1165">Example</span></span>

<span data-ttu-id="d36ad-1166">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1166">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="d36ad-1167">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="d36ad-1167">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="d36ad-1168">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1168">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-1169">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1169">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d36ad-1170">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1170">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-1171">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-1171">Parameters</span></span>

|<span data-ttu-id="d36ad-1172">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-1172">Name</span></span>|<span data-ttu-id="d36ad-1173">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-1173">Type</span></span>|<span data-ttu-id="d36ad-1174">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-1174">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="d36ad-1175">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-1175">String</span></span>|<span data-ttu-id="d36ad-1176">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1176">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-1177">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1177">Requirements</span></span>

|<span data-ttu-id="d36ad-1178">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1178">Requirement</span></span>|<span data-ttu-id="d36ad-1179">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1179">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-1180">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-1180">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-1181">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-1181">1.0</span></span>|
|[<span data-ttu-id="d36ad-1182">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1182">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-1183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-1183">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-1184">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1184">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-1185">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-1185">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d36ad-1186">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d36ad-1186">Returns:</span></span>

<span data-ttu-id="d36ad-p164">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="d36ad-1189">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="d36ad-1189">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="d36ad-1190">、Office.context.mailbox.item.getinitializationcontextasync ([オプション], [callback])</span><span class="sxs-lookup"><span data-stu-id="d36ad-1190">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="d36ad-1191">[アクション可能なメッセージによってアドインがアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されたときに渡される初期化データを取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1191">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-1192">このメソッドは、Outlook 2016 以降の Windows (16.0.8413.1000 より後のバージョン) および Outlook on the Office 365 でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1192">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-1193">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-1193">Parameters</span></span>

|<span data-ttu-id="d36ad-1194">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-1194">Name</span></span>|<span data-ttu-id="d36ad-1195">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-1195">Type</span></span>|<span data-ttu-id="d36ad-1196">属性</span><span class="sxs-lookup"><span data-stu-id="d36ad-1196">Attributes</span></span>|<span data-ttu-id="d36ad-1197">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-1197">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="d36ad-1198">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-1198">Object</span></span>|<span data-ttu-id="d36ad-1199">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1199">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1200">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1200">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d36ad-1201">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-1201">Object</span></span>|<span data-ttu-id="d36ad-1202">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1202">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1203">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1203">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d36ad-1204">function</span><span class="sxs-lookup"><span data-stu-id="d36ad-1204">function</span></span>|<span data-ttu-id="d36ad-1205">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1205">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1206">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1206">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d36ad-1207">成功すると、初期化データが文字列とし`asyncResult.value`てプロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1207">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="d36ad-1208">初期化コンテキストがない場合、 `asyncResult`オブジェクトには、 `Error` `code`プロパティがに`9020`設定されたオブジェクトと`name`プロパティがに`GenericResponseError`設定されたオブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1208">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-1209">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1209">Requirements</span></span>

|<span data-ttu-id="d36ad-1210">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1210">Requirement</span></span>|<span data-ttu-id="d36ad-1211">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1211">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-1212">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-1212">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-1213">プレビュー</span><span class="sxs-lookup"><span data-stu-id="d36ad-1213">Preview</span></span>|
|[<span data-ttu-id="d36ad-1214">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1214">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-1215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-1215">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-1216">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-1217">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-1217">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-1218">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-1218">Example</span></span>

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

#### <a name="getregexmatches--object"></a><span data-ttu-id="d36ad-1219">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d36ad-1219">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="d36ad-1220">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1220">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-1221">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1221">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d36ad-p165">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p165">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d36ad-1225">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1225">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d36ad-1226">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1226">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="d36ad-p166">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-1230">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1230">Requirements</span></span>

|<span data-ttu-id="d36ad-1231">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1231">Requirement</span></span>|<span data-ttu-id="d36ad-1232">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1232">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-1233">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-1233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-1234">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-1234">1.0</span></span>|
|[<span data-ttu-id="d36ad-1235">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1235">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-1236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-1236">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-1237">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1237">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-1238">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-1238">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d36ad-1239">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d36ad-1239">Returns:</span></span>

<span data-ttu-id="d36ad-p167">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="d36ad-1242">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="d36ad-1242">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d36ad-1243">Object</span><span class="sxs-lookup"><span data-stu-id="d36ad-1243">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d36ad-1244">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-1244">Example</span></span>

<span data-ttu-id="d36ad-1245">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1245">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="d36ad-1246">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="d36ad-1246">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="d36ad-1247">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1247">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-1248">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1248">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d36ad-1249">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1249">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="d36ad-p168">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p168">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-1252">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-1252">Parameters</span></span>

|<span data-ttu-id="d36ad-1253">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-1253">Name</span></span>|<span data-ttu-id="d36ad-1254">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-1254">Type</span></span>|<span data-ttu-id="d36ad-1255">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-1255">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="d36ad-1256">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-1256">String</span></span>|<span data-ttu-id="d36ad-1257">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1257">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-1258">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1258">Requirements</span></span>

|<span data-ttu-id="d36ad-1259">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1259">Requirement</span></span>|<span data-ttu-id="d36ad-1260">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1260">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-1261">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-1261">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-1262">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-1262">1.0</span></span>|
|[<span data-ttu-id="d36ad-1263">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1263">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-1264">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-1264">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-1265">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1265">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-1266">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-1266">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d36ad-1267">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d36ad-1267">Returns:</span></span>

<span data-ttu-id="d36ad-1268">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1268">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="d36ad-1269">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="d36ad-1269">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d36ad-1270">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="d36ad-1270">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d36ad-1271">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-1271">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="d36ad-1272">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="d36ad-1272">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="d36ad-1273">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1273">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="d36ad-p169">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p169">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-1276">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-1276">Parameters</span></span>

|<span data-ttu-id="d36ad-1277">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-1277">Name</span></span>|<span data-ttu-id="d36ad-1278">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-1278">Type</span></span>|<span data-ttu-id="d36ad-1279">属性</span><span class="sxs-lookup"><span data-stu-id="d36ad-1279">Attributes</span></span>|<span data-ttu-id="d36ad-1280">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-1280">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="d36ad-1281">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d36ad-1281">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="d36ad-p170">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p170">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="d36ad-1285">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-1285">Object</span></span>|<span data-ttu-id="d36ad-1286">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1286">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1287">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1287">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d36ad-1288">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-1288">Object</span></span>|<span data-ttu-id="d36ad-1289">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1289">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1290">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1290">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d36ad-1291">function</span><span class="sxs-lookup"><span data-stu-id="d36ad-1291">function</span></span>||<span data-ttu-id="d36ad-1292">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1292">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d36ad-1293">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1293">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="d36ad-1294">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1294">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-1295">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1295">Requirements</span></span>

|<span data-ttu-id="d36ad-1296">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1296">Requirement</span></span>|<span data-ttu-id="d36ad-1297">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1297">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-1298">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-1298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-1299">1.2</span><span class="sxs-lookup"><span data-stu-id="d36ad-1299">1.2</span></span>|
|[<span data-ttu-id="d36ad-1300">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1300">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-1301">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-1301">ReadWriteItem</span></span>|
|[<span data-ttu-id="d36ad-1302">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1302">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-1303">作成</span><span class="sxs-lookup"><span data-stu-id="d36ad-1303">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="d36ad-1304">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d36ad-1304">Returns:</span></span>

<span data-ttu-id="d36ad-1305">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1305">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="d36ad-1306">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="d36ad-1306">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d36ad-1307">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-1307">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d36ad-1308">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-1308">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="d36ad-1309">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="d36ad-1309">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="d36ad-1310">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1310">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="d36ad-1311">強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1311">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-1312">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1312">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-1313">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1313">Requirements</span></span>

|<span data-ttu-id="d36ad-1314">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1314">Requirement</span></span>|<span data-ttu-id="d36ad-1315">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1315">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-1316">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-1316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-1317">1.6</span><span class="sxs-lookup"><span data-stu-id="d36ad-1317">1.6</span></span>|
|[<span data-ttu-id="d36ad-1318">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-1319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-1319">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-1320">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-1321">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-1321">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d36ad-1322">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d36ad-1322">Returns:</span></span>

<span data-ttu-id="d36ad-1323">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="d36ad-1323">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="d36ad-1324">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-1324">Example</span></span>

<span data-ttu-id="d36ad-1325">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1325">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="d36ad-1326">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d36ad-1326">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="d36ad-p173">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p173">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-1329">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1329">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d36ad-p174">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p174">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d36ad-1333">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1333">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d36ad-1334">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1334">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="d36ad-p175">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p175">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d36ad-1338">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1338">Requirements</span></span>

|<span data-ttu-id="d36ad-1339">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1339">Requirement</span></span>|<span data-ttu-id="d36ad-1340">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1340">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-1341">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-1341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-1342">1.6</span><span class="sxs-lookup"><span data-stu-id="d36ad-1342">1.6</span></span>|
|[<span data-ttu-id="d36ad-1343">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-1344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-1344">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-1345">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-1346">読み取り</span><span class="sxs-lookup"><span data-stu-id="d36ad-1346">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d36ad-1347">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d36ad-1347">Returns:</span></span>

<span data-ttu-id="d36ad-p176">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p176">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="d36ad-1350">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-1350">Example</span></span>

<span data-ttu-id="d36ad-1351">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1351">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="d36ad-1352">getSharedPropertiesAsync ([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d36ad-1352">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="d36ad-1353">共有フォルダー、予定表、またはメールボックス内の選択した予定またはメッセージのプロパティを取得します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1353">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-1354">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-1354">Parameters</span></span>

|<span data-ttu-id="d36ad-1355">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-1355">Name</span></span>|<span data-ttu-id="d36ad-1356">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-1356">Type</span></span>|<span data-ttu-id="d36ad-1357">属性</span><span class="sxs-lookup"><span data-stu-id="d36ad-1357">Attributes</span></span>|<span data-ttu-id="d36ad-1358">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-1358">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="d36ad-1359">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-1359">Object</span></span>|<span data-ttu-id="d36ad-1360">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1360">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1361">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1361">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d36ad-1362">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-1362">Object</span></span>|<span data-ttu-id="d36ad-1363">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1363">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1364">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1364">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d36ad-1365">function</span><span class="sxs-lookup"><span data-stu-id="d36ad-1365">function</span></span>||<span data-ttu-id="d36ad-1366">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1366">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d36ad-1367">共有プロパティは、 [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) `asyncResult.value`プロパティのオブジェクトとして提供されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1367">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="d36ad-1368">このオブジェクトは、アイテムの共有プロパティを取得するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1368">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-1369">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1369">Requirements</span></span>

|<span data-ttu-id="d36ad-1370">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1370">Requirement</span></span>|<span data-ttu-id="d36ad-1371">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1371">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-1372">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-1372">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-1373">プレビュー</span><span class="sxs-lookup"><span data-stu-id="d36ad-1373">Preview</span></span>|
|[<span data-ttu-id="d36ad-1374">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1374">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-1375">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-1375">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-1376">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1376">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-1377">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-1377">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-1378">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-1378">Example</span></span>

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="d36ad-1379">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d36ad-1379">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="d36ad-1380">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1380">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="d36ad-p178">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p178">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-1384">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-1384">Parameters</span></span>

|<span data-ttu-id="d36ad-1385">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-1385">Name</span></span>|<span data-ttu-id="d36ad-1386">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-1386">Type</span></span>|<span data-ttu-id="d36ad-1387">属性</span><span class="sxs-lookup"><span data-stu-id="d36ad-1387">Attributes</span></span>|<span data-ttu-id="d36ad-1388">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-1388">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="d36ad-1389">function</span><span class="sxs-lookup"><span data-stu-id="d36ad-1389">function</span></span>||<span data-ttu-id="d36ad-1390">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1390">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d36ad-1391">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1391">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="d36ad-1392">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1392">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="d36ad-1393">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-1393">Object</span></span>|<span data-ttu-id="d36ad-1394">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1394">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1395">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1395">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="d36ad-1396">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1396">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-1397">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1397">Requirements</span></span>

|<span data-ttu-id="d36ad-1398">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1398">Requirement</span></span>|<span data-ttu-id="d36ad-1399">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1399">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-1400">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-1400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-1401">1.0</span><span class="sxs-lookup"><span data-stu-id="d36ad-1401">1.0</span></span>|
|[<span data-ttu-id="d36ad-1402">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1402">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-1403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-1403">ReadItem</span></span>|
|[<span data-ttu-id="d36ad-1404">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1404">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-1405">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-1405">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-1406">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-1406">Example</span></span>

<span data-ttu-id="d36ad-p181">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p181">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="d36ad-1410">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d36ad-1410">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="d36ad-1411">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1411">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="d36ad-1412">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1412">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="d36ad-1413">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1413">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="d36ad-1414">Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1414">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="d36ad-1415">ユーザーがアプリを閉じたとき、またはインラインフォームの作成が開始されたときに、別のウィンドウで続行するためにフォームをポップアウトした後、セッションが終了します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1415">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-1416">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-1416">Parameters</span></span>

|<span data-ttu-id="d36ad-1417">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-1417">Name</span></span>|<span data-ttu-id="d36ad-1418">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-1418">Type</span></span>|<span data-ttu-id="d36ad-1419">属性</span><span class="sxs-lookup"><span data-stu-id="d36ad-1419">Attributes</span></span>|<span data-ttu-id="d36ad-1420">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-1420">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="d36ad-1421">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-1421">String</span></span>||<span data-ttu-id="d36ad-1422">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1422">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="d36ad-1423">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-1423">Object</span></span>|<span data-ttu-id="d36ad-1424">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1424">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1425">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1425">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d36ad-1426">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-1426">Object</span></span>|<span data-ttu-id="d36ad-1427">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1427">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1428">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1428">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d36ad-1429">function</span><span class="sxs-lookup"><span data-stu-id="d36ad-1429">function</span></span>|<span data-ttu-id="d36ad-1430">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1430">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1431">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1431">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d36ad-1432">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1432">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d36ad-1433">エラー</span><span class="sxs-lookup"><span data-stu-id="d36ad-1433">Errors</span></span>

|<span data-ttu-id="d36ad-1434">エラー コード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1434">Error code</span></span>|<span data-ttu-id="d36ad-1435">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-1435">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="d36ad-1436">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1436">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-1437">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1437">Requirements</span></span>

|<span data-ttu-id="d36ad-1438">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1438">Requirement</span></span>|<span data-ttu-id="d36ad-1439">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1439">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-1440">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-1440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-1441">1.1</span><span class="sxs-lookup"><span data-stu-id="d36ad-1441">1.1</span></span>|
|[<span data-ttu-id="d36ad-1442">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1442">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-1443">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-1443">ReadWriteItem</span></span>|
|[<span data-ttu-id="d36ad-1444">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1444">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-1445">作成</span><span class="sxs-lookup"><span data-stu-id="d36ad-1445">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-1446">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-1446">Example</span></span>

<span data-ttu-id="d36ad-1447">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1447">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="d36ad-1448">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d36ad-1448">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="d36ad-1449">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1449">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="d36ad-1450">現在、サポートされて`Office.EventType.AttachmentsChanged`いる`Office.EventType.AppointmentTimeChanged`イベント`Office.EventType.EnhancedLocationsChanged`の`Office.EventType.RecipientsChanged`種類は`Office.EventType.RecurrenceChanged`、、、、、です。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1450">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-1451">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-1451">Parameters</span></span>

| <span data-ttu-id="d36ad-1452">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-1452">Name</span></span> | <span data-ttu-id="d36ad-1453">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-1453">Type</span></span> | <span data-ttu-id="d36ad-1454">属性</span><span class="sxs-lookup"><span data-stu-id="d36ad-1454">Attributes</span></span> | <span data-ttu-id="d36ad-1455">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-1455">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d36ad-1456">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d36ad-1456">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d36ad-1457">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1457">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="d36ad-1458">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-1458">Object</span></span> | <span data-ttu-id="d36ad-1459">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1459">&lt;optional&gt;</span></span> | <span data-ttu-id="d36ad-1460">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1460">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d36ad-1461">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-1461">Object</span></span> | <span data-ttu-id="d36ad-1462">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1462">&lt;optional&gt;</span></span> | <span data-ttu-id="d36ad-1463">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1463">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d36ad-1464">function</span><span class="sxs-lookup"><span data-stu-id="d36ad-1464">function</span></span>| <span data-ttu-id="d36ad-1465">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1465">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1466">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1466">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-1467">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1467">Requirements</span></span>

|<span data-ttu-id="d36ad-1468">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1468">Requirement</span></span>| <span data-ttu-id="d36ad-1469">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1469">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-1470">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-1470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d36ad-1471">1.7</span><span class="sxs-lookup"><span data-stu-id="d36ad-1471">1.7</span></span> |
|[<span data-ttu-id="d36ad-1472">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d36ad-1473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-1473">ReadItem</span></span> |
|[<span data-ttu-id="d36ad-1474">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d36ad-1475">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d36ad-1475">Compose or Read</span></span> |

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="d36ad-1476">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d36ad-1476">saveAsync([options], callback)</span></span>

<span data-ttu-id="d36ad-1477">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1477">Asynchronously saves an item.</span></span>

<span data-ttu-id="d36ad-p183">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p183">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-1481">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1481">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="d36ad-1482">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1482">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="d36ad-p185">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p185">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="d36ad-1486">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1486">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="d36ad-1487">Outlook for Mac は、新規`saveAsync`作成モードの会議ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1487">Outlook for Mac does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="d36ad-1488">そのため、この`saveAsync`シナリオで呼び出しを行うと、エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1488">As such, calling `saveAsync` in that scenario returns an error.</span></span> <span data-ttu-id="d36ad-1489">回避策については[、「OFFICE JS API を使用して Outlook For Mac で会議を下書きとして保存できません](https://support.microsoft.com/help/4505745)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1489">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="d36ad-1490">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1490">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-1491">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-1491">Parameters</span></span>

|<span data-ttu-id="d36ad-1492">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-1492">Name</span></span>|<span data-ttu-id="d36ad-1493">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-1493">Type</span></span>|<span data-ttu-id="d36ad-1494">属性</span><span class="sxs-lookup"><span data-stu-id="d36ad-1494">Attributes</span></span>|<span data-ttu-id="d36ad-1495">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-1495">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="d36ad-1496">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-1496">Object</span></span>|<span data-ttu-id="d36ad-1497">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1497">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1498">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1498">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d36ad-1499">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-1499">Object</span></span>|<span data-ttu-id="d36ad-1500">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1500">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1501">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1501">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d36ad-1502">関数</span><span class="sxs-lookup"><span data-stu-id="d36ad-1502">function</span></span>||<span data-ttu-id="d36ad-1503">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1503">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d36ad-1504">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1504">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-1505">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1505">Requirements</span></span>

|<span data-ttu-id="d36ad-1506">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1506">Requirement</span></span>|<span data-ttu-id="d36ad-1507">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1507">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-1508">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-1508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-1509">1.3</span><span class="sxs-lookup"><span data-stu-id="d36ad-1509">1.3</span></span>|
|[<span data-ttu-id="d36ad-1510">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-1511">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-1511">ReadWriteItem</span></span>|
|[<span data-ttu-id="d36ad-1512">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-1513">作成</span><span class="sxs-lookup"><span data-stu-id="d36ad-1513">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d36ad-1514">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-1514">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="d36ad-p187">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p187">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="d36ad-1517">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="d36ad-1517">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="d36ad-1518">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1518">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="d36ad-p188">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p188">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d36ad-1522">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d36ad-1522">Parameters</span></span>

|<span data-ttu-id="d36ad-1523">名前</span><span class="sxs-lookup"><span data-stu-id="d36ad-1523">Name</span></span>|<span data-ttu-id="d36ad-1524">型</span><span class="sxs-lookup"><span data-stu-id="d36ad-1524">Type</span></span>|<span data-ttu-id="d36ad-1525">属性</span><span class="sxs-lookup"><span data-stu-id="d36ad-1525">Attributes</span></span>|<span data-ttu-id="d36ad-1526">説明</span><span class="sxs-lookup"><span data-stu-id="d36ad-1526">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="d36ad-1527">String</span><span class="sxs-lookup"><span data-stu-id="d36ad-1527">String</span></span>||<span data-ttu-id="d36ad-p189">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p189">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="d36ad-1531">Object</span><span class="sxs-lookup"><span data-stu-id="d36ad-1531">Object</span></span>|<span data-ttu-id="d36ad-1532">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1532">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1533">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1533">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d36ad-1534">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d36ad-1534">Object</span></span>|<span data-ttu-id="d36ad-1535">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1535">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-1536">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1536">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="d36ad-1537">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d36ad-1537">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="d36ad-1538">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d36ad-1538">&lt;optional&gt;</span></span>|<span data-ttu-id="d36ad-p190">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p190">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="d36ad-p191">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-p191">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="d36ad-1543">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1543">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="d36ad-1544">function</span><span class="sxs-lookup"><span data-stu-id="d36ad-1544">function</span></span>||<span data-ttu-id="d36ad-1545">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d36ad-1545">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d36ad-1546">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1546">Requirements</span></span>

|<span data-ttu-id="d36ad-1547">要件</span><span class="sxs-lookup"><span data-stu-id="d36ad-1547">Requirement</span></span>|<span data-ttu-id="d36ad-1548">値</span><span class="sxs-lookup"><span data-stu-id="d36ad-1548">Value</span></span>|
|---|---|
|[<span data-ttu-id="d36ad-1549">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d36ad-1549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d36ad-1550">1.2</span><span class="sxs-lookup"><span data-stu-id="d36ad-1550">1.2</span></span>|
|[<span data-ttu-id="d36ad-1551">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d36ad-1551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d36ad-1552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d36ad-1552">ReadWriteItem</span></span>|
|[<span data-ttu-id="d36ad-1553">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d36ad-1553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d36ad-1554">作成</span><span class="sxs-lookup"><span data-stu-id="d36ad-1554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d36ad-1555">例</span><span class="sxs-lookup"><span data-stu-id="d36ad-1555">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
