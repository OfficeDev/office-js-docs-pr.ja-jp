---
title: Office.context.mailbox.item - プレビュー要件セット
description: ''
ms.date: 01/16/2019
localization_priority: Normal
ms.openlocfilehash: b4b2ec9c735270d9b1bfca3d1c24ef6b0f1ca1cb
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389600"
---
# <a name="item"></a><span data-ttu-id="35291-102">item</span><span class="sxs-lookup"><span data-stu-id="35291-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="35291-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="35291-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="35291-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="35291-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-106">要件</span><span class="sxs-lookup"><span data-stu-id="35291-106">Requirements</span></span>

|<span data-ttu-id="35291-107">要件</span><span class="sxs-lookup"><span data-stu-id="35291-107">Requirement</span></span>|<span data-ttu-id="35291-108">値</span><span class="sxs-lookup"><span data-stu-id="35291-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-110">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-110">1.0</span></span>|
|[<span data-ttu-id="35291-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="35291-112">Restricted</span></span>|
|[<span data-ttu-id="35291-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-114">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="35291-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="35291-115">Members and methods</span></span>

| <span data-ttu-id="35291-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-116">Member</span></span> | <span data-ttu-id="35291-117">種類</span><span class="sxs-lookup"><span data-stu-id="35291-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="35291-118">attachments</span><span class="sxs-lookup"><span data-stu-id="35291-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="35291-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-119">Member</span></span> |
| [<span data-ttu-id="35291-120">bcc</span><span class="sxs-lookup"><span data-stu-id="35291-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="35291-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-121">Member</span></span> |
| [<span data-ttu-id="35291-122">body</span><span class="sxs-lookup"><span data-stu-id="35291-122">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="35291-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-123">Member</span></span> |
| [<span data-ttu-id="35291-124">cc</span><span class="sxs-lookup"><span data-stu-id="35291-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="35291-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-125">Member</span></span> |
| [<span data-ttu-id="35291-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="35291-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="35291-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-127">Member</span></span> |
| [<span data-ttu-id="35291-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="35291-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="35291-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-129">Member</span></span> |
| [<span data-ttu-id="35291-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="35291-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="35291-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-131">Member</span></span> |
| [<span data-ttu-id="35291-132">end</span><span class="sxs-lookup"><span data-stu-id="35291-132">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="35291-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-133">Member</span></span> |
| [<span data-ttu-id="35291-134">from</span><span class="sxs-lookup"><span data-stu-id="35291-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="35291-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-135">Member</span></span> |
| [<span data-ttu-id="35291-136">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="35291-136">internetHeaders</span></span>](#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) | <span data-ttu-id="35291-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-137">Member</span></span> |
| [<span data-ttu-id="35291-138">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="35291-138">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="35291-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-139">Member</span></span> |
| [<span data-ttu-id="35291-140">itemClass</span><span class="sxs-lookup"><span data-stu-id="35291-140">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="35291-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-141">Member</span></span> |
| [<span data-ttu-id="35291-142">itemId</span><span class="sxs-lookup"><span data-stu-id="35291-142">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="35291-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-143">Member</span></span> |
| [<span data-ttu-id="35291-144">itemType</span><span class="sxs-lookup"><span data-stu-id="35291-144">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="35291-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-145">Member</span></span> |
| [<span data-ttu-id="35291-146">location</span><span class="sxs-lookup"><span data-stu-id="35291-146">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="35291-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-147">Member</span></span> |
| [<span data-ttu-id="35291-148">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="35291-148">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="35291-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-149">Member</span></span> |
| [<span data-ttu-id="35291-150">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="35291-150">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="35291-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-151">Member</span></span> |
| [<span data-ttu-id="35291-152">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="35291-152">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="35291-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-153">Member</span></span> |
| [<span data-ttu-id="35291-154">organizer</span><span class="sxs-lookup"><span data-stu-id="35291-154">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="35291-155">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-155">Member</span></span> |
| [<span data-ttu-id="35291-156">recurrence</span><span class="sxs-lookup"><span data-stu-id="35291-156">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="35291-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-157">Member</span></span> |
| [<span data-ttu-id="35291-158">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="35291-158">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="35291-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-159">Member</span></span> |
| [<span data-ttu-id="35291-160">sender</span><span class="sxs-lookup"><span data-stu-id="35291-160">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="35291-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-161">Member</span></span> |
| [<span data-ttu-id="35291-162">seriesId</span><span class="sxs-lookup"><span data-stu-id="35291-162">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="35291-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-163">Member</span></span> |
| [<span data-ttu-id="35291-164">start</span><span class="sxs-lookup"><span data-stu-id="35291-164">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="35291-165">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-165">Member</span></span> |
| [<span data-ttu-id="35291-166">subject</span><span class="sxs-lookup"><span data-stu-id="35291-166">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="35291-167">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-167">Member</span></span> |
| [<span data-ttu-id="35291-168">to</span><span class="sxs-lookup"><span data-stu-id="35291-168">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="35291-169">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-169">Member</span></span> |
| [<span data-ttu-id="35291-170">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="35291-170">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="35291-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-171">Method</span></span> |
| [<span data-ttu-id="35291-172">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="35291-172">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="35291-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-173">Method</span></span> |
| [<span data-ttu-id="35291-174">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="35291-174">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="35291-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-175">Method</span></span> |
| [<span data-ttu-id="35291-176">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="35291-176">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="35291-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-177">Method</span></span> |
| [<span data-ttu-id="35291-178">close</span><span class="sxs-lookup"><span data-stu-id="35291-178">close</span></span>](#close) | <span data-ttu-id="35291-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-179">Method</span></span> |
| [<span data-ttu-id="35291-180">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="35291-180">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="35291-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-181">Method</span></span> |
| [<span data-ttu-id="35291-182">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="35291-182">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="35291-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-183">Method</span></span> |
| [<span data-ttu-id="35291-184">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="35291-184">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) | <span data-ttu-id="35291-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-185">Method</span></span> |
| [<span data-ttu-id="35291-186">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="35291-186">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="35291-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-187">Method</span></span> |
| [<span data-ttu-id="35291-188">getEntities</span><span class="sxs-lookup"><span data-stu-id="35291-188">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="35291-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-189">Method</span></span> |
| [<span data-ttu-id="35291-190">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="35291-190">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="35291-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-191">Method</span></span> |
| [<span data-ttu-id="35291-192">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="35291-192">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="35291-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-193">Method</span></span> |
| [<span data-ttu-id="35291-194">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="35291-194">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="35291-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-195">Method</span></span> |
| [<span data-ttu-id="35291-196">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="35291-196">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="35291-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-197">Method</span></span> |
| [<span data-ttu-id="35291-198">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="35291-198">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="35291-199">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-199">Method</span></span> |
| [<span data-ttu-id="35291-200">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="35291-200">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="35291-201">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-201">Method</span></span> |
| [<span data-ttu-id="35291-202">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="35291-202">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="35291-203">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-203">Method</span></span> |
| [<span data-ttu-id="35291-204">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="35291-204">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="35291-205">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-205">Method</span></span> |
| [<span data-ttu-id="35291-206">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="35291-206">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="35291-207">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-207">Method</span></span> |
| [<span data-ttu-id="35291-208">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="35291-208">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="35291-209">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-209">Method</span></span> |
| [<span data-ttu-id="35291-210">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="35291-210">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="35291-211">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-211">Method</span></span> |
| [<span data-ttu-id="35291-212">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="35291-212">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="35291-213">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-213">Method</span></span> |
| [<span data-ttu-id="35291-214">saveAsync</span><span class="sxs-lookup"><span data-stu-id="35291-214">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="35291-215">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-215">Method</span></span> |
| [<span data-ttu-id="35291-216">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="35291-216">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="35291-217">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-217">Method</span></span> |

### <a name="example"></a><span data-ttu-id="35291-218">例</span><span class="sxs-lookup"><span data-stu-id="35291-218">Example</span></span>

<span data-ttu-id="35291-219">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="35291-219">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
}
```

### <a name="members"></a><span data-ttu-id="35291-220">メンバー</span><span class="sxs-lookup"><span data-stu-id="35291-220">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="35291-221">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="35291-221">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="35291-222">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-222">Gets the item's attachments as an array.</span></span> <span data-ttu-id="35291-223">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="35291-223">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="35291-224">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="35291-224">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="35291-225">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="35291-225">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="35291-226">型:</span><span class="sxs-lookup"><span data-stu-id="35291-226">Type:</span></span>

*   <span data-ttu-id="35291-227">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="35291-227">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-228">要件</span><span class="sxs-lookup"><span data-stu-id="35291-228">Requirements</span></span>

|<span data-ttu-id="35291-229">要件</span><span class="sxs-lookup"><span data-stu-id="35291-229">Requirement</span></span>|<span data-ttu-id="35291-230">値</span><span class="sxs-lookup"><span data-stu-id="35291-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-231">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-231">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-232">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-232">1.0</span></span>|
|[<span data-ttu-id="35291-233">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-233">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-234">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-234">ReadItem</span></span>|
|[<span data-ttu-id="35291-235">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-235">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-236">読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-236">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-237">例</span><span class="sxs-lookup"><span data-stu-id="35291-237">Example</span></span>

<span data-ttu-id="35291-238">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="35291-238">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="35291-239">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="35291-239">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="35291-240">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-240">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="35291-241">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="35291-241">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-242">型:</span><span class="sxs-lookup"><span data-stu-id="35291-242">Type:</span></span>

*   [<span data-ttu-id="35291-243">Recipients</span><span class="sxs-lookup"><span data-stu-id="35291-243">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="35291-244">要件</span><span class="sxs-lookup"><span data-stu-id="35291-244">Requirements</span></span>

|<span data-ttu-id="35291-245">要件</span><span class="sxs-lookup"><span data-stu-id="35291-245">Requirement</span></span>|<span data-ttu-id="35291-246">値</span><span class="sxs-lookup"><span data-stu-id="35291-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-247">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-247">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-248">1.1</span><span class="sxs-lookup"><span data-stu-id="35291-248">1.1</span></span>|
|[<span data-ttu-id="35291-249">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-249">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-250">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-250">ReadItem</span></span>|
|[<span data-ttu-id="35291-251">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-251">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-252">作成</span><span class="sxs-lookup"><span data-stu-id="35291-252">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-253">例</span><span class="sxs-lookup"><span data-stu-id="35291-253">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="35291-254">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="35291-254">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="35291-255">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-255">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-256">型:</span><span class="sxs-lookup"><span data-stu-id="35291-256">Type:</span></span>

*   [<span data-ttu-id="35291-257">Body</span><span class="sxs-lookup"><span data-stu-id="35291-257">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="35291-258">要件</span><span class="sxs-lookup"><span data-stu-id="35291-258">Requirements</span></span>

|<span data-ttu-id="35291-259">要件</span><span class="sxs-lookup"><span data-stu-id="35291-259">Requirement</span></span>|<span data-ttu-id="35291-260">値</span><span class="sxs-lookup"><span data-stu-id="35291-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-261">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-261">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-262">1.1</span><span class="sxs-lookup"><span data-stu-id="35291-262">1.1</span></span>|
|[<span data-ttu-id="35291-263">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-263">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-264">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-264">ReadItem</span></span>|
|[<span data-ttu-id="35291-265">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-265">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-266">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-266">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="35291-267">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="35291-267">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="35291-268">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="35291-268">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="35291-269">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="35291-269">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="35291-270">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="35291-270">Read mode</span></span>

<span data-ttu-id="35291-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="35291-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="35291-273">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="35291-273">Compose mode</span></span>

<span data-ttu-id="35291-274">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-274">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-275">型:</span><span class="sxs-lookup"><span data-stu-id="35291-275">Type:</span></span>

*   <span data-ttu-id="35291-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="35291-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-277">要件</span><span class="sxs-lookup"><span data-stu-id="35291-277">Requirements</span></span>

|<span data-ttu-id="35291-278">要件</span><span class="sxs-lookup"><span data-stu-id="35291-278">Requirement</span></span>|<span data-ttu-id="35291-279">値</span><span class="sxs-lookup"><span data-stu-id="35291-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-280">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-281">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-281">1.0</span></span>|
|[<span data-ttu-id="35291-282">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-283">ReadItem</span></span>|
|[<span data-ttu-id="35291-284">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-285">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-285">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-286">例</span><span class="sxs-lookup"><span data-stu-id="35291-286">Example</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="35291-287">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="35291-287">(nullable) conversationId :String</span></span>

<span data-ttu-id="35291-288">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-288">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="35291-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="35291-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="35291-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="35291-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-293">型:</span><span class="sxs-lookup"><span data-stu-id="35291-293">Type:</span></span>

*   <span data-ttu-id="35291-294">String</span><span class="sxs-lookup"><span data-stu-id="35291-294">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-295">要件</span><span class="sxs-lookup"><span data-stu-id="35291-295">Requirements</span></span>

|<span data-ttu-id="35291-296">要件</span><span class="sxs-lookup"><span data-stu-id="35291-296">Requirement</span></span>|<span data-ttu-id="35291-297">値</span><span class="sxs-lookup"><span data-stu-id="35291-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-298">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-299">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-299">1.0</span></span>|
|[<span data-ttu-id="35291-300">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-300">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-301">ReadItem</span></span>|
|[<span data-ttu-id="35291-302">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-302">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-303">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-303">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="35291-304">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="35291-304">dateTimeCreated :Date</span></span>

<span data-ttu-id="35291-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="35291-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-307">型:</span><span class="sxs-lookup"><span data-stu-id="35291-307">Type:</span></span>

*   <span data-ttu-id="35291-308">日付</span><span class="sxs-lookup"><span data-stu-id="35291-308">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-309">要件</span><span class="sxs-lookup"><span data-stu-id="35291-309">Requirements</span></span>

|<span data-ttu-id="35291-310">要件</span><span class="sxs-lookup"><span data-stu-id="35291-310">Requirement</span></span>|<span data-ttu-id="35291-311">値</span><span class="sxs-lookup"><span data-stu-id="35291-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-312">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-313">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-313">1.0</span></span>|
|[<span data-ttu-id="35291-314">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-315">ReadItem</span></span>|
|[<span data-ttu-id="35291-316">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-317">読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-317">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-318">例</span><span class="sxs-lookup"><span data-stu-id="35291-318">Example</span></span>

```javascript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="35291-319">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="35291-319">dateTimeModified :Date</span></span>

<span data-ttu-id="35291-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="35291-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="35291-322">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="35291-322">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-323">型:</span><span class="sxs-lookup"><span data-stu-id="35291-323">Type:</span></span>

*   <span data-ttu-id="35291-324">日付</span><span class="sxs-lookup"><span data-stu-id="35291-324">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-325">要件</span><span class="sxs-lookup"><span data-stu-id="35291-325">Requirements</span></span>

|<span data-ttu-id="35291-326">要件</span><span class="sxs-lookup"><span data-stu-id="35291-326">Requirement</span></span>|<span data-ttu-id="35291-327">値</span><span class="sxs-lookup"><span data-stu-id="35291-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-328">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-329">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-329">1.0</span></span>|
|[<span data-ttu-id="35291-330">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-330">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-331">ReadItem</span></span>|
|[<span data-ttu-id="35291-332">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-332">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-333">読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-334">例</span><span class="sxs-lookup"><span data-stu-id="35291-334">Example</span></span>

```javascript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="35291-335">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="35291-335">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="35291-336">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="35291-336">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="35291-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="35291-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="35291-339">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="35291-339">Read mode</span></span>

<span data-ttu-id="35291-340">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-340">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="35291-341">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="35291-341">Compose mode</span></span>

<span data-ttu-id="35291-342">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-342">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="35291-343">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="35291-343">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-344">型:</span><span class="sxs-lookup"><span data-stu-id="35291-344">Type:</span></span>

*   <span data-ttu-id="35291-345">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="35291-345">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-346">要件</span><span class="sxs-lookup"><span data-stu-id="35291-346">Requirements</span></span>

|<span data-ttu-id="35291-347">要件</span><span class="sxs-lookup"><span data-stu-id="35291-347">Requirement</span></span>|<span data-ttu-id="35291-348">値</span><span class="sxs-lookup"><span data-stu-id="35291-348">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-349">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-349">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-350">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-350">1.0</span></span>|
|[<span data-ttu-id="35291-351">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-351">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-352">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-352">ReadItem</span></span>|
|[<span data-ttu-id="35291-353">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-353">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-354">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-354">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-355">例</span><span class="sxs-lookup"><span data-stu-id="35291-355">Example</span></span>

<span data-ttu-id="35291-356">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="35291-356">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="35291-357">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="35291-357">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="35291-358">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-358">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="35291-p112">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="35291-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="35291-361">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="35291-361">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="35291-362">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="35291-362">Read mode</span></span>

<span data-ttu-id="35291-363">`from` プロパティは `EmailAddressDetails` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-363">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="35291-364">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="35291-364">Compose mode</span></span>

<span data-ttu-id="35291-365">`from` プロパティは From 値を取得するメソッドを提供する `From` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-365">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="35291-366">型:</span><span class="sxs-lookup"><span data-stu-id="35291-366">Type:</span></span>

*   <span data-ttu-id="35291-367">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="35291-367">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-368">要件</span><span class="sxs-lookup"><span data-stu-id="35291-368">Requirements</span></span>

|<span data-ttu-id="35291-369">要件</span><span class="sxs-lookup"><span data-stu-id="35291-369">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="35291-370">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-370">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-371">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-371">1.0</span></span>|<span data-ttu-id="35291-372">1.7</span><span class="sxs-lookup"><span data-stu-id="35291-372">1.7</span></span>|
|[<span data-ttu-id="35291-373">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-373">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-374">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-374">ReadItem</span></span>|<span data-ttu-id="35291-375">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="35291-375">ReadWriteItem</span></span>|
|[<span data-ttu-id="35291-376">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-376">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-377">Read</span><span class="sxs-lookup"><span data-stu-id="35291-377">Read</span></span>|<span data-ttu-id="35291-378">Compose</span><span class="sxs-lookup"><span data-stu-id="35291-378">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="35291-379">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="35291-379">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="35291-380">メッセージのインターネット ヘッダーを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="35291-380">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-381">型:</span><span class="sxs-lookup"><span data-stu-id="35291-381">Type:</span></span>

*   [<span data-ttu-id="35291-382">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="35291-382">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="35291-383">要件</span><span class="sxs-lookup"><span data-stu-id="35291-383">Requirements</span></span>

|<span data-ttu-id="35291-384">要件</span><span class="sxs-lookup"><span data-stu-id="35291-384">Requirement</span></span>|<span data-ttu-id="35291-385">値</span><span class="sxs-lookup"><span data-stu-id="35291-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-386">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-387">プレビュー</span><span class="sxs-lookup"><span data-stu-id="35291-387">Preview</span></span>|
|[<span data-ttu-id="35291-388">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-389">ReadItem</span></span>|
|[<span data-ttu-id="35291-390">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-391">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="35291-391">Compose or read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="35291-392">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="35291-392">internetMessageId :String</span></span>

<span data-ttu-id="35291-p113">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="35291-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-395">型:</span><span class="sxs-lookup"><span data-stu-id="35291-395">Type:</span></span>

*   <span data-ttu-id="35291-396">String</span><span class="sxs-lookup"><span data-stu-id="35291-396">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-397">要件</span><span class="sxs-lookup"><span data-stu-id="35291-397">Requirements</span></span>

|<span data-ttu-id="35291-398">要件</span><span class="sxs-lookup"><span data-stu-id="35291-398">Requirement</span></span>|<span data-ttu-id="35291-399">値</span><span class="sxs-lookup"><span data-stu-id="35291-399">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-400">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-401">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-401">1.0</span></span>|
|[<span data-ttu-id="35291-402">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-402">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-403">ReadItem</span></span>|
|[<span data-ttu-id="35291-404">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-404">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-405">読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-405">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-406">例</span><span class="sxs-lookup"><span data-stu-id="35291-406">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="35291-407">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="35291-407">itemClass :String</span></span>

<span data-ttu-id="35291-p114">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="35291-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="35291-p115">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="35291-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="35291-412">型</span><span class="sxs-lookup"><span data-stu-id="35291-412">Type</span></span>|<span data-ttu-id="35291-413">説明</span><span class="sxs-lookup"><span data-stu-id="35291-413">Description</span></span>|<span data-ttu-id="35291-414">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="35291-414">item class</span></span>|
|---|---|---|
|<span data-ttu-id="35291-415">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="35291-415">Appointment items</span></span>|<span data-ttu-id="35291-416">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="35291-416">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="35291-417">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="35291-417">Message items</span></span>|<span data-ttu-id="35291-418">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="35291-418">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="35291-419">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="35291-419">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-420">型:</span><span class="sxs-lookup"><span data-stu-id="35291-420">Type:</span></span>

*   <span data-ttu-id="35291-421">String</span><span class="sxs-lookup"><span data-stu-id="35291-421">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-422">要件</span><span class="sxs-lookup"><span data-stu-id="35291-422">Requirements</span></span>

|<span data-ttu-id="35291-423">要件</span><span class="sxs-lookup"><span data-stu-id="35291-423">Requirement</span></span>|<span data-ttu-id="35291-424">値</span><span class="sxs-lookup"><span data-stu-id="35291-424">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-425">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-425">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-426">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-426">1.0</span></span>|
|[<span data-ttu-id="35291-427">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-427">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-428">ReadItem</span></span>|
|[<span data-ttu-id="35291-429">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-429">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-430">読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-430">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-431">例</span><span class="sxs-lookup"><span data-stu-id="35291-431">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="35291-432">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="35291-432">(nullable) itemId :String</span></span>

<span data-ttu-id="35291-p116">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="35291-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="35291-435">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="35291-435">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="35291-436">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="35291-436">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="35291-437">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="35291-437">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="35291-438">詳細は、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="35291-438">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="35291-p118">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="35291-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-441">種類:</span><span class="sxs-lookup"><span data-stu-id="35291-441">Type:</span></span>

*   <span data-ttu-id="35291-442">String</span><span class="sxs-lookup"><span data-stu-id="35291-442">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-443">要件</span><span class="sxs-lookup"><span data-stu-id="35291-443">Requirements</span></span>

|<span data-ttu-id="35291-444">要件</span><span class="sxs-lookup"><span data-stu-id="35291-444">Requirement</span></span>|<span data-ttu-id="35291-445">値</span><span class="sxs-lookup"><span data-stu-id="35291-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-446">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-447">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-447">1.0</span></span>|
|[<span data-ttu-id="35291-448">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-449">ReadItem</span></span>|
|[<span data-ttu-id="35291-450">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-451">読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-452">例</span><span class="sxs-lookup"><span data-stu-id="35291-452">Example</span></span>

<span data-ttu-id="35291-p119">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="35291-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="35291-455">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="35291-455">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="35291-456">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-456">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="35291-457">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="35291-457">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-458">型:</span><span class="sxs-lookup"><span data-stu-id="35291-458">Type:</span></span>

*   [<span data-ttu-id="35291-459">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="35291-459">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="35291-460">要件</span><span class="sxs-lookup"><span data-stu-id="35291-460">Requirements</span></span>

|<span data-ttu-id="35291-461">要件</span><span class="sxs-lookup"><span data-stu-id="35291-461">Requirement</span></span>|<span data-ttu-id="35291-462">値</span><span class="sxs-lookup"><span data-stu-id="35291-462">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-463">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-463">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-464">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-464">1.0</span></span>|
|[<span data-ttu-id="35291-465">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-465">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-466">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-466">ReadItem</span></span>|
|[<span data-ttu-id="35291-467">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-467">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-468">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-468">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-469">例</span><span class="sxs-lookup"><span data-stu-id="35291-469">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="35291-470">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="35291-470">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="35291-471">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="35291-471">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="35291-472">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="35291-472">Read mode</span></span>

<span data-ttu-id="35291-473">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="35291-473">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="35291-474">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="35291-474">Compose mode</span></span>

<span data-ttu-id="35291-475">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-475">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-476">型:</span><span class="sxs-lookup"><span data-stu-id="35291-476">Type:</span></span>

*   <span data-ttu-id="35291-477">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="35291-477">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-478">要件</span><span class="sxs-lookup"><span data-stu-id="35291-478">Requirements</span></span>

|<span data-ttu-id="35291-479">要件</span><span class="sxs-lookup"><span data-stu-id="35291-479">Requirement</span></span>|<span data-ttu-id="35291-480">値</span><span class="sxs-lookup"><span data-stu-id="35291-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-481">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-482">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-482">1.0</span></span>|
|[<span data-ttu-id="35291-483">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-484">ReadItem</span></span>|
|[<span data-ttu-id="35291-485">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-486">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-487">例</span><span class="sxs-lookup"><span data-stu-id="35291-487">Example</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="35291-488">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="35291-488">normalizedSubject :String</span></span>

<span data-ttu-id="35291-p120">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="35291-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="35291-p121">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="35291-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-493">型:</span><span class="sxs-lookup"><span data-stu-id="35291-493">Type:</span></span>

*   <span data-ttu-id="35291-494">String</span><span class="sxs-lookup"><span data-stu-id="35291-494">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-495">要件</span><span class="sxs-lookup"><span data-stu-id="35291-495">Requirements</span></span>

|<span data-ttu-id="35291-496">要件</span><span class="sxs-lookup"><span data-stu-id="35291-496">Requirement</span></span>|<span data-ttu-id="35291-497">値</span><span class="sxs-lookup"><span data-stu-id="35291-497">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-498">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-498">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-499">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-499">1.0</span></span>|
|[<span data-ttu-id="35291-500">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-500">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-501">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-501">ReadItem</span></span>|
|[<span data-ttu-id="35291-502">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-502">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-503">読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-503">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-504">例</span><span class="sxs-lookup"><span data-stu-id="35291-504">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="35291-505">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="35291-505">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="35291-506">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-506">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-507">型:</span><span class="sxs-lookup"><span data-stu-id="35291-507">Type:</span></span>

*   [<span data-ttu-id="35291-508">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="35291-508">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="35291-509">要件</span><span class="sxs-lookup"><span data-stu-id="35291-509">Requirements</span></span>

|<span data-ttu-id="35291-510">要件</span><span class="sxs-lookup"><span data-stu-id="35291-510">Requirement</span></span>|<span data-ttu-id="35291-511">値</span><span class="sxs-lookup"><span data-stu-id="35291-511">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-512">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-512">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-513">1.3</span><span class="sxs-lookup"><span data-stu-id="35291-513">1.3</span></span>|
|[<span data-ttu-id="35291-514">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-514">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-515">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-515">ReadItem</span></span>|
|[<span data-ttu-id="35291-516">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-516">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-517">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-517">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="35291-518">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="35291-518">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="35291-519">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="35291-519">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="35291-520">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="35291-520">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="35291-521">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="35291-521">Read mode</span></span>

<span data-ttu-id="35291-522">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="35291-522">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="35291-523">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="35291-523">Compose mode</span></span>

<span data-ttu-id="35291-524">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-524">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-525">型:</span><span class="sxs-lookup"><span data-stu-id="35291-525">Type:</span></span>

*   <span data-ttu-id="35291-526">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="35291-526">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-527">要件</span><span class="sxs-lookup"><span data-stu-id="35291-527">Requirements</span></span>

|<span data-ttu-id="35291-528">要件</span><span class="sxs-lookup"><span data-stu-id="35291-528">Requirement</span></span>|<span data-ttu-id="35291-529">値</span><span class="sxs-lookup"><span data-stu-id="35291-529">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-530">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-530">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-531">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-531">1.0</span></span>|
|[<span data-ttu-id="35291-532">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-532">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-533">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-533">ReadItem</span></span>|
|[<span data-ttu-id="35291-534">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-534">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-535">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-535">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-536">例</span><span class="sxs-lookup"><span data-stu-id="35291-536">Example</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="35291-537">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="35291-537">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="35291-538">指定の会議の開催者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-538">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="35291-539">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="35291-539">Read mode</span></span>

<span data-ttu-id="35291-540">`organizer` プロパティは、会議開催者を表す [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-540">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="35291-541">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="35291-541">Compose mode</span></span>

<span data-ttu-id="35291-542">`organizer` プロパティは Organizer 値を取得するメソッドを提供する [Organizer](/javascript/api/outlook/office.organizer) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-542">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-543">型:</span><span class="sxs-lookup"><span data-stu-id="35291-543">Type:</span></span>

*   <span data-ttu-id="35291-544">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="35291-544">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-545">要件</span><span class="sxs-lookup"><span data-stu-id="35291-545">Requirements</span></span>

|<span data-ttu-id="35291-546">要件</span><span class="sxs-lookup"><span data-stu-id="35291-546">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="35291-547">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-547">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-548">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-548">1.0</span></span>|<span data-ttu-id="35291-549">1.7</span><span class="sxs-lookup"><span data-stu-id="35291-549">1.7</span></span>|
|[<span data-ttu-id="35291-550">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-550">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-551">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-551">ReadItem</span></span>|<span data-ttu-id="35291-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="35291-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="35291-553">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-553">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-554">Read</span><span class="sxs-lookup"><span data-stu-id="35291-554">Read</span></span>|<span data-ttu-id="35291-555">Compose</span><span class="sxs-lookup"><span data-stu-id="35291-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-556">例</span><span class="sxs-lookup"><span data-stu-id="35291-556">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="35291-557">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="35291-557">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="35291-558">予定の定期的なパターンを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="35291-558">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="35291-559">会議出席依頼の定期的なパターンを取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-559">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="35291-560">予定アイテムの閲覧モードと新規作成モード。</span><span class="sxs-lookup"><span data-stu-id="35291-560">Read and compose modes for appointment items.</span></span> <span data-ttu-id="35291-561">会議出席依頼アイテムの閲覧モード。</span><span class="sxs-lookup"><span data-stu-id="35291-561">Read mode for meeting request items.</span></span>

<span data-ttu-id="35291-562">`recurrence` プロパティは、アイテムがシリーズか、シリーズに含まれるインスタンスの場合、定期的な予定または会議出席依頼に対して [recurrence](/javascript/api/outlook/office.recurrence) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-562">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="35291-563">`null` は、単発の予定および単発の予定の会議出席依頼に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="35291-563">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="35291-564">`undefined` は、会議出席依頼ではないメッセージに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="35291-564">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="35291-565">注: 会議出席依頼の `itemClass` 値は IPM.Schedule.Meeting.Request です。</span><span class="sxs-lookup"><span data-stu-id="35291-565">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="35291-566">注: recurrence オブジェクトが `null` の場合、オブジェクトがシリーズの一部ではなく、1 つの単発の予定または 1 つの単発の予定の会議出席依頼であることを示します。</span><span class="sxs-lookup"><span data-stu-id="35291-566">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-567">型:</span><span class="sxs-lookup"><span data-stu-id="35291-567">Type:</span></span>

* [<span data-ttu-id="35291-568">Recurrence</span><span class="sxs-lookup"><span data-stu-id="35291-568">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="35291-569">要件</span><span class="sxs-lookup"><span data-stu-id="35291-569">Requirement</span></span>|<span data-ttu-id="35291-570">値</span><span class="sxs-lookup"><span data-stu-id="35291-570">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-571">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-571">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-572">1.7</span><span class="sxs-lookup"><span data-stu-id="35291-572">1.7</span></span>|
|[<span data-ttu-id="35291-573">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-573">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-574">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-574">ReadItem</span></span>|
|[<span data-ttu-id="35291-575">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-575">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-576">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="35291-576">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="35291-577">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="35291-577">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="35291-578">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="35291-578">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="35291-579">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="35291-579">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="35291-580">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="35291-580">Read mode</span></span>

<span data-ttu-id="35291-581">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="35291-581">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="35291-582">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="35291-582">Compose mode</span></span>

<span data-ttu-id="35291-583">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-583">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-584">型:</span><span class="sxs-lookup"><span data-stu-id="35291-584">Type:</span></span>

*   <span data-ttu-id="35291-585">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="35291-585">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-586">要件</span><span class="sxs-lookup"><span data-stu-id="35291-586">Requirements</span></span>

|<span data-ttu-id="35291-587">要件</span><span class="sxs-lookup"><span data-stu-id="35291-587">Requirement</span></span>|<span data-ttu-id="35291-588">値</span><span class="sxs-lookup"><span data-stu-id="35291-588">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-589">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-589">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-590">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-590">1.0</span></span>|
|[<span data-ttu-id="35291-591">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-591">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-592">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-592">ReadItem</span></span>|
|[<span data-ttu-id="35291-593">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-593">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-594">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-594">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-595">例</span><span class="sxs-lookup"><span data-stu-id="35291-595">Example</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="35291-596">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="35291-596">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="35291-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="35291-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="35291-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="35291-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="35291-601">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="35291-601">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-602">型:</span><span class="sxs-lookup"><span data-stu-id="35291-602">Type:</span></span>

*   [<span data-ttu-id="35291-603">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="35291-603">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="35291-604">要件</span><span class="sxs-lookup"><span data-stu-id="35291-604">Requirements</span></span>

|<span data-ttu-id="35291-605">要件</span><span class="sxs-lookup"><span data-stu-id="35291-605">Requirement</span></span>|<span data-ttu-id="35291-606">値</span><span class="sxs-lookup"><span data-stu-id="35291-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-607">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-608">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-608">1.0</span></span>|
|[<span data-ttu-id="35291-609">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-610">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-610">ReadItem</span></span>|
|[<span data-ttu-id="35291-611">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-612">読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-612">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-613">例</span><span class="sxs-lookup"><span data-stu-id="35291-613">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="35291-614">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="35291-614">(nullable) seriesId :String</span></span>

<span data-ttu-id="35291-615">あるインスタンスが属するシリーズの ID を取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-615">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="35291-616">OWA と Outlook では、`seriesId` はこのアイテムが属する親 (シリーズ) アイテムの Exchange Web Services (EWS) ID を返します。</span><span class="sxs-lookup"><span data-stu-id="35291-616">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="35291-617">ただし、iOS と Android の場合、`seriesId` は親アイテムの REST ID を返します。</span><span class="sxs-lookup"><span data-stu-id="35291-617">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="35291-618">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="35291-618">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="35291-619">`seriesId` プロパティは、Outlook REST API で使用される Outlook ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="35291-619">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="35291-620">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="35291-620">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="35291-621">詳細については、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="35291-621">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="35291-622">`seriesId` プロパティは、単発の予定、シリーズ アイテム、会議出席依頼など、親アイテムを持たないアイテムに対して `null` を返し、会議出席依頼ではないその他のアイテムに対して `undefined` を返します。</span><span class="sxs-lookup"><span data-stu-id="35291-622">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-623">型:</span><span class="sxs-lookup"><span data-stu-id="35291-623">Type:</span></span>

* <span data-ttu-id="35291-624">String</span><span class="sxs-lookup"><span data-stu-id="35291-624">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-625">要件</span><span class="sxs-lookup"><span data-stu-id="35291-625">Requirements</span></span>

|<span data-ttu-id="35291-626">要件</span><span class="sxs-lookup"><span data-stu-id="35291-626">Requirement</span></span>|<span data-ttu-id="35291-627">値</span><span class="sxs-lookup"><span data-stu-id="35291-627">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-628">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-628">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-629">1.7</span><span class="sxs-lookup"><span data-stu-id="35291-629">1.7</span></span>|
|[<span data-ttu-id="35291-630">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-630">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-631">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-631">ReadItem</span></span>|
|[<span data-ttu-id="35291-632">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-632">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-633">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-633">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-634">例</span><span class="sxs-lookup"><span data-stu-id="35291-634">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="35291-635">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="35291-635">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="35291-636">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="35291-636">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="35291-p130">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="35291-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="35291-639">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="35291-639">Read mode</span></span>

<span data-ttu-id="35291-640">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-640">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="35291-641">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="35291-641">Compose mode</span></span>

<span data-ttu-id="35291-642">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-642">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="35291-643">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="35291-643">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-644">型:</span><span class="sxs-lookup"><span data-stu-id="35291-644">Type:</span></span>

*   <span data-ttu-id="35291-645">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="35291-645">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-646">要件</span><span class="sxs-lookup"><span data-stu-id="35291-646">Requirements</span></span>

|<span data-ttu-id="35291-647">要件</span><span class="sxs-lookup"><span data-stu-id="35291-647">Requirement</span></span>|<span data-ttu-id="35291-648">値</span><span class="sxs-lookup"><span data-stu-id="35291-648">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-649">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-649">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-650">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-650">1.0</span></span>|
|[<span data-ttu-id="35291-651">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-651">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-652">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-652">ReadItem</span></span>|
|[<span data-ttu-id="35291-653">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-653">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-654">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-654">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-655">例</span><span class="sxs-lookup"><span data-stu-id="35291-655">Example</span></span>

<span data-ttu-id="35291-656">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="35291-656">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="35291-657">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="35291-657">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="35291-658">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="35291-658">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="35291-659">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="35291-659">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="35291-660">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="35291-660">Read mode</span></span>

<span data-ttu-id="35291-p131">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="35291-663">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="35291-663">Compose mode</span></span>

<span data-ttu-id="35291-664">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-664">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="35291-665">型:</span><span class="sxs-lookup"><span data-stu-id="35291-665">Type:</span></span>

*   <span data-ttu-id="35291-666">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="35291-666">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-667">要件</span><span class="sxs-lookup"><span data-stu-id="35291-667">Requirements</span></span>

|<span data-ttu-id="35291-668">要件</span><span class="sxs-lookup"><span data-stu-id="35291-668">Requirement</span></span>|<span data-ttu-id="35291-669">値</span><span class="sxs-lookup"><span data-stu-id="35291-669">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-670">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-670">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-671">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-671">1.0</span></span>|
|[<span data-ttu-id="35291-672">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-672">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-673">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-673">ReadItem</span></span>|
|[<span data-ttu-id="35291-674">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-674">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-675">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-675">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="35291-676">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="35291-676">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="35291-677">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="35291-677">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="35291-678">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="35291-678">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="35291-679">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="35291-679">Read mode</span></span>

<span data-ttu-id="35291-p133">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="35291-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="35291-682">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="35291-682">Compose mode</span></span>

<span data-ttu-id="35291-683">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-683">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="35291-684">型:</span><span class="sxs-lookup"><span data-stu-id="35291-684">Type:</span></span>

*   <span data-ttu-id="35291-685">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="35291-685">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-686">要件</span><span class="sxs-lookup"><span data-stu-id="35291-686">Requirements</span></span>

|<span data-ttu-id="35291-687">要件</span><span class="sxs-lookup"><span data-stu-id="35291-687">Requirement</span></span>|<span data-ttu-id="35291-688">値</span><span class="sxs-lookup"><span data-stu-id="35291-688">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-689">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-689">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-690">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-690">1.0</span></span>|
|[<span data-ttu-id="35291-691">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-691">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-692">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-692">ReadItem</span></span>|
|[<span data-ttu-id="35291-693">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-693">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-694">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-694">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-695">例</span><span class="sxs-lookup"><span data-stu-id="35291-695">Example</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="35291-696">メソッド</span><span class="sxs-lookup"><span data-stu-id="35291-696">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="35291-697">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="35291-697">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="35291-698">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="35291-698">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="35291-699">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="35291-699">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="35291-700">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="35291-700">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-701">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-701">Parameters:</span></span>
|<span data-ttu-id="35291-702">名前</span><span class="sxs-lookup"><span data-stu-id="35291-702">Name</span></span>|<span data-ttu-id="35291-703">型</span><span class="sxs-lookup"><span data-stu-id="35291-703">Type</span></span>|<span data-ttu-id="35291-704">属性</span><span class="sxs-lookup"><span data-stu-id="35291-704">Attributes</span></span>|<span data-ttu-id="35291-705">説明</span><span class="sxs-lookup"><span data-stu-id="35291-705">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="35291-706">String</span><span class="sxs-lookup"><span data-stu-id="35291-706">String</span></span>||<span data-ttu-id="35291-p134">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="35291-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="35291-709">String</span><span class="sxs-lookup"><span data-stu-id="35291-709">String</span></span>||<span data-ttu-id="35291-p135">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="35291-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="35291-712">Object</span><span class="sxs-lookup"><span data-stu-id="35291-712">Object</span></span>|<span data-ttu-id="35291-713">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-713">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-714">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="35291-714">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="35291-715">Object</span><span class="sxs-lookup"><span data-stu-id="35291-715">Object</span></span>|<span data-ttu-id="35291-716">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-716">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-717">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="35291-717">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="35291-718">Boolean</span><span class="sxs-lookup"><span data-stu-id="35291-718">Boolean</span></span>|<span data-ttu-id="35291-719">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-719">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-720">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="35291-720">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="35291-721">function</span><span class="sxs-lookup"><span data-stu-id="35291-721">function</span></span>|<span data-ttu-id="35291-722">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-722">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-723">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="35291-723">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="35291-724">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="35291-724">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="35291-725">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="35291-725">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="35291-726">エラー</span><span class="sxs-lookup"><span data-stu-id="35291-726">Errors</span></span>

|<span data-ttu-id="35291-727">エラー コード</span><span class="sxs-lookup"><span data-stu-id="35291-727">Error code</span></span>|<span data-ttu-id="35291-728">説明</span><span class="sxs-lookup"><span data-stu-id="35291-728">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="35291-729">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="35291-729">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="35291-730">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="35291-730">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="35291-731">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="35291-731">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-732">要件</span><span class="sxs-lookup"><span data-stu-id="35291-732">Requirements</span></span>

|<span data-ttu-id="35291-733">要件</span><span class="sxs-lookup"><span data-stu-id="35291-733">Requirement</span></span>|<span data-ttu-id="35291-734">値</span><span class="sxs-lookup"><span data-stu-id="35291-734">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-735">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-735">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-736">1.1</span><span class="sxs-lookup"><span data-stu-id="35291-736">1.1</span></span>|
|[<span data-ttu-id="35291-737">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-737">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-738">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="35291-738">ReadWriteItem</span></span>|
|[<span data-ttu-id="35291-739">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-739">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-740">作成</span><span class="sxs-lookup"><span data-stu-id="35291-740">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="35291-741">例</span><span class="sxs-lookup"><span data-stu-id="35291-741">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="35291-742">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="35291-742">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
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
        
      }
    );
  }
);
```

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="35291-743">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="35291-743">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="35291-744">ファイルを添付ファイルとして base64 エンコーディングからメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="35291-744">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="35291-745">`addFileAttachmentFromBase64Async` メソッドは、base64 エンコーディングからファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="35291-745">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="35291-746">このメソッドによって、AsyncResult.value オブジェクトの添付ファイル識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="35291-746">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="35291-747">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="35291-747">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-748">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-748">Parameters:</span></span>
|<span data-ttu-id="35291-749">名前</span><span class="sxs-lookup"><span data-stu-id="35291-749">Name</span></span>|<span data-ttu-id="35291-750">型</span><span class="sxs-lookup"><span data-stu-id="35291-750">Type</span></span>|<span data-ttu-id="35291-751">属性</span><span class="sxs-lookup"><span data-stu-id="35291-751">Attributes</span></span>|<span data-ttu-id="35291-752">説明</span><span class="sxs-lookup"><span data-stu-id="35291-752">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="35291-753">String</span><span class="sxs-lookup"><span data-stu-id="35291-753">String</span></span>||<span data-ttu-id="35291-754">電子メールまたはイベントに追加する画像またはファイルの base64 でエンコードされたコンテンツ。</span><span class="sxs-lookup"><span data-stu-id="35291-754">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="35291-755">String</span><span class="sxs-lookup"><span data-stu-id="35291-755">String</span></span>||<span data-ttu-id="35291-p137">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="35291-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="35291-758">Object</span><span class="sxs-lookup"><span data-stu-id="35291-758">Object</span></span>|<span data-ttu-id="35291-759">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-759">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-760">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="35291-760">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="35291-761">Object</span><span class="sxs-lookup"><span data-stu-id="35291-761">Object</span></span>|<span data-ttu-id="35291-762">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-762">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-763">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="35291-763">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="35291-764">Boolean</span><span class="sxs-lookup"><span data-stu-id="35291-764">Boolean</span></span>|<span data-ttu-id="35291-765">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-765">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-766">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="35291-766">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="35291-767">function</span><span class="sxs-lookup"><span data-stu-id="35291-767">function</span></span>|<span data-ttu-id="35291-768">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-768">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-769">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="35291-769">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="35291-770">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="35291-770">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="35291-771">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="35291-771">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="35291-772">エラー</span><span class="sxs-lookup"><span data-stu-id="35291-772">Errors</span></span>

|<span data-ttu-id="35291-773">エラー コード</span><span class="sxs-lookup"><span data-stu-id="35291-773">Error code</span></span>|<span data-ttu-id="35291-774">説明</span><span class="sxs-lookup"><span data-stu-id="35291-774">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="35291-775">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="35291-775">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="35291-776">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="35291-776">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="35291-777">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="35291-777">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-778">要件</span><span class="sxs-lookup"><span data-stu-id="35291-778">Requirements</span></span>

|<span data-ttu-id="35291-779">要件</span><span class="sxs-lookup"><span data-stu-id="35291-779">Requirement</span></span>|<span data-ttu-id="35291-780">値</span><span class="sxs-lookup"><span data-stu-id="35291-780">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-781">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-782">プレビュー</span><span class="sxs-lookup"><span data-stu-id="35291-782">Preview</span></span>|
|[<span data-ttu-id="35291-783">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-783">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-784">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="35291-784">ReadWriteItem</span></span>|
|[<span data-ttu-id="35291-785">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-785">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-786">作成</span><span class="sxs-lookup"><span data-stu-id="35291-786">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="35291-787">例</span><span class="sxs-lookup"><span data-stu-id="35291-787">Examples</span></span>

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
      }
    );
  }
);
```

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="35291-788">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="35291-788">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="35291-789">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="35291-789">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="35291-790">現在、サポートされているイベントの種類は `Office.EventType.AttachmentsChanged`、`Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged`、`Office.EventType.RecurrenceChanged` です。</span><span class="sxs-lookup"><span data-stu-id="35291-790">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-791">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-791">Parameters:</span></span>

| <span data-ttu-id="35291-792">名前</span><span class="sxs-lookup"><span data-stu-id="35291-792">Name</span></span> | <span data-ttu-id="35291-793">型</span><span class="sxs-lookup"><span data-stu-id="35291-793">Type</span></span> | <span data-ttu-id="35291-794">属性</span><span class="sxs-lookup"><span data-stu-id="35291-794">Attributes</span></span> | <span data-ttu-id="35291-795">説明</span><span class="sxs-lookup"><span data-stu-id="35291-795">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="35291-796">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="35291-796">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="35291-797">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="35291-797">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="35291-798">Function</span><span class="sxs-lookup"><span data-stu-id="35291-798">Function</span></span> || <span data-ttu-id="35291-p138">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="35291-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="35291-802">Object</span><span class="sxs-lookup"><span data-stu-id="35291-802">Object</span></span> | <span data-ttu-id="35291-803">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-803">&lt;optional&gt;</span></span> | <span data-ttu-id="35291-804">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="35291-804">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="35291-805">Object</span><span class="sxs-lookup"><span data-stu-id="35291-805">Object</span></span> | <span data-ttu-id="35291-806">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-806">&lt;optional&gt;</span></span> | <span data-ttu-id="35291-807">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="35291-807">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="35291-808">function</span><span class="sxs-lookup"><span data-stu-id="35291-808">function</span></span>| <span data-ttu-id="35291-809">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-809">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-810">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="35291-810">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-811">要件</span><span class="sxs-lookup"><span data-stu-id="35291-811">Requirements</span></span>

|<span data-ttu-id="35291-812">要件</span><span class="sxs-lookup"><span data-stu-id="35291-812">Requirement</span></span>| <span data-ttu-id="35291-813">値</span><span class="sxs-lookup"><span data-stu-id="35291-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-814">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35291-815">1.7</span><span class="sxs-lookup"><span data-stu-id="35291-815">1.7</span></span> |
|[<span data-ttu-id="35291-816">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35291-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-817">ReadItem</span></span> |
|[<span data-ttu-id="35291-818">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35291-819">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="35291-819">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="35291-820">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="35291-820">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="35291-821">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="35291-821">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="35291-p139">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="35291-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="35291-825">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="35291-825">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="35291-826">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="35291-826">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-827">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-827">Parameters:</span></span>

|<span data-ttu-id="35291-828">名前</span><span class="sxs-lookup"><span data-stu-id="35291-828">Name</span></span>|<span data-ttu-id="35291-829">型</span><span class="sxs-lookup"><span data-stu-id="35291-829">Type</span></span>|<span data-ttu-id="35291-830">属性</span><span class="sxs-lookup"><span data-stu-id="35291-830">Attributes</span></span>|<span data-ttu-id="35291-831">説明</span><span class="sxs-lookup"><span data-stu-id="35291-831">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="35291-832">String</span><span class="sxs-lookup"><span data-stu-id="35291-832">String</span></span>||<span data-ttu-id="35291-p140">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="35291-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="35291-835">String</span><span class="sxs-lookup"><span data-stu-id="35291-835">String</span></span>||<span data-ttu-id="35291-p141">添付するアイテムの件名。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="35291-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="35291-838">Object</span><span class="sxs-lookup"><span data-stu-id="35291-838">Object</span></span>|<span data-ttu-id="35291-839">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-839">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-840">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="35291-840">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="35291-841">Object</span><span class="sxs-lookup"><span data-stu-id="35291-841">Object</span></span>|<span data-ttu-id="35291-842">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-842">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-843">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="35291-843">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="35291-844">function</span><span class="sxs-lookup"><span data-stu-id="35291-844">function</span></span>|<span data-ttu-id="35291-845">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-845">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-846">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="35291-846">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="35291-847">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="35291-847">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="35291-848">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="35291-848">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="35291-849">エラー</span><span class="sxs-lookup"><span data-stu-id="35291-849">Errors</span></span>

|<span data-ttu-id="35291-850">エラー コード</span><span class="sxs-lookup"><span data-stu-id="35291-850">Error code</span></span>|<span data-ttu-id="35291-851">説明</span><span class="sxs-lookup"><span data-stu-id="35291-851">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="35291-852">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="35291-852">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-853">要件</span><span class="sxs-lookup"><span data-stu-id="35291-853">Requirements</span></span>

|<span data-ttu-id="35291-854">要件</span><span class="sxs-lookup"><span data-stu-id="35291-854">Requirement</span></span>|<span data-ttu-id="35291-855">値</span><span class="sxs-lookup"><span data-stu-id="35291-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-856">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-856">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-857">1.1</span><span class="sxs-lookup"><span data-stu-id="35291-857">1.1</span></span>|
|[<span data-ttu-id="35291-858">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-858">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-859">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="35291-859">ReadWriteItem</span></span>|
|[<span data-ttu-id="35291-860">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-860">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-861">作成</span><span class="sxs-lookup"><span data-stu-id="35291-861">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-862">例</span><span class="sxs-lookup"><span data-stu-id="35291-862">Example</span></span>

<span data-ttu-id="35291-863">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="35291-863">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```javascript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="35291-864">close()</span><span class="sxs-lookup"><span data-stu-id="35291-864">close()</span></span>

<span data-ttu-id="35291-865">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="35291-865">Closes the current item that is being composed.</span></span>

<span data-ttu-id="35291-p142">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="35291-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="35291-868">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="35291-868">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="35291-869">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="35291-869">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-870">要件</span><span class="sxs-lookup"><span data-stu-id="35291-870">Requirements</span></span>

|<span data-ttu-id="35291-871">要件</span><span class="sxs-lookup"><span data-stu-id="35291-871">Requirement</span></span>|<span data-ttu-id="35291-872">値</span><span class="sxs-lookup"><span data-stu-id="35291-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-873">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-873">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-874">1.3</span><span class="sxs-lookup"><span data-stu-id="35291-874">1.3</span></span>|
|[<span data-ttu-id="35291-875">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-876">制限あり</span><span class="sxs-lookup"><span data-stu-id="35291-876">Restricted</span></span>|
|[<span data-ttu-id="35291-877">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-878">作成</span><span class="sxs-lookup"><span data-stu-id="35291-878">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="35291-879">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="35291-879">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="35291-880">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="35291-880">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="35291-881">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="35291-881">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="35291-882">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="35291-882">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="35291-883">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="35291-883">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="35291-p143">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="35291-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-887">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-887">Parameters:</span></span>

|<span data-ttu-id="35291-888">名前</span><span class="sxs-lookup"><span data-stu-id="35291-888">Name</span></span>|<span data-ttu-id="35291-889">型</span><span class="sxs-lookup"><span data-stu-id="35291-889">Type</span></span>|<span data-ttu-id="35291-890">属性</span><span class="sxs-lookup"><span data-stu-id="35291-890">Attributes</span></span>|<span data-ttu-id="35291-891">説明</span><span class="sxs-lookup"><span data-stu-id="35291-891">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="35291-892">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="35291-892">String &#124; Object</span></span>||<span data-ttu-id="35291-p144">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="35291-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="35291-895">**または**</span><span class="sxs-lookup"><span data-stu-id="35291-895">**OR**</span></span><br/><span data-ttu-id="35291-p145">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="35291-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="35291-898">String</span><span class="sxs-lookup"><span data-stu-id="35291-898">String</span></span>|<span data-ttu-id="35291-899">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-899">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="35291-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="35291-902">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-902">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="35291-903">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-903">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-904">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="35291-904">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="35291-905">String</span><span class="sxs-lookup"><span data-stu-id="35291-905">String</span></span>||<span data-ttu-id="35291-p147">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="35291-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="35291-908">String</span><span class="sxs-lookup"><span data-stu-id="35291-908">String</span></span>||<span data-ttu-id="35291-909">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="35291-909">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="35291-910">String</span><span class="sxs-lookup"><span data-stu-id="35291-910">String</span></span>||<span data-ttu-id="35291-p148">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="35291-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="35291-913">Boolean</span><span class="sxs-lookup"><span data-stu-id="35291-913">Boolean</span></span>||<span data-ttu-id="35291-p149">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="35291-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="35291-916">String</span><span class="sxs-lookup"><span data-stu-id="35291-916">String</span></span>||<span data-ttu-id="35291-p150">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="35291-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="35291-920">function</span><span class="sxs-lookup"><span data-stu-id="35291-920">function</span></span>|<span data-ttu-id="35291-921">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-921">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-922">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="35291-922">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-923">要件</span><span class="sxs-lookup"><span data-stu-id="35291-923">Requirements</span></span>

|<span data-ttu-id="35291-924">要件</span><span class="sxs-lookup"><span data-stu-id="35291-924">Requirement</span></span>|<span data-ttu-id="35291-925">値</span><span class="sxs-lookup"><span data-stu-id="35291-925">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-926">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-926">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-927">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-927">1.0</span></span>|
|[<span data-ttu-id="35291-928">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-928">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-929">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-929">ReadItem</span></span>|
|[<span data-ttu-id="35291-930">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-930">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-931">読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-931">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="35291-932">例</span><span class="sxs-lookup"><span data-stu-id="35291-932">Examples</span></span>

<span data-ttu-id="35291-933">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="35291-933">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="35291-934">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="35291-934">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="35291-935">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="35291-935">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="35291-936">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="35291-936">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="35291-937">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="35291-937">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="35291-938">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="35291-938">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="35291-939">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="35291-939">displayReplyForm(formData)</span></span>

<span data-ttu-id="35291-940">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="35291-940">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="35291-941">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="35291-941">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="35291-942">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="35291-942">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="35291-943">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="35291-943">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="35291-p151">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="35291-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-947">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-947">Parameters:</span></span>

|<span data-ttu-id="35291-948">名前</span><span class="sxs-lookup"><span data-stu-id="35291-948">Name</span></span>|<span data-ttu-id="35291-949">型</span><span class="sxs-lookup"><span data-stu-id="35291-949">Type</span></span>|<span data-ttu-id="35291-950">属性</span><span class="sxs-lookup"><span data-stu-id="35291-950">Attributes</span></span>|<span data-ttu-id="35291-951">説明</span><span class="sxs-lookup"><span data-stu-id="35291-951">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="35291-952">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="35291-952">String &#124; Object</span></span>||<span data-ttu-id="35291-p152">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="35291-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="35291-955">**または**</span><span class="sxs-lookup"><span data-stu-id="35291-955">**OR**</span></span><br/><span data-ttu-id="35291-p153">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="35291-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="35291-958">String</span><span class="sxs-lookup"><span data-stu-id="35291-958">String</span></span>|<span data-ttu-id="35291-959">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-959">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-p154">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="35291-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="35291-962">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-962">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="35291-963">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-963">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-964">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="35291-964">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="35291-965">String</span><span class="sxs-lookup"><span data-stu-id="35291-965">String</span></span>||<span data-ttu-id="35291-p155">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="35291-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="35291-968">String</span><span class="sxs-lookup"><span data-stu-id="35291-968">String</span></span>||<span data-ttu-id="35291-969">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="35291-969">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="35291-970">String</span><span class="sxs-lookup"><span data-stu-id="35291-970">String</span></span>||<span data-ttu-id="35291-p156">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="35291-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="35291-973">Boolean</span><span class="sxs-lookup"><span data-stu-id="35291-973">Boolean</span></span>||<span data-ttu-id="35291-p157">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="35291-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="35291-976">String</span><span class="sxs-lookup"><span data-stu-id="35291-976">String</span></span>||<span data-ttu-id="35291-p158">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="35291-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="35291-980">function</span><span class="sxs-lookup"><span data-stu-id="35291-980">function</span></span>|<span data-ttu-id="35291-981">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-981">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-982">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="35291-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-983">要件</span><span class="sxs-lookup"><span data-stu-id="35291-983">Requirements</span></span>

|<span data-ttu-id="35291-984">要件</span><span class="sxs-lookup"><span data-stu-id="35291-984">Requirement</span></span>|<span data-ttu-id="35291-985">値</span><span class="sxs-lookup"><span data-stu-id="35291-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-986">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-987">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-987">1.0</span></span>|
|[<span data-ttu-id="35291-988">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-989">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-989">ReadItem</span></span>|
|[<span data-ttu-id="35291-990">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-991">読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-991">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="35291-992">例</span><span class="sxs-lookup"><span data-stu-id="35291-992">Examples</span></span>

<span data-ttu-id="35291-993">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="35291-993">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="35291-994">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="35291-994">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="35291-995">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="35291-995">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="35291-996">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="35291-996">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="35291-997">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="35291-997">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="35291-998">本文、ファイルの添付ファイル、アイテムの添付ファイル、コールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="35291-998">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="35291-999">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="35291-999">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="35291-1000">メッセージまたは予定から指定の添付ファイルを取得し、それを `AttachmentContent` オブジェクトとして返されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1000">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="35291-1001">`getAttachmentContentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-1001">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="35291-1002">ベスト プラクティスとして、識別子を使用し、`getAttachmentsAsync` または `item.attachments` 呼び出しで attachmentIds を取得した同じセッションで添付ファイルを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="35291-1002">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="35291-1003">Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="35291-1003">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="35291-1004">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始し、フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="35291-1004">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-1005">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-1005">Parameters:</span></span>

|<span data-ttu-id="35291-1006">名前</span><span class="sxs-lookup"><span data-stu-id="35291-1006">Name</span></span>|<span data-ttu-id="35291-1007">型</span><span class="sxs-lookup"><span data-stu-id="35291-1007">Type</span></span>|<span data-ttu-id="35291-1008">属性</span><span class="sxs-lookup"><span data-stu-id="35291-1008">Attributes</span></span>|<span data-ttu-id="35291-1009">説明</span><span class="sxs-lookup"><span data-stu-id="35291-1009">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="35291-1010">String</span><span class="sxs-lookup"><span data-stu-id="35291-1010">String</span></span>||<span data-ttu-id="35291-1011">取得する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="35291-1011">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="35291-1012">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="35291-1012">Object</span></span>|<span data-ttu-id="35291-1013">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1013">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1014">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="35291-1014">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="35291-1015">Object</span><span class="sxs-lookup"><span data-stu-id="35291-1015">Object</span></span>|<span data-ttu-id="35291-1016">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1016">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1017">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="35291-1017">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="35291-1018">function</span><span class="sxs-lookup"><span data-stu-id="35291-1018">function</span></span>|<span data-ttu-id="35291-1019">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1019">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1020">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1020">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-1021">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1021">Requirements</span></span>

|<span data-ttu-id="35291-1022">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1022">Requirement</span></span>|<span data-ttu-id="35291-1023">値</span><span class="sxs-lookup"><span data-stu-id="35291-1023">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-1024">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-1024">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-1025">プレビュー</span><span class="sxs-lookup"><span data-stu-id="35291-1025">Preview</span></span>|
|[<span data-ttu-id="35291-1026">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-1026">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-1027">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-1027">ReadItem</span></span>|
|[<span data-ttu-id="35291-1028">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-1028">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-1029">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-1029">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35291-1030">戻り値:</span><span class="sxs-lookup"><span data-stu-id="35291-1030">Returns:</span></span>

<span data-ttu-id="35291-1031">型: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="35291-1031">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="35291-1032">例</span><span class="sxs-lookup"><span data-stu-id="35291-1032">Example</span></span>

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
    // parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file
    if (result.format == Office.MailboxEnums.AttachmentContentFormat.Base64) {
        // handle file attachment
    }
    else if (result.format == Office.MailboxEnums.AttachmentContentFormat.Eml) {
        // handle item attachment
    }
    else if (result.format == Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
        // handle .icalender attachment
    }
    else {
        // handle cloud attachment  
    }
}
```

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="35291-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="35291-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="35291-1034">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-1034">Gets the item's attachments as an array.</span></span> <span data-ttu-id="35291-1035">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="35291-1035">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-1036">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-1036">Parameters:</span></span>

|<span data-ttu-id="35291-1037">名前</span><span class="sxs-lookup"><span data-stu-id="35291-1037">Name</span></span>|<span data-ttu-id="35291-1038">型</span><span class="sxs-lookup"><span data-stu-id="35291-1038">Type</span></span>|<span data-ttu-id="35291-1039">属性</span><span class="sxs-lookup"><span data-stu-id="35291-1039">Attributes</span></span>|<span data-ttu-id="35291-1040">説明</span><span class="sxs-lookup"><span data-stu-id="35291-1040">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="35291-1041">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="35291-1041">Object</span></span>|<span data-ttu-id="35291-1042">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1043">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="35291-1043">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="35291-1044">Object</span><span class="sxs-lookup"><span data-stu-id="35291-1044">Object</span></span>|<span data-ttu-id="35291-1045">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1046">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="35291-1046">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="35291-1047">function</span><span class="sxs-lookup"><span data-stu-id="35291-1047">function</span></span>|<span data-ttu-id="35291-1048">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1048">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1049">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1049">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-1050">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1050">Requirements</span></span>

|<span data-ttu-id="35291-1051">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1051">Requirement</span></span>|<span data-ttu-id="35291-1052">値</span><span class="sxs-lookup"><span data-stu-id="35291-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-1053">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-1053">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-1054">プレビュー</span><span class="sxs-lookup"><span data-stu-id="35291-1054">Preview</span></span>|
|[<span data-ttu-id="35291-1055">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-1056">ReadItem</span></span>|
|[<span data-ttu-id="35291-1057">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-1058">作成</span><span class="sxs-lookup"><span data-stu-id="35291-1058">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="35291-1059">戻り値:</span><span class="sxs-lookup"><span data-stu-id="35291-1059">Returns:</span></span>

<span data-ttu-id="35291-1060">型: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="35291-1060">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="35291-1061">例</span><span class="sxs-lookup"><span data-stu-id="35291-1061">Example</span></span>

<span data-ttu-id="35291-1062">次の例では、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="35291-1062">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);  
function callback(result) {
    if (result.value.length > 0) {
        for (i = 0 ; i < result.value.length ; i++) {
            var _att = result.value [i];
            outputString += "<BR>" + i + ". Name: ";
            outputString += _att.name;
            outputString += "<BR>ID: " + _att.id;
            outputString += "<BR>contentType: " + _att.contentType;
            outputString += "<BR>size: " + _att.size;
            outputString += "<BR>attachmentType: " + _att.attachmentType;
            outputString += "<BR>isInline: " + _att.isInline;
        }
    }
}
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="35291-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="35291-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="35291-1064">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-1064">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="35291-1065">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="35291-1065">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-1066">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1066">Requirements</span></span>

|<span data-ttu-id="35291-1067">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1067">Requirement</span></span>|<span data-ttu-id="35291-1068">値</span><span class="sxs-lookup"><span data-stu-id="35291-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-1069">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-1070">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-1070">1.0</span></span>|
|[<span data-ttu-id="35291-1071">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-1071">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-1072">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-1072">ReadItem</span></span>|
|[<span data-ttu-id="35291-1073">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-1073">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-1074">読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-1074">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35291-1075">戻り値:</span><span class="sxs-lookup"><span data-stu-id="35291-1075">Returns:</span></span>

<span data-ttu-id="35291-1076">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="35291-1076">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="35291-1077">例</span><span class="sxs-lookup"><span data-stu-id="35291-1077">Example</span></span>

<span data-ttu-id="35291-1078">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="35291-1078">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="35291-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="35291-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="35291-1080">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-1080">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="35291-1081">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="35291-1081">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-1082">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-1082">Parameters:</span></span>

|<span data-ttu-id="35291-1083">名前</span><span class="sxs-lookup"><span data-stu-id="35291-1083">Name</span></span>|<span data-ttu-id="35291-1084">型</span><span class="sxs-lookup"><span data-stu-id="35291-1084">Type</span></span>|<span data-ttu-id="35291-1085">説明</span><span class="sxs-lookup"><span data-stu-id="35291-1085">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="35291-1086">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="35291-1086">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="35291-1087">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="35291-1087">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-1088">Requirements</span><span class="sxs-lookup"><span data-stu-id="35291-1088">Requirements</span></span>

|<span data-ttu-id="35291-1089">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1089">Requirement</span></span>|<span data-ttu-id="35291-1090">値</span><span class="sxs-lookup"><span data-stu-id="35291-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-1091">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-1092">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-1092">1.0</span></span>|
|[<span data-ttu-id="35291-1093">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-1093">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-1094">制限あり</span><span class="sxs-lookup"><span data-stu-id="35291-1094">Restricted</span></span>|
|[<span data-ttu-id="35291-1095">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-1095">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-1096">読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-1096">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35291-1097">戻り値:</span><span class="sxs-lookup"><span data-stu-id="35291-1097">Returns:</span></span>

<span data-ttu-id="35291-1098">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="35291-1098">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="35291-1099">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="35291-1099">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="35291-1100">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="35291-1100">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="35291-1101">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="35291-1101">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="35291-1102">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="35291-1102">Value of `entityType`</span></span>|<span data-ttu-id="35291-1103">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="35291-1103">Type of objects in returned array</span></span>|<span data-ttu-id="35291-1104">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="35291-1104">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="35291-1105">文字列</span><span class="sxs-lookup"><span data-stu-id="35291-1105">String</span></span>|<span data-ttu-id="35291-1106">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="35291-1106">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="35291-1107">連絡先</span><span class="sxs-lookup"><span data-stu-id="35291-1107">Contact</span></span>|<span data-ttu-id="35291-1108">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="35291-1108">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="35291-1109">文字列</span><span class="sxs-lookup"><span data-stu-id="35291-1109">String</span></span>|<span data-ttu-id="35291-1110">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="35291-1110">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="35291-1111">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="35291-1111">MeetingSuggestion</span></span>|<span data-ttu-id="35291-1112">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="35291-1112">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="35291-1113">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="35291-1113">PhoneNumber</span></span>|<span data-ttu-id="35291-1114">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="35291-1114">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="35291-1115">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="35291-1115">TaskSuggestion</span></span>|<span data-ttu-id="35291-1116">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="35291-1116">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="35291-1117">文字列</span><span class="sxs-lookup"><span data-stu-id="35291-1117">String</span></span>|<span data-ttu-id="35291-1118">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="35291-1118">**Restricted**</span></span>|

<span data-ttu-id="35291-1119">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="35291-1119">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="35291-1120">例</span><span class="sxs-lookup"><span data-stu-id="35291-1120">Example</span></span>

<span data-ttu-id="35291-1121">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="35291-1121">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="35291-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="35291-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="35291-1123">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-1123">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="35291-1124">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="35291-1124">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="35291-1125">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-1125">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-1126">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-1126">Parameters:</span></span>

|<span data-ttu-id="35291-1127">名前</span><span class="sxs-lookup"><span data-stu-id="35291-1127">Name</span></span>|<span data-ttu-id="35291-1128">型</span><span class="sxs-lookup"><span data-stu-id="35291-1128">Type</span></span>|<span data-ttu-id="35291-1129">説明</span><span class="sxs-lookup"><span data-stu-id="35291-1129">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="35291-1130">String</span><span class="sxs-lookup"><span data-stu-id="35291-1130">String</span></span>|<span data-ttu-id="35291-1131">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="35291-1131">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-1132">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1132">Requirements</span></span>

|<span data-ttu-id="35291-1133">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1133">Requirement</span></span>|<span data-ttu-id="35291-1134">値</span><span class="sxs-lookup"><span data-stu-id="35291-1134">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-1135">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-1135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-1136">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-1136">1.0</span></span>|
|[<span data-ttu-id="35291-1137">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-1137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-1138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-1138">ReadItem</span></span>|
|[<span data-ttu-id="35291-1139">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-1139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-1140">読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-1140">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35291-1141">戻り値:</span><span class="sxs-lookup"><span data-stu-id="35291-1141">Returns:</span></span>

<span data-ttu-id="35291-p162">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="35291-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="35291-1144">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="35291-1144">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="35291-1145">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="35291-1145">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="35291-1146">アドインが[操作可能メッセージによってアクティブ化](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-1146">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="35291-1147">このメソッドは、Outlook 2016 for Windows 以降 (16.0.8413.1000 以降のクイック実行バージョン) および Outlook on the web for Office 365 でのみサポートされます。</span><span class="sxs-lookup"><span data-stu-id="35291-1147">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-1148">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-1148">Parameters:</span></span>
|<span data-ttu-id="35291-1149">名前</span><span class="sxs-lookup"><span data-stu-id="35291-1149">Name</span></span>|<span data-ttu-id="35291-1150">型</span><span class="sxs-lookup"><span data-stu-id="35291-1150">Type</span></span>|<span data-ttu-id="35291-1151">属性</span><span class="sxs-lookup"><span data-stu-id="35291-1151">Attributes</span></span>|<span data-ttu-id="35291-1152">説明</span><span class="sxs-lookup"><span data-stu-id="35291-1152">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="35291-1153">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="35291-1153">Object</span></span>|<span data-ttu-id="35291-1154">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1154">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1155">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="35291-1155">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="35291-1156">Object</span><span class="sxs-lookup"><span data-stu-id="35291-1156">Object</span></span>|<span data-ttu-id="35291-1157">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1158">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="35291-1158">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="35291-1159">function</span><span class="sxs-lookup"><span data-stu-id="35291-1159">function</span></span>|<span data-ttu-id="35291-1160">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1161">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1161">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="35291-1162">成功すると、初期化データが文字列として `asyncResult.value` プロパティで指定されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1162">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="35291-1163">初期化コンテキストがない場合、`asyncResult` オブジェクトには、`code` プロパティが `9020`、`name` プロパティが `GenericResponseError` に設定された `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="35291-1163">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-1164">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1164">Requirements</span></span>

|<span data-ttu-id="35291-1165">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1165">Requirement</span></span>|<span data-ttu-id="35291-1166">値</span><span class="sxs-lookup"><span data-stu-id="35291-1166">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-1167">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-1167">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-1168">プレビュー</span><span class="sxs-lookup"><span data-stu-id="35291-1168">Preview</span></span>|
|[<span data-ttu-id="35291-1169">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-1169">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-1170">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-1170">ReadItem</span></span>|
|[<span data-ttu-id="35291-1171">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-1171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-1172">読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-1172">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-1173">例</span><span class="sxs-lookup"><span data-stu-id="35291-1173">Example</span></span>

```javascript
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a><span data-ttu-id="35291-1174">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="35291-1174">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="35291-1175">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="35291-1175">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="35291-1176">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="35291-1176">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="35291-p163">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="35291-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="35291-1180">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="35291-1180">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="35291-1181">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="35291-1181">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="35291-p164">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-1185">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1185">Requirements</span></span>

|<span data-ttu-id="35291-1186">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1186">Requirement</span></span>|<span data-ttu-id="35291-1187">値</span><span class="sxs-lookup"><span data-stu-id="35291-1187">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-1188">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-1188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-1189">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-1189">1.0</span></span>|
|[<span data-ttu-id="35291-1190">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-1190">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-1191">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-1191">ReadItem</span></span>|
|[<span data-ttu-id="35291-1192">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-1192">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-1193">読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-1193">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35291-1194">戻り値:</span><span class="sxs-lookup"><span data-stu-id="35291-1194">Returns:</span></span>

<span data-ttu-id="35291-p165">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="35291-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="35291-1197">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="35291-1197">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="35291-1198">Object</span><span class="sxs-lookup"><span data-stu-id="35291-1198">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="35291-1199">例</span><span class="sxs-lookup"><span data-stu-id="35291-1199">Example</span></span>

<span data-ttu-id="35291-1200">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="35291-1200">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="35291-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="35291-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="35291-1202">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="35291-1202">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="35291-1203">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="35291-1203">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="35291-1204">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="35291-1204">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="35291-p166">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="35291-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-1207">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-1207">Parameters:</span></span>

|<span data-ttu-id="35291-1208">名前</span><span class="sxs-lookup"><span data-stu-id="35291-1208">Name</span></span>|<span data-ttu-id="35291-1209">型</span><span class="sxs-lookup"><span data-stu-id="35291-1209">Type</span></span>|<span data-ttu-id="35291-1210">説明</span><span class="sxs-lookup"><span data-stu-id="35291-1210">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="35291-1211">String</span><span class="sxs-lookup"><span data-stu-id="35291-1211">String</span></span>|<span data-ttu-id="35291-1212">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="35291-1212">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-1213">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1213">Requirements</span></span>

|<span data-ttu-id="35291-1214">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1214">Requirement</span></span>|<span data-ttu-id="35291-1215">値</span><span class="sxs-lookup"><span data-stu-id="35291-1215">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-1216">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-1216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-1217">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-1217">1.0</span></span>|
|[<span data-ttu-id="35291-1218">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-1218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-1219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-1219">ReadItem</span></span>|
|[<span data-ttu-id="35291-1220">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-1220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-1221">読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-1221">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35291-1222">戻り値:</span><span class="sxs-lookup"><span data-stu-id="35291-1222">Returns:</span></span>

<span data-ttu-id="35291-1223">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="35291-1223">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="35291-1224">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="35291-1224">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="35291-1225">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="35291-1225">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="35291-1226">例</span><span class="sxs-lookup"><span data-stu-id="35291-1226">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="35291-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="35291-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="35291-1228">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-1228">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="35291-p167">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="35291-p167">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-1231">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-1231">Parameters:</span></span>

|<span data-ttu-id="35291-1232">名前</span><span class="sxs-lookup"><span data-stu-id="35291-1232">Name</span></span>|<span data-ttu-id="35291-1233">型</span><span class="sxs-lookup"><span data-stu-id="35291-1233">Type</span></span>|<span data-ttu-id="35291-1234">属性</span><span class="sxs-lookup"><span data-stu-id="35291-1234">Attributes</span></span>|<span data-ttu-id="35291-1235">説明</span><span class="sxs-lookup"><span data-stu-id="35291-1235">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="35291-1236">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="35291-1236">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="35291-p168">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="35291-p168">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="35291-1240">Object</span><span class="sxs-lookup"><span data-stu-id="35291-1240">Object</span></span>|<span data-ttu-id="35291-1241">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1241">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1242">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="35291-1242">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="35291-1243">Object</span><span class="sxs-lookup"><span data-stu-id="35291-1243">Object</span></span>|<span data-ttu-id="35291-1244">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1244">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1245">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="35291-1245">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="35291-1246">function</span><span class="sxs-lookup"><span data-stu-id="35291-1246">function</span></span>||<span data-ttu-id="35291-1247">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1247">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="35291-1248">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="35291-1248">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="35291-1249">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="35291-1249">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-1250">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1250">Requirements</span></span>

|<span data-ttu-id="35291-1251">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1251">Requirement</span></span>|<span data-ttu-id="35291-1252">値</span><span class="sxs-lookup"><span data-stu-id="35291-1252">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-1253">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-1253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-1254">1.2</span><span class="sxs-lookup"><span data-stu-id="35291-1254">1.2</span></span>|
|[<span data-ttu-id="35291-1255">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-1255">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-1256">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="35291-1256">ReadWriteItem</span></span>|
|[<span data-ttu-id="35291-1257">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-1257">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-1258">作成</span><span class="sxs-lookup"><span data-stu-id="35291-1258">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="35291-1259">戻り値:</span><span class="sxs-lookup"><span data-stu-id="35291-1259">Returns:</span></span>

<span data-ttu-id="35291-1260">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="35291-1260">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="35291-1261">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="35291-1261">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="35291-1262">String</span><span class="sxs-lookup"><span data-stu-id="35291-1262">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="35291-1263">例</span><span class="sxs-lookup"><span data-stu-id="35291-1263">Example</span></span>

```javascript
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="35291-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="35291-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="35291-p170">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="35291-p170">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="35291-1267">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="35291-1267">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-1268">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1268">Requirements</span></span>

|<span data-ttu-id="35291-1269">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1269">Requirement</span></span>|<span data-ttu-id="35291-1270">値</span><span class="sxs-lookup"><span data-stu-id="35291-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-1271">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-1272">1.6</span><span class="sxs-lookup"><span data-stu-id="35291-1272">1.6</span></span>|
|[<span data-ttu-id="35291-1273">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-1274">ReadItem</span></span>|
|[<span data-ttu-id="35291-1275">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-1276">読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35291-1277">戻り値:</span><span class="sxs-lookup"><span data-stu-id="35291-1277">Returns:</span></span>

<span data-ttu-id="35291-1278">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="35291-1278">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="35291-1279">例</span><span class="sxs-lookup"><span data-stu-id="35291-1279">Example</span></span>

<span data-ttu-id="35291-1280">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="35291-1280">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="35291-1281">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="35291-1281">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="35291-p171">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="35291-p171">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="35291-1284">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="35291-1284">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="35291-p172">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="35291-p172">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="35291-1288">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="35291-1288">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="35291-1289">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="35291-1289">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="35291-p173">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-p173">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="35291-1293">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1293">Requirements</span></span>

|<span data-ttu-id="35291-1294">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1294">Requirement</span></span>|<span data-ttu-id="35291-1295">値</span><span class="sxs-lookup"><span data-stu-id="35291-1295">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-1296">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-1296">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-1297">1.6</span><span class="sxs-lookup"><span data-stu-id="35291-1297">1.6</span></span>|
|[<span data-ttu-id="35291-1298">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-1298">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-1299">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-1299">ReadItem</span></span>|
|[<span data-ttu-id="35291-1300">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-1300">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-1301">読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-1301">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35291-1302">戻り値:</span><span class="sxs-lookup"><span data-stu-id="35291-1302">Returns:</span></span>

<span data-ttu-id="35291-p174">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="35291-p174">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="35291-1305">例</span><span class="sxs-lookup"><span data-stu-id="35291-1305">Example</span></span>

<span data-ttu-id="35291-1306">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="35291-1306">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="35291-1307">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="35291-1307">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="35291-1308">共有フォルダー、カレンダー、メールボックスで選択した予定またはメッセージのプロパティを取得します。</span><span class="sxs-lookup"><span data-stu-id="35291-1308">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-1309">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-1309">Parameters:</span></span>

|<span data-ttu-id="35291-1310">名前</span><span class="sxs-lookup"><span data-stu-id="35291-1310">Name</span></span>|<span data-ttu-id="35291-1311">型</span><span class="sxs-lookup"><span data-stu-id="35291-1311">Type</span></span>|<span data-ttu-id="35291-1312">属性</span><span class="sxs-lookup"><span data-stu-id="35291-1312">Attributes</span></span>|<span data-ttu-id="35291-1313">説明</span><span class="sxs-lookup"><span data-stu-id="35291-1313">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="35291-1314">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="35291-1314">Object</span></span>|<span data-ttu-id="35291-1315">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1315">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1316">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="35291-1316">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="35291-1317">Object</span><span class="sxs-lookup"><span data-stu-id="35291-1317">Object</span></span>|<span data-ttu-id="35291-1318">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1318">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1319">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="35291-1319">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="35291-1320">function</span><span class="sxs-lookup"><span data-stu-id="35291-1320">function</span></span>||<span data-ttu-id="35291-1321">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1321">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="35291-1322">共有プロパティは `asyncResult.value` プロパティの [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1322">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="35291-1323">このオブジェクトは、アイテムの共有プロパティの取得に使用できます。</span><span class="sxs-lookup"><span data-stu-id="35291-1323">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-1324">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1324">Requirements</span></span>

|<span data-ttu-id="35291-1325">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1325">Requirement</span></span>|<span data-ttu-id="35291-1326">値</span><span class="sxs-lookup"><span data-stu-id="35291-1326">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-1327">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-1327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-1328">プレビュー</span><span class="sxs-lookup"><span data-stu-id="35291-1328">Preview</span></span>|
|[<span data-ttu-id="35291-1329">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-1329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-1330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-1330">ReadItem</span></span>|
|[<span data-ttu-id="35291-1331">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-1331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-1332">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-1332">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-1333">例</span><span class="sxs-lookup"><span data-stu-id="35291-1333">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="35291-1334">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="35291-1334">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="35291-1335">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="35291-1335">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="35291-p176">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="35291-p176">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-1339">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-1339">Parameters:</span></span>

|<span data-ttu-id="35291-1340">名前</span><span class="sxs-lookup"><span data-stu-id="35291-1340">Name</span></span>|<span data-ttu-id="35291-1341">型</span><span class="sxs-lookup"><span data-stu-id="35291-1341">Type</span></span>|<span data-ttu-id="35291-1342">属性</span><span class="sxs-lookup"><span data-stu-id="35291-1342">Attributes</span></span>|<span data-ttu-id="35291-1343">説明</span><span class="sxs-lookup"><span data-stu-id="35291-1343">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="35291-1344">function</span><span class="sxs-lookup"><span data-stu-id="35291-1344">function</span></span>||<span data-ttu-id="35291-1345">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1345">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="35291-1346">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1346">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="35291-1347">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="35291-1347">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="35291-1348">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="35291-1348">Object</span></span>|<span data-ttu-id="35291-1349">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1349">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1350">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="35291-1350">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="35291-1351">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="35291-1351">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-1352">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1352">Requirements</span></span>

|<span data-ttu-id="35291-1353">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1353">Requirement</span></span>|<span data-ttu-id="35291-1354">値</span><span class="sxs-lookup"><span data-stu-id="35291-1354">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-1355">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-1355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-1356">1.0</span><span class="sxs-lookup"><span data-stu-id="35291-1356">1.0</span></span>|
|[<span data-ttu-id="35291-1357">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-1357">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-1358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-1358">ReadItem</span></span>|
|[<span data-ttu-id="35291-1359">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-1359">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-1360">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="35291-1360">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-1361">例</span><span class="sxs-lookup"><span data-stu-id="35291-1361">Example</span></span>

<span data-ttu-id="35291-p179">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="35291-p179">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="35291-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="35291-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="35291-1366">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="35291-1366">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="35291-1367">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="35291-1367">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="35291-1368">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="35291-1368">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="35291-1369">Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="35291-1369">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="35291-1370">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始し、フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="35291-1370">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-1371">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-1371">Parameters:</span></span>

|<span data-ttu-id="35291-1372">名前</span><span class="sxs-lookup"><span data-stu-id="35291-1372">Name</span></span>|<span data-ttu-id="35291-1373">型</span><span class="sxs-lookup"><span data-stu-id="35291-1373">Type</span></span>|<span data-ttu-id="35291-1374">属性</span><span class="sxs-lookup"><span data-stu-id="35291-1374">Attributes</span></span>|<span data-ttu-id="35291-1375">説明</span><span class="sxs-lookup"><span data-stu-id="35291-1375">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="35291-1376">String</span><span class="sxs-lookup"><span data-stu-id="35291-1376">String</span></span>||<span data-ttu-id="35291-1377">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="35291-1377">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="35291-1378">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="35291-1378">Object</span></span>|<span data-ttu-id="35291-1379">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1379">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1380">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="35291-1380">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="35291-1381">Object</span><span class="sxs-lookup"><span data-stu-id="35291-1381">Object</span></span>|<span data-ttu-id="35291-1382">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1382">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1383">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="35291-1383">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="35291-1384">function</span><span class="sxs-lookup"><span data-stu-id="35291-1384">function</span></span>|<span data-ttu-id="35291-1385">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1385">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1386">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1386">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="35291-1387">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="35291-1387">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="35291-1388">エラー</span><span class="sxs-lookup"><span data-stu-id="35291-1388">Errors</span></span>

|<span data-ttu-id="35291-1389">エラー コード</span><span class="sxs-lookup"><span data-stu-id="35291-1389">Error code</span></span>|<span data-ttu-id="35291-1390">説明</span><span class="sxs-lookup"><span data-stu-id="35291-1390">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="35291-1391">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="35291-1391">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-1392">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1392">Requirements</span></span>

|<span data-ttu-id="35291-1393">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1393">Requirement</span></span>|<span data-ttu-id="35291-1394">値</span><span class="sxs-lookup"><span data-stu-id="35291-1394">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-1395">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-1395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-1396">1.1</span><span class="sxs-lookup"><span data-stu-id="35291-1396">1.1</span></span>|
|[<span data-ttu-id="35291-1397">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-1397">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-1398">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="35291-1398">ReadWriteItem</span></span>|
|[<span data-ttu-id="35291-1399">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-1399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-1400">作成</span><span class="sxs-lookup"><span data-stu-id="35291-1400">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-1401">例</span><span class="sxs-lookup"><span data-stu-id="35291-1401">Example</span></span>

<span data-ttu-id="35291-1402">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="35291-1402">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="35291-1403">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="35291-1403">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="35291-1404">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="35291-1404">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="35291-1405">現在、サポートされているイベントの種類は `Office.EventType.AttachmentsChanged`、`Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged`、`Office.EventType.RecurrenceChanged` です。</span><span class="sxs-lookup"><span data-stu-id="35291-1405">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-1406">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-1406">Parameters:</span></span>

| <span data-ttu-id="35291-1407">名前</span><span class="sxs-lookup"><span data-stu-id="35291-1407">Name</span></span> | <span data-ttu-id="35291-1408">型</span><span class="sxs-lookup"><span data-stu-id="35291-1408">Type</span></span> | <span data-ttu-id="35291-1409">属性</span><span class="sxs-lookup"><span data-stu-id="35291-1409">Attributes</span></span> | <span data-ttu-id="35291-1410">説明</span><span class="sxs-lookup"><span data-stu-id="35291-1410">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="35291-1411">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="35291-1411">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="35291-1412">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="35291-1412">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="35291-1413">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="35291-1413">Object</span></span> | <span data-ttu-id="35291-1414">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1414">&lt;optional&gt;</span></span> | <span data-ttu-id="35291-1415">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="35291-1415">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="35291-1416">Object</span><span class="sxs-lookup"><span data-stu-id="35291-1416">Object</span></span> | <span data-ttu-id="35291-1417">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1417">&lt;optional&gt;</span></span> | <span data-ttu-id="35291-1418">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="35291-1418">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="35291-1419">function</span><span class="sxs-lookup"><span data-stu-id="35291-1419">function</span></span>| <span data-ttu-id="35291-1420">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1420">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1421">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1421">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-1422">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1422">Requirements</span></span>

|<span data-ttu-id="35291-1423">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1423">Requirement</span></span>| <span data-ttu-id="35291-1424">値</span><span class="sxs-lookup"><span data-stu-id="35291-1424">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-1425">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-1425">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35291-1426">1.7</span><span class="sxs-lookup"><span data-stu-id="35291-1426">1.7</span></span> |
|[<span data-ttu-id="35291-1427">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-1427">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35291-1428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35291-1428">ReadItem</span></span> |
|[<span data-ttu-id="35291-1429">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-1429">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35291-1430">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="35291-1430">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="35291-1431">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="35291-1431">saveAsync([options], callback)</span></span>

<span data-ttu-id="35291-1432">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="35291-1432">Asynchronously saves an item.</span></span>

<span data-ttu-id="35291-p181">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="35291-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="35291-1436">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="35291-1436">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="35291-1437">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1437">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="35291-p183">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="35291-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="35291-1441">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="35291-1441">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="35291-1442">Mac Outlook では、新規作成モードの会議で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="35291-1442">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="35291-1443">Mac Outlook では、会議で `saveAsync` を呼び出すとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1443">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="35291-1444">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1444">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-1445">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-1445">Parameters:</span></span>

|<span data-ttu-id="35291-1446">名前</span><span class="sxs-lookup"><span data-stu-id="35291-1446">Name</span></span>|<span data-ttu-id="35291-1447">型</span><span class="sxs-lookup"><span data-stu-id="35291-1447">Type</span></span>|<span data-ttu-id="35291-1448">属性</span><span class="sxs-lookup"><span data-stu-id="35291-1448">Attributes</span></span>|<span data-ttu-id="35291-1449">説明</span><span class="sxs-lookup"><span data-stu-id="35291-1449">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="35291-1450">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="35291-1450">Object</span></span>|<span data-ttu-id="35291-1451">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1451">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1452">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="35291-1452">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="35291-1453">Object</span><span class="sxs-lookup"><span data-stu-id="35291-1453">Object</span></span>|<span data-ttu-id="35291-1454">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1454">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1455">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="35291-1455">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="35291-1456">function</span><span class="sxs-lookup"><span data-stu-id="35291-1456">function</span></span>||<span data-ttu-id="35291-1457">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1457">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="35291-1458">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1458">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-1459">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1459">Requirements</span></span>

|<span data-ttu-id="35291-1460">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1460">Requirement</span></span>|<span data-ttu-id="35291-1461">値</span><span class="sxs-lookup"><span data-stu-id="35291-1461">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-1462">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-1462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-1463">1.3</span><span class="sxs-lookup"><span data-stu-id="35291-1463">1.3</span></span>|
|[<span data-ttu-id="35291-1464">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-1464">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-1465">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="35291-1465">ReadWriteItem</span></span>|
|[<span data-ttu-id="35291-1466">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-1466">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-1467">作成</span><span class="sxs-lookup"><span data-stu-id="35291-1467">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="35291-1468">例</span><span class="sxs-lookup"><span data-stu-id="35291-1468">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="35291-p185">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="35291-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="35291-1471">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="35291-1471">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="35291-1472">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="35291-1472">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="35291-p186">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="35291-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35291-1476">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="35291-1476">Parameters:</span></span>

|<span data-ttu-id="35291-1477">名前</span><span class="sxs-lookup"><span data-stu-id="35291-1477">Name</span></span>|<span data-ttu-id="35291-1478">型</span><span class="sxs-lookup"><span data-stu-id="35291-1478">Type</span></span>|<span data-ttu-id="35291-1479">属性</span><span class="sxs-lookup"><span data-stu-id="35291-1479">Attributes</span></span>|<span data-ttu-id="35291-1480">説明</span><span class="sxs-lookup"><span data-stu-id="35291-1480">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="35291-1481">String</span><span class="sxs-lookup"><span data-stu-id="35291-1481">String</span></span>||<span data-ttu-id="35291-p187">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="35291-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="35291-1485">Object</span><span class="sxs-lookup"><span data-stu-id="35291-1485">Object</span></span>|<span data-ttu-id="35291-1486">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1486">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1487">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="35291-1487">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="35291-1488">Object</span><span class="sxs-lookup"><span data-stu-id="35291-1488">Object</span></span>|<span data-ttu-id="35291-1489">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1489">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-1490">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="35291-1490">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="35291-1491">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="35291-1491">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="35291-1492">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="35291-1492">&lt;optional&gt;</span></span>|<span data-ttu-id="35291-p188">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="35291-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="35291-p189">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="35291-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="35291-1497">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1497">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="35291-1498">function</span><span class="sxs-lookup"><span data-stu-id="35291-1498">function</span></span>||<span data-ttu-id="35291-1499">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="35291-1499">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35291-1500">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1500">Requirements</span></span>

|<span data-ttu-id="35291-1501">要件</span><span class="sxs-lookup"><span data-stu-id="35291-1501">Requirement</span></span>|<span data-ttu-id="35291-1502">値</span><span class="sxs-lookup"><span data-stu-id="35291-1502">Value</span></span>|
|---|---|
|[<span data-ttu-id="35291-1503">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="35291-1503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="35291-1504">1.2</span><span class="sxs-lookup"><span data-stu-id="35291-1504">1.2</span></span>|
|[<span data-ttu-id="35291-1505">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="35291-1505">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="35291-1506">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="35291-1506">ReadWriteItem</span></span>|
|[<span data-ttu-id="35291-1507">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="35291-1507">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="35291-1508">作成</span><span class="sxs-lookup"><span data-stu-id="35291-1508">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="35291-1509">例</span><span class="sxs-lookup"><span data-stu-id="35291-1509">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
