---
title: Office. アイテム-プレビュー要件セット
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: b74b3aa3c455d33d17767163c960adef7cf783fa
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838586"
---
# <a name="item"></a><span data-ttu-id="167d3-102">item</span><span class="sxs-lookup"><span data-stu-id="167d3-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="167d3-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="167d3-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="167d3-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-106">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-106">Requirements</span></span>

|<span data-ttu-id="167d3-107">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-107">Requirement</span></span>|<span data-ttu-id="167d3-108">値</span><span class="sxs-lookup"><span data-stu-id="167d3-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-110">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-110">1.0</span></span>|
|[<span data-ttu-id="167d3-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="167d3-112">Restricted</span></span>|
|[<span data-ttu-id="167d3-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="167d3-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-115">Members and methods</span></span>

| <span data-ttu-id="167d3-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-116">Member</span></span> | <span data-ttu-id="167d3-117">種類</span><span class="sxs-lookup"><span data-stu-id="167d3-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="167d3-118">attachments</span><span class="sxs-lookup"><span data-stu-id="167d3-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="167d3-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-119">Member</span></span> |
| [<span data-ttu-id="167d3-120">bcc</span><span class="sxs-lookup"><span data-stu-id="167d3-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="167d3-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-121">Member</span></span> |
| [<span data-ttu-id="167d3-122">body</span><span class="sxs-lookup"><span data-stu-id="167d3-122">body</span></span>](#body-body) | <span data-ttu-id="167d3-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-123">Member</span></span> |
| [<span data-ttu-id="167d3-124">cc</span><span class="sxs-lookup"><span data-stu-id="167d3-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="167d3-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-125">Member</span></span> |
| [<span data-ttu-id="167d3-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="167d3-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="167d3-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-127">Member</span></span> |
| [<span data-ttu-id="167d3-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="167d3-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="167d3-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-129">Member</span></span> |
| [<span data-ttu-id="167d3-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="167d3-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="167d3-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-131">Member</span></span> |
| [<span data-ttu-id="167d3-132">end</span><span class="sxs-lookup"><span data-stu-id="167d3-132">end</span></span>](#end-datetime) | <span data-ttu-id="167d3-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-133">Member</span></span> |
| [<span data-ttu-id="167d3-134">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="167d3-134">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="167d3-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-135">Member</span></span> |
| [<span data-ttu-id="167d3-136">from</span><span class="sxs-lookup"><span data-stu-id="167d3-136">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="167d3-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-137">Member</span></span> |
| [<span data-ttu-id="167d3-138">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="167d3-138">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="167d3-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-139">Member</span></span> |
| [<span data-ttu-id="167d3-140">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="167d3-140">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="167d3-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-141">Member</span></span> |
| [<span data-ttu-id="167d3-142">itemClass</span><span class="sxs-lookup"><span data-stu-id="167d3-142">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="167d3-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-143">Member</span></span> |
| [<span data-ttu-id="167d3-144">itemId</span><span class="sxs-lookup"><span data-stu-id="167d3-144">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="167d3-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-145">Member</span></span> |
| [<span data-ttu-id="167d3-146">itemType</span><span class="sxs-lookup"><span data-stu-id="167d3-146">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="167d3-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-147">Member</span></span> |
| [<span data-ttu-id="167d3-148">location</span><span class="sxs-lookup"><span data-stu-id="167d3-148">location</span></span>](#location-stringlocation) | <span data-ttu-id="167d3-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-149">Member</span></span> |
| [<span data-ttu-id="167d3-150">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="167d3-150">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="167d3-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-151">Member</span></span> |
| [<span data-ttu-id="167d3-152">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="167d3-152">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="167d3-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-153">Member</span></span> |
| [<span data-ttu-id="167d3-154">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="167d3-154">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="167d3-155">Member</span><span class="sxs-lookup"><span data-stu-id="167d3-155">Member</span></span> |
| [<span data-ttu-id="167d3-156">organizer</span><span class="sxs-lookup"><span data-stu-id="167d3-156">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="167d3-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-157">Member</span></span> |
| [<span data-ttu-id="167d3-158">繰り返さ</span><span class="sxs-lookup"><span data-stu-id="167d3-158">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="167d3-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-159">Member</span></span> |
| [<span data-ttu-id="167d3-160">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="167d3-160">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="167d3-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-161">Member</span></span> |
| [<span data-ttu-id="167d3-162">sender</span><span class="sxs-lookup"><span data-stu-id="167d3-162">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="167d3-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-163">Member</span></span> |
| [<span data-ttu-id="167d3-164">系列 id</span><span class="sxs-lookup"><span data-stu-id="167d3-164">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="167d3-165">Member</span><span class="sxs-lookup"><span data-stu-id="167d3-165">Member</span></span> |
| [<span data-ttu-id="167d3-166">start</span><span class="sxs-lookup"><span data-stu-id="167d3-166">start</span></span>](#start-datetime) | <span data-ttu-id="167d3-167">Member</span><span class="sxs-lookup"><span data-stu-id="167d3-167">Member</span></span> |
| [<span data-ttu-id="167d3-168">subject</span><span class="sxs-lookup"><span data-stu-id="167d3-168">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="167d3-169">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-169">Member</span></span> |
| [<span data-ttu-id="167d3-170">to</span><span class="sxs-lookup"><span data-stu-id="167d3-170">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="167d3-171">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-171">Member</span></span> |
| [<span data-ttu-id="167d3-172">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="167d3-172">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="167d3-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-173">Method</span></span> |
| [<span data-ttu-id="167d3-174">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="167d3-174">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="167d3-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-175">Method</span></span> |
| [<span data-ttu-id="167d3-176">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="167d3-176">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="167d3-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-177">Method</span></span> |
| [<span data-ttu-id="167d3-178">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="167d3-178">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="167d3-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-179">Method</span></span> |
| [<span data-ttu-id="167d3-180">close</span><span class="sxs-lookup"><span data-stu-id="167d3-180">close</span></span>](#close) | <span data-ttu-id="167d3-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-181">Method</span></span> |
| [<span data-ttu-id="167d3-182">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="167d3-182">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="167d3-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-183">Method</span></span> |
| [<span data-ttu-id="167d3-184">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="167d3-184">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="167d3-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-185">Method</span></span> |
| [<span data-ttu-id="167d3-186">getattachmentcontentasync</span><span class="sxs-lookup"><span data-stu-id="167d3-186">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="167d3-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-187">Method</span></span> |
| [<span data-ttu-id="167d3-188">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="167d3-188">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="167d3-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-189">Method</span></span> |
| [<span data-ttu-id="167d3-190">getEntities</span><span class="sxs-lookup"><span data-stu-id="167d3-190">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="167d3-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-191">Method</span></span> |
| [<span data-ttu-id="167d3-192">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="167d3-192">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="167d3-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-193">Method</span></span> |
| [<span data-ttu-id="167d3-194">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="167d3-194">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="167d3-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-195">Method</span></span> |
| [<span data-ttu-id="167d3-196">、office.context.mailbox.item.getinitializationcontextasync</span><span class="sxs-lookup"><span data-stu-id="167d3-196">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="167d3-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-197">Method</span></span> |
| [<span data-ttu-id="167d3-198">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="167d3-198">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="167d3-199">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-199">Method</span></span> |
| [<span data-ttu-id="167d3-200">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="167d3-200">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="167d3-201">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-201">Method</span></span> |
| [<span data-ttu-id="167d3-202">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="167d3-202">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="167d3-203">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-203">Method</span></span> |
| [<span data-ttu-id="167d3-204">office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="167d3-204">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="167d3-205">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-205">Method</span></span> |
| [<span data-ttu-id="167d3-206">office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="167d3-206">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="167d3-207">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-207">Method</span></span> |
| [<span data-ttu-id="167d3-208">getsharedpropertiesasync</span><span class="sxs-lookup"><span data-stu-id="167d3-208">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="167d3-209">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-209">Method</span></span> |
| [<span data-ttu-id="167d3-210">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="167d3-210">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="167d3-211">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-211">Method</span></span> |
| [<span data-ttu-id="167d3-212">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="167d3-212">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="167d3-213">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-213">Method</span></span> |
| [<span data-ttu-id="167d3-214">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="167d3-214">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="167d3-215">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-215">Method</span></span> |
| [<span data-ttu-id="167d3-216">saveAsync</span><span class="sxs-lookup"><span data-stu-id="167d3-216">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="167d3-217">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-217">Method</span></span> |
| [<span data-ttu-id="167d3-218">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="167d3-218">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="167d3-219">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-219">Method</span></span> |

### <a name="example"></a><span data-ttu-id="167d3-220">例</span><span class="sxs-lookup"><span data-stu-id="167d3-220">Example</span></span>

<span data-ttu-id="167d3-221">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="167d3-221">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="167d3-222">メンバー</span><span class="sxs-lookup"><span data-stu-id="167d3-222">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="167d3-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="167d3-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="167d3-224">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-224">Gets the item's attachments as an array.</span></span> <span data-ttu-id="167d3-225">Read mode only.</span><span class="sxs-lookup"><span data-stu-id="167d3-225">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-226">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="167d3-226">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="167d3-227">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="167d3-227">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="167d3-228">タイプ</span><span class="sxs-lookup"><span data-stu-id="167d3-228">Type</span></span>

*   <span data-ttu-id="167d3-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="167d3-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-230">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-230">Requirements</span></span>

|<span data-ttu-id="167d3-231">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-231">Requirement</span></span>|<span data-ttu-id="167d3-232">値</span><span class="sxs-lookup"><span data-stu-id="167d3-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-233">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-234">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-234">1.0</span></span>|
|[<span data-ttu-id="167d3-235">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-235">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-236">ReadItem</span></span>|
|[<span data-ttu-id="167d3-237">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-237">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-238">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-238">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-239">例</span><span class="sxs-lookup"><span data-stu-id="167d3-239">Example</span></span>

<span data-ttu-id="167d3-240">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="167d3-240">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="167d3-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="167d3-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="167d3-242">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-242">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="167d3-243">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="167d3-243">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="167d3-244">タイプ</span><span class="sxs-lookup"><span data-stu-id="167d3-244">Type</span></span>

*   [<span data-ttu-id="167d3-245">受信者</span><span class="sxs-lookup"><span data-stu-id="167d3-245">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="167d3-246">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-246">Requirements</span></span>

|<span data-ttu-id="167d3-247">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-247">Requirement</span></span>|<span data-ttu-id="167d3-248">値</span><span class="sxs-lookup"><span data-stu-id="167d3-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-249">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-250">1.1</span><span class="sxs-lookup"><span data-stu-id="167d3-250">1.1</span></span>|
|[<span data-ttu-id="167d3-251">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-251">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-252">ReadItem</span></span>|
|[<span data-ttu-id="167d3-253">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-253">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-254">作成</span><span class="sxs-lookup"><span data-stu-id="167d3-254">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-255">例</span><span class="sxs-lookup"><span data-stu-id="167d3-255">Example</span></span>

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

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="167d3-256">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="167d3-256">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="167d3-257">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-257">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="167d3-258">型</span><span class="sxs-lookup"><span data-stu-id="167d3-258">Type</span></span>

*   [<span data-ttu-id="167d3-259">Body</span><span class="sxs-lookup"><span data-stu-id="167d3-259">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="167d3-260">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-260">Requirements</span></span>

|<span data-ttu-id="167d3-261">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-261">Requirement</span></span>|<span data-ttu-id="167d3-262">値</span><span class="sxs-lookup"><span data-stu-id="167d3-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-263">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-264">1.1</span><span class="sxs-lookup"><span data-stu-id="167d3-264">1.1</span></span>|
|[<span data-ttu-id="167d3-265">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-266">ReadItem</span></span>|
|[<span data-ttu-id="167d3-267">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-268">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-268">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-269">例</span><span class="sxs-lookup"><span data-stu-id="167d3-269">Example</span></span>

<span data-ttu-id="167d3-270">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-270">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="167d3-271">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="167d3-271">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

---
---

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="167d3-272">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="167d3-272">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="167d3-273">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="167d3-273">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="167d3-274">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="167d3-274">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="167d3-275">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="167d3-275">Read mode</span></span>

<span data-ttu-id="167d3-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="167d3-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="167d3-278">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="167d3-278">Compose mode</span></span>

<span data-ttu-id="167d3-279">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-279">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="167d3-280">タイプ</span><span class="sxs-lookup"><span data-stu-id="167d3-280">Type</span></span>

*   <span data-ttu-id="167d3-281">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="167d3-281">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-282">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-282">Requirements</span></span>

|<span data-ttu-id="167d3-283">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-283">Requirement</span></span>|<span data-ttu-id="167d3-284">値</span><span class="sxs-lookup"><span data-stu-id="167d3-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-285">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-285">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-286">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-286">1.0</span></span>|
|[<span data-ttu-id="167d3-287">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-287">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-288">ReadItem</span></span>|
|[<span data-ttu-id="167d3-289">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-289">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-290">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-290">Compose or Read</span></span>|

---
---

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="167d3-291">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="167d3-291">(nullable) conversationId :String</span></span>

<span data-ttu-id="167d3-292">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-292">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="167d3-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="167d3-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="167d3-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="167d3-297">Type</span><span class="sxs-lookup"><span data-stu-id="167d3-297">Type</span></span>

*   <span data-ttu-id="167d3-298">String</span><span class="sxs-lookup"><span data-stu-id="167d3-298">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-299">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-299">Requirements</span></span>

|<span data-ttu-id="167d3-300">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-300">Requirement</span></span>|<span data-ttu-id="167d3-301">値</span><span class="sxs-lookup"><span data-stu-id="167d3-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-302">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-303">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-303">1.0</span></span>|
|[<span data-ttu-id="167d3-304">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-305">ReadItem</span></span>|
|[<span data-ttu-id="167d3-306">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-307">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-307">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-308">例</span><span class="sxs-lookup"><span data-stu-id="167d3-308">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="167d3-309">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="167d3-309">dateTimeCreated :Date</span></span>

<span data-ttu-id="167d3-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="167d3-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="167d3-312">タイプ</span><span class="sxs-lookup"><span data-stu-id="167d3-312">Type</span></span>

*   <span data-ttu-id="167d3-313">日付</span><span class="sxs-lookup"><span data-stu-id="167d3-313">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-314">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-314">Requirements</span></span>

|<span data-ttu-id="167d3-315">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-315">Requirement</span></span>|<span data-ttu-id="167d3-316">値</span><span class="sxs-lookup"><span data-stu-id="167d3-316">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-317">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-317">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-318">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-318">1.0</span></span>|
|[<span data-ttu-id="167d3-319">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-319">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-320">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-320">ReadItem</span></span>|
|[<span data-ttu-id="167d3-321">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-321">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-322">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-322">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-323">例</span><span class="sxs-lookup"><span data-stu-id="167d3-323">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="167d3-324">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="167d3-324">dateTimeModified :Date</span></span>

<span data-ttu-id="167d3-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="167d3-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-327">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="167d3-327">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="167d3-328">タイプ</span><span class="sxs-lookup"><span data-stu-id="167d3-328">Type</span></span>

*   <span data-ttu-id="167d3-329">日付</span><span class="sxs-lookup"><span data-stu-id="167d3-329">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-330">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-330">Requirements</span></span>

|<span data-ttu-id="167d3-331">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-331">Requirement</span></span>|<span data-ttu-id="167d3-332">値</span><span class="sxs-lookup"><span data-stu-id="167d3-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-333">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-334">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-334">1.0</span></span>|
|[<span data-ttu-id="167d3-335">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-336">ReadItem</span></span>|
|[<span data-ttu-id="167d3-337">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-338">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-338">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-339">例</span><span class="sxs-lookup"><span data-stu-id="167d3-339">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

---
---

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="167d3-340">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="167d3-340">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="167d3-341">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="167d3-341">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="167d3-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="167d3-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="167d3-344">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="167d3-344">Read mode</span></span>

<span data-ttu-id="167d3-345">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-345">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="167d3-346">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="167d3-346">Compose mode</span></span>

<span data-ttu-id="167d3-347">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-347">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="167d3-348">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="167d3-348">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="167d3-349">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="167d3-349">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="167d3-350">型</span><span class="sxs-lookup"><span data-stu-id="167d3-350">Type</span></span>

*   <span data-ttu-id="167d3-351">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="167d3-351">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-352">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-352">Requirements</span></span>

|<span data-ttu-id="167d3-353">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-353">Requirement</span></span>|<span data-ttu-id="167d3-354">値</span><span class="sxs-lookup"><span data-stu-id="167d3-354">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-355">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-356">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-356">1.0</span></span>|
|[<span data-ttu-id="167d3-357">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-357">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-358">ReadItem</span></span>|
|[<span data-ttu-id="167d3-359">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-359">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-360">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-360">Compose or Read</span></span>|

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="167d3-361">enhancedLocation:[enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="167d3-361">enhancedLocation :[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="167d3-362">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="167d3-362">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="167d3-363">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="167d3-363">Read mode</span></span>

<span data-ttu-id="167d3-364">この`enhancedLocation`プロパティは、予定に関連付けられている場所 ( [locationdetails](/javascript/api/outlook/office.locationdetails)オブジェクトで表される) のセットを取得できる[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-364">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="167d3-365">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="167d3-365">Compose mode</span></span>

<span data-ttu-id="167d3-366">この`enhancedLocation`プロパティは、予定の場所を取得、削除、または追加するためのメソッドを提供する[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-366">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="167d3-367">タイプ</span><span class="sxs-lookup"><span data-stu-id="167d3-367">Type</span></span>

*   [<span data-ttu-id="167d3-368">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="167d3-368">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="167d3-369">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-369">Requirements</span></span>

|<span data-ttu-id="167d3-370">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-370">Requirement</span></span>|<span data-ttu-id="167d3-371">値</span><span class="sxs-lookup"><span data-stu-id="167d3-371">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-372">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-372">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-373">プレビュー</span><span class="sxs-lookup"><span data-stu-id="167d3-373">Preview</span></span>|
|[<span data-ttu-id="167d3-374">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-374">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-375">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-375">ReadItem</span></span>|
|[<span data-ttu-id="167d3-376">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-376">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-377">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-377">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-378">例</span><span class="sxs-lookup"><span data-stu-id="167d3-378">Example</span></span>

<span data-ttu-id="167d3-379">次の例では、予定に関連付けられている現在の場所を取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-379">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="167d3-380">from:[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)|[from](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="167d3-380">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="167d3-381">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-381">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="167d3-p112">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="167d3-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-384">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="167d3-384">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="167d3-385">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="167d3-385">Read mode</span></span>

<span data-ttu-id="167d3-386">プロパティ`from`は`EmailAddressDetails`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-386">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="167d3-387">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="167d3-387">Compose mode</span></span>

<span data-ttu-id="167d3-388">プロパティ`from`は、from `From`値を取得するメソッドを提供するオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-388">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="167d3-389">タイプ</span><span class="sxs-lookup"><span data-stu-id="167d3-389">Type</span></span>

*   <span data-ttu-id="167d3-390">[電子メールアドレス](/javascript/api/outlook/office.emailaddressdetails) | [の](/javascript/api/outlook/office.from)詳細</span><span class="sxs-lookup"><span data-stu-id="167d3-390">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-391">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-391">Requirements</span></span>

|<span data-ttu-id="167d3-392">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-392">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="167d3-393">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-393">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-394">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-394">1.0</span></span>|<span data-ttu-id="167d3-395">1.7</span><span class="sxs-lookup"><span data-stu-id="167d3-395">1.7</span></span>|
|[<span data-ttu-id="167d3-396">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-396">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-397">ReadItem</span></span>|<span data-ttu-id="167d3-398">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="167d3-398">ReadWriteItem</span></span>|
|[<span data-ttu-id="167d3-399">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-399">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-400">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-400">Read</span></span>|<span data-ttu-id="167d3-401">作成</span><span class="sxs-lookup"><span data-stu-id="167d3-401">Compose</span></span>|

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="167d3-402">internetHeaders:[internetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="167d3-402">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="167d3-403">メッセージのインターネットヘッダーを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="167d3-403">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="167d3-404">タイプ</span><span class="sxs-lookup"><span data-stu-id="167d3-404">Type</span></span>

*   [<span data-ttu-id="167d3-405">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="167d3-405">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="167d3-406">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-406">Requirements</span></span>

|<span data-ttu-id="167d3-407">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-407">Requirement</span></span>|<span data-ttu-id="167d3-408">値</span><span class="sxs-lookup"><span data-stu-id="167d3-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-409">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-410">プレビュー</span><span class="sxs-lookup"><span data-stu-id="167d3-410">Preview</span></span>|
|[<span data-ttu-id="167d3-411">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-412">ReadItem</span></span>|
|[<span data-ttu-id="167d3-413">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-414">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-414">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-415">例</span><span class="sxs-lookup"><span data-stu-id="167d3-415">Example</span></span>

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="167d3-416">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="167d3-416">internetMessageId :String</span></span>

<span data-ttu-id="167d3-p113">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="167d3-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="167d3-419">Type</span><span class="sxs-lookup"><span data-stu-id="167d3-419">Type</span></span>

*   <span data-ttu-id="167d3-420">String</span><span class="sxs-lookup"><span data-stu-id="167d3-420">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-421">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-421">Requirements</span></span>

|<span data-ttu-id="167d3-422">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-422">Requirement</span></span>|<span data-ttu-id="167d3-423">値</span><span class="sxs-lookup"><span data-stu-id="167d3-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-424">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-425">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-425">1.0</span></span>|
|[<span data-ttu-id="167d3-426">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-426">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-427">ReadItem</span></span>|
|[<span data-ttu-id="167d3-428">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-429">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-430">例</span><span class="sxs-lookup"><span data-stu-id="167d3-430">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="167d3-431">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="167d3-431">itemClass :String</span></span>

<span data-ttu-id="167d3-p114">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="167d3-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="167d3-p115">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="167d3-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="167d3-436">型</span><span class="sxs-lookup"><span data-stu-id="167d3-436">Type</span></span>|<span data-ttu-id="167d3-437">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-437">Description</span></span>|<span data-ttu-id="167d3-438">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="167d3-438">item class</span></span>|
|---|---|---|
|<span data-ttu-id="167d3-439">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="167d3-439">Appointment items</span></span>|<span data-ttu-id="167d3-440">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="167d3-440">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="167d3-441">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="167d3-441">Message items</span></span>|<span data-ttu-id="167d3-442">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="167d3-442">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="167d3-443">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-443">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="167d3-444">Type</span><span class="sxs-lookup"><span data-stu-id="167d3-444">Type</span></span>

*   <span data-ttu-id="167d3-445">String</span><span class="sxs-lookup"><span data-stu-id="167d3-445">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-446">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-446">Requirements</span></span>

|<span data-ttu-id="167d3-447">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-447">Requirement</span></span>|<span data-ttu-id="167d3-448">値</span><span class="sxs-lookup"><span data-stu-id="167d3-448">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-449">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-449">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-450">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-450">1.0</span></span>|
|[<span data-ttu-id="167d3-451">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-451">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-452">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-452">ReadItem</span></span>|
|[<span data-ttu-id="167d3-453">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-453">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-454">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-454">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-455">例</span><span class="sxs-lookup"><span data-stu-id="167d3-455">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="167d3-456">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="167d3-456">(nullable) itemId :String</span></span>

<span data-ttu-id="167d3-p116">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="167d3-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-459">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="167d3-459">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="167d3-460">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="167d3-460">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="167d3-461">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="167d3-461">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="167d3-462">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="167d3-462">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="167d3-p118">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="167d3-465">Type</span><span class="sxs-lookup"><span data-stu-id="167d3-465">Type</span></span>

*   <span data-ttu-id="167d3-466">String</span><span class="sxs-lookup"><span data-stu-id="167d3-466">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-467">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-467">Requirements</span></span>

|<span data-ttu-id="167d3-468">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-468">Requirement</span></span>|<span data-ttu-id="167d3-469">値</span><span class="sxs-lookup"><span data-stu-id="167d3-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-470">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-471">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-471">1.0</span></span>|
|[<span data-ttu-id="167d3-472">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-473">ReadItem</span></span>|
|[<span data-ttu-id="167d3-474">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-475">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-475">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-476">例</span><span class="sxs-lookup"><span data-stu-id="167d3-476">Example</span></span>

<span data-ttu-id="167d3-p119">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="167d3-479">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="167d3-479">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="167d3-480">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-480">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="167d3-481">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="167d3-481">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="167d3-482">型</span><span class="sxs-lookup"><span data-stu-id="167d3-482">Type</span></span>

*   [<span data-ttu-id="167d3-483">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="167d3-483">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="167d3-484">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-484">Requirements</span></span>

|<span data-ttu-id="167d3-485">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-485">Requirement</span></span>|<span data-ttu-id="167d3-486">値</span><span class="sxs-lookup"><span data-stu-id="167d3-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-487">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-487">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-488">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-488">1.0</span></span>|
|[<span data-ttu-id="167d3-489">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-489">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-490">ReadItem</span></span>|
|[<span data-ttu-id="167d3-491">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-491">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-492">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-492">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-493">例</span><span class="sxs-lookup"><span data-stu-id="167d3-493">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

---
---

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="167d3-494">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="167d3-494">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="167d3-495">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="167d3-495">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="167d3-496">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="167d3-496">Read mode</span></span>

<span data-ttu-id="167d3-497">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-497">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="167d3-498">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="167d3-498">Compose mode</span></span>

<span data-ttu-id="167d3-499">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-499">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="167d3-500">型</span><span class="sxs-lookup"><span data-stu-id="167d3-500">Type</span></span>

*   <span data-ttu-id="167d3-501">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="167d3-501">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-502">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-502">Requirements</span></span>

|<span data-ttu-id="167d3-503">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-503">Requirement</span></span>|<span data-ttu-id="167d3-504">値</span><span class="sxs-lookup"><span data-stu-id="167d3-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-505">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-506">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-506">1.0</span></span>|
|[<span data-ttu-id="167d3-507">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-508">ReadItem</span></span>|
|[<span data-ttu-id="167d3-509">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-510">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-510">Compose or Read</span></span>|

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="167d3-511">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="167d3-511">normalizedSubject :String</span></span>

<span data-ttu-id="167d3-p120">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="167d3-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="167d3-p121">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="167d3-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="167d3-516">型</span><span class="sxs-lookup"><span data-stu-id="167d3-516">Type</span></span>

*   <span data-ttu-id="167d3-517">String</span><span class="sxs-lookup"><span data-stu-id="167d3-517">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-518">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-518">Requirements</span></span>

|<span data-ttu-id="167d3-519">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-519">Requirement</span></span>|<span data-ttu-id="167d3-520">値</span><span class="sxs-lookup"><span data-stu-id="167d3-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-521">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-522">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-522">1.0</span></span>|
|[<span data-ttu-id="167d3-523">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-523">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-524">ReadItem</span></span>|
|[<span data-ttu-id="167d3-525">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-525">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-526">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-526">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-527">例</span><span class="sxs-lookup"><span data-stu-id="167d3-527">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

---
---

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="167d3-528">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="167d3-528">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="167d3-529">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-529">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="167d3-530">型</span><span class="sxs-lookup"><span data-stu-id="167d3-530">Type</span></span>

*   [<span data-ttu-id="167d3-531">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="167d3-531">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="167d3-532">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-532">Requirements</span></span>

|<span data-ttu-id="167d3-533">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-533">Requirement</span></span>|<span data-ttu-id="167d3-534">値</span><span class="sxs-lookup"><span data-stu-id="167d3-534">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-535">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-535">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-536">1.3</span><span class="sxs-lookup"><span data-stu-id="167d3-536">1.3</span></span>|
|[<span data-ttu-id="167d3-537">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-537">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-538">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-538">ReadItem</span></span>|
|[<span data-ttu-id="167d3-539">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-539">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-540">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-540">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-541">例</span><span class="sxs-lookup"><span data-stu-id="167d3-541">Example</span></span>

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

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="167d3-542">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="167d3-542">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="167d3-543">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="167d3-543">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="167d3-544">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="167d3-544">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="167d3-545">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="167d3-545">Read mode</span></span>

<span data-ttu-id="167d3-546">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-546">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="167d3-547">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="167d3-547">Compose mode</span></span>

<span data-ttu-id="167d3-548">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-548">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="167d3-549">型</span><span class="sxs-lookup"><span data-stu-id="167d3-549">Type</span></span>

*   <span data-ttu-id="167d3-550">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="167d3-550">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-551">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-551">Requirements</span></span>

|<span data-ttu-id="167d3-552">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-552">Requirement</span></span>|<span data-ttu-id="167d3-553">値</span><span class="sxs-lookup"><span data-stu-id="167d3-553">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-554">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-554">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-555">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-555">1.0</span></span>|
|[<span data-ttu-id="167d3-556">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-556">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-557">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-557">ReadItem</span></span>|
|[<span data-ttu-id="167d3-558">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-558">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-559">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-559">Compose or Read</span></span>|

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="167d3-560">開催者:[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)|[開催者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="167d3-560">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="167d3-561">指定した会議の開催者の電子メールアドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-561">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="167d3-562">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="167d3-562">Read mode</span></span>

<span data-ttu-id="167d3-563">プロパティ`organizer`は、会議の開催者を表す[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-563">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="167d3-564">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="167d3-564">Compose mode</span></span>

<span data-ttu-id="167d3-565">プロパティ`organizer`は、開催者の値を取得するためのメソッドを提供する[オーガナイザー](/javascript/api/outlook/office.organizer)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-565">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="167d3-566">タイプ</span><span class="sxs-lookup"><span data-stu-id="167d3-566">Type</span></span>

*   <span data-ttu-id="167d3-567">[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails) | [開催者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="167d3-567">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-568">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-568">Requirements</span></span>

|<span data-ttu-id="167d3-569">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-569">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="167d3-570">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-571">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-571">1.0</span></span>|<span data-ttu-id="167d3-572">1.7</span><span class="sxs-lookup"><span data-stu-id="167d3-572">1.7</span></span>|
|[<span data-ttu-id="167d3-573">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-573">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-574">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-574">ReadItem</span></span>|<span data-ttu-id="167d3-575">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="167d3-575">ReadWriteItem</span></span>|
|[<span data-ttu-id="167d3-576">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-576">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-577">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-577">Read</span></span>|<span data-ttu-id="167d3-578">作成</span><span class="sxs-lookup"><span data-stu-id="167d3-578">Compose</span></span>|

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="167d3-579">(nullable) 定期的なスケジュール:[定期的](/javascript/api/outlook/office.recurrence)なアイテム</span><span class="sxs-lookup"><span data-stu-id="167d3-579">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="167d3-580">予定の定期的なパターンを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="167d3-580">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="167d3-581">会議出席依頼の定期的なパターンを取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-581">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="167d3-582">予定アイテムの読み取りおよび作成モード。</span><span class="sxs-lookup"><span data-stu-id="167d3-582">Read and compose modes for appointment items.</span></span> <span data-ttu-id="167d3-583">会議出席依頼アイテムの閲覧モード。</span><span class="sxs-lookup"><span data-stu-id="167d3-583">Read mode for meeting request items.</span></span>

<span data-ttu-id="167d3-584">この`recurrence`プロパティは、アイテムが series または series 内のインスタンスの場合、定期的な予定または会議出席依頼に対して[定期的](/javascript/api/outlook/office.recurrence)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-584">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="167d3-585">`null`は、単一の予定および1つの予定の会議出席依頼に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-585">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="167d3-586">`undefined`は、会議出席依頼ではないメッセージに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-586">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="167d3-587">注: 会議出席依頼に`itemClass`は、IPM という値があります。出席依頼。</span><span class="sxs-lookup"><span data-stu-id="167d3-587">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="167d3-588">注: 定期的なオブジェクトが`null`の場合は、そのオブジェクトが単一の予定または1つの予定の会議出席依頼であり、データ系列の一部ではないことを示します。</span><span class="sxs-lookup"><span data-stu-id="167d3-588">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="167d3-589">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="167d3-589">Read mode</span></span>

<span data-ttu-id="167d3-590">この`recurrence`プロパティは、定期的な予定を表す[定期的](/javascript/api/outlook/office.recurrence)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-590">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="167d3-591">これは、予定および会議出席依頼に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-591">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="167d3-592">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="167d3-592">Compose mode</span></span>

<span data-ttu-id="167d3-593">この`recurrence`プロパティは、予定の繰り返しを管理するためのメソッドを提供する[定期的](/javascript/api/outlook/office.recurrence)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-593">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="167d3-594">これは予定に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-594">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="167d3-595">タイプ</span><span class="sxs-lookup"><span data-stu-id="167d3-595">Type</span></span>

* [<span data-ttu-id="167d3-596">Recurrence</span><span class="sxs-lookup"><span data-stu-id="167d3-596">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="167d3-597">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-597">Requirement</span></span>|<span data-ttu-id="167d3-598">値</span><span class="sxs-lookup"><span data-stu-id="167d3-598">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-599">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-599">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-600">1.7</span><span class="sxs-lookup"><span data-stu-id="167d3-600">1.7</span></span>|
|[<span data-ttu-id="167d3-601">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-601">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-602">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-602">ReadItem</span></span>|
|[<span data-ttu-id="167d3-603">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-603">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-604">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-604">Compose or Read</span></span>|

---
---

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="167d3-605">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="167d3-605">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="167d3-606">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="167d3-606">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="167d3-607">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="167d3-607">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="167d3-608">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="167d3-608">Read mode</span></span>

<span data-ttu-id="167d3-609">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-609">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="167d3-610">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="167d3-610">Compose mode</span></span>

<span data-ttu-id="167d3-611">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-611">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="167d3-612">タイプ</span><span class="sxs-lookup"><span data-stu-id="167d3-612">Type</span></span>

*   <span data-ttu-id="167d3-613">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="167d3-613">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-614">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-614">Requirements</span></span>

|<span data-ttu-id="167d3-615">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-615">Requirement</span></span>|<span data-ttu-id="167d3-616">値</span><span class="sxs-lookup"><span data-stu-id="167d3-616">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-617">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-617">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-618">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-618">1.0</span></span>|
|[<span data-ttu-id="167d3-619">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-619">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-620">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-620">ReadItem</span></span>|
|[<span data-ttu-id="167d3-621">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-621">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-622">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-622">Compose or Read</span></span>|

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="167d3-623">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="167d3-623">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="167d3-p128">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="167d3-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="167d3-p129">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsfrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="167d3-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-628">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="167d3-628">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="167d3-629">型</span><span class="sxs-lookup"><span data-stu-id="167d3-629">Type</span></span>

*   [<span data-ttu-id="167d3-630">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="167d3-630">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="167d3-631">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-631">Requirements</span></span>

|<span data-ttu-id="167d3-632">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-632">Requirement</span></span>|<span data-ttu-id="167d3-633">値</span><span class="sxs-lookup"><span data-stu-id="167d3-633">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-634">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-634">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-635">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-635">1.0</span></span>|
|[<span data-ttu-id="167d3-636">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-636">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-637">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-637">ReadItem</span></span>|
|[<span data-ttu-id="167d3-638">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-638">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-639">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-639">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-640">例</span><span class="sxs-lookup"><span data-stu-id="167d3-640">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="167d3-641">(nullable) 系列 id: String</span><span class="sxs-lookup"><span data-stu-id="167d3-641">(nullable) seriesId :String</span></span>

<span data-ttu-id="167d3-642">インスタンスが属する系列の id を取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-642">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="167d3-643">OWA および Outlook で、は`seriesId` 、このアイテムが属する親 (シリーズ) アイテムの Exchange Web サービス (EWS) ID を返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-643">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="167d3-644">ただし、iOS と Android では、 `seriesId`は親アイテムの REST ID を返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-644">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-645">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="167d3-645">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="167d3-646">`seriesId`プロパティが outlook REST API で使用される outlook id と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="167d3-646">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="167d3-647">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="167d3-647">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="167d3-648">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="167d3-648">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="167d3-649">この`seriesId`プロパティは`null` 、単一の予定、系列のアイテム、会議出席依頼などの親アイテムを持たないアイテムに`undefined`対して、会議出席依頼以外のアイテムに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-649">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="167d3-650">Type</span><span class="sxs-lookup"><span data-stu-id="167d3-650">Type</span></span>

* <span data-ttu-id="167d3-651">String</span><span class="sxs-lookup"><span data-stu-id="167d3-651">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-652">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-652">Requirements</span></span>

|<span data-ttu-id="167d3-653">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-653">Requirement</span></span>|<span data-ttu-id="167d3-654">値</span><span class="sxs-lookup"><span data-stu-id="167d3-654">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-655">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-655">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-656">1.7</span><span class="sxs-lookup"><span data-stu-id="167d3-656">1.7</span></span>|
|[<span data-ttu-id="167d3-657">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-657">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-658">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-658">ReadItem</span></span>|
|[<span data-ttu-id="167d3-659">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-659">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-660">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-660">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-661">例</span><span class="sxs-lookup"><span data-stu-id="167d3-661">Example</span></span>

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

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="167d3-662">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="167d3-662">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="167d3-663">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="167d3-663">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="167d3-p132">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="167d3-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="167d3-666">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="167d3-666">Read mode</span></span>

<span data-ttu-id="167d3-667">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-667">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="167d3-668">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="167d3-668">Compose mode</span></span>

<span data-ttu-id="167d3-669">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-669">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="167d3-670">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="167d3-670">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="167d3-671">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="167d3-671">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="167d3-672">型</span><span class="sxs-lookup"><span data-stu-id="167d3-672">Type</span></span>

*   <span data-ttu-id="167d3-673">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="167d3-673">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-674">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-674">Requirements</span></span>

|<span data-ttu-id="167d3-675">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-675">Requirement</span></span>|<span data-ttu-id="167d3-676">値</span><span class="sxs-lookup"><span data-stu-id="167d3-676">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-677">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-677">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-678">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-678">1.0</span></span>|
|[<span data-ttu-id="167d3-679">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-679">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-680">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-680">ReadItem</span></span>|
|[<span data-ttu-id="167d3-681">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-681">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-682">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-682">Compose or Read</span></span>|

---
---

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="167d3-683">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="167d3-683">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="167d3-684">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="167d3-684">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="167d3-685">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="167d3-685">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="167d3-686">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="167d3-686">Read mode</span></span>

<span data-ttu-id="167d3-p133">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="167d3-689">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="167d3-689">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="167d3-690">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="167d3-690">Compose mode</span></span>
<span data-ttu-id="167d3-691">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-691">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="167d3-692">型</span><span class="sxs-lookup"><span data-stu-id="167d3-692">Type</span></span>

*   <span data-ttu-id="167d3-693">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="167d3-693">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-694">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-694">Requirements</span></span>

|<span data-ttu-id="167d3-695">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-695">Requirement</span></span>|<span data-ttu-id="167d3-696">値</span><span class="sxs-lookup"><span data-stu-id="167d3-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-697">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-698">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-698">1.0</span></span>|
|[<span data-ttu-id="167d3-699">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-699">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-700">ReadItem</span></span>|
|[<span data-ttu-id="167d3-701">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-701">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-702">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-702">Compose or Read</span></span>|

---
---

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="167d3-703">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="167d3-703">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="167d3-704">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="167d3-704">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="167d3-705">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="167d3-705">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="167d3-706">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="167d3-706">Read mode</span></span>

<span data-ttu-id="167d3-p135">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="167d3-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="167d3-709">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="167d3-709">Compose mode</span></span>

<span data-ttu-id="167d3-710">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-710">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="167d3-711">型</span><span class="sxs-lookup"><span data-stu-id="167d3-711">Type</span></span>

*   <span data-ttu-id="167d3-712">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="167d3-712">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-713">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-713">Requirements</span></span>

|<span data-ttu-id="167d3-714">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-714">Requirement</span></span>|<span data-ttu-id="167d3-715">値</span><span class="sxs-lookup"><span data-stu-id="167d3-715">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-716">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-716">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-717">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-717">1.0</span></span>|
|[<span data-ttu-id="167d3-718">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-718">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-719">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-719">ReadItem</span></span>|
|[<span data-ttu-id="167d3-720">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-720">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-721">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-721">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="167d3-722">メソッド</span><span class="sxs-lookup"><span data-stu-id="167d3-722">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="167d3-723">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="167d3-723">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="167d3-724">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="167d3-724">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="167d3-725">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="167d3-725">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="167d3-726">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-726">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-727">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-727">Parameters</span></span>
|<span data-ttu-id="167d3-728">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-728">Name</span></span>|<span data-ttu-id="167d3-729">種類</span><span class="sxs-lookup"><span data-stu-id="167d3-729">Type</span></span>|<span data-ttu-id="167d3-730">属性</span><span class="sxs-lookup"><span data-stu-id="167d3-730">Attributes</span></span>|<span data-ttu-id="167d3-731">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-731">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="167d3-732">文字列</span><span class="sxs-lookup"><span data-stu-id="167d3-732">String</span></span>||<span data-ttu-id="167d3-p136">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="167d3-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="167d3-735">String</span><span class="sxs-lookup"><span data-stu-id="167d3-735">String</span></span>||<span data-ttu-id="167d3-p137">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="167d3-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="167d3-738">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-738">Object</span></span>|<span data-ttu-id="167d3-739">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-739">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-740">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="167d3-740">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="167d3-741">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-741">Object</span></span>|<span data-ttu-id="167d3-742">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-742">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-743">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-743">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="167d3-744">Boolean</span><span class="sxs-lookup"><span data-stu-id="167d3-744">Boolean</span></span>|<span data-ttu-id="167d3-745">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-745">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-746">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="167d3-746">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="167d3-747">function</span><span class="sxs-lookup"><span data-stu-id="167d3-747">function</span></span>|<span data-ttu-id="167d3-748">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-748">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-749">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-749">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="167d3-750">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-750">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="167d3-751">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="167d3-751">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="167d3-752">エラー</span><span class="sxs-lookup"><span data-stu-id="167d3-752">Errors</span></span>

|<span data-ttu-id="167d3-753">エラー コード</span><span class="sxs-lookup"><span data-stu-id="167d3-753">Error code</span></span>|<span data-ttu-id="167d3-754">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-754">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="167d3-755">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="167d3-755">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="167d3-756">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="167d3-756">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="167d3-757">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="167d3-757">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-758">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-758">Requirements</span></span>

|<span data-ttu-id="167d3-759">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-759">Requirement</span></span>|<span data-ttu-id="167d3-760">値</span><span class="sxs-lookup"><span data-stu-id="167d3-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-761">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-761">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-762">1.1</span><span class="sxs-lookup"><span data-stu-id="167d3-762">1.1</span></span>|
|[<span data-ttu-id="167d3-763">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-763">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-764">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="167d3-764">ReadWriteItem</span></span>|
|[<span data-ttu-id="167d3-765">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-765">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-766">作成</span><span class="sxs-lookup"><span data-stu-id="167d3-766">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="167d3-767">例</span><span class="sxs-lookup"><span data-stu-id="167d3-767">Examples</span></span>

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

<span data-ttu-id="167d3-768">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="167d3-768">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="167d3-769">addFileAttachmentFromBase64Async (base64File, attachmentname, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="167d3-769">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="167d3-770">base64 エンコードのファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="167d3-770">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="167d3-771">この`addFileAttachmentFromBase64Async`メソッドは、base64 エンコードからファイルをアップロードし、新規作成フォームのアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="167d3-771">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="167d3-772">このメソッドは、AsyncResult オブジェクトの添付ファイル識別子を返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-772">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="167d3-773">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-773">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-774">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-774">Parameters</span></span>

|<span data-ttu-id="167d3-775">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-775">Name</span></span>|<span data-ttu-id="167d3-776">種類</span><span class="sxs-lookup"><span data-stu-id="167d3-776">Type</span></span>|<span data-ttu-id="167d3-777">属性</span><span class="sxs-lookup"><span data-stu-id="167d3-777">Attributes</span></span>|<span data-ttu-id="167d3-778">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-778">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="167d3-779">文字列</span><span class="sxs-lookup"><span data-stu-id="167d3-779">String</span></span>||<span data-ttu-id="167d3-780">電子メールまたはイベントに追加する画像またはファイルの、base64 でエンコードされたコンテンツ。</span><span class="sxs-lookup"><span data-stu-id="167d3-780">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="167d3-781">String</span><span class="sxs-lookup"><span data-stu-id="167d3-781">String</span></span>||<span data-ttu-id="167d3-p139">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="167d3-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="167d3-784">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-784">Object</span></span>|<span data-ttu-id="167d3-785">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-785">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-786">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="167d3-786">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="167d3-787">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-787">Object</span></span>|<span data-ttu-id="167d3-788">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-788">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-789">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-789">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="167d3-790">Boolean</span><span class="sxs-lookup"><span data-stu-id="167d3-790">Boolean</span></span>|<span data-ttu-id="167d3-791">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-791">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-792">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="167d3-792">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="167d3-793">function</span><span class="sxs-lookup"><span data-stu-id="167d3-793">function</span></span>|<span data-ttu-id="167d3-794">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-794">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-795">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-795">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="167d3-796">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-796">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="167d3-797">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="167d3-797">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="167d3-798">エラー</span><span class="sxs-lookup"><span data-stu-id="167d3-798">Errors</span></span>

|<span data-ttu-id="167d3-799">エラー コード</span><span class="sxs-lookup"><span data-stu-id="167d3-799">Error code</span></span>|<span data-ttu-id="167d3-800">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-800">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="167d3-801">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="167d3-801">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="167d3-802">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="167d3-802">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="167d3-803">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="167d3-803">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-804">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-804">Requirements</span></span>

|<span data-ttu-id="167d3-805">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-805">Requirement</span></span>|<span data-ttu-id="167d3-806">値</span><span class="sxs-lookup"><span data-stu-id="167d3-806">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-807">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-807">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-808">プレビュー</span><span class="sxs-lookup"><span data-stu-id="167d3-808">Preview</span></span>|
|[<span data-ttu-id="167d3-809">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-809">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-810">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="167d3-810">ReadWriteItem</span></span>|
|[<span data-ttu-id="167d3-811">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-811">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-812">作成</span><span class="sxs-lookup"><span data-stu-id="167d3-812">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="167d3-813">例</span><span class="sxs-lookup"><span data-stu-id="167d3-813">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="167d3-814">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="167d3-814">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="167d3-815">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="167d3-815">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="167d3-816">現在、サポートされて`Office.EventType.AttachmentsChanged`いる`Office.EventType.AppointmentTimeChanged`イベント`Office.EventType.EnhancedLocationsChanged`の`Office.EventType.RecipientsChanged`種類は`Office.EventType.RecurrenceChanged`、、、、、です。</span><span class="sxs-lookup"><span data-stu-id="167d3-816">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-817">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-817">Parameters</span></span>

| <span data-ttu-id="167d3-818">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-818">Name</span></span> | <span data-ttu-id="167d3-819">型</span><span class="sxs-lookup"><span data-stu-id="167d3-819">Type</span></span> | <span data-ttu-id="167d3-820">属性</span><span class="sxs-lookup"><span data-stu-id="167d3-820">Attributes</span></span> | <span data-ttu-id="167d3-821">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-821">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="167d3-822">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="167d3-822">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="167d3-823">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="167d3-823">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="167d3-824">関数</span><span class="sxs-lookup"><span data-stu-id="167d3-824">Function</span></span> || <span data-ttu-id="167d3-p140">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="167d3-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="167d3-828">Object</span><span class="sxs-lookup"><span data-stu-id="167d3-828">Object</span></span> | <span data-ttu-id="167d3-829">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-829">&lt;optional&gt;</span></span> | <span data-ttu-id="167d3-830">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="167d3-830">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="167d3-831">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-831">Object</span></span> | <span data-ttu-id="167d3-832">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-832">&lt;optional&gt;</span></span> | <span data-ttu-id="167d3-833">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-833">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="167d3-834">function</span><span class="sxs-lookup"><span data-stu-id="167d3-834">function</span></span>| <span data-ttu-id="167d3-835">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-835">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-836">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-836">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-837">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-837">Requirements</span></span>

|<span data-ttu-id="167d3-838">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-838">Requirement</span></span>| <span data-ttu-id="167d3-839">値</span><span class="sxs-lookup"><span data-stu-id="167d3-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-840">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="167d3-841">1.7</span><span class="sxs-lookup"><span data-stu-id="167d3-841">1.7</span></span> |
|[<span data-ttu-id="167d3-842">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-842">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="167d3-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-843">ReadItem</span></span> |
|[<span data-ttu-id="167d3-844">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-844">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="167d3-845">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-845">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="167d3-846">例</span><span class="sxs-lookup"><span data-stu-id="167d3-846">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="167d3-847">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="167d3-847">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="167d3-848">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="167d3-848">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="167d3-p141">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="167d3-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="167d3-852">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-852">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="167d3-853">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="167d3-853">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-854">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-854">Parameters</span></span>

|<span data-ttu-id="167d3-855">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-855">Name</span></span>|<span data-ttu-id="167d3-856">型</span><span class="sxs-lookup"><span data-stu-id="167d3-856">Type</span></span>|<span data-ttu-id="167d3-857">属性</span><span class="sxs-lookup"><span data-stu-id="167d3-857">Attributes</span></span>|<span data-ttu-id="167d3-858">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-858">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="167d3-859">文字列</span><span class="sxs-lookup"><span data-stu-id="167d3-859">String</span></span>||<span data-ttu-id="167d3-p142">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="167d3-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="167d3-862">String</span><span class="sxs-lookup"><span data-stu-id="167d3-862">String</span></span>||<span data-ttu-id="167d3-863">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="167d3-863">The subject of the item to be attached.</span></span> <span data-ttu-id="167d3-864">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="167d3-864">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="167d3-865">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-865">Object</span></span>|<span data-ttu-id="167d3-866">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-866">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-867">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="167d3-867">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="167d3-868">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-868">Object</span></span>|<span data-ttu-id="167d3-869">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-869">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-870">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-870">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="167d3-871">function</span><span class="sxs-lookup"><span data-stu-id="167d3-871">function</span></span>|<span data-ttu-id="167d3-872">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-872">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-873">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-873">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="167d3-874">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-874">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="167d3-875">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="167d3-875">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="167d3-876">エラー</span><span class="sxs-lookup"><span data-stu-id="167d3-876">Errors</span></span>

|<span data-ttu-id="167d3-877">エラー コード</span><span class="sxs-lookup"><span data-stu-id="167d3-877">Error code</span></span>|<span data-ttu-id="167d3-878">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-878">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="167d3-879">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="167d3-879">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-880">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-880">Requirements</span></span>

|<span data-ttu-id="167d3-881">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-881">Requirement</span></span>|<span data-ttu-id="167d3-882">値</span><span class="sxs-lookup"><span data-stu-id="167d3-882">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-883">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-883">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-884">1.1</span><span class="sxs-lookup"><span data-stu-id="167d3-884">1.1</span></span>|
|[<span data-ttu-id="167d3-885">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-885">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-886">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="167d3-886">ReadWriteItem</span></span>|
|[<span data-ttu-id="167d3-887">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-887">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-888">作成</span><span class="sxs-lookup"><span data-stu-id="167d3-888">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-889">例</span><span class="sxs-lookup"><span data-stu-id="167d3-889">Example</span></span>

<span data-ttu-id="167d3-890">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-890">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="167d3-891">close()</span><span class="sxs-lookup"><span data-stu-id="167d3-891">close()</span></span>

<span data-ttu-id="167d3-892">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="167d3-892">Closes the current item that is being composed.</span></span>

<span data-ttu-id="167d3-p144">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="167d3-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-895">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-895">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="167d3-896">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="167d3-896">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-897">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-897">Requirements</span></span>

|<span data-ttu-id="167d3-898">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-898">Requirement</span></span>|<span data-ttu-id="167d3-899">値</span><span class="sxs-lookup"><span data-stu-id="167d3-899">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-900">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-900">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-901">1.3</span><span class="sxs-lookup"><span data-stu-id="167d3-901">1.3</span></span>|
|[<span data-ttu-id="167d3-902">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-902">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-903">制限あり</span><span class="sxs-lookup"><span data-stu-id="167d3-903">Restricted</span></span>|
|[<span data-ttu-id="167d3-904">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-904">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-905">新規作成</span><span class="sxs-lookup"><span data-stu-id="167d3-905">Compose</span></span>|

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="167d3-906">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="167d3-906">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="167d3-907">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-907">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-908">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="167d3-908">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="167d3-909">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-909">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="167d3-910">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="167d3-910">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="167d3-p145">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="167d3-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-914">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-914">Parameters</span></span>

|<span data-ttu-id="167d3-915">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-915">Name</span></span>|<span data-ttu-id="167d3-916">型</span><span class="sxs-lookup"><span data-stu-id="167d3-916">Type</span></span>|<span data-ttu-id="167d3-917">属性</span><span class="sxs-lookup"><span data-stu-id="167d3-917">Attributes</span></span>|<span data-ttu-id="167d3-918">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-918">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="167d3-919">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="167d3-919">String &#124; Object</span></span>||<span data-ttu-id="167d3-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="167d3-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="167d3-922">**または**</span><span class="sxs-lookup"><span data-stu-id="167d3-922">**OR**</span></span><br/><span data-ttu-id="167d3-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="167d3-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="167d3-925">String</span><span class="sxs-lookup"><span data-stu-id="167d3-925">String</span></span>|<span data-ttu-id="167d3-926">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-926">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="167d3-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="167d3-929">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-929">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="167d3-930">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-930">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-931">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="167d3-931">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="167d3-932">文字列</span><span class="sxs-lookup"><span data-stu-id="167d3-932">String</span></span>||<span data-ttu-id="167d3-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="167d3-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="167d3-935">String</span><span class="sxs-lookup"><span data-stu-id="167d3-935">String</span></span>||<span data-ttu-id="167d3-936">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="167d3-936">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="167d3-937">文字列</span><span class="sxs-lookup"><span data-stu-id="167d3-937">String</span></span>||<span data-ttu-id="167d3-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="167d3-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="167d3-940">ブール値</span><span class="sxs-lookup"><span data-stu-id="167d3-940">Boolean</span></span>||<span data-ttu-id="167d3-p151">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="167d3-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="167d3-943">String</span><span class="sxs-lookup"><span data-stu-id="167d3-943">String</span></span>||<span data-ttu-id="167d3-p152">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="167d3-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="167d3-947">function</span><span class="sxs-lookup"><span data-stu-id="167d3-947">function</span></span>|<span data-ttu-id="167d3-948">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-948">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-949">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-949">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-950">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-950">Requirements</span></span>

|<span data-ttu-id="167d3-951">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-951">Requirement</span></span>|<span data-ttu-id="167d3-952">値</span><span class="sxs-lookup"><span data-stu-id="167d3-952">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-953">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-953">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-954">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-954">1.0</span></span>|
|[<span data-ttu-id="167d3-955">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-955">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-956">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-956">ReadItem</span></span>|
|[<span data-ttu-id="167d3-957">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-957">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-958">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-958">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="167d3-959">例</span><span class="sxs-lookup"><span data-stu-id="167d3-959">Examples</span></span>

<span data-ttu-id="167d3-960">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="167d3-960">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="167d3-961">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="167d3-961">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="167d3-962">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="167d3-962">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="167d3-963">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="167d3-963">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="167d3-964">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="167d3-964">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="167d3-965">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="167d3-965">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="167d3-966">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="167d3-966">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="167d3-967">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-967">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-968">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="167d3-968">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="167d3-969">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-969">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="167d3-970">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="167d3-970">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="167d3-p153">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="167d3-p153">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-974">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-974">Parameters</span></span>

|<span data-ttu-id="167d3-975">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-975">Name</span></span>|<span data-ttu-id="167d3-976">型</span><span class="sxs-lookup"><span data-stu-id="167d3-976">Type</span></span>|<span data-ttu-id="167d3-977">属性</span><span class="sxs-lookup"><span data-stu-id="167d3-977">Attributes</span></span>|<span data-ttu-id="167d3-978">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-978">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="167d3-979">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="167d3-979">String &#124; Object</span></span>||<span data-ttu-id="167d3-p154">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="167d3-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="167d3-982">**または**</span><span class="sxs-lookup"><span data-stu-id="167d3-982">**OR**</span></span><br/><span data-ttu-id="167d3-p155">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="167d3-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="167d3-985">String</span><span class="sxs-lookup"><span data-stu-id="167d3-985">String</span></span>|<span data-ttu-id="167d3-986">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-986">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-p156">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="167d3-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="167d3-989">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-989">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="167d3-990">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-990">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-991">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="167d3-991">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="167d3-992">文字列</span><span class="sxs-lookup"><span data-stu-id="167d3-992">String</span></span>||<span data-ttu-id="167d3-p157">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="167d3-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="167d3-995">String</span><span class="sxs-lookup"><span data-stu-id="167d3-995">String</span></span>||<span data-ttu-id="167d3-996">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="167d3-996">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="167d3-997">文字列</span><span class="sxs-lookup"><span data-stu-id="167d3-997">String</span></span>||<span data-ttu-id="167d3-p158">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="167d3-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="167d3-1000">ブール値</span><span class="sxs-lookup"><span data-stu-id="167d3-1000">Boolean</span></span>||<span data-ttu-id="167d3-p159">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="167d3-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="167d3-1003">String</span><span class="sxs-lookup"><span data-stu-id="167d3-1003">String</span></span>||<span data-ttu-id="167d3-p160">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="167d3-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="167d3-1007">function</span><span class="sxs-lookup"><span data-stu-id="167d3-1007">function</span></span>|<span data-ttu-id="167d3-1008">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1008">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1009">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1009">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-1010">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1010">Requirements</span></span>

|<span data-ttu-id="167d3-1011">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1011">Requirement</span></span>|<span data-ttu-id="167d3-1012">値</span><span class="sxs-lookup"><span data-stu-id="167d3-1012">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-1013">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-1013">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-1014">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-1014">1.0</span></span>|
|[<span data-ttu-id="167d3-1015">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1015">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-1016">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-1016">ReadItem</span></span>|
|[<span data-ttu-id="167d3-1017">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-1017">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-1018">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-1018">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="167d3-1019">例</span><span class="sxs-lookup"><span data-stu-id="167d3-1019">Examples</span></span>

<span data-ttu-id="167d3-1020">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1020">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="167d3-1021">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1021">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="167d3-1022">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1022">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="167d3-1023">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1023">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="167d3-1024">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1024">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="167d3-1025">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1025">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="167d3-1026">getattachmentcontentasync (attachmentId, [options], [callback]) > [attachmentcontent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="167d3-1026">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="167d3-1027">メッセージまたは予定から指定された添付ファイルを取得し`AttachmentContent` 、それをオブジェクトとして返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1027">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="167d3-1028">メソッド`getAttachmentContentAsync`は、指定された id の添付ファイルをアイテムから取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1028">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="167d3-1029">ベストプラクティスとして、識別子を使用して、または`getAttachmentsAsync` `item.attachments`の呼び出しで attachmentIds を取得したのと同じセッションの添付ファイルを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="167d3-1029">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="167d3-1030">Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="167d3-1030">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="167d3-1031">ユーザーがアプリを閉じたとき、またはインラインフォームの作成が開始されたときに、別のウィンドウで続行するためにフォームをポップアウトした後、セッションが終了します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1031">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-1032">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-1032">Parameters</span></span>

|<span data-ttu-id="167d3-1033">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-1033">Name</span></span>|<span data-ttu-id="167d3-1034">型</span><span class="sxs-lookup"><span data-stu-id="167d3-1034">Type</span></span>|<span data-ttu-id="167d3-1035">属性</span><span class="sxs-lookup"><span data-stu-id="167d3-1035">Attributes</span></span>|<span data-ttu-id="167d3-1036">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-1036">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="167d3-1037">String</span><span class="sxs-lookup"><span data-stu-id="167d3-1037">String</span></span>||<span data-ttu-id="167d3-1038">取得する添付ファイルの識別子を指定します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1038">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="167d3-1039">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-1039">Object</span></span>|<span data-ttu-id="167d3-1040">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1041">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="167d3-1041">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="167d3-1042">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-1042">Object</span></span>|<span data-ttu-id="167d3-1043">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1043">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1044">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1044">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="167d3-1045">関数</span><span class="sxs-lookup"><span data-stu-id="167d3-1045">function</span></span>|<span data-ttu-id="167d3-1046">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1046">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1047">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1047">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-1048">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1048">Requirements</span></span>

|<span data-ttu-id="167d3-1049">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1049">Requirement</span></span>|<span data-ttu-id="167d3-1050">値</span><span class="sxs-lookup"><span data-stu-id="167d3-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-1051">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-1051">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-1052">プレビュー</span><span class="sxs-lookup"><span data-stu-id="167d3-1052">Preview</span></span>|
|[<span data-ttu-id="167d3-1053">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1053">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-1054">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-1054">ReadItem</span></span>|
|[<span data-ttu-id="167d3-1055">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-1055">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-1056">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-1056">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="167d3-1057">戻り値:</span><span class="sxs-lookup"><span data-stu-id="167d3-1057">Returns:</span></span>

<span data-ttu-id="167d3-1058">型: [attachmentcontent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="167d3-1058">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="167d3-1059">例</span><span class="sxs-lookup"><span data-stu-id="167d3-1059">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="167d3-1060">getAttachmentsAsync ([オプション], [callback])] > <[attachmentdetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="167d3-1060">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="167d3-1061">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1061">Gets the item's attachments as an array.</span></span> <span data-ttu-id="167d3-1062">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="167d3-1062">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-1063">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-1063">Parameters</span></span>

|<span data-ttu-id="167d3-1064">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-1064">Name</span></span>|<span data-ttu-id="167d3-1065">型</span><span class="sxs-lookup"><span data-stu-id="167d3-1065">Type</span></span>|<span data-ttu-id="167d3-1066">属性</span><span class="sxs-lookup"><span data-stu-id="167d3-1066">Attributes</span></span>|<span data-ttu-id="167d3-1067">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-1067">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="167d3-1068">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-1068">Object</span></span>|<span data-ttu-id="167d3-1069">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1070">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="167d3-1070">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="167d3-1071">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-1071">Object</span></span>|<span data-ttu-id="167d3-1072">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1072">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1073">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1073">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="167d3-1074">function</span><span class="sxs-lookup"><span data-stu-id="167d3-1074">function</span></span>|<span data-ttu-id="167d3-1075">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1075">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1076">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1076">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-1077">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1077">Requirements</span></span>

|<span data-ttu-id="167d3-1078">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1078">Requirement</span></span>|<span data-ttu-id="167d3-1079">値</span><span class="sxs-lookup"><span data-stu-id="167d3-1079">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-1080">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-1080">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-1081">プレビュー</span><span class="sxs-lookup"><span data-stu-id="167d3-1081">Preview</span></span>|
|[<span data-ttu-id="167d3-1082">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1082">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-1083">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-1083">ReadItem</span></span>|
|[<span data-ttu-id="167d3-1084">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-1084">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-1085">作成</span><span class="sxs-lookup"><span data-stu-id="167d3-1085">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="167d3-1086">戻り値:</span><span class="sxs-lookup"><span data-stu-id="167d3-1086">Returns:</span></span>

<span data-ttu-id="167d3-1087">型: <[attachmentdetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="167d3-1087">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="167d3-1088">例</span><span class="sxs-lookup"><span data-stu-id="167d3-1088">Example</span></span>

<span data-ttu-id="167d3-1089">次の例では、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1089">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="167d3-1090">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="167d3-1090">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="167d3-1091">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1091">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-1092">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="167d3-1092">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-1093">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1093">Requirements</span></span>

|<span data-ttu-id="167d3-1094">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1094">Requirement</span></span>|<span data-ttu-id="167d3-1095">値</span><span class="sxs-lookup"><span data-stu-id="167d3-1095">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-1096">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-1096">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-1097">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-1097">1.0</span></span>|
|[<span data-ttu-id="167d3-1098">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1098">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-1099">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-1099">ReadItem</span></span>|
|[<span data-ttu-id="167d3-1100">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-1100">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-1101">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-1101">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="167d3-1102">戻り値:</span><span class="sxs-lookup"><span data-stu-id="167d3-1102">Returns:</span></span>

<span data-ttu-id="167d3-1103">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="167d3-1103">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="167d3-1104">例</span><span class="sxs-lookup"><span data-stu-id="167d3-1104">Example</span></span>

<span data-ttu-id="167d3-1105">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="167d3-1105">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="167d3-1106">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="167d3-1106">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="167d3-1107">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1107">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-1108">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="167d3-1108">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-1109">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-1109">Parameters</span></span>

|<span data-ttu-id="167d3-1110">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-1110">Name</span></span>|<span data-ttu-id="167d3-1111">型</span><span class="sxs-lookup"><span data-stu-id="167d3-1111">Type</span></span>|<span data-ttu-id="167d3-1112">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-1112">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="167d3-1113">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="167d3-1113">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="167d3-1114">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="167d3-1114">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-1115">Requirements</span><span class="sxs-lookup"><span data-stu-id="167d3-1115">Requirements</span></span>

|<span data-ttu-id="167d3-1116">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1116">Requirement</span></span>|<span data-ttu-id="167d3-1117">値</span><span class="sxs-lookup"><span data-stu-id="167d3-1117">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-1118">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-1118">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-1119">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-1119">1.0</span></span>|
|[<span data-ttu-id="167d3-1120">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1120">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-1121">制限あり</span><span class="sxs-lookup"><span data-stu-id="167d3-1121">Restricted</span></span>|
|[<span data-ttu-id="167d3-1122">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-1122">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-1123">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-1123">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="167d3-1124">戻り値:</span><span class="sxs-lookup"><span data-stu-id="167d3-1124">Returns:</span></span>

<span data-ttu-id="167d3-1125">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1125">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="167d3-1126">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1126">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="167d3-1127">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="167d3-1127">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="167d3-1128">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="167d3-1128">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="167d3-1129">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="167d3-1129">Value of `entityType`</span></span>|<span data-ttu-id="167d3-1130">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="167d3-1130">Type of objects in returned array</span></span>|<span data-ttu-id="167d3-1131">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1131">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="167d3-1132">String</span><span class="sxs-lookup"><span data-stu-id="167d3-1132">String</span></span>|<span data-ttu-id="167d3-1133">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="167d3-1133">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="167d3-1134">連絡先</span><span class="sxs-lookup"><span data-stu-id="167d3-1134">Contact</span></span>|<span data-ttu-id="167d3-1135">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="167d3-1135">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="167d3-1136">文字列</span><span class="sxs-lookup"><span data-stu-id="167d3-1136">String</span></span>|<span data-ttu-id="167d3-1137">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="167d3-1137">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="167d3-1138">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="167d3-1138">MeetingSuggestion</span></span>|<span data-ttu-id="167d3-1139">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="167d3-1139">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="167d3-1140">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="167d3-1140">PhoneNumber</span></span>|<span data-ttu-id="167d3-1141">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="167d3-1141">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="167d3-1142">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="167d3-1142">TaskSuggestion</span></span>|<span data-ttu-id="167d3-1143">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="167d3-1143">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="167d3-1144">文字列</span><span class="sxs-lookup"><span data-stu-id="167d3-1144">String</span></span>|<span data-ttu-id="167d3-1145">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="167d3-1145">**Restricted**</span></span>|

<span data-ttu-id="167d3-1146">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="167d3-1146">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="167d3-1147">例</span><span class="sxs-lookup"><span data-stu-id="167d3-1147">Example</span></span>

<span data-ttu-id="167d3-1148">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1148">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="167d3-1149">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="167d3-1149">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="167d3-1150">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1150">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-1151">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="167d3-1151">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="167d3-1152">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1152">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-1153">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-1153">Parameters</span></span>

|<span data-ttu-id="167d3-1154">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-1154">Name</span></span>|<span data-ttu-id="167d3-1155">型</span><span class="sxs-lookup"><span data-stu-id="167d3-1155">Type</span></span>|<span data-ttu-id="167d3-1156">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-1156">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="167d3-1157">String</span><span class="sxs-lookup"><span data-stu-id="167d3-1157">String</span></span>|<span data-ttu-id="167d3-1158">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="167d3-1158">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-1159">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1159">Requirements</span></span>

|<span data-ttu-id="167d3-1160">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1160">Requirement</span></span>|<span data-ttu-id="167d3-1161">値</span><span class="sxs-lookup"><span data-stu-id="167d3-1161">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-1162">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-1162">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-1163">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-1163">1.0</span></span>|
|[<span data-ttu-id="167d3-1164">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1164">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-1165">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-1165">ReadItem</span></span>|
|[<span data-ttu-id="167d3-1166">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-1166">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-1167">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-1167">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="167d3-1168">戻り値:</span><span class="sxs-lookup"><span data-stu-id="167d3-1168">Returns:</span></span>

<span data-ttu-id="167d3-p164">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="167d3-1171">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="167d3-1171">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="167d3-1172">、office.context.mailbox.item.getinitializationcontextasync ([オプション], [callback])</span><span class="sxs-lookup"><span data-stu-id="167d3-1172">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="167d3-1173">[アクション可能なメッセージによってアドインがアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されたときに渡される初期化データを取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1173">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-1174">このメソッドは、outlook 2016 またはそれ以降のバージョンの Windows (16.0.8413.1000 より後のバージョン) および outlook on the Office 365 でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="167d3-1174">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-1175">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-1175">Parameters</span></span>

|<span data-ttu-id="167d3-1176">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-1176">Name</span></span>|<span data-ttu-id="167d3-1177">型</span><span class="sxs-lookup"><span data-stu-id="167d3-1177">Type</span></span>|<span data-ttu-id="167d3-1178">属性</span><span class="sxs-lookup"><span data-stu-id="167d3-1178">Attributes</span></span>|<span data-ttu-id="167d3-1179">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-1179">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="167d3-1180">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-1180">Object</span></span>|<span data-ttu-id="167d3-1181">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1181">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1182">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="167d3-1182">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="167d3-1183">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-1183">Object</span></span>|<span data-ttu-id="167d3-1184">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1184">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1185">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1185">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="167d3-1186">function</span><span class="sxs-lookup"><span data-stu-id="167d3-1186">function</span></span>|<span data-ttu-id="167d3-1187">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1187">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1188">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1188">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="167d3-1189">成功すると、初期化データが文字列とし`asyncResult.value`てプロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1189">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="167d3-1190">初期化コンテキストがない場合、 `asyncResult`オブジェクトには、 `Error` `code`プロパティがに`9020`設定されたオブジェクトと`name`プロパティがに`GenericResponseError`設定されたオブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1190">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-1191">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1191">Requirements</span></span>

|<span data-ttu-id="167d3-1192">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1192">Requirement</span></span>|<span data-ttu-id="167d3-1193">値</span><span class="sxs-lookup"><span data-stu-id="167d3-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-1194">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-1195">プレビュー</span><span class="sxs-lookup"><span data-stu-id="167d3-1195">Preview</span></span>|
|[<span data-ttu-id="167d3-1196">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1196">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-1197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-1197">ReadItem</span></span>|
|[<span data-ttu-id="167d3-1198">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-1198">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-1199">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-1199">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-1200">例</span><span class="sxs-lookup"><span data-stu-id="167d3-1200">Example</span></span>

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

#### <a name="getregexmatches--object"></a><span data-ttu-id="167d3-1201">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="167d3-1201">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="167d3-1202">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1202">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-1203">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="167d3-1203">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="167d3-p165">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="167d3-p165">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="167d3-1207">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="167d3-1207">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="167d3-1208">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="167d3-1208">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="167d3-p166">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-1212">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1212">Requirements</span></span>

|<span data-ttu-id="167d3-1213">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1213">Requirement</span></span>|<span data-ttu-id="167d3-1214">値</span><span class="sxs-lookup"><span data-stu-id="167d3-1214">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-1215">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-1215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-1216">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-1216">1.0</span></span>|
|[<span data-ttu-id="167d3-1217">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1217">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-1218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-1218">ReadItem</span></span>|
|[<span data-ttu-id="167d3-1219">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-1219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-1220">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-1220">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="167d3-1221">戻り値:</span><span class="sxs-lookup"><span data-stu-id="167d3-1221">Returns:</span></span>

<span data-ttu-id="167d3-p167">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="167d3-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="167d3-1224">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="167d3-1224">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="167d3-1225">Object</span><span class="sxs-lookup"><span data-stu-id="167d3-1225">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="167d3-1226">例</span><span class="sxs-lookup"><span data-stu-id="167d3-1226">Example</span></span>

<span data-ttu-id="167d3-1227">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="167d3-1227">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="167d3-1228">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="167d3-1228">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="167d3-1229">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1229">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-1230">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="167d3-1230">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="167d3-1231">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="167d3-1231">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="167d3-p168">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="167d3-p168">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-1234">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-1234">Parameters</span></span>

|<span data-ttu-id="167d3-1235">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-1235">Name</span></span>|<span data-ttu-id="167d3-1236">型</span><span class="sxs-lookup"><span data-stu-id="167d3-1236">Type</span></span>|<span data-ttu-id="167d3-1237">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-1237">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="167d3-1238">String</span><span class="sxs-lookup"><span data-stu-id="167d3-1238">String</span></span>|<span data-ttu-id="167d3-1239">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="167d3-1239">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-1240">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1240">Requirements</span></span>

|<span data-ttu-id="167d3-1241">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1241">Requirement</span></span>|<span data-ttu-id="167d3-1242">値</span><span class="sxs-lookup"><span data-stu-id="167d3-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-1243">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-1243">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-1244">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-1244">1.0</span></span>|
|[<span data-ttu-id="167d3-1245">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1245">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-1246">ReadItem</span></span>|
|[<span data-ttu-id="167d3-1247">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-1247">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-1248">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-1248">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="167d3-1249">戻り値:</span><span class="sxs-lookup"><span data-stu-id="167d3-1249">Returns:</span></span>

<span data-ttu-id="167d3-1250">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="167d3-1250">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="167d3-1251">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="167d3-1251">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="167d3-1252">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="167d3-1252">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="167d3-1253">例</span><span class="sxs-lookup"><span data-stu-id="167d3-1253">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

---
---

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="167d3-1254">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="167d3-1254">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="167d3-1255">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1255">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="167d3-p169">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-p169">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-1258">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-1258">Parameters</span></span>

|<span data-ttu-id="167d3-1259">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-1259">Name</span></span>|<span data-ttu-id="167d3-1260">型</span><span class="sxs-lookup"><span data-stu-id="167d3-1260">Type</span></span>|<span data-ttu-id="167d3-1261">属性</span><span class="sxs-lookup"><span data-stu-id="167d3-1261">Attributes</span></span>|<span data-ttu-id="167d3-1262">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-1262">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="167d3-1263">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="167d3-1263">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="167d3-p170">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="167d3-p170">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="167d3-1267">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-1267">Object</span></span>|<span data-ttu-id="167d3-1268">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1268">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1269">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="167d3-1269">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="167d3-1270">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-1270">Object</span></span>|<span data-ttu-id="167d3-1271">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1271">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1272">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1272">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="167d3-1273">function</span><span class="sxs-lookup"><span data-stu-id="167d3-1273">function</span></span>||<span data-ttu-id="167d3-1274">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1274">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="167d3-1275">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1275">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="167d3-1276">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="167d3-1276">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-1277">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1277">Requirements</span></span>

|<span data-ttu-id="167d3-1278">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1278">Requirement</span></span>|<span data-ttu-id="167d3-1279">値</span><span class="sxs-lookup"><span data-stu-id="167d3-1279">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-1280">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-1280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-1281">1.2</span><span class="sxs-lookup"><span data-stu-id="167d3-1281">1.2</span></span>|
|[<span data-ttu-id="167d3-1282">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-1283">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="167d3-1283">ReadWriteItem</span></span>|
|[<span data-ttu-id="167d3-1284">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-1284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-1285">作成</span><span class="sxs-lookup"><span data-stu-id="167d3-1285">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="167d3-1286">戻り値:</span><span class="sxs-lookup"><span data-stu-id="167d3-1286">Returns:</span></span>

<span data-ttu-id="167d3-1287">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="167d3-1287">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="167d3-1288">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="167d3-1288">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="167d3-1289">String</span><span class="sxs-lookup"><span data-stu-id="167d3-1289">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="167d3-1290">例</span><span class="sxs-lookup"><span data-stu-id="167d3-1290">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="167d3-1291">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="167d3-1291">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="167d3-1292">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1292">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="167d3-1293">強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1293">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-1294">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="167d3-1294">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-1295">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1295">Requirements</span></span>

|<span data-ttu-id="167d3-1296">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1296">Requirement</span></span>|<span data-ttu-id="167d3-1297">値</span><span class="sxs-lookup"><span data-stu-id="167d3-1297">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-1298">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-1298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-1299">1.6</span><span class="sxs-lookup"><span data-stu-id="167d3-1299">1.6</span></span>|
|[<span data-ttu-id="167d3-1300">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1300">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-1301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-1301">ReadItem</span></span>|
|[<span data-ttu-id="167d3-1302">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-1302">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-1303">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-1303">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="167d3-1304">戻り値:</span><span class="sxs-lookup"><span data-stu-id="167d3-1304">Returns:</span></span>

<span data-ttu-id="167d3-1305">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="167d3-1305">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="167d3-1306">例</span><span class="sxs-lookup"><span data-stu-id="167d3-1306">Example</span></span>

<span data-ttu-id="167d3-1307">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="167d3-1307">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="167d3-1308">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="167d3-1308">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="167d3-p173">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-p173">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-1311">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="167d3-1311">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="167d3-p174">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="167d3-p174">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="167d3-1315">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="167d3-1315">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="167d3-1316">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="167d3-1316">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="167d3-p175">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-p175">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="167d3-1320">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1320">Requirements</span></span>

|<span data-ttu-id="167d3-1321">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1321">Requirement</span></span>|<span data-ttu-id="167d3-1322">値</span><span class="sxs-lookup"><span data-stu-id="167d3-1322">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-1323">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-1323">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-1324">1.6</span><span class="sxs-lookup"><span data-stu-id="167d3-1324">1.6</span></span>|
|[<span data-ttu-id="167d3-1325">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1325">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-1326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-1326">ReadItem</span></span>|
|[<span data-ttu-id="167d3-1327">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-1327">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-1328">読み取り</span><span class="sxs-lookup"><span data-stu-id="167d3-1328">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="167d3-1329">戻り値:</span><span class="sxs-lookup"><span data-stu-id="167d3-1329">Returns:</span></span>

<span data-ttu-id="167d3-p176">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="167d3-p176">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="167d3-1332">例</span><span class="sxs-lookup"><span data-stu-id="167d3-1332">Example</span></span>

<span data-ttu-id="167d3-1333">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="167d3-1333">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="167d3-1334">getsharedpropertiesasync ([options], callback)</span><span class="sxs-lookup"><span data-stu-id="167d3-1334">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="167d3-1335">共有フォルダー、予定表、またはメールボックス内の選択した予定またはメッセージのプロパティを取得します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1335">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-1336">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-1336">Parameters</span></span>

|<span data-ttu-id="167d3-1337">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-1337">Name</span></span>|<span data-ttu-id="167d3-1338">型</span><span class="sxs-lookup"><span data-stu-id="167d3-1338">Type</span></span>|<span data-ttu-id="167d3-1339">属性</span><span class="sxs-lookup"><span data-stu-id="167d3-1339">Attributes</span></span>|<span data-ttu-id="167d3-1340">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-1340">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="167d3-1341">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-1341">Object</span></span>|<span data-ttu-id="167d3-1342">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1342">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1343">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="167d3-1343">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="167d3-1344">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-1344">Object</span></span>|<span data-ttu-id="167d3-1345">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1345">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1346">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1346">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="167d3-1347">function</span><span class="sxs-lookup"><span data-stu-id="167d3-1347">function</span></span>||<span data-ttu-id="167d3-1348">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1348">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="167d3-1349">共有プロパティは、 [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) `asyncResult.value`プロパティのオブジェクトとして提供されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1349">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="167d3-1350">このオブジェクトは、アイテムの共有プロパティを取得するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1350">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-1351">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1351">Requirements</span></span>

|<span data-ttu-id="167d3-1352">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1352">Requirement</span></span>|<span data-ttu-id="167d3-1353">値</span><span class="sxs-lookup"><span data-stu-id="167d3-1353">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-1354">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-1354">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-1355">プレビュー</span><span class="sxs-lookup"><span data-stu-id="167d3-1355">Preview</span></span>|
|[<span data-ttu-id="167d3-1356">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1356">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-1357">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-1357">ReadItem</span></span>|
|[<span data-ttu-id="167d3-1358">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-1358">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-1359">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-1359">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-1360">例</span><span class="sxs-lookup"><span data-stu-id="167d3-1360">Example</span></span>

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

---
---

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="167d3-1361">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="167d3-1361">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="167d3-1362">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1362">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="167d3-p178">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="167d3-p178">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-1366">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-1366">Parameters</span></span>

|<span data-ttu-id="167d3-1367">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-1367">Name</span></span>|<span data-ttu-id="167d3-1368">型</span><span class="sxs-lookup"><span data-stu-id="167d3-1368">Type</span></span>|<span data-ttu-id="167d3-1369">属性</span><span class="sxs-lookup"><span data-stu-id="167d3-1369">Attributes</span></span>|<span data-ttu-id="167d3-1370">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-1370">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="167d3-1371">function</span><span class="sxs-lookup"><span data-stu-id="167d3-1371">function</span></span>||<span data-ttu-id="167d3-1372">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1372">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="167d3-1373">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1373">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="167d3-1374">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1374">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="167d3-1375">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-1375">Object</span></span>|<span data-ttu-id="167d3-1376">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1376">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1377">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1377">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="167d3-1378">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1378">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-1379">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1379">Requirements</span></span>

|<span data-ttu-id="167d3-1380">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1380">Requirement</span></span>|<span data-ttu-id="167d3-1381">値</span><span class="sxs-lookup"><span data-stu-id="167d3-1381">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-1382">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-1382">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-1383">1.0</span><span class="sxs-lookup"><span data-stu-id="167d3-1383">1.0</span></span>|
|[<span data-ttu-id="167d3-1384">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1384">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-1385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-1385">ReadItem</span></span>|
|[<span data-ttu-id="167d3-1386">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-1386">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-1387">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-1387">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-1388">例</span><span class="sxs-lookup"><span data-stu-id="167d3-1388">Example</span></span>

<span data-ttu-id="167d3-p181">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="167d3-p181">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="167d3-1392">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="167d3-1392">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="167d3-1393">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1393">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="167d3-1394">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1394">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="167d3-1395">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="167d3-1395">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="167d3-1396">Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="167d3-1396">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="167d3-1397">ユーザーがアプリを閉じたとき、またはインラインフォームの作成が開始されたときに、別のウィンドウで続行するためにフォームをポップアウトした後、セッションが終了します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1397">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-1398">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-1398">Parameters</span></span>

|<span data-ttu-id="167d3-1399">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-1399">Name</span></span>|<span data-ttu-id="167d3-1400">型</span><span class="sxs-lookup"><span data-stu-id="167d3-1400">Type</span></span>|<span data-ttu-id="167d3-1401">属性</span><span class="sxs-lookup"><span data-stu-id="167d3-1401">Attributes</span></span>|<span data-ttu-id="167d3-1402">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-1402">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="167d3-1403">String</span><span class="sxs-lookup"><span data-stu-id="167d3-1403">String</span></span>||<span data-ttu-id="167d3-1404">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="167d3-1404">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="167d3-1405">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-1405">Object</span></span>|<span data-ttu-id="167d3-1406">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1406">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1407">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="167d3-1407">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="167d3-1408">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-1408">Object</span></span>|<span data-ttu-id="167d3-1409">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1409">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1410">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1410">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="167d3-1411">function</span><span class="sxs-lookup"><span data-stu-id="167d3-1411">function</span></span>|<span data-ttu-id="167d3-1412">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1412">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1413">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1413">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="167d3-1414">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1414">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="167d3-1415">エラー</span><span class="sxs-lookup"><span data-stu-id="167d3-1415">Errors</span></span>

|<span data-ttu-id="167d3-1416">エラー コード</span><span class="sxs-lookup"><span data-stu-id="167d3-1416">Error code</span></span>|<span data-ttu-id="167d3-1417">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-1417">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="167d3-1418">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="167d3-1418">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-1419">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1419">Requirements</span></span>

|<span data-ttu-id="167d3-1420">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1420">Requirement</span></span>|<span data-ttu-id="167d3-1421">値</span><span class="sxs-lookup"><span data-stu-id="167d3-1421">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-1422">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-1422">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-1423">1.1</span><span class="sxs-lookup"><span data-stu-id="167d3-1423">1.1</span></span>|
|[<span data-ttu-id="167d3-1424">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1424">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-1425">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="167d3-1425">ReadWriteItem</span></span>|
|[<span data-ttu-id="167d3-1426">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-1426">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-1427">作成</span><span class="sxs-lookup"><span data-stu-id="167d3-1427">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-1428">例</span><span class="sxs-lookup"><span data-stu-id="167d3-1428">Example</span></span>

<span data-ttu-id="167d3-1429">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1429">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="167d3-1430">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="167d3-1430">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="167d3-1431">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1431">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="167d3-1432">現在、サポートされて`Office.EventType.AttachmentsChanged`いる`Office.EventType.AppointmentTimeChanged`イベント`Office.EventType.EnhancedLocationsChanged`の`Office.EventType.RecipientsChanged`種類は`Office.EventType.RecurrenceChanged`、、、、、です。</span><span class="sxs-lookup"><span data-stu-id="167d3-1432">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-1433">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-1433">Parameters</span></span>

| <span data-ttu-id="167d3-1434">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-1434">Name</span></span> | <span data-ttu-id="167d3-1435">型</span><span class="sxs-lookup"><span data-stu-id="167d3-1435">Type</span></span> | <span data-ttu-id="167d3-1436">属性</span><span class="sxs-lookup"><span data-stu-id="167d3-1436">Attributes</span></span> | <span data-ttu-id="167d3-1437">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-1437">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="167d3-1438">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="167d3-1438">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="167d3-1439">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="167d3-1439">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="167d3-1440">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-1440">Object</span></span> | <span data-ttu-id="167d3-1441">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1441">&lt;optional&gt;</span></span> | <span data-ttu-id="167d3-1442">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="167d3-1442">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="167d3-1443">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-1443">Object</span></span> | <span data-ttu-id="167d3-1444">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1444">&lt;optional&gt;</span></span> | <span data-ttu-id="167d3-1445">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1445">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="167d3-1446">function</span><span class="sxs-lookup"><span data-stu-id="167d3-1446">function</span></span>| <span data-ttu-id="167d3-1447">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1447">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1448">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1448">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-1449">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1449">Requirements</span></span>

|<span data-ttu-id="167d3-1450">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1450">Requirement</span></span>| <span data-ttu-id="167d3-1451">値</span><span class="sxs-lookup"><span data-stu-id="167d3-1451">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-1452">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-1452">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="167d3-1453">1.7</span><span class="sxs-lookup"><span data-stu-id="167d3-1453">1.7</span></span> |
|[<span data-ttu-id="167d3-1454">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1454">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="167d3-1455">ReadItem</span><span class="sxs-lookup"><span data-stu-id="167d3-1455">ReadItem</span></span> |
|[<span data-ttu-id="167d3-1456">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-1456">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="167d3-1457">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="167d3-1457">Compose or Read</span></span> |

---
---

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="167d3-1458">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="167d3-1458">saveAsync([options], callback)</span></span>

<span data-ttu-id="167d3-1459">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1459">Asynchronously saves an item.</span></span>

<span data-ttu-id="167d3-p183">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-p183">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-1463">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="167d3-1463">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="167d3-1464">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1464">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="167d3-p185">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-p185">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="167d3-1468">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="167d3-1468">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="167d3-1469">Mac Outlook では、新規作成モードの会議で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="167d3-1469">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="167d3-1470">Mac Outlook では、会議で `saveAsync` を呼び出すとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1470">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="167d3-1471">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1471">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-1472">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-1472">Parameters</span></span>

|<span data-ttu-id="167d3-1473">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-1473">Name</span></span>|<span data-ttu-id="167d3-1474">型</span><span class="sxs-lookup"><span data-stu-id="167d3-1474">Type</span></span>|<span data-ttu-id="167d3-1475">属性</span><span class="sxs-lookup"><span data-stu-id="167d3-1475">Attributes</span></span>|<span data-ttu-id="167d3-1476">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-1476">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="167d3-1477">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-1477">Object</span></span>|<span data-ttu-id="167d3-1478">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1478">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1479">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="167d3-1479">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="167d3-1480">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-1480">Object</span></span>|<span data-ttu-id="167d3-1481">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1481">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1482">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1482">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="167d3-1483">function</span><span class="sxs-lookup"><span data-stu-id="167d3-1483">function</span></span>||<span data-ttu-id="167d3-1484">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1484">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="167d3-1485">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1485">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-1486">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1486">Requirements</span></span>

|<span data-ttu-id="167d3-1487">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1487">Requirement</span></span>|<span data-ttu-id="167d3-1488">値</span><span class="sxs-lookup"><span data-stu-id="167d3-1488">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-1489">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-1489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-1490">1.3</span><span class="sxs-lookup"><span data-stu-id="167d3-1490">1.3</span></span>|
|[<span data-ttu-id="167d3-1491">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1491">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-1492">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="167d3-1492">ReadWriteItem</span></span>|
|[<span data-ttu-id="167d3-1493">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-1493">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-1494">作成</span><span class="sxs-lookup"><span data-stu-id="167d3-1494">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="167d3-1495">例</span><span class="sxs-lookup"><span data-stu-id="167d3-1495">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="167d3-p187">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="167d3-p187">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="167d3-1498">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="167d3-1498">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="167d3-1499">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="167d3-1499">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="167d3-p188">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="167d3-p188">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="167d3-1503">パラメーター</span><span class="sxs-lookup"><span data-stu-id="167d3-1503">Parameters</span></span>

|<span data-ttu-id="167d3-1504">名前</span><span class="sxs-lookup"><span data-stu-id="167d3-1504">Name</span></span>|<span data-ttu-id="167d3-1505">型</span><span class="sxs-lookup"><span data-stu-id="167d3-1505">Type</span></span>|<span data-ttu-id="167d3-1506">属性</span><span class="sxs-lookup"><span data-stu-id="167d3-1506">Attributes</span></span>|<span data-ttu-id="167d3-1507">説明</span><span class="sxs-lookup"><span data-stu-id="167d3-1507">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="167d3-1508">String</span><span class="sxs-lookup"><span data-stu-id="167d3-1508">String</span></span>||<span data-ttu-id="167d3-p189">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="167d3-p189">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="167d3-1512">Object</span><span class="sxs-lookup"><span data-stu-id="167d3-1512">Object</span></span>|<span data-ttu-id="167d3-1513">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1513">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1514">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="167d3-1514">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="167d3-1515">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="167d3-1515">Object</span></span>|<span data-ttu-id="167d3-1516">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1516">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-1517">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1517">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="167d3-1518">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="167d3-1518">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="167d3-1519">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="167d3-1519">&lt;optional&gt;</span></span>|<span data-ttu-id="167d3-p190">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-p190">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="167d3-p191">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-p191">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="167d3-1524">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1524">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="167d3-1525">function</span><span class="sxs-lookup"><span data-stu-id="167d3-1525">function</span></span>||<span data-ttu-id="167d3-1526">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="167d3-1526">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="167d3-1527">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1527">Requirements</span></span>

|<span data-ttu-id="167d3-1528">要件</span><span class="sxs-lookup"><span data-stu-id="167d3-1528">Requirement</span></span>|<span data-ttu-id="167d3-1529">値</span><span class="sxs-lookup"><span data-stu-id="167d3-1529">Value</span></span>|
|---|---|
|[<span data-ttu-id="167d3-1530">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="167d3-1530">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="167d3-1531">1.2</span><span class="sxs-lookup"><span data-stu-id="167d3-1531">1.2</span></span>|
|[<span data-ttu-id="167d3-1532">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="167d3-1532">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="167d3-1533">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="167d3-1533">ReadWriteItem</span></span>|
|[<span data-ttu-id="167d3-1534">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="167d3-1534">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="167d3-1535">作成</span><span class="sxs-lookup"><span data-stu-id="167d3-1535">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="167d3-1536">例</span><span class="sxs-lookup"><span data-stu-id="167d3-1536">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
