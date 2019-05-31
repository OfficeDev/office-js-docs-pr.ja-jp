---
title: Office. メールボックス-要件セット1.7
description: ''
ms.date: 05/30/2019
localization_priority: Normal
ms.openlocfilehash: 8895a27938bc6c72d929cf9423a377eaa2cc02a9
ms.sourcegitcommit: b299b8a5dfffb6102cb14b431bdde4861abfb47f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/30/2019
ms.locfileid: "34589189"
---
# <a name="item"></a><span data-ttu-id="063a2-102">item</span><span class="sxs-lookup"><span data-stu-id="063a2-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="063a2-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="063a2-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="063a2-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="063a2-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-106">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-106">Requirements</span></span>

|<span data-ttu-id="063a2-107">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-107">Requirement</span></span>|<span data-ttu-id="063a2-108">値</span><span class="sxs-lookup"><span data-stu-id="063a2-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-110">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-110">1.0</span></span>|
|[<span data-ttu-id="063a2-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="063a2-112">Restricted</span></span>|
|[<span data-ttu-id="063a2-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="063a2-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="063a2-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-115">Members and methods</span></span>

| <span data-ttu-id="063a2-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-116">Member</span></span> | <span data-ttu-id="063a2-117">種類</span><span class="sxs-lookup"><span data-stu-id="063a2-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="063a2-118">attachments</span><span class="sxs-lookup"><span data-stu-id="063a2-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="063a2-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-119">Member</span></span> |
| [<span data-ttu-id="063a2-120">bcc</span><span class="sxs-lookup"><span data-stu-id="063a2-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="063a2-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-121">Member</span></span> |
| [<span data-ttu-id="063a2-122">body</span><span class="sxs-lookup"><span data-stu-id="063a2-122">body</span></span>](#body-body) | <span data-ttu-id="063a2-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-123">Member</span></span> |
| [<span data-ttu-id="063a2-124">cc</span><span class="sxs-lookup"><span data-stu-id="063a2-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="063a2-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-125">Member</span></span> |
| [<span data-ttu-id="063a2-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="063a2-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="063a2-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-127">Member</span></span> |
| [<span data-ttu-id="063a2-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="063a2-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="063a2-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-129">Member</span></span> |
| [<span data-ttu-id="063a2-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="063a2-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="063a2-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-131">Member</span></span> |
| [<span data-ttu-id="063a2-132">end</span><span class="sxs-lookup"><span data-stu-id="063a2-132">end</span></span>](#end-datetime) | <span data-ttu-id="063a2-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-133">Member</span></span> |
| [<span data-ttu-id="063a2-134">from</span><span class="sxs-lookup"><span data-stu-id="063a2-134">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="063a2-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-135">Member</span></span> |
| [<span data-ttu-id="063a2-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="063a2-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="063a2-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-137">Member</span></span> |
| [<span data-ttu-id="063a2-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="063a2-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="063a2-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-139">Member</span></span> |
| [<span data-ttu-id="063a2-140">itemId</span><span class="sxs-lookup"><span data-stu-id="063a2-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="063a2-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-141">Member</span></span> |
| [<span data-ttu-id="063a2-142">itemType</span><span class="sxs-lookup"><span data-stu-id="063a2-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="063a2-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-143">Member</span></span> |
| [<span data-ttu-id="063a2-144">location</span><span class="sxs-lookup"><span data-stu-id="063a2-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="063a2-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-145">Member</span></span> |
| [<span data-ttu-id="063a2-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="063a2-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="063a2-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-147">Member</span></span> |
| [<span data-ttu-id="063a2-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="063a2-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="063a2-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-149">Member</span></span> |
| [<span data-ttu-id="063a2-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="063a2-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="063a2-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-151">Member</span></span> |
| [<span data-ttu-id="063a2-152">organizer</span><span class="sxs-lookup"><span data-stu-id="063a2-152">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="063a2-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-153">Member</span></span> |
| [<span data-ttu-id="063a2-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="063a2-154">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="063a2-155">Member</span><span class="sxs-lookup"><span data-stu-id="063a2-155">Member</span></span> |
| [<span data-ttu-id="063a2-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="063a2-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="063a2-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-157">Member</span></span> |
| [<span data-ttu-id="063a2-158">sender</span><span class="sxs-lookup"><span data-stu-id="063a2-158">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="063a2-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-159">Member</span></span> |
| [<span data-ttu-id="063a2-160">系列 Id</span><span class="sxs-lookup"><span data-stu-id="063a2-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="063a2-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-161">Member</span></span> |
| [<span data-ttu-id="063a2-162">start</span><span class="sxs-lookup"><span data-stu-id="063a2-162">start</span></span>](#start-datetime) | <span data-ttu-id="063a2-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-163">Member</span></span> |
| [<span data-ttu-id="063a2-164">subject</span><span class="sxs-lookup"><span data-stu-id="063a2-164">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="063a2-165">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-165">Member</span></span> |
| [<span data-ttu-id="063a2-166">to</span><span class="sxs-lookup"><span data-stu-id="063a2-166">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="063a2-167">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-167">Member</span></span> |
| [<span data-ttu-id="063a2-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="063a2-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="063a2-169">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-169">Method</span></span> |
| [<span data-ttu-id="063a2-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="063a2-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="063a2-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-171">Method</span></span> |
| [<span data-ttu-id="063a2-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="063a2-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="063a2-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-173">Method</span></span> |
| [<span data-ttu-id="063a2-174">close</span><span class="sxs-lookup"><span data-stu-id="063a2-174">close</span></span>](#close) | <span data-ttu-id="063a2-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-175">Method</span></span> |
| [<span data-ttu-id="063a2-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="063a2-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="063a2-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-177">Method</span></span> |
| [<span data-ttu-id="063a2-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="063a2-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="063a2-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-179">Method</span></span> |
| [<span data-ttu-id="063a2-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="063a2-180">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="063a2-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-181">Method</span></span> |
| [<span data-ttu-id="063a2-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="063a2-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="063a2-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-183">Method</span></span> |
| [<span data-ttu-id="063a2-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="063a2-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="063a2-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-185">Method</span></span> |
| [<span data-ttu-id="063a2-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="063a2-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="063a2-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-187">Method</span></span> |
| [<span data-ttu-id="063a2-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="063a2-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="063a2-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-189">Method</span></span> |
| [<span data-ttu-id="063a2-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="063a2-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="063a2-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-191">Method</span></span> |
| [<span data-ttu-id="063a2-192">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="063a2-192">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="063a2-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-193">Method</span></span> |
| [<span data-ttu-id="063a2-194">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="063a2-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="063a2-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-195">Method</span></span> |
| [<span data-ttu-id="063a2-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="063a2-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="063a2-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-197">Method</span></span> |
| [<span data-ttu-id="063a2-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="063a2-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="063a2-199">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-199">Method</span></span> |
| [<span data-ttu-id="063a2-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="063a2-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="063a2-201">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-201">Method</span></span> |
| [<span data-ttu-id="063a2-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="063a2-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="063a2-203">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-203">Method</span></span> |
| [<span data-ttu-id="063a2-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="063a2-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="063a2-205">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="063a2-206">例</span><span class="sxs-lookup"><span data-stu-id="063a2-206">Example</span></span>

<span data-ttu-id="063a2-207">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="063a2-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="063a2-208">メンバー</span><span class="sxs-lookup"><span data-stu-id="063a2-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="063a2-209">添付ファイル: <[Attachmentdetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="063a2-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="063a2-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="063a2-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="063a2-212">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="063a2-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="063a2-213">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="063a2-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="063a2-214">型</span><span class="sxs-lookup"><span data-stu-id="063a2-214">Type</span></span>

*   <span data-ttu-id="063a2-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="063a2-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-216">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-216">Requirements</span></span>

|<span data-ttu-id="063a2-217">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-217">Requirement</span></span>|<span data-ttu-id="063a2-218">値</span><span class="sxs-lookup"><span data-stu-id="063a2-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-219">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-220">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-220">1.0</span></span>|
|[<span data-ttu-id="063a2-221">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-222">ReadItem</span></span>|
|[<span data-ttu-id="063a2-223">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-224">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="063a2-225">例</span><span class="sxs-lookup"><span data-stu-id="063a2-225">Example</span></span>

<span data-ttu-id="063a2-226">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="063a2-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="063a2-227">bcc:[受信者](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="063a2-227">bcc: [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="063a2-228">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="063a2-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="063a2-229">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="063a2-229">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="063a2-230">型</span><span class="sxs-lookup"><span data-stu-id="063a2-230">Type</span></span>

*   [<span data-ttu-id="063a2-231">受信者</span><span class="sxs-lookup"><span data-stu-id="063a2-231">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="063a2-232">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-232">Requirements</span></span>

|<span data-ttu-id="063a2-233">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-233">Requirement</span></span>|<span data-ttu-id="063a2-234">値</span><span class="sxs-lookup"><span data-stu-id="063a2-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-235">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-236">1.1</span><span class="sxs-lookup"><span data-stu-id="063a2-236">1.1</span></span>|
|[<span data-ttu-id="063a2-237">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-237">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-238">ReadItem</span></span>|
|[<span data-ttu-id="063a2-239">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-239">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-240">作成</span><span class="sxs-lookup"><span data-stu-id="063a2-240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="063a2-241">例</span><span class="sxs-lookup"><span data-stu-id="063a2-241">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="063a2-242">本文:[本文](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="063a2-242">body: [Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="063a2-243">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="063a2-243">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="063a2-244">型</span><span class="sxs-lookup"><span data-stu-id="063a2-244">Type</span></span>

*   [<span data-ttu-id="063a2-245">Body</span><span class="sxs-lookup"><span data-stu-id="063a2-245">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="063a2-246">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-246">Requirements</span></span>

|<span data-ttu-id="063a2-247">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-247">Requirement</span></span>|<span data-ttu-id="063a2-248">値</span><span class="sxs-lookup"><span data-stu-id="063a2-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-249">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-250">1.1</span><span class="sxs-lookup"><span data-stu-id="063a2-250">1.1</span></span>|
|[<span data-ttu-id="063a2-251">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-251">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-252">ReadItem</span></span>|
|[<span data-ttu-id="063a2-253">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-253">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-254">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="063a2-254">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="063a2-255">例</span><span class="sxs-lookup"><span data-stu-id="063a2-255">Example</span></span>

<span data-ttu-id="063a2-256">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="063a2-256">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="063a2-257">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="063a2-257">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="063a2-258">cc: <[emailaddressdetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="063a2-258">cc: Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="063a2-259">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="063a2-259">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="063a2-260">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="063a2-260">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="063a2-261">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="063a2-261">Read mode</span></span>

<span data-ttu-id="063a2-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="063a2-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="063a2-264">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="063a2-264">Compose mode</span></span>

<span data-ttu-id="063a2-265">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-265">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="063a2-266">型</span><span class="sxs-lookup"><span data-stu-id="063a2-266">Type</span></span>

*   <span data-ttu-id="063a2-267">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="063a2-267">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-268">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-268">Requirements</span></span>

|<span data-ttu-id="063a2-269">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-269">Requirement</span></span>|<span data-ttu-id="063a2-270">値</span><span class="sxs-lookup"><span data-stu-id="063a2-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-271">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-272">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-272">1.0</span></span>|
|[<span data-ttu-id="063a2-273">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-274">ReadItem</span></span>|
|[<span data-ttu-id="063a2-275">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-276">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="063a2-276">Compose or Read</span></span>|

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="063a2-277">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="063a2-277">(nullable) conversationId: String</span></span>

<span data-ttu-id="063a2-278">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="063a2-278">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="063a2-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="063a2-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="063a2-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="063a2-283">Type</span><span class="sxs-lookup"><span data-stu-id="063a2-283">Type</span></span>

*   <span data-ttu-id="063a2-284">String</span><span class="sxs-lookup"><span data-stu-id="063a2-284">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-285">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-285">Requirements</span></span>

|<span data-ttu-id="063a2-286">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-286">Requirement</span></span>|<span data-ttu-id="063a2-287">値</span><span class="sxs-lookup"><span data-stu-id="063a2-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-288">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-289">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-289">1.0</span></span>|
|[<span data-ttu-id="063a2-290">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-290">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-291">ReadItem</span></span>|
|[<span data-ttu-id="063a2-292">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-293">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="063a2-293">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="063a2-294">例</span><span class="sxs-lookup"><span data-stu-id="063a2-294">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="063a2-295">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="063a2-295">dateTimeCreated: Date</span></span>

<span data-ttu-id="063a2-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="063a2-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="063a2-298">型</span><span class="sxs-lookup"><span data-stu-id="063a2-298">Type</span></span>

*   <span data-ttu-id="063a2-299">日付</span><span class="sxs-lookup"><span data-stu-id="063a2-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-300">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-300">Requirements</span></span>

|<span data-ttu-id="063a2-301">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-301">Requirement</span></span>|<span data-ttu-id="063a2-302">値</span><span class="sxs-lookup"><span data-stu-id="063a2-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-303">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-304">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-304">1.0</span></span>|
|[<span data-ttu-id="063a2-305">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-305">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-306">ReadItem</span></span>|
|[<span data-ttu-id="063a2-307">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-307">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-308">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="063a2-309">例</span><span class="sxs-lookup"><span data-stu-id="063a2-309">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="063a2-310">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="063a2-310">dateTimeModified: Date</span></span>

<span data-ttu-id="063a2-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="063a2-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="063a2-313">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="063a2-313">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="063a2-314">型</span><span class="sxs-lookup"><span data-stu-id="063a2-314">Type</span></span>

*   <span data-ttu-id="063a2-315">日付</span><span class="sxs-lookup"><span data-stu-id="063a2-315">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-316">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-316">Requirements</span></span>

|<span data-ttu-id="063a2-317">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-317">Requirement</span></span>|<span data-ttu-id="063a2-318">値</span><span class="sxs-lookup"><span data-stu-id="063a2-318">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-319">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-320">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-320">1.0</span></span>|
|[<span data-ttu-id="063a2-321">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-321">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-322">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-322">ReadItem</span></span>|
|[<span data-ttu-id="063a2-323">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-323">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-324">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-324">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="063a2-325">例</span><span class="sxs-lookup"><span data-stu-id="063a2-325">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

---
---

#### <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="063a2-326">終了: 日付 |[時間](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="063a2-326">end: Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="063a2-327">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="063a2-327">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="063a2-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="063a2-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="063a2-330">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="063a2-330">Read mode</span></span>

<span data-ttu-id="063a2-331">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-331">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="063a2-332">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="063a2-332">Compose mode</span></span>

<span data-ttu-id="063a2-333">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-333">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="063a2-334">[`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="063a2-334">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="063a2-335">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="063a2-335">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="063a2-336">型</span><span class="sxs-lookup"><span data-stu-id="063a2-336">Type</span></span>

*   <span data-ttu-id="063a2-337">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="063a2-337">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-338">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-338">Requirements</span></span>

|<span data-ttu-id="063a2-339">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-339">Requirement</span></span>|<span data-ttu-id="063a2-340">値</span><span class="sxs-lookup"><span data-stu-id="063a2-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-341">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-342">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-342">1.0</span></span>|
|[<span data-ttu-id="063a2-343">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-344">ReadItem</span></span>|
|[<span data-ttu-id="063a2-345">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-346">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="063a2-346">Compose or Read</span></span>|

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="063a2-347">from: [emailaddressdetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[from](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="063a2-347">from: [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="063a2-348">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="063a2-348">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="063a2-p112">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="063a2-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="063a2-351">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="063a2-351">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="063a2-352">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="063a2-352">Read mode</span></span>

<span data-ttu-id="063a2-353">プロパティ`from`は`EmailAddressDetails`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-353">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="063a2-354">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="063a2-354">Compose mode</span></span>

<span data-ttu-id="063a2-355">プロパティ`from`は、from `From`値を取得するメソッドを提供するオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-355">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="063a2-356">型</span><span class="sxs-lookup"><span data-stu-id="063a2-356">Type</span></span>

*   <span data-ttu-id="063a2-357">[電子メールアドレス](/javascript/api/outlook_1_7/office.emailaddressdetails) | [の](/javascript/api/outlook_1_7/office.from)詳細</span><span class="sxs-lookup"><span data-stu-id="063a2-357">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-358">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-358">Requirements</span></span>

|<span data-ttu-id="063a2-359">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-359">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="063a2-360">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-361">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-361">1.0</span></span>|<span data-ttu-id="063a2-362">1.7</span><span class="sxs-lookup"><span data-stu-id="063a2-362">1.7</span></span>|
|[<span data-ttu-id="063a2-363">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-364">ReadItem</span></span>|<span data-ttu-id="063a2-365">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="063a2-365">ReadWriteItem</span></span>|
|[<span data-ttu-id="063a2-366">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-367">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-367">Read</span></span>|<span data-ttu-id="063a2-368">作成</span><span class="sxs-lookup"><span data-stu-id="063a2-368">Compose</span></span>|

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="063a2-369">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="063a2-369">internetMessageId: String</span></span>

<span data-ttu-id="063a2-p113">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="063a2-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="063a2-372">Type</span><span class="sxs-lookup"><span data-stu-id="063a2-372">Type</span></span>

*   <span data-ttu-id="063a2-373">String</span><span class="sxs-lookup"><span data-stu-id="063a2-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-374">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-374">Requirements</span></span>

|<span data-ttu-id="063a2-375">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-375">Requirement</span></span>|<span data-ttu-id="063a2-376">値</span><span class="sxs-lookup"><span data-stu-id="063a2-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-377">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-378">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-378">1.0</span></span>|
|[<span data-ttu-id="063a2-379">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-380">ReadItem</span></span>|
|[<span data-ttu-id="063a2-381">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-382">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="063a2-383">例</span><span class="sxs-lookup"><span data-stu-id="063a2-383">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="063a2-384">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="063a2-384">itemClass: String</span></span>

<span data-ttu-id="063a2-p114">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="063a2-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="063a2-p115">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="063a2-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="063a2-389">型</span><span class="sxs-lookup"><span data-stu-id="063a2-389">Type</span></span>|<span data-ttu-id="063a2-390">説明</span><span class="sxs-lookup"><span data-stu-id="063a2-390">Description</span></span>|<span data-ttu-id="063a2-391">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="063a2-391">item class</span></span>|
|---|---|---|
|<span data-ttu-id="063a2-392">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="063a2-392">Appointment items</span></span>|<span data-ttu-id="063a2-393">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="063a2-393">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="063a2-394">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="063a2-394">Message items</span></span>|<span data-ttu-id="063a2-395">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="063a2-395">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="063a2-396">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="063a2-396">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="063a2-397">Type</span><span class="sxs-lookup"><span data-stu-id="063a2-397">Type</span></span>

*   <span data-ttu-id="063a2-398">String</span><span class="sxs-lookup"><span data-stu-id="063a2-398">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-399">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-399">Requirements</span></span>

|<span data-ttu-id="063a2-400">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-400">Requirement</span></span>|<span data-ttu-id="063a2-401">値</span><span class="sxs-lookup"><span data-stu-id="063a2-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-402">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-403">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-403">1.0</span></span>|
|[<span data-ttu-id="063a2-404">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-404">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-405">ReadItem</span></span>|
|[<span data-ttu-id="063a2-406">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-406">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-407">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-407">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="063a2-408">例</span><span class="sxs-lookup"><span data-stu-id="063a2-408">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="063a2-409">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="063a2-409">(nullable) itemId: String</span></span>

<span data-ttu-id="063a2-p116">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="063a2-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="063a2-412">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="063a2-412">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="063a2-413">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="063a2-413">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="063a2-414">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="063a2-414">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="063a2-415">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="063a2-415">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="063a2-p118">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="063a2-418">Type</span><span class="sxs-lookup"><span data-stu-id="063a2-418">Type</span></span>

*   <span data-ttu-id="063a2-419">String</span><span class="sxs-lookup"><span data-stu-id="063a2-419">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-420">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-420">Requirements</span></span>

|<span data-ttu-id="063a2-421">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-421">Requirement</span></span>|<span data-ttu-id="063a2-422">値</span><span class="sxs-lookup"><span data-stu-id="063a2-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-423">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-424">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-424">1.0</span></span>|
|[<span data-ttu-id="063a2-425">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-425">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-426">ReadItem</span></span>|
|[<span data-ttu-id="063a2-427">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-427">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-428">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="063a2-429">例</span><span class="sxs-lookup"><span data-stu-id="063a2-429">Example</span></span>

<span data-ttu-id="063a2-p119">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="063a2-432">itemType: [MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="063a2-432">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="063a2-433">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="063a2-433">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="063a2-434">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="063a2-434">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="063a2-435">型</span><span class="sxs-lookup"><span data-stu-id="063a2-435">Type</span></span>

*   [<span data-ttu-id="063a2-436">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="063a2-436">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="063a2-437">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-437">Requirements</span></span>

|<span data-ttu-id="063a2-438">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-438">Requirement</span></span>|<span data-ttu-id="063a2-439">値</span><span class="sxs-lookup"><span data-stu-id="063a2-439">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-440">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-441">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-441">1.0</span></span>|
|[<span data-ttu-id="063a2-442">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-442">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-443">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-443">ReadItem</span></span>|
|[<span data-ttu-id="063a2-444">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-444">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-445">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="063a2-445">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="063a2-446">例</span><span class="sxs-lookup"><span data-stu-id="063a2-446">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

---
---

#### <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="063a2-447">場所: String |[場所](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="063a2-447">location: String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="063a2-448">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="063a2-448">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="063a2-449">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="063a2-449">Read mode</span></span>

<span data-ttu-id="063a2-450">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-450">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="063a2-451">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="063a2-451">Compose mode</span></span>

<span data-ttu-id="063a2-452">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-452">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="063a2-453">型</span><span class="sxs-lookup"><span data-stu-id="063a2-453">Type</span></span>

*   <span data-ttu-id="063a2-454">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="063a2-454">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-455">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-455">Requirements</span></span>

|<span data-ttu-id="063a2-456">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-456">Requirement</span></span>|<span data-ttu-id="063a2-457">値</span><span class="sxs-lookup"><span data-stu-id="063a2-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-458">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-459">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-459">1.0</span></span>|
|[<span data-ttu-id="063a2-460">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-461">ReadItem</span></span>|
|[<span data-ttu-id="063a2-462">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-463">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="063a2-463">Compose or Read</span></span>|

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="063a2-464">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="063a2-464">normalizedSubject: String</span></span>

<span data-ttu-id="063a2-p120">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="063a2-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="063a2-p121">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="063a2-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="063a2-469">型</span><span class="sxs-lookup"><span data-stu-id="063a2-469">Type</span></span>

*   <span data-ttu-id="063a2-470">String</span><span class="sxs-lookup"><span data-stu-id="063a2-470">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-471">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-471">Requirements</span></span>

|<span data-ttu-id="063a2-472">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-472">Requirement</span></span>|<span data-ttu-id="063a2-473">値</span><span class="sxs-lookup"><span data-stu-id="063a2-473">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-474">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-474">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-475">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-475">1.0</span></span>|
|[<span data-ttu-id="063a2-476">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-476">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-477">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-477">ReadItem</span></span>|
|[<span data-ttu-id="063a2-478">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-478">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-479">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-479">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="063a2-480">例</span><span class="sxs-lookup"><span data-stu-id="063a2-480">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="063a2-481">notificationMessages: [Notificationmessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="063a2-481">notificationMessages: [NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="063a2-482">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="063a2-482">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="063a2-483">型</span><span class="sxs-lookup"><span data-stu-id="063a2-483">Type</span></span>

*   [<span data-ttu-id="063a2-484">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="063a2-484">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="063a2-485">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-485">Requirements</span></span>

|<span data-ttu-id="063a2-486">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-486">Requirement</span></span>|<span data-ttu-id="063a2-487">値</span><span class="sxs-lookup"><span data-stu-id="063a2-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-488">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-489">1.3</span><span class="sxs-lookup"><span data-stu-id="063a2-489">1.3</span></span>|
|[<span data-ttu-id="063a2-490">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-491">ReadItem</span></span>|
|[<span data-ttu-id="063a2-492">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-493">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="063a2-493">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="063a2-494">例</span><span class="sxs-lookup"><span data-stu-id="063a2-494">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="063a2-495">任意出席者: <[emailaddressdetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="063a2-495">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="063a2-496">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="063a2-496">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="063a2-497">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="063a2-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="063a2-498">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="063a2-498">Read mode</span></span>

<span data-ttu-id="063a2-499">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-499">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="063a2-500">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="063a2-500">Compose mode</span></span>

<span data-ttu-id="063a2-501">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-501">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="063a2-502">型</span><span class="sxs-lookup"><span data-stu-id="063a2-502">Type</span></span>

*   <span data-ttu-id="063a2-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="063a2-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-504">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-504">Requirements</span></span>

|<span data-ttu-id="063a2-505">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-505">Requirement</span></span>|<span data-ttu-id="063a2-506">値</span><span class="sxs-lookup"><span data-stu-id="063a2-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-507">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-508">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-508">1.0</span></span>|
|[<span data-ttu-id="063a2-509">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-510">ReadItem</span></span>|
|[<span data-ttu-id="063a2-511">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-512">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="063a2-512">Compose or Read</span></span>|

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="063a2-513">開催者: [emailaddressdetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[開催者](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="063a2-513">organizer: [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="063a2-514">指定した会議の開催者の電子メールアドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="063a2-514">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="063a2-515">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="063a2-515">Read mode</span></span>

<span data-ttu-id="063a2-516">プロパティ`organizer`は、会議の開催者を表す[emailaddressdetails](/javascript/api/outlook_1_7/office.emailaddressdetails)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-516">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="063a2-517">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="063a2-517">Compose mode</span></span>

<span data-ttu-id="063a2-518">プロパティ`organizer`は、開催者の値を取得するためのメソッドを提供する[オーガナイザー](/javascript/api/outlook_1_7/office.organizer)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-518">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="063a2-519">型</span><span class="sxs-lookup"><span data-stu-id="063a2-519">Type</span></span>

*   <span data-ttu-id="063a2-520">[Emailaddressdetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [開催者](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="063a2-520">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-521">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-521">Requirements</span></span>

|<span data-ttu-id="063a2-522">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-522">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="063a2-523">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-524">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-524">1.0</span></span>|<span data-ttu-id="063a2-525">1.7</span><span class="sxs-lookup"><span data-stu-id="063a2-525">1.7</span></span>|
|[<span data-ttu-id="063a2-526">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-527">ReadItem</span></span>|<span data-ttu-id="063a2-528">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="063a2-528">ReadWriteItem</span></span>|
|[<span data-ttu-id="063a2-529">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-529">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-530">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-530">Read</span></span>|<span data-ttu-id="063a2-531">作成</span><span class="sxs-lookup"><span data-stu-id="063a2-531">Compose</span></span>|

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="063a2-532">(nullable) 定期的なスケジュール:[定期的](/javascript/api/outlook_1_7/office.recurrence)なアイテム</span><span class="sxs-lookup"><span data-stu-id="063a2-532">(nullable) recurrence: [Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="063a2-533">予定の定期的なパターンを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="063a2-533">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="063a2-534">会議出席依頼の定期的なパターンを取得します。</span><span class="sxs-lookup"><span data-stu-id="063a2-534">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="063a2-535">予定アイテムの読み取りおよび作成モード。</span><span class="sxs-lookup"><span data-stu-id="063a2-535">Read and compose modes for appointment items.</span></span> <span data-ttu-id="063a2-536">会議出席依頼アイテムの閲覧モード。</span><span class="sxs-lookup"><span data-stu-id="063a2-536">Read mode for meeting request items.</span></span>

<span data-ttu-id="063a2-537">この`recurrence`プロパティは、アイテムが series または series 内のインスタンスの場合、定期的な予定または会議出席依頼に対して[定期的](/javascript/api/outlook_1_7/office.recurrence)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-537">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="063a2-538">`null`は、単一の予定および1つの予定の会議出席依頼に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-538">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="063a2-539">`undefined`は、会議出席依頼ではないメッセージに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-539">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="063a2-540">注: 会議出席依頼に`itemClass`は、IPM という値があります。出席依頼。</span><span class="sxs-lookup"><span data-stu-id="063a2-540">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="063a2-541">注: 定期的なオブジェクトが`null`の場合は、そのオブジェクトが単一の予定または1つの予定の会議出席依頼であり、データ系列の一部ではないことを示します。</span><span class="sxs-lookup"><span data-stu-id="063a2-541">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="063a2-542">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="063a2-542">Read mode</span></span>

<span data-ttu-id="063a2-543">この`recurrence`プロパティは、定期的な予定を表す[定期的](/javascript/api/outlook_1_7/office.recurrence)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-543">The `recurrence` property returns a [Recurrence](/javascript/api/outlook_1_7/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="063a2-544">これは、予定および会議出席依頼に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="063a2-544">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="063a2-545">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="063a2-545">Compose mode</span></span>

<span data-ttu-id="063a2-546">この`recurrence`プロパティは、予定の繰り返しを管理するためのメソッドを提供する[定期的](/javascript/api/outlook_1_7/office.recurrence)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-546">The `recurrence` property returns a [Recurrence](/javascript/api/outlook_1_7/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="063a2-547">これは予定に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="063a2-547">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="063a2-548">型</span><span class="sxs-lookup"><span data-stu-id="063a2-548">Type</span></span>

* [<span data-ttu-id="063a2-549">繰り返さ</span><span class="sxs-lookup"><span data-stu-id="063a2-549">Recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="063a2-550">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-550">Requirement</span></span>|<span data-ttu-id="063a2-551">値</span><span class="sxs-lookup"><span data-stu-id="063a2-551">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-552">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-552">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-553">1.7</span><span class="sxs-lookup"><span data-stu-id="063a2-553">1.7</span></span>|
|[<span data-ttu-id="063a2-554">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-554">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-555">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-555">ReadItem</span></span>|
|[<span data-ttu-id="063a2-556">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-556">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-557">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="063a2-557">Compose or Read</span></span>|

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="063a2-558">requiredat< dees: 配列.[emailaddressdetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="063a2-558">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="063a2-559">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="063a2-559">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="063a2-560">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="063a2-560">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="063a2-561">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="063a2-561">Read mode</span></span>

<span data-ttu-id="063a2-562">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-562">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="063a2-563">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="063a2-563">Compose mode</span></span>

<span data-ttu-id="063a2-564">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-564">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="063a2-565">型</span><span class="sxs-lookup"><span data-stu-id="063a2-565">Type</span></span>

*   <span data-ttu-id="063a2-566">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="063a2-566">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-567">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-567">Requirements</span></span>

|<span data-ttu-id="063a2-568">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-568">Requirement</span></span>|<span data-ttu-id="063a2-569">値</span><span class="sxs-lookup"><span data-stu-id="063a2-569">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-570">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-571">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-571">1.0</span></span>|
|[<span data-ttu-id="063a2-572">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-572">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-573">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-573">ReadItem</span></span>|
|[<span data-ttu-id="063a2-574">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-574">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-575">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="063a2-575">Compose or Read</span></span>|

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="063a2-576">sender: [Emailaddressdetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="063a2-576">sender: [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="063a2-p128">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="063a2-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="063a2-p129">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsfrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="063a2-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="063a2-581">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="063a2-581">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="063a2-582">型</span><span class="sxs-lookup"><span data-stu-id="063a2-582">Type</span></span>

*   [<span data-ttu-id="063a2-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="063a2-583">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="063a2-584">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-584">Requirements</span></span>

|<span data-ttu-id="063a2-585">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-585">Requirement</span></span>|<span data-ttu-id="063a2-586">値</span><span class="sxs-lookup"><span data-stu-id="063a2-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-587">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-588">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-588">1.0</span></span>|
|[<span data-ttu-id="063a2-589">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-589">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-590">ReadItem</span></span>|
|[<span data-ttu-id="063a2-591">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-591">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-592">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="063a2-593">例</span><span class="sxs-lookup"><span data-stu-id="063a2-593">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="063a2-594">(nullable) 系列 Id: String</span><span class="sxs-lookup"><span data-stu-id="063a2-594">(nullable) seriesId: String</span></span>

<span data-ttu-id="063a2-595">インスタンスが属する系列の id を取得します。</span><span class="sxs-lookup"><span data-stu-id="063a2-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="063a2-596">OWA および Outlook で、は`seriesId` 、このアイテムが属する親 (シリーズ) アイテムの Exchange Web サービス (EWS) ID を返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-596">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="063a2-597">ただし、iOS と Android では、 `seriesId`は親アイテムの REST ID を返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="063a2-598">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="063a2-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="063a2-599">`seriesId`プロパティが OUTLOOK REST API で使用される outlook id と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="063a2-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="063a2-600">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="063a2-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="063a2-601">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="063a2-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="063a2-602">この`seriesId`プロパティは`null` 、単一の予定、系列のアイテム、会議出席依頼などの親アイテムを持たないアイテムに`undefined`対して、会議出席依頼以外のアイテムに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="063a2-603">Type</span><span class="sxs-lookup"><span data-stu-id="063a2-603">Type</span></span>

* <span data-ttu-id="063a2-604">String</span><span class="sxs-lookup"><span data-stu-id="063a2-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-605">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-605">Requirements</span></span>

|<span data-ttu-id="063a2-606">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-606">Requirement</span></span>|<span data-ttu-id="063a2-607">値</span><span class="sxs-lookup"><span data-stu-id="063a2-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-608">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-609">1.7</span><span class="sxs-lookup"><span data-stu-id="063a2-609">1.7</span></span>|
|[<span data-ttu-id="063a2-610">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-610">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-611">ReadItem</span></span>|
|[<span data-ttu-id="063a2-612">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-613">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="063a2-613">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="063a2-614">例</span><span class="sxs-lookup"><span data-stu-id="063a2-614">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="063a2-615">開始: 日付 |[時間](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="063a2-615">start: Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="063a2-616">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="063a2-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="063a2-p132">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="063a2-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="063a2-619">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="063a2-619">Read mode</span></span>

<span data-ttu-id="063a2-620">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-620">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="063a2-621">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="063a2-621">Compose mode</span></span>

<span data-ttu-id="063a2-622">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="063a2-623">[`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="063a2-623">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="063a2-624">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="063a2-624">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="063a2-625">型</span><span class="sxs-lookup"><span data-stu-id="063a2-625">Type</span></span>

*   <span data-ttu-id="063a2-626">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="063a2-626">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-627">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-627">Requirements</span></span>

|<span data-ttu-id="063a2-628">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-628">Requirement</span></span>|<span data-ttu-id="063a2-629">値</span><span class="sxs-lookup"><span data-stu-id="063a2-629">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-630">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-630">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-631">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-631">1.0</span></span>|
|[<span data-ttu-id="063a2-632">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-632">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-633">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-633">ReadItem</span></span>|
|[<span data-ttu-id="063a2-634">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-634">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-635">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="063a2-635">Compose or Read</span></span>|

---
---

#### <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="063a2-636">subject: String |[件名](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="063a2-636">subject: String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="063a2-637">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="063a2-637">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="063a2-638">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="063a2-638">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="063a2-639">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="063a2-639">Read mode</span></span>

<span data-ttu-id="063a2-p133">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="063a2-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="063a2-642">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="063a2-642">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="063a2-643">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="063a2-643">Compose mode</span></span>

<span data-ttu-id="063a2-644">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="063a2-645">型</span><span class="sxs-lookup"><span data-stu-id="063a2-645">Type</span></span>

*   <span data-ttu-id="063a2-646">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="063a2-646">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-647">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-647">Requirements</span></span>

|<span data-ttu-id="063a2-648">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-648">Requirement</span></span>|<span data-ttu-id="063a2-649">値</span><span class="sxs-lookup"><span data-stu-id="063a2-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-650">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-651">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-651">1.0</span></span>|
|[<span data-ttu-id="063a2-652">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-653">ReadItem</span></span>|
|[<span data-ttu-id="063a2-654">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-655">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="063a2-655">Compose or Read</span></span>|

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="063a2-656">宛先: <[emailaddressdetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="063a2-656">to: Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="063a2-657">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="063a2-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="063a2-658">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="063a2-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="063a2-659">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="063a2-659">Read mode</span></span>

<span data-ttu-id="063a2-p135">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="063a2-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="063a2-662">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="063a2-662">Compose mode</span></span>

<span data-ttu-id="063a2-663">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-663">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="063a2-664">型</span><span class="sxs-lookup"><span data-stu-id="063a2-664">Type</span></span>

*   <span data-ttu-id="063a2-665">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="063a2-665">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-666">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-666">Requirements</span></span>

|<span data-ttu-id="063a2-667">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-667">Requirement</span></span>|<span data-ttu-id="063a2-668">値</span><span class="sxs-lookup"><span data-stu-id="063a2-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-669">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-670">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-670">1.0</span></span>|
|[<span data-ttu-id="063a2-671">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-671">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-672">ReadItem</span></span>|
|[<span data-ttu-id="063a2-673">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-673">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-674">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="063a2-674">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="063a2-675">メソッド</span><span class="sxs-lookup"><span data-stu-id="063a2-675">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="063a2-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="063a2-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="063a2-677">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="063a2-677">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="063a2-678">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="063a2-678">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="063a2-679">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="063a2-679">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="063a2-680">パラメーター</span><span class="sxs-lookup"><span data-stu-id="063a2-680">Parameters</span></span>
|<span data-ttu-id="063a2-681">名前</span><span class="sxs-lookup"><span data-stu-id="063a2-681">Name</span></span>|<span data-ttu-id="063a2-682">種類</span><span class="sxs-lookup"><span data-stu-id="063a2-682">Type</span></span>|<span data-ttu-id="063a2-683">属性</span><span class="sxs-lookup"><span data-stu-id="063a2-683">Attributes</span></span>|<span data-ttu-id="063a2-684">説明</span><span class="sxs-lookup"><span data-stu-id="063a2-684">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="063a2-685">String</span><span class="sxs-lookup"><span data-stu-id="063a2-685">String</span></span>||<span data-ttu-id="063a2-p136">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="063a2-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="063a2-688">String</span><span class="sxs-lookup"><span data-stu-id="063a2-688">String</span></span>||<span data-ttu-id="063a2-p137">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="063a2-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="063a2-691">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="063a2-691">Object</span></span>|<span data-ttu-id="063a2-692">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-692">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-693">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="063a2-693">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="063a2-694">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="063a2-694">Object</span></span>|<span data-ttu-id="063a2-695">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-695">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-696">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="063a2-696">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="063a2-697">Boolean</span><span class="sxs-lookup"><span data-stu-id="063a2-697">Boolean</span></span>|<span data-ttu-id="063a2-698">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-698">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-699">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="063a2-699">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="063a2-700">function</span><span class="sxs-lookup"><span data-stu-id="063a2-700">function</span></span>|<span data-ttu-id="063a2-701">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-701">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-702">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-702">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="063a2-703">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-703">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="063a2-704">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="063a2-704">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="063a2-705">エラー</span><span class="sxs-lookup"><span data-stu-id="063a2-705">Errors</span></span>

|<span data-ttu-id="063a2-706">エラー コード</span><span class="sxs-lookup"><span data-stu-id="063a2-706">Error code</span></span>|<span data-ttu-id="063a2-707">説明</span><span class="sxs-lookup"><span data-stu-id="063a2-707">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="063a2-708">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="063a2-708">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="063a2-709">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="063a2-709">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="063a2-710">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="063a2-710">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="063a2-711">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-711">Requirements</span></span>

|<span data-ttu-id="063a2-712">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-712">Requirement</span></span>|<span data-ttu-id="063a2-713">値</span><span class="sxs-lookup"><span data-stu-id="063a2-713">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-714">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-714">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-715">1.1</span><span class="sxs-lookup"><span data-stu-id="063a2-715">1.1</span></span>|
|[<span data-ttu-id="063a2-716">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-716">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-717">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="063a2-717">ReadWriteItem</span></span>|
|[<span data-ttu-id="063a2-718">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-718">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-719">作成</span><span class="sxs-lookup"><span data-stu-id="063a2-719">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="063a2-720">例</span><span class="sxs-lookup"><span data-stu-id="063a2-720">Examples</span></span>

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

<span data-ttu-id="063a2-721">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="063a2-721">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="063a2-722">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="063a2-722">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="063a2-723">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="063a2-723">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="063a2-724">現在、サポートされて`Office.EventType.AppointmentTimeChanged`いる`Office.EventType.RecipientsChanged`イベントの種類は、、、です。`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="063a2-724">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="063a2-725">パラメーター</span><span class="sxs-lookup"><span data-stu-id="063a2-725">Parameters</span></span>

| <span data-ttu-id="063a2-726">名前</span><span class="sxs-lookup"><span data-stu-id="063a2-726">Name</span></span> | <span data-ttu-id="063a2-727">型</span><span class="sxs-lookup"><span data-stu-id="063a2-727">Type</span></span> | <span data-ttu-id="063a2-728">属性</span><span class="sxs-lookup"><span data-stu-id="063a2-728">Attributes</span></span> | <span data-ttu-id="063a2-729">説明</span><span class="sxs-lookup"><span data-stu-id="063a2-729">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="063a2-730">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="063a2-730">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="063a2-731">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="063a2-731">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="063a2-732">Function</span><span class="sxs-lookup"><span data-stu-id="063a2-732">Function</span></span> || <span data-ttu-id="063a2-p138">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="063a2-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="063a2-736">Object</span><span class="sxs-lookup"><span data-stu-id="063a2-736">Object</span></span> | <span data-ttu-id="063a2-737">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-737">&lt;optional&gt;</span></span> | <span data-ttu-id="063a2-738">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="063a2-738">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="063a2-739">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="063a2-739">Object</span></span> | <span data-ttu-id="063a2-740">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-740">&lt;optional&gt;</span></span> | <span data-ttu-id="063a2-741">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="063a2-741">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="063a2-742">function</span><span class="sxs-lookup"><span data-stu-id="063a2-742">function</span></span>| <span data-ttu-id="063a2-743">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-743">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-744">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-744">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="063a2-745">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-745">Requirements</span></span>

|<span data-ttu-id="063a2-746">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-746">Requirement</span></span>| <span data-ttu-id="063a2-747">値</span><span class="sxs-lookup"><span data-stu-id="063a2-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-748">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="063a2-749">1.7</span><span class="sxs-lookup"><span data-stu-id="063a2-749">1.7</span></span> |
|[<span data-ttu-id="063a2-750">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-750">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="063a2-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-751">ReadItem</span></span> |
|[<span data-ttu-id="063a2-752">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-752">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="063a2-753">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="063a2-753">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="063a2-754">例</span><span class="sxs-lookup"><span data-stu-id="063a2-754">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="063a2-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="063a2-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="063a2-756">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="063a2-756">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="063a2-p139">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="063a2-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="063a2-760">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="063a2-760">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="063a2-761">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="063a2-761">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="063a2-762">パラメーター</span><span class="sxs-lookup"><span data-stu-id="063a2-762">Parameters</span></span>

|<span data-ttu-id="063a2-763">名前</span><span class="sxs-lookup"><span data-stu-id="063a2-763">Name</span></span>|<span data-ttu-id="063a2-764">型</span><span class="sxs-lookup"><span data-stu-id="063a2-764">Type</span></span>|<span data-ttu-id="063a2-765">属性</span><span class="sxs-lookup"><span data-stu-id="063a2-765">Attributes</span></span>|<span data-ttu-id="063a2-766">説明</span><span class="sxs-lookup"><span data-stu-id="063a2-766">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="063a2-767">String</span><span class="sxs-lookup"><span data-stu-id="063a2-767">String</span></span>||<span data-ttu-id="063a2-p140">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="063a2-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="063a2-770">String</span><span class="sxs-lookup"><span data-stu-id="063a2-770">String</span></span>||<span data-ttu-id="063a2-771">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="063a2-771">The subject of the item to be attached.</span></span> <span data-ttu-id="063a2-772">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="063a2-772">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="063a2-773">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="063a2-773">Object</span></span>|<span data-ttu-id="063a2-774">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-774">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-775">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="063a2-775">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="063a2-776">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="063a2-776">Object</span></span>|<span data-ttu-id="063a2-777">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-777">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-778">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="063a2-778">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="063a2-779">function</span><span class="sxs-lookup"><span data-stu-id="063a2-779">function</span></span>|<span data-ttu-id="063a2-780">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-780">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-781">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-781">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="063a2-782">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-782">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="063a2-783">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="063a2-783">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="063a2-784">エラー</span><span class="sxs-lookup"><span data-stu-id="063a2-784">Errors</span></span>

|<span data-ttu-id="063a2-785">エラー コード</span><span class="sxs-lookup"><span data-stu-id="063a2-785">Error code</span></span>|<span data-ttu-id="063a2-786">説明</span><span class="sxs-lookup"><span data-stu-id="063a2-786">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="063a2-787">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="063a2-787">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="063a2-788">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-788">Requirements</span></span>

|<span data-ttu-id="063a2-789">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-789">Requirement</span></span>|<span data-ttu-id="063a2-790">値</span><span class="sxs-lookup"><span data-stu-id="063a2-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-791">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-792">1.1</span><span class="sxs-lookup"><span data-stu-id="063a2-792">1.1</span></span>|
|[<span data-ttu-id="063a2-793">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-793">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-794">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="063a2-794">ReadWriteItem</span></span>|
|[<span data-ttu-id="063a2-795">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-795">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-796">作成</span><span class="sxs-lookup"><span data-stu-id="063a2-796">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="063a2-797">例</span><span class="sxs-lookup"><span data-stu-id="063a2-797">Example</span></span>

<span data-ttu-id="063a2-798">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-798">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="063a2-799">close()</span><span class="sxs-lookup"><span data-stu-id="063a2-799">close()</span></span>

<span data-ttu-id="063a2-800">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="063a2-800">Closes the current item that is being composed.</span></span>

<span data-ttu-id="063a2-p142">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="063a2-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="063a2-803">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-803">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="063a2-804">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="063a2-804">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-805">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-805">Requirements</span></span>

|<span data-ttu-id="063a2-806">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-806">Requirement</span></span>|<span data-ttu-id="063a2-807">値</span><span class="sxs-lookup"><span data-stu-id="063a2-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-808">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-809">1.3</span><span class="sxs-lookup"><span data-stu-id="063a2-809">1.3</span></span>|
|[<span data-ttu-id="063a2-810">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-811">制限あり</span><span class="sxs-lookup"><span data-stu-id="063a2-811">Restricted</span></span>|
|[<span data-ttu-id="063a2-812">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-813">新規作成</span><span class="sxs-lookup"><span data-stu-id="063a2-813">Compose</span></span>|

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="063a2-814">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="063a2-814">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="063a2-815">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-815">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="063a2-816">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="063a2-816">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="063a2-817">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-817">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="063a2-818">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="063a2-818">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="063a2-p143">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="063a2-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="063a2-822">パラメーター</span><span class="sxs-lookup"><span data-stu-id="063a2-822">Parameters</span></span>

|<span data-ttu-id="063a2-823">名前</span><span class="sxs-lookup"><span data-stu-id="063a2-823">Name</span></span>|<span data-ttu-id="063a2-824">型</span><span class="sxs-lookup"><span data-stu-id="063a2-824">Type</span></span>|<span data-ttu-id="063a2-825">属性</span><span class="sxs-lookup"><span data-stu-id="063a2-825">Attributes</span></span>|<span data-ttu-id="063a2-826">説明</span><span class="sxs-lookup"><span data-stu-id="063a2-826">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="063a2-827">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="063a2-827">String &#124; Object</span></span>||<span data-ttu-id="063a2-p144">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="063a2-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="063a2-830">**または**</span><span class="sxs-lookup"><span data-stu-id="063a2-830">**OR**</span></span><br/><span data-ttu-id="063a2-p145">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="063a2-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="063a2-833">String</span><span class="sxs-lookup"><span data-stu-id="063a2-833">String</span></span>|<span data-ttu-id="063a2-834">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-834">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="063a2-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="063a2-837">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-837">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="063a2-838">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-838">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-839">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="063a2-839">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="063a2-840">String</span><span class="sxs-lookup"><span data-stu-id="063a2-840">String</span></span>||<span data-ttu-id="063a2-p147">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="063a2-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="063a2-843">String</span><span class="sxs-lookup"><span data-stu-id="063a2-843">String</span></span>||<span data-ttu-id="063a2-844">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="063a2-844">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="063a2-845">文字列</span><span class="sxs-lookup"><span data-stu-id="063a2-845">String</span></span>||<span data-ttu-id="063a2-p148">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="063a2-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="063a2-848">ブール値</span><span class="sxs-lookup"><span data-stu-id="063a2-848">Boolean</span></span>||<span data-ttu-id="063a2-p149">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="063a2-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="063a2-851">String</span><span class="sxs-lookup"><span data-stu-id="063a2-851">String</span></span>||<span data-ttu-id="063a2-p150">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="063a2-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="063a2-855">function</span><span class="sxs-lookup"><span data-stu-id="063a2-855">function</span></span>|<span data-ttu-id="063a2-856">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-856">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-857">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-857">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="063a2-858">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-858">Requirements</span></span>

|<span data-ttu-id="063a2-859">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-859">Requirement</span></span>|<span data-ttu-id="063a2-860">値</span><span class="sxs-lookup"><span data-stu-id="063a2-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-861">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-862">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-862">1.0</span></span>|
|[<span data-ttu-id="063a2-863">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-863">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-864">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-864">ReadItem</span></span>|
|[<span data-ttu-id="063a2-865">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-865">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-866">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-866">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="063a2-867">例</span><span class="sxs-lookup"><span data-stu-id="063a2-867">Examples</span></span>

<span data-ttu-id="063a2-868">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="063a2-868">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="063a2-869">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="063a2-869">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="063a2-870">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="063a2-870">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="063a2-871">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="063a2-871">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="063a2-872">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="063a2-872">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="063a2-873">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="063a2-873">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="063a2-874">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="063a2-874">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="063a2-875">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-875">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="063a2-876">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="063a2-876">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="063a2-877">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-877">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="063a2-878">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="063a2-878">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="063a2-p151">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="063a2-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="063a2-882">パラメーター</span><span class="sxs-lookup"><span data-stu-id="063a2-882">Parameters</span></span>

|<span data-ttu-id="063a2-883">名前</span><span class="sxs-lookup"><span data-stu-id="063a2-883">Name</span></span>|<span data-ttu-id="063a2-884">型</span><span class="sxs-lookup"><span data-stu-id="063a2-884">Type</span></span>|<span data-ttu-id="063a2-885">属性</span><span class="sxs-lookup"><span data-stu-id="063a2-885">Attributes</span></span>|<span data-ttu-id="063a2-886">説明</span><span class="sxs-lookup"><span data-stu-id="063a2-886">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="063a2-887">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="063a2-887">String &#124; Object</span></span>||<span data-ttu-id="063a2-p152">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="063a2-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="063a2-890">**または**</span><span class="sxs-lookup"><span data-stu-id="063a2-890">**OR**</span></span><br/><span data-ttu-id="063a2-p153">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="063a2-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="063a2-893">String</span><span class="sxs-lookup"><span data-stu-id="063a2-893">String</span></span>|<span data-ttu-id="063a2-894">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-894">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-p154">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="063a2-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="063a2-897">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-897">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="063a2-898">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-898">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-899">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="063a2-899">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="063a2-900">String</span><span class="sxs-lookup"><span data-stu-id="063a2-900">String</span></span>||<span data-ttu-id="063a2-p155">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="063a2-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="063a2-903">String</span><span class="sxs-lookup"><span data-stu-id="063a2-903">String</span></span>||<span data-ttu-id="063a2-904">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="063a2-904">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="063a2-905">文字列</span><span class="sxs-lookup"><span data-stu-id="063a2-905">String</span></span>||<span data-ttu-id="063a2-p156">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="063a2-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="063a2-908">ブール値</span><span class="sxs-lookup"><span data-stu-id="063a2-908">Boolean</span></span>||<span data-ttu-id="063a2-p157">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="063a2-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="063a2-911">String</span><span class="sxs-lookup"><span data-stu-id="063a2-911">String</span></span>||<span data-ttu-id="063a2-p158">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="063a2-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="063a2-915">function</span><span class="sxs-lookup"><span data-stu-id="063a2-915">function</span></span>|<span data-ttu-id="063a2-916">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-916">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-917">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-917">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="063a2-918">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-918">Requirements</span></span>

|<span data-ttu-id="063a2-919">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-919">Requirement</span></span>|<span data-ttu-id="063a2-920">値</span><span class="sxs-lookup"><span data-stu-id="063a2-920">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-921">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-921">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-922">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-922">1.0</span></span>|
|[<span data-ttu-id="063a2-923">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-923">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-924">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-924">ReadItem</span></span>|
|[<span data-ttu-id="063a2-925">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-925">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-926">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-926">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="063a2-927">例</span><span class="sxs-lookup"><span data-stu-id="063a2-927">Examples</span></span>

<span data-ttu-id="063a2-928">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="063a2-928">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="063a2-929">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="063a2-929">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="063a2-930">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="063a2-930">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="063a2-931">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="063a2-931">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="063a2-932">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="063a2-932">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="063a2-933">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="063a2-933">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="063a2-934">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="063a2-934">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="063a2-935">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="063a2-935">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="063a2-936">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="063a2-936">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-937">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-937">Requirements</span></span>

|<span data-ttu-id="063a2-938">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-938">Requirement</span></span>|<span data-ttu-id="063a2-939">値</span><span class="sxs-lookup"><span data-stu-id="063a2-939">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-940">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-940">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-941">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-941">1.0</span></span>|
|[<span data-ttu-id="063a2-942">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-942">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-943">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-943">ReadItem</span></span>|
|[<span data-ttu-id="063a2-944">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-944">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-945">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-945">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="063a2-946">戻り値:</span><span class="sxs-lookup"><span data-stu-id="063a2-946">Returns:</span></span>

<span data-ttu-id="063a2-947">型:[Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="063a2-947">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="063a2-948">例</span><span class="sxs-lookup"><span data-stu-id="063a2-948">Example</span></span>

<span data-ttu-id="063a2-949">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="063a2-949">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="063a2-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="063a2-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="063a2-951">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="063a2-951">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="063a2-952">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="063a2-952">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="063a2-953">パラメーター</span><span class="sxs-lookup"><span data-stu-id="063a2-953">Parameters</span></span>

|<span data-ttu-id="063a2-954">名前</span><span class="sxs-lookup"><span data-stu-id="063a2-954">Name</span></span>|<span data-ttu-id="063a2-955">型</span><span class="sxs-lookup"><span data-stu-id="063a2-955">Type</span></span>|<span data-ttu-id="063a2-956">説明</span><span class="sxs-lookup"><span data-stu-id="063a2-956">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="063a2-957">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="063a2-957">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="063a2-958">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="063a2-958">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="063a2-959">Requirements</span><span class="sxs-lookup"><span data-stu-id="063a2-959">Requirements</span></span>

|<span data-ttu-id="063a2-960">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-960">Requirement</span></span>|<span data-ttu-id="063a2-961">値</span><span class="sxs-lookup"><span data-stu-id="063a2-961">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-962">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-962">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-963">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-963">1.0</span></span>|
|[<span data-ttu-id="063a2-964">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-964">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-965">制限あり</span><span class="sxs-lookup"><span data-stu-id="063a2-965">Restricted</span></span>|
|[<span data-ttu-id="063a2-966">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-966">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-967">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-967">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="063a2-968">戻り値:</span><span class="sxs-lookup"><span data-stu-id="063a2-968">Returns:</span></span>

<span data-ttu-id="063a2-969">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-969">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="063a2-970">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-970">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="063a2-971">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="063a2-971">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="063a2-972">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="063a2-972">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="063a2-973">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="063a2-973">Value of `entityType`</span></span>|<span data-ttu-id="063a2-974">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="063a2-974">Type of objects in returned array</span></span>|<span data-ttu-id="063a2-975">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="063a2-975">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="063a2-976">文字列</span><span class="sxs-lookup"><span data-stu-id="063a2-976">String</span></span>|<span data-ttu-id="063a2-977">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="063a2-977">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="063a2-978">連絡先</span><span class="sxs-lookup"><span data-stu-id="063a2-978">Contact</span></span>|<span data-ttu-id="063a2-979">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="063a2-979">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="063a2-980">文字列</span><span class="sxs-lookup"><span data-stu-id="063a2-980">String</span></span>|<span data-ttu-id="063a2-981">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="063a2-981">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="063a2-982">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="063a2-982">MeetingSuggestion</span></span>|<span data-ttu-id="063a2-983">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="063a2-983">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="063a2-984">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="063a2-984">PhoneNumber</span></span>|<span data-ttu-id="063a2-985">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="063a2-985">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="063a2-986">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="063a2-986">TaskSuggestion</span></span>|<span data-ttu-id="063a2-987">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="063a2-987">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="063a2-988">文字列</span><span class="sxs-lookup"><span data-stu-id="063a2-988">String</span></span>|<span data-ttu-id="063a2-989">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="063a2-989">**Restricted**</span></span>|

<span data-ttu-id="063a2-990">型:Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="063a2-990">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="063a2-991">例</span><span class="sxs-lookup"><span data-stu-id="063a2-991">Example</span></span>

<span data-ttu-id="063a2-992">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="063a2-992">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="063a2-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="063a2-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="063a2-994">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-994">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="063a2-995">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="063a2-995">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="063a2-996">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-996">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="063a2-997">パラメーター</span><span class="sxs-lookup"><span data-stu-id="063a2-997">Parameters</span></span>

|<span data-ttu-id="063a2-998">名前</span><span class="sxs-lookup"><span data-stu-id="063a2-998">Name</span></span>|<span data-ttu-id="063a2-999">型</span><span class="sxs-lookup"><span data-stu-id="063a2-999">Type</span></span>|<span data-ttu-id="063a2-1000">説明</span><span class="sxs-lookup"><span data-stu-id="063a2-1000">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="063a2-1001">String</span><span class="sxs-lookup"><span data-stu-id="063a2-1001">String</span></span>|<span data-ttu-id="063a2-1002">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="063a2-1002">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="063a2-1003">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1003">Requirements</span></span>

|<span data-ttu-id="063a2-1004">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1004">Requirement</span></span>|<span data-ttu-id="063a2-1005">値</span><span class="sxs-lookup"><span data-stu-id="063a2-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-1006">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-1006">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-1007">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-1007">1.0</span></span>|
|[<span data-ttu-id="063a2-1008">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-1008">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-1009">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-1009">ReadItem</span></span>|
|[<span data-ttu-id="063a2-1010">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-1010">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-1011">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-1011">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="063a2-1012">戻り値:</span><span class="sxs-lookup"><span data-stu-id="063a2-1012">Returns:</span></span>

<span data-ttu-id="063a2-p160">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="063a2-1015">型:Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="063a2-1015">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="063a2-1016">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="063a2-1016">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="063a2-1017">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-1017">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="063a2-1018">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="063a2-1018">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="063a2-p161">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="063a2-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="063a2-1022">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="063a2-1022">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="063a2-1023">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="063a2-1023">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="063a2-p162">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="063a2-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-1027">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1027">Requirements</span></span>

|<span data-ttu-id="063a2-1028">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1028">Requirement</span></span>|<span data-ttu-id="063a2-1029">値</span><span class="sxs-lookup"><span data-stu-id="063a2-1029">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-1030">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-1030">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-1031">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-1031">1.0</span></span>|
|[<span data-ttu-id="063a2-1032">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-1032">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-1033">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-1033">ReadItem</span></span>|
|[<span data-ttu-id="063a2-1034">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-1034">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-1035">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-1035">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="063a2-1036">戻り値:</span><span class="sxs-lookup"><span data-stu-id="063a2-1036">Returns:</span></span>

<span data-ttu-id="063a2-p163">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="063a2-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="063a2-1039">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="063a2-1039">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="063a2-1040">Object</span><span class="sxs-lookup"><span data-stu-id="063a2-1040">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="063a2-1041">例</span><span class="sxs-lookup"><span data-stu-id="063a2-1041">Example</span></span>

<span data-ttu-id="063a2-1042">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="063a2-1042">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="063a2-1043">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="063a2-1043">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="063a2-1044">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-1044">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="063a2-1045">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="063a2-1045">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="063a2-1046">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="063a2-1046">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="063a2-p164">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="063a2-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="063a2-1049">パラメーター</span><span class="sxs-lookup"><span data-stu-id="063a2-1049">Parameters</span></span>

|<span data-ttu-id="063a2-1050">名前</span><span class="sxs-lookup"><span data-stu-id="063a2-1050">Name</span></span>|<span data-ttu-id="063a2-1051">型</span><span class="sxs-lookup"><span data-stu-id="063a2-1051">Type</span></span>|<span data-ttu-id="063a2-1052">説明</span><span class="sxs-lookup"><span data-stu-id="063a2-1052">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="063a2-1053">String</span><span class="sxs-lookup"><span data-stu-id="063a2-1053">String</span></span>|<span data-ttu-id="063a2-1054">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="063a2-1054">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="063a2-1055">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1055">Requirements</span></span>

|<span data-ttu-id="063a2-1056">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1056">Requirement</span></span>|<span data-ttu-id="063a2-1057">値</span><span class="sxs-lookup"><span data-stu-id="063a2-1057">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-1058">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-1058">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-1059">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-1059">1.0</span></span>|
|[<span data-ttu-id="063a2-1060">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-1060">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-1061">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-1061">ReadItem</span></span>|
|[<span data-ttu-id="063a2-1062">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-1062">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-1063">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-1063">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="063a2-1064">戻り値:</span><span class="sxs-lookup"><span data-stu-id="063a2-1064">Returns:</span></span>

<span data-ttu-id="063a2-1065">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="063a2-1065">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="063a2-1066">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="063a2-1066">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="063a2-1067">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="063a2-1067">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="063a2-1068">例</span><span class="sxs-lookup"><span data-stu-id="063a2-1068">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="063a2-1069">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="063a2-1069">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="063a2-1070">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-1070">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="063a2-p165">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="063a2-1073">パラメーター</span><span class="sxs-lookup"><span data-stu-id="063a2-1073">Parameters</span></span>

|<span data-ttu-id="063a2-1074">名前</span><span class="sxs-lookup"><span data-stu-id="063a2-1074">Name</span></span>|<span data-ttu-id="063a2-1075">型</span><span class="sxs-lookup"><span data-stu-id="063a2-1075">Type</span></span>|<span data-ttu-id="063a2-1076">属性</span><span class="sxs-lookup"><span data-stu-id="063a2-1076">Attributes</span></span>|<span data-ttu-id="063a2-1077">説明</span><span class="sxs-lookup"><span data-stu-id="063a2-1077">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="063a2-1078">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="063a2-1078">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="063a2-p166">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="063a2-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="063a2-1082">Object</span><span class="sxs-lookup"><span data-stu-id="063a2-1082">Object</span></span>|<span data-ttu-id="063a2-1083">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-1083">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-1084">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="063a2-1084">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="063a2-1085">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="063a2-1085">Object</span></span>|<span data-ttu-id="063a2-1086">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-1086">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-1087">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1087">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="063a2-1088">function</span><span class="sxs-lookup"><span data-stu-id="063a2-1088">function</span></span>||<span data-ttu-id="063a2-1089">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1089">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="063a2-1090">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="063a2-1090">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="063a2-1091">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="063a2-1091">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="063a2-1092">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1092">Requirements</span></span>

|<span data-ttu-id="063a2-1093">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1093">Requirement</span></span>|<span data-ttu-id="063a2-1094">値</span><span class="sxs-lookup"><span data-stu-id="063a2-1094">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-1095">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-1095">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-1096">1.2</span><span class="sxs-lookup"><span data-stu-id="063a2-1096">1.2</span></span>|
|[<span data-ttu-id="063a2-1097">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-1097">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-1098">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="063a2-1098">ReadWriteItem</span></span>|
|[<span data-ttu-id="063a2-1099">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-1099">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-1100">作成</span><span class="sxs-lookup"><span data-stu-id="063a2-1100">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="063a2-1101">戻り値:</span><span class="sxs-lookup"><span data-stu-id="063a2-1101">Returns:</span></span>

<span data-ttu-id="063a2-1102">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="063a2-1102">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="063a2-1103">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="063a2-1103">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="063a2-1104">String</span><span class="sxs-lookup"><span data-stu-id="063a2-1104">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="063a2-1105">例</span><span class="sxs-lookup"><span data-stu-id="063a2-1105">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="063a2-1106">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="063a2-1106">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="063a2-1107">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="063a2-1107">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="063a2-1108">強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1108">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="063a2-1109">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="063a2-1109">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-1110">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1110">Requirements</span></span>

|<span data-ttu-id="063a2-1111">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1111">Requirement</span></span>|<span data-ttu-id="063a2-1112">値</span><span class="sxs-lookup"><span data-stu-id="063a2-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-1113">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-1114">1.6</span><span class="sxs-lookup"><span data-stu-id="063a2-1114">1.6</span></span>|
|[<span data-ttu-id="063a2-1115">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-1115">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-1116">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-1116">ReadItem</span></span>|
|[<span data-ttu-id="063a2-1117">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-1117">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-1118">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-1118">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="063a2-1119">戻り値:</span><span class="sxs-lookup"><span data-stu-id="063a2-1119">Returns:</span></span>

<span data-ttu-id="063a2-1120">型:[Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="063a2-1120">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="063a2-1121">例</span><span class="sxs-lookup"><span data-stu-id="063a2-1121">Example</span></span>

<span data-ttu-id="063a2-1122">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="063a2-1122">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="063a2-1123">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="063a2-1123">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="063a2-p169">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="063a2-1126">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="063a2-1126">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="063a2-p170">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="063a2-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="063a2-1130">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="063a2-1130">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="063a2-1131">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="063a2-1131">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="063a2-p171">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="063a2-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="063a2-1135">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1135">Requirements</span></span>

|<span data-ttu-id="063a2-1136">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1136">Requirement</span></span>|<span data-ttu-id="063a2-1137">値</span><span class="sxs-lookup"><span data-stu-id="063a2-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-1138">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-1139">1.6</span><span class="sxs-lookup"><span data-stu-id="063a2-1139">1.6</span></span>|
|[<span data-ttu-id="063a2-1140">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-1141">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-1141">ReadItem</span></span>|
|[<span data-ttu-id="063a2-1142">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-1143">読み取り</span><span class="sxs-lookup"><span data-stu-id="063a2-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="063a2-1144">戻り値:</span><span class="sxs-lookup"><span data-stu-id="063a2-1144">Returns:</span></span>

<span data-ttu-id="063a2-p172">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="063a2-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="063a2-1147">例</span><span class="sxs-lookup"><span data-stu-id="063a2-1147">Example</span></span>

<span data-ttu-id="063a2-1148">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="063a2-1148">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="063a2-1149">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="063a2-1149">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="063a2-1150">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1150">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="063a2-p173">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="063a2-p173">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="063a2-1154">パラメーター</span><span class="sxs-lookup"><span data-stu-id="063a2-1154">Parameters</span></span>

|<span data-ttu-id="063a2-1155">名前</span><span class="sxs-lookup"><span data-stu-id="063a2-1155">Name</span></span>|<span data-ttu-id="063a2-1156">型</span><span class="sxs-lookup"><span data-stu-id="063a2-1156">Type</span></span>|<span data-ttu-id="063a2-1157">属性</span><span class="sxs-lookup"><span data-stu-id="063a2-1157">Attributes</span></span>|<span data-ttu-id="063a2-1158">説明</span><span class="sxs-lookup"><span data-stu-id="063a2-1158">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="063a2-1159">function</span><span class="sxs-lookup"><span data-stu-id="063a2-1159">function</span></span>||<span data-ttu-id="063a2-1160">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1160">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="063a2-1161">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1161">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="063a2-1162">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1162">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="063a2-1163">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="063a2-1163">Object</span></span>|<span data-ttu-id="063a2-1164">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-1164">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-1165">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1165">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="063a2-1166">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1166">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="063a2-1167">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1167">Requirements</span></span>

|<span data-ttu-id="063a2-1168">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1168">Requirement</span></span>|<span data-ttu-id="063a2-1169">値</span><span class="sxs-lookup"><span data-stu-id="063a2-1169">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-1170">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-1170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-1171">1.0</span><span class="sxs-lookup"><span data-stu-id="063a2-1171">1.0</span></span>|
|[<span data-ttu-id="063a2-1172">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-1172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-1173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-1173">ReadItem</span></span>|
|[<span data-ttu-id="063a2-1174">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-1174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-1175">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="063a2-1175">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="063a2-1176">例</span><span class="sxs-lookup"><span data-stu-id="063a2-1176">Example</span></span>

<span data-ttu-id="063a2-p176">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="063a2-p176">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="063a2-1180">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="063a2-1180">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="063a2-1181">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="063a2-1181">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="063a2-p177">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="063a2-p177">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="063a2-1186">パラメーター</span><span class="sxs-lookup"><span data-stu-id="063a2-1186">Parameters</span></span>

|<span data-ttu-id="063a2-1187">名前</span><span class="sxs-lookup"><span data-stu-id="063a2-1187">Name</span></span>|<span data-ttu-id="063a2-1188">型</span><span class="sxs-lookup"><span data-stu-id="063a2-1188">Type</span></span>|<span data-ttu-id="063a2-1189">属性</span><span class="sxs-lookup"><span data-stu-id="063a2-1189">Attributes</span></span>|<span data-ttu-id="063a2-1190">説明</span><span class="sxs-lookup"><span data-stu-id="063a2-1190">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="063a2-1191">String</span><span class="sxs-lookup"><span data-stu-id="063a2-1191">String</span></span>||<span data-ttu-id="063a2-1192">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="063a2-1192">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="063a2-1193">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="063a2-1193">Object</span></span>|<span data-ttu-id="063a2-1194">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-1194">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-1195">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="063a2-1195">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="063a2-1196">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="063a2-1196">Object</span></span>|<span data-ttu-id="063a2-1197">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-1197">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-1198">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1198">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="063a2-1199">function</span><span class="sxs-lookup"><span data-stu-id="063a2-1199">function</span></span>|<span data-ttu-id="063a2-1200">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-1200">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-1201">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1201">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="063a2-1202">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1202">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="063a2-1203">エラー</span><span class="sxs-lookup"><span data-stu-id="063a2-1203">Errors</span></span>

|<span data-ttu-id="063a2-1204">エラー コード</span><span class="sxs-lookup"><span data-stu-id="063a2-1204">Error code</span></span>|<span data-ttu-id="063a2-1205">説明</span><span class="sxs-lookup"><span data-stu-id="063a2-1205">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="063a2-1206">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="063a2-1206">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="063a2-1207">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1207">Requirements</span></span>

|<span data-ttu-id="063a2-1208">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1208">Requirement</span></span>|<span data-ttu-id="063a2-1209">値</span><span class="sxs-lookup"><span data-stu-id="063a2-1209">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-1210">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-1210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-1211">1.1</span><span class="sxs-lookup"><span data-stu-id="063a2-1211">1.1</span></span>|
|[<span data-ttu-id="063a2-1212">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-1212">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-1213">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="063a2-1213">ReadWriteItem</span></span>|
|[<span data-ttu-id="063a2-1214">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-1214">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-1215">作成</span><span class="sxs-lookup"><span data-stu-id="063a2-1215">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="063a2-1216">例</span><span class="sxs-lookup"><span data-stu-id="063a2-1216">Example</span></span>

<span data-ttu-id="063a2-1217">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="063a2-1217">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="063a2-1218">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="063a2-1218">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="063a2-1219">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="063a2-1219">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="063a2-1220">現在、サポートされて`Office.EventType.AppointmentTimeChanged`いる`Office.EventType.RecipientsChanged`イベントの種類は、、、です。`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="063a2-1220">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="063a2-1221">パラメーター</span><span class="sxs-lookup"><span data-stu-id="063a2-1221">Parameters</span></span>

| <span data-ttu-id="063a2-1222">名前</span><span class="sxs-lookup"><span data-stu-id="063a2-1222">Name</span></span> | <span data-ttu-id="063a2-1223">型</span><span class="sxs-lookup"><span data-stu-id="063a2-1223">Type</span></span> | <span data-ttu-id="063a2-1224">属性</span><span class="sxs-lookup"><span data-stu-id="063a2-1224">Attributes</span></span> | <span data-ttu-id="063a2-1225">説明</span><span class="sxs-lookup"><span data-stu-id="063a2-1225">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="063a2-1226">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="063a2-1226">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="063a2-1227">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="063a2-1227">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="063a2-1228">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="063a2-1228">Object</span></span> | <span data-ttu-id="063a2-1229">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="063a2-1230">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="063a2-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="063a2-1231">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="063a2-1231">Object</span></span> | <span data-ttu-id="063a2-1232">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="063a2-1233">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="063a2-1234">関数</span><span class="sxs-lookup"><span data-stu-id="063a2-1234">function</span></span>| <span data-ttu-id="063a2-1235">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-1236">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="063a2-1237">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1237">Requirements</span></span>

|<span data-ttu-id="063a2-1238">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1238">Requirement</span></span>| <span data-ttu-id="063a2-1239">値</span><span class="sxs-lookup"><span data-stu-id="063a2-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-1240">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-1240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="063a2-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="063a2-1241">1.7</span></span> |
|[<span data-ttu-id="063a2-1242">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-1242">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="063a2-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="063a2-1243">ReadItem</span></span> |
|[<span data-ttu-id="063a2-1244">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-1244">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="063a2-1245">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="063a2-1245">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="063a2-1246">例</span><span class="sxs-lookup"><span data-stu-id="063a2-1246">Example</span></span>

```javascript
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrenceChanged, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};
```

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="063a2-1247">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="063a2-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="063a2-1248">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="063a2-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="063a2-p178">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-p178">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="063a2-1252">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="063a2-1252">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="063a2-1253">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="063a2-p180">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="063a2-1257">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="063a2-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="063a2-1258">Outlook for Mac は、新規`saveAsync`作成モードの会議ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="063a2-1258">Outlook for Mac does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="063a2-1259">そのため、この`saveAsync`シナリオで呼び出しを行うと、エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1259">As such, calling `saveAsync` in that scenario returns an error.</span></span> <span data-ttu-id="063a2-1260">回避策については[、「OFFICE JS API を使用して Outlook For Mac で会議を下書きとして保存できません](https://support.microsoft.com/help/4505745)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="063a2-1260">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="063a2-1261">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1261">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="063a2-1262">パラメーター</span><span class="sxs-lookup"><span data-stu-id="063a2-1262">Parameters</span></span>

|<span data-ttu-id="063a2-1263">名前</span><span class="sxs-lookup"><span data-stu-id="063a2-1263">Name</span></span>|<span data-ttu-id="063a2-1264">型</span><span class="sxs-lookup"><span data-stu-id="063a2-1264">Type</span></span>|<span data-ttu-id="063a2-1265">属性</span><span class="sxs-lookup"><span data-stu-id="063a2-1265">Attributes</span></span>|<span data-ttu-id="063a2-1266">説明</span><span class="sxs-lookup"><span data-stu-id="063a2-1266">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="063a2-1267">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="063a2-1267">Object</span></span>|<span data-ttu-id="063a2-1268">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-1268">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-1269">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="063a2-1269">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="063a2-1270">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="063a2-1270">Object</span></span>|<span data-ttu-id="063a2-1271">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-1271">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-1272">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1272">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="063a2-1273">関数</span><span class="sxs-lookup"><span data-stu-id="063a2-1273">function</span></span>||<span data-ttu-id="063a2-1274">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1274">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="063a2-1275">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1275">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="063a2-1276">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1276">Requirements</span></span>

|<span data-ttu-id="063a2-1277">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1277">Requirement</span></span>|<span data-ttu-id="063a2-1278">値</span><span class="sxs-lookup"><span data-stu-id="063a2-1278">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-1279">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-1279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-1280">1.3</span><span class="sxs-lookup"><span data-stu-id="063a2-1280">1.3</span></span>|
|[<span data-ttu-id="063a2-1281">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-1281">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-1282">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="063a2-1282">ReadWriteItem</span></span>|
|[<span data-ttu-id="063a2-1283">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-1283">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-1284">作成</span><span class="sxs-lookup"><span data-stu-id="063a2-1284">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="063a2-1285">例</span><span class="sxs-lookup"><span data-stu-id="063a2-1285">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="063a2-p182">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="063a2-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="063a2-1288">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="063a2-1288">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="063a2-1289">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="063a2-1289">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="063a2-p183">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="063a2-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="063a2-1293">パラメーター</span><span class="sxs-lookup"><span data-stu-id="063a2-1293">Parameters</span></span>

|<span data-ttu-id="063a2-1294">名前</span><span class="sxs-lookup"><span data-stu-id="063a2-1294">Name</span></span>|<span data-ttu-id="063a2-1295">型</span><span class="sxs-lookup"><span data-stu-id="063a2-1295">Type</span></span>|<span data-ttu-id="063a2-1296">属性</span><span class="sxs-lookup"><span data-stu-id="063a2-1296">Attributes</span></span>|<span data-ttu-id="063a2-1297">説明</span><span class="sxs-lookup"><span data-stu-id="063a2-1297">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="063a2-1298">String</span><span class="sxs-lookup"><span data-stu-id="063a2-1298">String</span></span>||<span data-ttu-id="063a2-p184">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="063a2-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="063a2-1302">Object</span><span class="sxs-lookup"><span data-stu-id="063a2-1302">Object</span></span>|<span data-ttu-id="063a2-1303">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-1303">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-1304">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="063a2-1304">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="063a2-1305">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="063a2-1305">Object</span></span>|<span data-ttu-id="063a2-1306">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-1306">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-1307">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1307">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="063a2-1308">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="063a2-1308">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="063a2-1309">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="063a2-1309">&lt;optional&gt;</span></span>|<span data-ttu-id="063a2-p185">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-p185">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="063a2-p186">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-p186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="063a2-1314">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="063a2-1314">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="063a2-1315">function</span><span class="sxs-lookup"><span data-stu-id="063a2-1315">function</span></span>||<span data-ttu-id="063a2-1316">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="063a2-1316">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="063a2-1317">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1317">Requirements</span></span>

|<span data-ttu-id="063a2-1318">要件</span><span class="sxs-lookup"><span data-stu-id="063a2-1318">Requirement</span></span>|<span data-ttu-id="063a2-1319">値</span><span class="sxs-lookup"><span data-stu-id="063a2-1319">Value</span></span>|
|---|---|
|[<span data-ttu-id="063a2-1320">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="063a2-1320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="063a2-1321">1.2</span><span class="sxs-lookup"><span data-stu-id="063a2-1321">1.2</span></span>|
|[<span data-ttu-id="063a2-1322">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="063a2-1322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="063a2-1323">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="063a2-1323">ReadWriteItem</span></span>|
|[<span data-ttu-id="063a2-1324">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="063a2-1324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="063a2-1325">作成</span><span class="sxs-lookup"><span data-stu-id="063a2-1325">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="063a2-1326">例</span><span class="sxs-lookup"><span data-stu-id="063a2-1326">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
