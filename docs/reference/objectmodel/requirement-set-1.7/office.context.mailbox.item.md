---
title: Office. メールボックス-要件セット1.7
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: d400765293449899eb2e26f3d87128bc88b70000
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629680"
---
# <a name="item"></a><span data-ttu-id="ce5c8-102">item</span><span class="sxs-lookup"><span data-stu-id="ce5c8-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="ce5c8-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="ce5c8-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="ce5c8-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-106">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-106">Requirements</span></span>

|<span data-ttu-id="ce5c8-107">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-107">Requirement</span></span>|<span data-ttu-id="ce5c8-108">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-110">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-110">1.0</span></span>|
|[<span data-ttu-id="ce5c8-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="ce5c8-112">Restricted</span></span>|
|[<span data-ttu-id="ce5c8-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ce5c8-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ce5c8-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-115">Members and methods</span></span>

| <span data-ttu-id="ce5c8-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-116">Member</span></span> | <span data-ttu-id="ce5c8-117">種類</span><span class="sxs-lookup"><span data-stu-id="ce5c8-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ce5c8-118">attachments</span><span class="sxs-lookup"><span data-stu-id="ce5c8-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="ce5c8-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-119">Member</span></span> |
| [<span data-ttu-id="ce5c8-120">bcc</span><span class="sxs-lookup"><span data-stu-id="ce5c8-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="ce5c8-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-121">Member</span></span> |
| [<span data-ttu-id="ce5c8-122">body</span><span class="sxs-lookup"><span data-stu-id="ce5c8-122">body</span></span>](#body-body) | <span data-ttu-id="ce5c8-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-123">Member</span></span> |
| [<span data-ttu-id="ce5c8-124">cc</span><span class="sxs-lookup"><span data-stu-id="ce5c8-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ce5c8-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-125">Member</span></span> |
| [<span data-ttu-id="ce5c8-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="ce5c8-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="ce5c8-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-127">Member</span></span> |
| [<span data-ttu-id="ce5c8-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="ce5c8-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="ce5c8-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-129">Member</span></span> |
| [<span data-ttu-id="ce5c8-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="ce5c8-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="ce5c8-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-131">Member</span></span> |
| [<span data-ttu-id="ce5c8-132">end</span><span class="sxs-lookup"><span data-stu-id="ce5c8-132">end</span></span>](#end-datetime) | <span data-ttu-id="ce5c8-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-133">Member</span></span> |
| [<span data-ttu-id="ce5c8-134">from</span><span class="sxs-lookup"><span data-stu-id="ce5c8-134">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="ce5c8-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-135">Member</span></span> |
| [<span data-ttu-id="ce5c8-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="ce5c8-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="ce5c8-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-137">Member</span></span> |
| [<span data-ttu-id="ce5c8-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="ce5c8-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="ce5c8-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-139">Member</span></span> |
| [<span data-ttu-id="ce5c8-140">itemId</span><span class="sxs-lookup"><span data-stu-id="ce5c8-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="ce5c8-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-141">Member</span></span> |
| [<span data-ttu-id="ce5c8-142">itemType</span><span class="sxs-lookup"><span data-stu-id="ce5c8-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="ce5c8-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-143">Member</span></span> |
| [<span data-ttu-id="ce5c8-144">location</span><span class="sxs-lookup"><span data-stu-id="ce5c8-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="ce5c8-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-145">Member</span></span> |
| [<span data-ttu-id="ce5c8-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="ce5c8-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="ce5c8-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-147">Member</span></span> |
| [<span data-ttu-id="ce5c8-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="ce5c8-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="ce5c8-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-149">Member</span></span> |
| [<span data-ttu-id="ce5c8-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="ce5c8-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ce5c8-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-151">Member</span></span> |
| [<span data-ttu-id="ce5c8-152">organizer</span><span class="sxs-lookup"><span data-stu-id="ce5c8-152">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="ce5c8-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-153">Member</span></span> |
| [<span data-ttu-id="ce5c8-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="ce5c8-154">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="ce5c8-155">Member</span><span class="sxs-lookup"><span data-stu-id="ce5c8-155">Member</span></span> |
| [<span data-ttu-id="ce5c8-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="ce5c8-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ce5c8-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-157">Member</span></span> |
| [<span data-ttu-id="ce5c8-158">sender</span><span class="sxs-lookup"><span data-stu-id="ce5c8-158">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="ce5c8-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-159">Member</span></span> |
| [<span data-ttu-id="ce5c8-160">系列 Id</span><span class="sxs-lookup"><span data-stu-id="ce5c8-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="ce5c8-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-161">Member</span></span> |
| [<span data-ttu-id="ce5c8-162">start</span><span class="sxs-lookup"><span data-stu-id="ce5c8-162">start</span></span>](#start-datetime) | <span data-ttu-id="ce5c8-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-163">Member</span></span> |
| [<span data-ttu-id="ce5c8-164">subject</span><span class="sxs-lookup"><span data-stu-id="ce5c8-164">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="ce5c8-165">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-165">Member</span></span> |
| [<span data-ttu-id="ce5c8-166">to</span><span class="sxs-lookup"><span data-stu-id="ce5c8-166">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ce5c8-167">メンバー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-167">Member</span></span> |
| [<span data-ttu-id="ce5c8-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ce5c8-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="ce5c8-169">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-169">Method</span></span> |
| [<span data-ttu-id="ce5c8-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="ce5c8-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="ce5c8-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-171">Method</span></span> |
| [<span data-ttu-id="ce5c8-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ce5c8-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="ce5c8-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-173">Method</span></span> |
| [<span data-ttu-id="ce5c8-174">close</span><span class="sxs-lookup"><span data-stu-id="ce5c8-174">close</span></span>](#close) | <span data-ttu-id="ce5c8-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-175">Method</span></span> |
| [<span data-ttu-id="ce5c8-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="ce5c8-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="ce5c8-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-177">Method</span></span> |
| [<span data-ttu-id="ce5c8-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="ce5c8-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="ce5c8-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-179">Method</span></span> |
| [<span data-ttu-id="ce5c8-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="ce5c8-180">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="ce5c8-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-181">Method</span></span> |
| [<span data-ttu-id="ce5c8-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="ce5c8-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="ce5c8-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-183">Method</span></span> |
| [<span data-ttu-id="ce5c8-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="ce5c8-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="ce5c8-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-185">Method</span></span> |
| [<span data-ttu-id="ce5c8-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="ce5c8-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="ce5c8-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-187">Method</span></span> |
| [<span data-ttu-id="ce5c8-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="ce5c8-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="ce5c8-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-189">Method</span></span> |
| [<span data-ttu-id="ce5c8-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="ce5c8-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="ce5c8-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-191">Method</span></span> |
| [<span data-ttu-id="ce5c8-192">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="ce5c8-192">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="ce5c8-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-193">Method</span></span> |
| [<span data-ttu-id="ce5c8-194">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="ce5c8-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="ce5c8-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-195">Method</span></span> |
| [<span data-ttu-id="ce5c8-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="ce5c8-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="ce5c8-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-197">Method</span></span> |
| [<span data-ttu-id="ce5c8-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ce5c8-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="ce5c8-199">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-199">Method</span></span> |
| [<span data-ttu-id="ce5c8-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="ce5c8-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="ce5c8-201">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-201">Method</span></span> |
| [<span data-ttu-id="ce5c8-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="ce5c8-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="ce5c8-203">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-203">Method</span></span> |
| [<span data-ttu-id="ce5c8-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="ce5c8-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="ce5c8-205">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="ce5c8-206">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-206">Example</span></span>

<span data-ttu-id="ce5c8-207">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="ce5c8-208">Members</span><span class="sxs-lookup"><span data-stu-id="ce5c8-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-17"></a><span data-ttu-id="ce5c8-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="ce5c8-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

<span data-ttu-id="ce5c8-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ce5c8-212">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="ce5c8-213">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="ce5c8-214">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-214">Type</span></span>

*   <span data-ttu-id="ce5c8-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="ce5c8-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-216">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-216">Requirements</span></span>

|<span data-ttu-id="ce5c8-217">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-217">Requirement</span></span>|<span data-ttu-id="ce5c8-218">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-219">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-220">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-220">1.0</span></span>|
|[<span data-ttu-id="ce5c8-221">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-222">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-223">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-224">読み取り</span><span class="sxs-lookup"><span data-stu-id="ce5c8-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce5c8-225">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-225">Example</span></span>

<span data-ttu-id="ce5c8-226">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="ce5c8-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="ce5c8-228">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="ce5c8-229">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-229">Compose mode only.</span></span>

<span data-ttu-id="ce5c8-230">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-230">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ce5c8-231">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-231">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ce5c8-232">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-232">Get 500 members maximum.</span></span>
- <span data-ttu-id="ce5c8-233">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-233">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="ce5c8-234">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-234">Type</span></span>

*   [<span data-ttu-id="ce5c8-235">受信者</span><span class="sxs-lookup"><span data-stu-id="ce5c8-235">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="ce5c8-236">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-236">Requirements</span></span>

|<span data-ttu-id="ce5c8-237">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-237">Requirement</span></span>|<span data-ttu-id="ce5c8-238">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-239">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-240">1.1</span><span class="sxs-lookup"><span data-stu-id="ce5c8-240">1.1</span></span>|
|[<span data-ttu-id="ce5c8-241">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-242">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-243">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-244">作成</span><span class="sxs-lookup"><span data-stu-id="ce5c8-244">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ce5c8-245">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-245">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-17"></a><span data-ttu-id="ce5c8-246">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-246">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span></span>

<span data-ttu-id="ce5c8-247">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-247">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="ce5c8-248">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-248">Type</span></span>

*   [<span data-ttu-id="ce5c8-249">Body</span><span class="sxs-lookup"><span data-stu-id="ce5c8-249">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="ce5c8-250">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-250">Requirements</span></span>

|<span data-ttu-id="ce5c8-251">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-251">Requirement</span></span>|<span data-ttu-id="ce5c8-252">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-253">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-254">1.1</span><span class="sxs-lookup"><span data-stu-id="ce5c8-254">1.1</span></span>|
|[<span data-ttu-id="ce5c8-255">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-256">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-257">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-258">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ce5c8-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce5c8-259">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-259">Example</span></span>

<span data-ttu-id="ce5c8-260">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-260">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="ce5c8-261">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-261">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="ce5c8-262">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-262">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="ce5c8-263">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-263">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="ce5c8-264">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-264">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce5c8-265">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-265">Read mode</span></span>

<span data-ttu-id="ce5c8-266">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-266">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="ce5c8-267">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-267">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ce5c8-268">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-268">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="ce5c8-269">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-269">Compose mode</span></span>

<span data-ttu-id="ce5c8-270">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-270">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="ce5c8-271">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-271">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ce5c8-272">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-272">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ce5c8-273">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-273">Get 500 members maximum.</span></span>
- <span data-ttu-id="ce5c8-274">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-274">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ce5c8-275">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-275">Type</span></span>

*   <span data-ttu-id="ce5c8-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-277">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-277">Requirements</span></span>

|<span data-ttu-id="ce5c8-278">必要条件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-278">Requirement</span></span>|<span data-ttu-id="ce5c8-279">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-280">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-281">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-281">1.0</span></span>|
|[<span data-ttu-id="ce5c8-282">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-283">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-284">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-285">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ce5c8-285">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="ce5c8-286">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-286">(nullable) conversationId: String</span></span>

<span data-ttu-id="ce5c8-287">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-287">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="ce5c8-p109">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="ce5c8-p110">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="ce5c8-292">Type</span><span class="sxs-lookup"><span data-stu-id="ce5c8-292">Type</span></span>

*   <span data-ttu-id="ce5c8-293">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-293">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-294">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-294">Requirements</span></span>

|<span data-ttu-id="ce5c8-295">必要条件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-295">Requirement</span></span>|<span data-ttu-id="ce5c8-296">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-297">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-297">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-298">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-298">1.0</span></span>|
|[<span data-ttu-id="ce5c8-299">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-299">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-300">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-300">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-301">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-301">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-302">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ce5c8-302">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce5c8-303">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-303">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="ce5c8-304">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="ce5c8-304">dateTimeCreated: Date</span></span>

<span data-ttu-id="ce5c8-p111">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ce5c8-307">種類</span><span class="sxs-lookup"><span data-stu-id="ce5c8-307">Type</span></span>

*   <span data-ttu-id="ce5c8-308">日付</span><span class="sxs-lookup"><span data-stu-id="ce5c8-308">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-309">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-309">Requirements</span></span>

|<span data-ttu-id="ce5c8-310">必要条件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-310">Requirement</span></span>|<span data-ttu-id="ce5c8-311">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-312">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-313">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-313">1.0</span></span>|
|[<span data-ttu-id="ce5c8-314">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-315">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-316">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-317">読み取り</span><span class="sxs-lookup"><span data-stu-id="ce5c8-317">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce5c8-318">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-318">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="ce5c8-319">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="ce5c8-319">dateTimeModified: Date</span></span>

<span data-ttu-id="ce5c8-p112">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ce5c8-322">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-322">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="ce5c8-323">種類</span><span class="sxs-lookup"><span data-stu-id="ce5c8-323">Type</span></span>

*   <span data-ttu-id="ce5c8-324">日付</span><span class="sxs-lookup"><span data-stu-id="ce5c8-324">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-325">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-325">Requirements</span></span>

|<span data-ttu-id="ce5c8-326">必要条件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-326">Requirement</span></span>|<span data-ttu-id="ce5c8-327">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-328">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-329">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-329">1.0</span></span>|
|[<span data-ttu-id="ce5c8-330">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-330">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-331">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-332">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-332">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-333">読み取り</span><span class="sxs-lookup"><span data-stu-id="ce5c8-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce5c8-334">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-334">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="ce5c8-335">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-335">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="ce5c8-336">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-336">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="ce5c8-p113">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce5c8-339">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-339">Read mode</span></span>

<span data-ttu-id="ce5c8-340">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-340">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="ce5c8-341">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-341">Compose mode</span></span>

<span data-ttu-id="ce5c8-342">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-342">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="ce5c8-343">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-343">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="ce5c8-344">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-344">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="ce5c8-345">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-345">Type</span></span>

*   <span data-ttu-id="ce5c8-346">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-346">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-347">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-347">Requirements</span></span>

|<span data-ttu-id="ce5c8-348">必要条件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-348">Requirement</span></span>|<span data-ttu-id="ce5c8-349">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-349">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-350">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-350">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-351">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-351">1.0</span></span>|
|[<span data-ttu-id="ce5c8-352">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-352">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-353">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-354">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-354">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-355">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ce5c8-355">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17fromjavascriptapioutlookofficefromviewoutlook-js-17"></a><span data-ttu-id="ce5c8-356">from: [emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[from](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-356">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

<span data-ttu-id="ce5c8-357">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-357">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="ce5c8-p114">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ce5c8-360">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-360">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce5c8-361">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-361">Read mode</span></span>

<span data-ttu-id="ce5c8-362">プロパティ`from`は`EmailAddressDetails`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-362">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="ce5c8-363">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-363">Compose mode</span></span>

<span data-ttu-id="ce5c8-364">プロパティ`from`は、from `From`値を取得するメソッドを提供するオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-364">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ce5c8-365">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-365">Type</span></span>

*   <span data-ttu-id="ce5c8-366">[電子メールアドレス](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [の](/javascript/api/outlook/office.from?view=outlook-js-1.7)詳細</span><span class="sxs-lookup"><span data-stu-id="ce5c8-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-367">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-367">Requirements</span></span>

|<span data-ttu-id="ce5c8-368">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-368">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="ce5c8-369">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-369">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-370">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-370">1.0</span></span>|<span data-ttu-id="ce5c8-371">1.7</span><span class="sxs-lookup"><span data-stu-id="ce5c8-371">1.7</span></span>|
|[<span data-ttu-id="ce5c8-372">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-372">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-373">ReadItem</span></span>|<span data-ttu-id="ce5c8-374">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-374">ReadWriteItem</span></span>|
|[<span data-ttu-id="ce5c8-375">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-375">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-376">Read</span><span class="sxs-lookup"><span data-stu-id="ce5c8-376">Read</span></span>|<span data-ttu-id="ce5c8-377">Compose</span><span class="sxs-lookup"><span data-stu-id="ce5c8-377">Compose</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="ce5c8-378">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-378">internetMessageId: String</span></span>

<span data-ttu-id="ce5c8-p115">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ce5c8-381">Type</span><span class="sxs-lookup"><span data-stu-id="ce5c8-381">Type</span></span>

*   <span data-ttu-id="ce5c8-382">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-383">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-383">Requirements</span></span>

|<span data-ttu-id="ce5c8-384">必要条件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-384">Requirement</span></span>|<span data-ttu-id="ce5c8-385">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-386">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-387">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-387">1.0</span></span>|
|[<span data-ttu-id="ce5c8-388">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-389">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-390">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-391">読み取り</span><span class="sxs-lookup"><span data-stu-id="ce5c8-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce5c8-392">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-392">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="ce5c8-393">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-393">itemClass: String</span></span>

<span data-ttu-id="ce5c8-p116">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="ce5c8-p117">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="ce5c8-398">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-398">Type</span></span>|<span data-ttu-id="ce5c8-399">説明</span><span class="sxs-lookup"><span data-stu-id="ce5c8-399">Description</span></span>|<span data-ttu-id="ce5c8-400">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="ce5c8-400">item class</span></span>|
|---|---|---|
|<span data-ttu-id="ce5c8-401">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="ce5c8-401">Appointment items</span></span>|<span data-ttu-id="ce5c8-402">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-402">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="ce5c8-403">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="ce5c8-403">Message items</span></span>|<span data-ttu-id="ce5c8-404">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-404">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="ce5c8-405">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-405">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="ce5c8-406">Type</span><span class="sxs-lookup"><span data-stu-id="ce5c8-406">Type</span></span>

*   <span data-ttu-id="ce5c8-407">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-408">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-408">Requirements</span></span>

|<span data-ttu-id="ce5c8-409">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-409">Requirement</span></span>|<span data-ttu-id="ce5c8-410">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-411">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-412">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-412">1.0</span></span>|
|[<span data-ttu-id="ce5c8-413">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-414">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-415">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-416">読み取り</span><span class="sxs-lookup"><span data-stu-id="ce5c8-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce5c8-417">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-417">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="ce5c8-418">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-418">(nullable) itemId: String</span></span>

<span data-ttu-id="ce5c8-p118">現在のアイテムの [Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p118">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ce5c8-421">`itemId` プロパティから返される識別子は、[Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) と同じです。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-421">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="ce5c8-422">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-422">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="ce5c8-423">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-423">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="ce5c8-424">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-424">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="ce5c8-p120">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p120">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="ce5c8-427">Type</span><span class="sxs-lookup"><span data-stu-id="ce5c8-427">Type</span></span>

*   <span data-ttu-id="ce5c8-428">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-428">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-429">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-429">Requirements</span></span>

|<span data-ttu-id="ce5c8-430">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-430">Requirement</span></span>|<span data-ttu-id="ce5c8-431">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-431">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-432">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-432">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-433">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-433">1.0</span></span>|
|[<span data-ttu-id="ce5c8-434">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-434">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-435">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-435">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-436">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-436">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-437">読み取り</span><span class="sxs-lookup"><span data-stu-id="ce5c8-437">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce5c8-438">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-438">Example</span></span>

<span data-ttu-id="ce5c8-p121">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p121">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-17"></a><span data-ttu-id="ce5c8-441">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-441">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span></span>

<span data-ttu-id="ce5c8-442">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-442">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="ce5c8-443">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-443">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="ce5c8-444">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-444">Type</span></span>

*   [<span data-ttu-id="ce5c8-445">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="ce5c8-445">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="ce5c8-446">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-446">Requirements</span></span>

|<span data-ttu-id="ce5c8-447">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-447">Requirement</span></span>|<span data-ttu-id="ce5c8-448">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-448">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-449">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-449">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-450">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-450">1.0</span></span>|
|[<span data-ttu-id="ce5c8-451">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-451">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-452">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-452">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-453">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-453">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-454">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ce5c8-454">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce5c8-455">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-455">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-17"></a><span data-ttu-id="ce5c8-456">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-456">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

<span data-ttu-id="ce5c8-457">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-457">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce5c8-458">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-458">Read mode</span></span>

<span data-ttu-id="ce5c8-459">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-459">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="ce5c8-460">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-460">Compose mode</span></span>

<span data-ttu-id="ce5c8-461">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-461">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ce5c8-462">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-462">Type</span></span>

*   <span data-ttu-id="ce5c8-463">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-463">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-464">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-464">Requirements</span></span>

|<span data-ttu-id="ce5c8-465">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-465">Requirement</span></span>|<span data-ttu-id="ce5c8-466">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-467">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-468">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-468">1.0</span></span>|
|[<span data-ttu-id="ce5c8-469">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-470">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-471">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-472">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ce5c8-472">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="ce5c8-473">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-473">normalizedSubject: String</span></span>

<span data-ttu-id="ce5c8-p122">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p122">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="ce5c8-p123">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p123">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="ce5c8-478">Type</span><span class="sxs-lookup"><span data-stu-id="ce5c8-478">Type</span></span>

*   <span data-ttu-id="ce5c8-479">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-479">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-480">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-480">Requirements</span></span>

|<span data-ttu-id="ce5c8-481">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-481">Requirement</span></span>|<span data-ttu-id="ce5c8-482">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-482">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-483">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-483">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-484">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-484">1.0</span></span>|
|[<span data-ttu-id="ce5c8-485">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-485">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-486">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-486">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-487">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-487">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-488">読み取り</span><span class="sxs-lookup"><span data-stu-id="ce5c8-488">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce5c8-489">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-489">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-17"></a><span data-ttu-id="ce5c8-490">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-490">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span></span>

<span data-ttu-id="ce5c8-491">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-491">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="ce5c8-492">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-492">Type</span></span>

*   [<span data-ttu-id="ce5c8-493">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="ce5c8-493">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="ce5c8-494">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-494">Requirements</span></span>

|<span data-ttu-id="ce5c8-495">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-495">Requirement</span></span>|<span data-ttu-id="ce5c8-496">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-496">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-497">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-497">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-498">1.3</span><span class="sxs-lookup"><span data-stu-id="ce5c8-498">1.3</span></span>|
|[<span data-ttu-id="ce5c8-499">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-499">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-500">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-500">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-501">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-501">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-502">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ce5c8-502">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce5c8-503">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-503">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="ce5c8-504">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-504">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="ce5c8-505">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-505">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="ce5c8-506">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-506">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce5c8-507">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-507">Read mode</span></span>

<span data-ttu-id="ce5c8-508">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-508">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="ce5c8-509">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-509">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ce5c8-510">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-510">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="ce5c8-511">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-511">Compose mode</span></span>

<span data-ttu-id="ce5c8-512">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-512">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="ce5c8-513">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-513">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ce5c8-514">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-514">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ce5c8-515">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-515">Get 500 members maximum.</span></span>
- <span data-ttu-id="ce5c8-516">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-516">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ce5c8-517">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-517">Type</span></span>

*   <span data-ttu-id="ce5c8-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-519">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-519">Requirements</span></span>

|<span data-ttu-id="ce5c8-520">必要条件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-520">Requirement</span></span>|<span data-ttu-id="ce5c8-521">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-522">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-523">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-523">1.0</span></span>|
|[<span data-ttu-id="ce5c8-524">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-525">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-526">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-527">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ce5c8-527">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17organizerjavascriptapioutlookofficeorganizerviewoutlook-js-17"></a><span data-ttu-id="ce5c8-528">開催者: [emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[開催者](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-528">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

<span data-ttu-id="ce5c8-529">指定した会議の開催者の電子メールアドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-529">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce5c8-530">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-530">Read mode</span></span>

<span data-ttu-id="ce5c8-531">プロパティ`organizer`は、会議の開催者を表す[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-531">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="ce5c8-532">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-532">Compose mode</span></span>

<span data-ttu-id="ce5c8-533">プロパティ`organizer`は、開催者の値を取得するためのメソッドを提供する[オーガナイザー](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-533">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="ce5c8-534">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-534">Type</span></span>

*   <span data-ttu-id="ce5c8-535">[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [開催者](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-535">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-536">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-536">Requirements</span></span>

|<span data-ttu-id="ce5c8-537">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-537">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="ce5c8-538">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-539">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-539">1.0</span></span>|<span data-ttu-id="ce5c8-540">1.7</span><span class="sxs-lookup"><span data-stu-id="ce5c8-540">1.7</span></span>|
|[<span data-ttu-id="ce5c8-541">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-542">ReadItem</span></span>|<span data-ttu-id="ce5c8-543">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-543">ReadWriteItem</span></span>|
|[<span data-ttu-id="ce5c8-544">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-544">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-545">Read</span><span class="sxs-lookup"><span data-stu-id="ce5c8-545">Read</span></span>|<span data-ttu-id="ce5c8-546">Compose</span><span class="sxs-lookup"><span data-stu-id="ce5c8-546">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-17"></a><span data-ttu-id="ce5c8-547">(nullable) 定期的なスケジュール:[定期的](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)なアイテム</span><span class="sxs-lookup"><span data-stu-id="ce5c8-547">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span></span>

<span data-ttu-id="ce5c8-548">予定の定期的なパターンを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-548">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="ce5c8-549">会議出席依頼の定期的なパターンを取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-549">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="ce5c8-550">予定アイテムの読み取りおよび作成モード。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-550">Read and compose modes for appointment items.</span></span> <span data-ttu-id="ce5c8-551">会議出席依頼アイテムの閲覧モード。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-551">Read mode for meeting request items.</span></span>

<span data-ttu-id="ce5c8-552">この`recurrence`プロパティは、アイテムが series または series 内のインスタンスの場合、定期的な予定または会議出席依頼に対して[定期的](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-552">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="ce5c8-553">`null`は、単一の予定および1つの予定の会議出席依頼に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-553">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="ce5c8-554">`undefined`は、会議出席依頼ではないメッセージに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-554">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="ce5c8-555">注: 会議出席依頼に`itemClass`は、IPM という値があります。出席依頼。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-555">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="ce5c8-556">注: 定期的なオブジェクトが`null`の場合は、そのオブジェクトが単一の予定または1つの予定の会議出席依頼であり、データ系列の一部ではないことを示します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-556">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce5c8-557">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-557">Read mode</span></span>

<span data-ttu-id="ce5c8-558">この`recurrence`プロパティは、定期的な予定を表す[定期的](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-558">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that represents the appointment recurrence.</span></span> <span data-ttu-id="ce5c8-559">これは、予定および会議出席依頼に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-559">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="ce5c8-560">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-560">Compose mode</span></span>

<span data-ttu-id="ce5c8-561">この`recurrence`プロパティは、予定の繰り返しを管理するためのメソッドを提供する[定期的](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-561">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="ce5c8-562">これは予定に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-562">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="ce5c8-563">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-563">Type</span></span>

* [<span data-ttu-id="ce5c8-564">繰り返さ</span><span class="sxs-lookup"><span data-stu-id="ce5c8-564">Recurrence</span></span>](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)

|<span data-ttu-id="ce5c8-565">必要条件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-565">Requirement</span></span>|<span data-ttu-id="ce5c8-566">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-567">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-568">1.7</span><span class="sxs-lookup"><span data-stu-id="ce5c8-568">1.7</span></span>|
|[<span data-ttu-id="ce5c8-569">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-570">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-571">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-572">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ce5c8-572">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="ce5c8-573">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-573">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="ce5c8-574">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-574">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="ce5c8-575">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-575">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce5c8-576">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-576">Read mode</span></span>

<span data-ttu-id="ce5c8-577">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-577">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="ce5c8-578">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-578">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ce5c8-579">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-579">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="ce5c8-580">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-580">Compose mode</span></span>

<span data-ttu-id="ce5c8-581">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-581">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="ce5c8-582">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-582">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ce5c8-583">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-583">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ce5c8-584">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-584">Get 500 members maximum.</span></span>
- <span data-ttu-id="ce5c8-585">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-585">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="ce5c8-586">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-586">Type</span></span>

*   <span data-ttu-id="ce5c8-587">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-587">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-588">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-588">Requirements</span></span>

|<span data-ttu-id="ce5c8-589">必要条件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-589">Requirement</span></span>|<span data-ttu-id="ce5c8-590">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-590">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-591">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-591">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-592">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-592">1.0</span></span>|
|[<span data-ttu-id="ce5c8-593">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-593">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-594">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-594">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-595">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-595">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-596">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ce5c8-596">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17"></a><span data-ttu-id="ce5c8-597">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-597">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span></span>

<span data-ttu-id="ce5c8-p134">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p134">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="ce5c8-p135">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsfrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p135">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ce5c8-602">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-602">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="ce5c8-603">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-603">Type</span></span>

*   [<span data-ttu-id="ce5c8-604">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ce5c8-604">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="ce5c8-605">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-605">Requirements</span></span>

|<span data-ttu-id="ce5c8-606">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-606">Requirement</span></span>|<span data-ttu-id="ce5c8-607">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-608">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-609">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-609">1.0</span></span>|
|[<span data-ttu-id="ce5c8-610">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-610">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-611">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-612">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-613">読み取り</span><span class="sxs-lookup"><span data-stu-id="ce5c8-613">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce5c8-614">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-614">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="ce5c8-615">(nullable) 系列 Id: String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-615">(nullable) seriesId: String</span></span>

<span data-ttu-id="ce5c8-616">インスタンスが属する系列の id を取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-616">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="ce5c8-617">Web 上の Outlook およびデスクトップクライアントでは、 `seriesId`は、このアイテムが属する親 (シリーズ) アイテムの Exchange web サービス (EWS) ID を返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-617">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="ce5c8-618">ただし、iOS と Android では、 `seriesId`は親アイテムの REST ID を返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-618">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="ce5c8-619">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-619">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="ce5c8-620">`seriesId`プロパティが OUTLOOK REST API で使用される outlook id と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-620">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="ce5c8-621">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-621">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="ce5c8-622">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-622">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="ce5c8-623">この`seriesId`プロパティは`null` 、単一の予定、系列のアイテム、会議出席依頼などの親アイテムを持たないアイテムに`undefined`対して、会議出席依頼以外のアイテムに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-623">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="ce5c8-624">Type</span><span class="sxs-lookup"><span data-stu-id="ce5c8-624">Type</span></span>

* <span data-ttu-id="ce5c8-625">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-625">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-626">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-626">Requirements</span></span>

|<span data-ttu-id="ce5c8-627">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-627">Requirement</span></span>|<span data-ttu-id="ce5c8-628">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-629">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-630">1.7</span><span class="sxs-lookup"><span data-stu-id="ce5c8-630">1.7</span></span>|
|[<span data-ttu-id="ce5c8-631">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-631">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-632">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-633">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-633">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-634">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ce5c8-634">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce5c8-635">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-635">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="ce5c8-636">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-636">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="ce5c8-637">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-637">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="ce5c8-p138">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p138">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce5c8-640">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-640">Read mode</span></span>

<span data-ttu-id="ce5c8-641">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-641">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="ce5c8-642">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-642">Compose mode</span></span>

<span data-ttu-id="ce5c8-643">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-643">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="ce5c8-644">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-644">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="ce5c8-645">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-645">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="ce5c8-646">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-646">Type</span></span>

*   <span data-ttu-id="ce5c8-647">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-647">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-648">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-648">Requirements</span></span>

|<span data-ttu-id="ce5c8-649">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-649">Requirement</span></span>|<span data-ttu-id="ce5c8-650">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-650">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-651">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-651">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-652">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-652">1.0</span></span>|
|[<span data-ttu-id="ce5c8-653">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-653">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-654">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-654">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-655">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-655">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-656">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ce5c8-656">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-17"></a><span data-ttu-id="ce5c8-657">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-657">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

<span data-ttu-id="ce5c8-658">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-658">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="ce5c8-659">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-659">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce5c8-660">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-660">Read mode</span></span>

<span data-ttu-id="ce5c8-p139">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p139">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="ce5c8-663">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-663">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="ce5c8-664">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-664">Compose mode</span></span>

<span data-ttu-id="ce5c8-665">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-665">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="ce5c8-666">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-666">Type</span></span>

*   <span data-ttu-id="ce5c8-667">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-667">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-668">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-668">Requirements</span></span>

|<span data-ttu-id="ce5c8-669">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-669">Requirement</span></span>|<span data-ttu-id="ce5c8-670">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-670">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-671">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-671">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-672">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-672">1.0</span></span>|
|[<span data-ttu-id="ce5c8-673">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-673">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-674">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-674">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-675">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-675">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-676">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ce5c8-676">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="ce5c8-677">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-677">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="ce5c8-678">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-678">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="ce5c8-679">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-679">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce5c8-680">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-680">Read mode</span></span>

<span data-ttu-id="ce5c8-681">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-681">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="ce5c8-682">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-682">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ce5c8-683">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-683">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="ce5c8-684">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-684">Compose mode</span></span>

<span data-ttu-id="ce5c8-685">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-685">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="ce5c8-686">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-686">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ce5c8-687">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-687">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ce5c8-688">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-688">Get 500 members maximum.</span></span>
- <span data-ttu-id="ce5c8-689">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-689">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ce5c8-690">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-690">Type</span></span>

*   <span data-ttu-id="ce5c8-691">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-691">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-692">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-692">Requirements</span></span>

|<span data-ttu-id="ce5c8-693">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-693">Requirement</span></span>|<span data-ttu-id="ce5c8-694">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-695">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-696">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-696">1.0</span></span>|
|[<span data-ttu-id="ce5c8-697">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-697">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-698">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-698">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-699">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-699">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-700">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ce5c8-700">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="ce5c8-701">メソッド</span><span class="sxs-lookup"><span data-stu-id="ce5c8-701">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="ce5c8-702">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ce5c8-702">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ce5c8-703">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-703">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="ce5c8-704">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-704">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="ce5c8-705">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-705">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce5c8-706">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ce5c8-706">Parameters</span></span>
|<span data-ttu-id="ce5c8-707">名前</span><span class="sxs-lookup"><span data-stu-id="ce5c8-707">Name</span></span>|<span data-ttu-id="ce5c8-708">種類</span><span class="sxs-lookup"><span data-stu-id="ce5c8-708">Type</span></span>|<span data-ttu-id="ce5c8-709">属性</span><span class="sxs-lookup"><span data-stu-id="ce5c8-709">Attributes</span></span>|<span data-ttu-id="ce5c8-710">説明</span><span class="sxs-lookup"><span data-stu-id="ce5c8-710">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="ce5c8-711">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-711">String</span></span>||<span data-ttu-id="ce5c8-p143">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p143">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="ce5c8-714">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-714">String</span></span>||<span data-ttu-id="ce5c8-p144">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p144">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="ce5c8-717">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ce5c8-717">Object</span></span>|<span data-ttu-id="ce5c8-718">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-718">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-719">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-719">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ce5c8-720">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ce5c8-720">Object</span></span>|<span data-ttu-id="ce5c8-721">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-721">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-722">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-722">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="ce5c8-723">Boolean</span><span class="sxs-lookup"><span data-stu-id="ce5c8-723">Boolean</span></span>|<span data-ttu-id="ce5c8-724">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-724">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-725">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-725">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="ce5c8-726">function</span><span class="sxs-lookup"><span data-stu-id="ce5c8-726">function</span></span>|<span data-ttu-id="ce5c8-727">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-727">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-728">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-728">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ce5c8-729">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-729">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ce5c8-730">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-730">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ce5c8-731">エラー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-731">Errors</span></span>

|<span data-ttu-id="ce5c8-732">エラー コード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-732">Error code</span></span>|<span data-ttu-id="ce5c8-733">説明</span><span class="sxs-lookup"><span data-stu-id="ce5c8-733">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="ce5c8-734">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-734">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="ce5c8-735">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-735">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="ce5c8-736">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-736">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce5c8-737">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-737">Requirements</span></span>

|<span data-ttu-id="ce5c8-738">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-738">Requirement</span></span>|<span data-ttu-id="ce5c8-739">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-739">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-740">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-740">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-741">1.1</span><span class="sxs-lookup"><span data-stu-id="ce5c8-741">1.1</span></span>|
|[<span data-ttu-id="ce5c8-742">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-742">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-743">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-743">ReadWriteItem</span></span>|
|[<span data-ttu-id="ce5c8-744">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-744">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-745">作成</span><span class="sxs-lookup"><span data-stu-id="ce5c8-745">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="ce5c8-746">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-746">Examples</span></span>

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

<span data-ttu-id="ce5c8-747">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-747">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="ce5c8-748">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ce5c8-748">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="ce5c8-749">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-749">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="ce5c8-750">現在、サポートされて`Office.EventType.AppointmentTimeChanged`いる`Office.EventType.RecipientsChanged`イベントの種類は、、、です。`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="ce5c8-750">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce5c8-751">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ce5c8-751">Parameters</span></span>

| <span data-ttu-id="ce5c8-752">名前</span><span class="sxs-lookup"><span data-stu-id="ce5c8-752">Name</span></span> | <span data-ttu-id="ce5c8-753">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-753">Type</span></span> | <span data-ttu-id="ce5c8-754">属性</span><span class="sxs-lookup"><span data-stu-id="ce5c8-754">Attributes</span></span> | <span data-ttu-id="ce5c8-755">説明</span><span class="sxs-lookup"><span data-stu-id="ce5c8-755">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="ce5c8-756">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="ce5c8-756">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="ce5c8-757">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-757">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="ce5c8-758">Function</span><span class="sxs-lookup"><span data-stu-id="ce5c8-758">Function</span></span> || <span data-ttu-id="ce5c8-p145">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p145">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="ce5c8-762">Object</span><span class="sxs-lookup"><span data-stu-id="ce5c8-762">Object</span></span> | <span data-ttu-id="ce5c8-763">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-763">&lt;optional&gt;</span></span> | <span data-ttu-id="ce5c8-764">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-764">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="ce5c8-765">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ce5c8-765">Object</span></span> | <span data-ttu-id="ce5c8-766">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-766">&lt;optional&gt;</span></span> | <span data-ttu-id="ce5c8-767">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-767">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="ce5c8-768">function</span><span class="sxs-lookup"><span data-stu-id="ce5c8-768">function</span></span>| <span data-ttu-id="ce5c8-769">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-769">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-770">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-770">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce5c8-771">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-771">Requirements</span></span>

|<span data-ttu-id="ce5c8-772">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-772">Requirement</span></span>| <span data-ttu-id="ce5c8-773">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-773">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-774">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-774">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce5c8-775">1.7</span><span class="sxs-lookup"><span data-stu-id="ce5c8-775">1.7</span></span> |
|[<span data-ttu-id="ce5c8-776">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-776">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce5c8-777">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-777">ReadItem</span></span> |
|[<span data-ttu-id="ce5c8-778">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-778">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce5c8-779">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ce5c8-779">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="ce5c8-780">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-780">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="ce5c8-781">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ce5c8-781">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ce5c8-782">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-782">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="ce5c8-p146">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p146">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="ce5c8-786">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-786">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="ce5c8-787">Office アドインを Outlook on the web で実行している場合、編集中のアイテム以外のアイテムに `addItemAttachmentAsync` メソッドでアイテムを添付できます。ただし、これはサポートされていないため、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-787">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce5c8-788">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ce5c8-788">Parameters</span></span>

|<span data-ttu-id="ce5c8-789">名前</span><span class="sxs-lookup"><span data-stu-id="ce5c8-789">Name</span></span>|<span data-ttu-id="ce5c8-790">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-790">Type</span></span>|<span data-ttu-id="ce5c8-791">属性</span><span class="sxs-lookup"><span data-stu-id="ce5c8-791">Attributes</span></span>|<span data-ttu-id="ce5c8-792">説明</span><span class="sxs-lookup"><span data-stu-id="ce5c8-792">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="ce5c8-793">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-793">String</span></span>||<span data-ttu-id="ce5c8-p147">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p147">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="ce5c8-796">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-796">String</span></span>||<span data-ttu-id="ce5c8-797">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-797">The subject of the item to be attached.</span></span> <span data-ttu-id="ce5c8-798">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-798">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="ce5c8-799">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ce5c8-799">Object</span></span>|<span data-ttu-id="ce5c8-800">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-800">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-801">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-801">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ce5c8-802">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ce5c8-802">Object</span></span>|<span data-ttu-id="ce5c8-803">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-803">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-804">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-804">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ce5c8-805">function</span><span class="sxs-lookup"><span data-stu-id="ce5c8-805">function</span></span>|<span data-ttu-id="ce5c8-806">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-806">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-807">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-807">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ce5c8-808">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-808">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ce5c8-809">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-809">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ce5c8-810">エラー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-810">Errors</span></span>

|<span data-ttu-id="ce5c8-811">エラー コード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-811">Error code</span></span>|<span data-ttu-id="ce5c8-812">説明</span><span class="sxs-lookup"><span data-stu-id="ce5c8-812">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="ce5c8-813">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-813">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce5c8-814">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-814">Requirements</span></span>

|<span data-ttu-id="ce5c8-815">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-815">Requirement</span></span>|<span data-ttu-id="ce5c8-816">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-817">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-818">1.1</span><span class="sxs-lookup"><span data-stu-id="ce5c8-818">1.1</span></span>|
|[<span data-ttu-id="ce5c8-819">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-820">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-820">ReadWriteItem</span></span>|
|[<span data-ttu-id="ce5c8-821">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-822">作成</span><span class="sxs-lookup"><span data-stu-id="ce5c8-822">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ce5c8-823">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-823">Example</span></span>

<span data-ttu-id="ce5c8-824">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-824">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="ce5c8-825">close()</span><span class="sxs-lookup"><span data-stu-id="ce5c8-825">close()</span></span>

<span data-ttu-id="ce5c8-826">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-826">Closes the current item that is being composed.</span></span>

<span data-ttu-id="ce5c8-p149">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p149">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="ce5c8-829">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-829">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="ce5c8-830">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-830">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-831">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-831">Requirements</span></span>

|<span data-ttu-id="ce5c8-832">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-832">Requirement</span></span>|<span data-ttu-id="ce5c8-833">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-833">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-834">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-834">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-835">1.3</span><span class="sxs-lookup"><span data-stu-id="ce5c8-835">1.3</span></span>|
|[<span data-ttu-id="ce5c8-836">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-836">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-837">制限あり</span><span class="sxs-lookup"><span data-stu-id="ce5c8-837">Restricted</span></span>|
|[<span data-ttu-id="ce5c8-838">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-838">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-839">新規作成</span><span class="sxs-lookup"><span data-stu-id="ce5c8-839">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="ce5c8-840">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="ce5c8-840">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="ce5c8-841">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-841">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ce5c8-842">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-842">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ce5c8-843">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-843">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ce5c8-844">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-844">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="ce5c8-p150">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p150">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce5c8-848">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ce5c8-848">Parameters</span></span>

|<span data-ttu-id="ce5c8-849">名前</span><span class="sxs-lookup"><span data-stu-id="ce5c8-849">Name</span></span>|<span data-ttu-id="ce5c8-850">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-850">Type</span></span>|<span data-ttu-id="ce5c8-851">属性</span><span class="sxs-lookup"><span data-stu-id="ce5c8-851">Attributes</span></span>|<span data-ttu-id="ce5c8-852">説明</span><span class="sxs-lookup"><span data-stu-id="ce5c8-852">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="ce5c8-853">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="ce5c8-853">String &#124; Object</span></span>||<span data-ttu-id="ce5c8-p151">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p151">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ce5c8-856">**または**</span><span class="sxs-lookup"><span data-stu-id="ce5c8-856">**OR**</span></span><br/><span data-ttu-id="ce5c8-p152">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p152">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="ce5c8-859">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-859">String</span></span>|<span data-ttu-id="ce5c8-860">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-860">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-p153">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="ce5c8-863">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-863">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="ce5c8-864">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-864">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-865">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-865">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="ce5c8-866">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-866">String</span></span>||<span data-ttu-id="ce5c8-p154">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p154">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="ce5c8-869">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-869">String</span></span>||<span data-ttu-id="ce5c8-870">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-870">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="ce5c8-871">文字列</span><span class="sxs-lookup"><span data-stu-id="ce5c8-871">String</span></span>||<span data-ttu-id="ce5c8-p155">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p155">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="ce5c8-874">ブール値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-874">Boolean</span></span>||<span data-ttu-id="ce5c8-p156">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p156">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="ce5c8-877">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-877">String</span></span>||<span data-ttu-id="ce5c8-p157">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p157">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="ce5c8-881">function</span><span class="sxs-lookup"><span data-stu-id="ce5c8-881">function</span></span>|<span data-ttu-id="ce5c8-882">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-882">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-883">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-883">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce5c8-884">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-884">Requirements</span></span>

|<span data-ttu-id="ce5c8-885">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-885">Requirement</span></span>|<span data-ttu-id="ce5c8-886">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-886">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-887">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-887">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-888">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-888">1.0</span></span>|
|[<span data-ttu-id="ce5c8-889">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-889">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-890">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-890">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-891">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-891">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-892">読み取り</span><span class="sxs-lookup"><span data-stu-id="ce5c8-892">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ce5c8-893">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-893">Examples</span></span>

<span data-ttu-id="ce5c8-894">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-894">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="ce5c8-895">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-895">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="ce5c8-896">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-896">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ce5c8-897">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-897">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="ce5c8-898">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-898">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="ce5c8-899">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-899">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="ce5c8-900">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="ce5c8-900">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="ce5c8-901">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-901">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ce5c8-902">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-902">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ce5c8-903">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-903">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ce5c8-904">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-904">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="ce5c8-p158">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p158">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce5c8-908">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ce5c8-908">Parameters</span></span>

|<span data-ttu-id="ce5c8-909">名前</span><span class="sxs-lookup"><span data-stu-id="ce5c8-909">Name</span></span>|<span data-ttu-id="ce5c8-910">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-910">Type</span></span>|<span data-ttu-id="ce5c8-911">属性</span><span class="sxs-lookup"><span data-stu-id="ce5c8-911">Attributes</span></span>|<span data-ttu-id="ce5c8-912">説明</span><span class="sxs-lookup"><span data-stu-id="ce5c8-912">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="ce5c8-913">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="ce5c8-913">String &#124; Object</span></span>||<span data-ttu-id="ce5c8-p159">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p159">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ce5c8-916">**または**</span><span class="sxs-lookup"><span data-stu-id="ce5c8-916">**OR**</span></span><br/><span data-ttu-id="ce5c8-p160">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p160">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="ce5c8-919">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-919">String</span></span>|<span data-ttu-id="ce5c8-920">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-920">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-p161">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p161">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="ce5c8-923">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-923">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="ce5c8-924">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-924">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-925">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-925">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="ce5c8-926">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-926">String</span></span>||<span data-ttu-id="ce5c8-p162">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p162">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="ce5c8-929">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-929">String</span></span>||<span data-ttu-id="ce5c8-930">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-930">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="ce5c8-931">文字列</span><span class="sxs-lookup"><span data-stu-id="ce5c8-931">String</span></span>||<span data-ttu-id="ce5c8-p163">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p163">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="ce5c8-934">ブール値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-934">Boolean</span></span>||<span data-ttu-id="ce5c8-p164">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p164">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="ce5c8-937">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-937">String</span></span>||<span data-ttu-id="ce5c8-p165">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p165">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="ce5c8-941">function</span><span class="sxs-lookup"><span data-stu-id="ce5c8-941">function</span></span>|<span data-ttu-id="ce5c8-942">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-942">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-943">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-943">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce5c8-944">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-944">Requirements</span></span>

|<span data-ttu-id="ce5c8-945">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-945">Requirement</span></span>|<span data-ttu-id="ce5c8-946">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-947">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-948">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-948">1.0</span></span>|
|[<span data-ttu-id="ce5c8-949">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-949">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-950">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-950">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-951">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-951">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-952">読み取り</span><span class="sxs-lookup"><span data-stu-id="ce5c8-952">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ce5c8-953">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-953">Examples</span></span>

<span data-ttu-id="ce5c8-954">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-954">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="ce5c8-955">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-955">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="ce5c8-956">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-956">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ce5c8-957">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-957">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="ce5c8-958">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-958">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="ce5c8-959">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-959">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="ce5c8-960">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="ce5c8-960">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="ce5c8-961">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-961">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="ce5c8-962">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-962">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-963">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-963">Requirements</span></span>

|<span data-ttu-id="ce5c8-964">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-964">Requirement</span></span>|<span data-ttu-id="ce5c8-965">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-966">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-966">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-967">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-967">1.0</span></span>|
|[<span data-ttu-id="ce5c8-968">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-968">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-969">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-970">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-970">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-971">読み取り</span><span class="sxs-lookup"><span data-stu-id="ce5c8-971">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ce5c8-972">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ce5c8-972">Returns:</span></span>

<span data-ttu-id="ce5c8-973">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-973">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="ce5c8-974">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-974">Example</span></span>

<span data-ttu-id="ce5c8-975">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-975">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="ce5c8-976">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="ce5c8-976">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="ce5c8-977">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-977">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="ce5c8-978">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-978">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce5c8-979">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ce5c8-979">Parameters</span></span>

|<span data-ttu-id="ce5c8-980">名前</span><span class="sxs-lookup"><span data-stu-id="ce5c8-980">Name</span></span>|<span data-ttu-id="ce5c8-981">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-981">Type</span></span>|<span data-ttu-id="ce5c8-982">説明</span><span class="sxs-lookup"><span data-stu-id="ce5c8-982">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="ce5c8-983">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="ce5c8-983">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.7)|<span data-ttu-id="ce5c8-984">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-984">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce5c8-985">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-985">Requirements</span></span>

|<span data-ttu-id="ce5c8-986">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-986">Requirement</span></span>|<span data-ttu-id="ce5c8-987">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-987">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-988">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-988">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-989">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-989">1.0</span></span>|
|[<span data-ttu-id="ce5c8-990">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-990">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-991">制限あり</span><span class="sxs-lookup"><span data-stu-id="ce5c8-991">Restricted</span></span>|
|[<span data-ttu-id="ce5c8-992">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-992">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-993">読み取り</span><span class="sxs-lookup"><span data-stu-id="ce5c8-993">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ce5c8-994">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ce5c8-994">Returns:</span></span>

<span data-ttu-id="ce5c8-995">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-995">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="ce5c8-996">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-996">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="ce5c8-997">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-997">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="ce5c8-998">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-998">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="ce5c8-999">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-999">Value of `entityType`</span></span>|<span data-ttu-id="ce5c8-1000">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1000">Type of objects in returned array</span></span>|<span data-ttu-id="ce5c8-1001">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1001">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="ce5c8-1002">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1002">String</span></span>|<span data-ttu-id="ce5c8-1003">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1003">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="ce5c8-1004">連絡先</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1004">Contact</span></span>|<span data-ttu-id="ce5c8-1005">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1005">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="ce5c8-1006">文字列</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1006">String</span></span>|<span data-ttu-id="ce5c8-1007">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1007">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="ce5c8-1008">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1008">MeetingSuggestion</span></span>|<span data-ttu-id="ce5c8-1009">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1009">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="ce5c8-1010">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1010">PhoneNumber</span></span>|<span data-ttu-id="ce5c8-1011">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1011">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="ce5c8-1012">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1012">TaskSuggestion</span></span>|<span data-ttu-id="ce5c8-1013">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1013">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="ce5c8-1014">文字列</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1014">String</span></span>|<span data-ttu-id="ce5c8-1015">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1015">**Restricted**</span></span>|

<span data-ttu-id="ce5c8-1016">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="ce5c8-1016">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

##### <a name="example"></a><span data-ttu-id="ce5c8-1017">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1017">Example</span></span>

<span data-ttu-id="ce5c8-1018">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1018">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="ce5c8-1019">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1019">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="ce5c8-1020">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1020">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ce5c8-1021">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1021">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ce5c8-1022">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1022">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce5c8-1023">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1023">Parameters</span></span>

|<span data-ttu-id="ce5c8-1024">名前</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1024">Name</span></span>|<span data-ttu-id="ce5c8-1025">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1025">Type</span></span>|<span data-ttu-id="ce5c8-1026">説明</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1026">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="ce5c8-1027">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1027">String</span></span>|<span data-ttu-id="ce5c8-1028">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1028">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce5c8-1029">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1029">Requirements</span></span>

|<span data-ttu-id="ce5c8-1030">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1030">Requirement</span></span>|<span data-ttu-id="ce5c8-1031">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1031">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-1032">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1032">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-1033">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1033">1.0</span></span>|
|[<span data-ttu-id="ce5c8-1034">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1034">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-1035">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1035">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-1036">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1036">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-1037">読み取り</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1037">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ce5c8-1038">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1038">Returns:</span></span>

<span data-ttu-id="ce5c8-p167">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p167">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="ce5c8-1041">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="ce5c8-1041">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="ce5c8-1042">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1042">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="ce5c8-1043">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1043">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ce5c8-1044">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1044">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ce5c8-p168">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p168">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="ce5c8-1048">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1048">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="ce5c8-1049">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1049">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="ce5c8-p169">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-1053">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1053">Requirements</span></span>

|<span data-ttu-id="ce5c8-1054">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1054">Requirement</span></span>|<span data-ttu-id="ce5c8-1055">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-1056">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-1057">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1057">1.0</span></span>|
|[<span data-ttu-id="ce5c8-1058">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1058">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-1059">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1059">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-1060">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1060">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-1061">読み取り</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1061">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ce5c8-1062">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1062">Returns:</span></span>

<span data-ttu-id="ce5c8-p170">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="ce5c8-1065">型: Object</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1065">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="ce5c8-1066">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1066">Example</span></span>

<span data-ttu-id="ce5c8-1067">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1067">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="ce5c8-1068">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1068">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="ce5c8-1069">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1069">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ce5c8-1070">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1070">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ce5c8-1071">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1071">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="ce5c8-p171">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce5c8-1074">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1074">Parameters</span></span>

|<span data-ttu-id="ce5c8-1075">名前</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1075">Name</span></span>|<span data-ttu-id="ce5c8-1076">種類</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1076">Type</span></span>|<span data-ttu-id="ce5c8-1077">説明</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1077">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="ce5c8-1078">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1078">String</span></span>|<span data-ttu-id="ce5c8-1079">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1079">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce5c8-1080">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1080">Requirements</span></span>

|<span data-ttu-id="ce5c8-1081">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1081">Requirement</span></span>|<span data-ttu-id="ce5c8-1082">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-1083">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1083">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-1084">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1084">1.0</span></span>|
|[<span data-ttu-id="ce5c8-1085">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1085">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1086">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-1087">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1087">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-1088">読み取り</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1088">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ce5c8-1089">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1089">Returns:</span></span>

<span data-ttu-id="ce5c8-1090">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1090">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="ce5c8-1091">型: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="ce5c8-1091">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="ce5c8-1092">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1092">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="ce5c8-1093">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1093">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="ce5c8-1094">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1094">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="ce5c8-p172">選択されていない状態でカーソルが本文または件名にある場合、メソッドは選択されたデータに対し空の文字列を返します。本文または件名以外のフィールドが選択されている場合には、メソッドは`InvalidSelection`エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p172">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce5c8-1097">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1097">Parameters</span></span>

|<span data-ttu-id="ce5c8-1098">名前</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1098">Name</span></span>|<span data-ttu-id="ce5c8-1099">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1099">Type</span></span>|<span data-ttu-id="ce5c8-1100">属性</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1100">Attributes</span></span>|<span data-ttu-id="ce5c8-1101">説明</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1101">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="ce5c8-1102">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1102">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="ce5c8-p173">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p173">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="ce5c8-1106">Object</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1106">Object</span></span>|<span data-ttu-id="ce5c8-1107">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1107">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-1108">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1108">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ce5c8-1109">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1109">Object</span></span>|<span data-ttu-id="ce5c8-1110">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1110">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-1111">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1111">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ce5c8-1112">function</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1112">function</span></span>||<span data-ttu-id="ce5c8-1113">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1113">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ce5c8-1114">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1114">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="ce5c8-1115">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1115">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce5c8-1116">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1116">Requirements</span></span>

|<span data-ttu-id="ce5c8-1117">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1117">Requirement</span></span>|<span data-ttu-id="ce5c8-1118">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1118">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-1119">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1119">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-1120">1.2</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1120">1.2</span></span>|
|[<span data-ttu-id="ce5c8-1121">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1121">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-1122">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1122">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-1123">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1123">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-1124">作成</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1124">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="ce5c8-1125">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1125">Returns:</span></span>

<span data-ttu-id="ce5c8-1126">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1126">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="ce5c8-1127">型:String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1127">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="ce5c8-1128">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1128">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  console.log("Selected text in " + prop + ": " + text);
}
```

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="ce5c8-1129">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1129">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="ce5c8-1130">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1130">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="ce5c8-1131">強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1131">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="ce5c8-1132">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1132">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-1133">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1133">Requirements</span></span>

|<span data-ttu-id="ce5c8-1134">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1134">Requirement</span></span>|<span data-ttu-id="ce5c8-1135">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1135">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-1136">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1136">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-1137">1.6</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1137">1.6</span></span>|
|[<span data-ttu-id="ce5c8-1138">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1138">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-1139">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1139">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-1140">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1140">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-1141">読み取り</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1141">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ce5c8-1142">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1142">Returns:</span></span>

<span data-ttu-id="ce5c8-1143">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1143">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="ce5c8-1144">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1144">Example</span></span>

<span data-ttu-id="ce5c8-1145">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1145">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="ce5c8-1146">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1146">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="ce5c8-p176">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p176">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="ce5c8-1149">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1149">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ce5c8-p177">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p177">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="ce5c8-1153">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1153">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="ce5c8-1154">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1154">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="ce5c8-p178">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p178">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce5c8-1158">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1158">Requirements</span></span>

|<span data-ttu-id="ce5c8-1159">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1159">Requirement</span></span>|<span data-ttu-id="ce5c8-1160">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1160">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-1161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-1162">1.6</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1162">1.6</span></span>|
|[<span data-ttu-id="ce5c8-1163">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-1164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1164">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-1165">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-1166">読み取り</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1166">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ce5c8-1167">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1167">Returns:</span></span>

<span data-ttu-id="ce5c8-p179">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p179">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="ce5c8-1170">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1170">Example</span></span>

<span data-ttu-id="ce5c8-1171">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1171">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="ce5c8-1172">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1172">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="ce5c8-1173">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1173">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="ce5c8-p180">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p180">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce5c8-1177">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1177">Parameters</span></span>

|<span data-ttu-id="ce5c8-1178">名前</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1178">Name</span></span>|<span data-ttu-id="ce5c8-1179">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1179">Type</span></span>|<span data-ttu-id="ce5c8-1180">属性</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1180">Attributes</span></span>|<span data-ttu-id="ce5c8-1181">説明</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1181">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="ce5c8-1182">function</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1182">function</span></span>||<span data-ttu-id="ce5c8-1183">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1183">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ce5c8-1184">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1184">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="ce5c8-1185">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1185">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="ce5c8-1186">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1186">Object</span></span>|<span data-ttu-id="ce5c8-1187">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1187">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-1188">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1188">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="ce5c8-1189">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1189">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce5c8-1190">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1190">Requirements</span></span>

|<span data-ttu-id="ce5c8-1191">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1191">Requirement</span></span>|<span data-ttu-id="ce5c8-1192">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1192">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-1193">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1193">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-1194">1.0</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1194">1.0</span></span>|
|[<span data-ttu-id="ce5c8-1195">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1195">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-1196">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1196">ReadItem</span></span>|
|[<span data-ttu-id="ce5c8-1197">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1197">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-1198">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1198">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce5c8-1199">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1199">Example</span></span>

<span data-ttu-id="ce5c8-p183">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p183">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="ce5c8-1203">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1203">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="ce5c8-1204">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1204">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="ce5c8-1205">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1205">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="ce5c8-1206">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1206">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="ce5c8-1207">Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1207">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="ce5c8-1208">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1208">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce5c8-1209">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1209">Parameters</span></span>

|<span data-ttu-id="ce5c8-1210">名前</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1210">Name</span></span>|<span data-ttu-id="ce5c8-1211">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1211">Type</span></span>|<span data-ttu-id="ce5c8-1212">属性</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1212">Attributes</span></span>|<span data-ttu-id="ce5c8-1213">説明</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1213">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="ce5c8-1214">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1214">String</span></span>||<span data-ttu-id="ce5c8-1215">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1215">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="ce5c8-1216">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1216">Object</span></span>|<span data-ttu-id="ce5c8-1217">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1217">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-1218">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1218">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ce5c8-1219">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1219">Object</span></span>|<span data-ttu-id="ce5c8-1220">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1220">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-1221">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1221">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ce5c8-1222">function</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1222">function</span></span>|<span data-ttu-id="ce5c8-1223">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1223">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-1224">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1224">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ce5c8-1225">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1225">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ce5c8-1226">エラー</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1226">Errors</span></span>

|<span data-ttu-id="ce5c8-1227">エラー コード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1227">Error code</span></span>|<span data-ttu-id="ce5c8-1228">説明</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1228">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="ce5c8-1229">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1229">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce5c8-1230">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1230">Requirements</span></span>

|<span data-ttu-id="ce5c8-1231">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1231">Requirement</span></span>|<span data-ttu-id="ce5c8-1232">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1232">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-1233">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-1234">1.1</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1234">1.1</span></span>|
|[<span data-ttu-id="ce5c8-1235">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1235">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-1236">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1236">ReadWriteItem</span></span>|
|[<span data-ttu-id="ce5c8-1237">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1237">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-1238">作成</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1238">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ce5c8-1239">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1239">Example</span></span>

<span data-ttu-id="ce5c8-1240">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1240">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="ce5c8-1241">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1241">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="ce5c8-1242">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1242">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="ce5c8-1243">現在、サポートされて`Office.EventType.AppointmentTimeChanged`いる`Office.EventType.RecipientsChanged`イベントの種類は、、、です。`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1243">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce5c8-1244">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1244">Parameters</span></span>

| <span data-ttu-id="ce5c8-1245">名前</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1245">Name</span></span> | <span data-ttu-id="ce5c8-1246">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1246">Type</span></span> | <span data-ttu-id="ce5c8-1247">属性</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1247">Attributes</span></span> | <span data-ttu-id="ce5c8-1248">説明</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1248">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="ce5c8-1249">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1249">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="ce5c8-1250">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1250">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="ce5c8-1251">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1251">Object</span></span> | <span data-ttu-id="ce5c8-1252">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1252">&lt;optional&gt;</span></span> | <span data-ttu-id="ce5c8-1253">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1253">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="ce5c8-1254">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1254">Object</span></span> | <span data-ttu-id="ce5c8-1255">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1255">&lt;optional&gt;</span></span> | <span data-ttu-id="ce5c8-1256">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1256">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="ce5c8-1257">function</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1257">function</span></span>| <span data-ttu-id="ce5c8-1258">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1258">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-1259">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1259">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce5c8-1260">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1260">Requirements</span></span>

|<span data-ttu-id="ce5c8-1261">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1261">Requirement</span></span>| <span data-ttu-id="ce5c8-1262">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1262">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-1263">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce5c8-1264">1.7</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1264">1.7</span></span> |
|[<span data-ttu-id="ce5c8-1265">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce5c8-1266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1266">ReadItem</span></span> |
|[<span data-ttu-id="ce5c8-1267">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce5c8-1268">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1268">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="ce5c8-1269">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1269">Example</span></span>

```js
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

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="ce5c8-1270">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1270">saveAsync([options], callback)</span></span>

<span data-ttu-id="ce5c8-1271">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1271">Asynchronously saves an item.</span></span>

<span data-ttu-id="ce5c8-1272">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1272">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="ce5c8-1273">Outlook on the web またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1273">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="ce5c8-1274">キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1274">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="ce5c8-1275">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1275">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="ce5c8-1276">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1276">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="ce5c8-p187">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p187">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="ce5c8-1280">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1280">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="ce5c8-1281">Outlook on Mac では、会議の保存はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1281">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="ce5c8-1282">`saveAsync` メソッドは、作成モードの会議から呼び出されると失敗します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1282">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="ce5c8-1283">回避策については、「[Office JS API を使用して Outlook for Mac で会議を下書きとして保存できない](https://support.microsoft.com/help/4505745)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1283">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="ce5c8-1284">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1284">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce5c8-1285">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1285">Parameters</span></span>

|<span data-ttu-id="ce5c8-1286">名前</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1286">Name</span></span>|<span data-ttu-id="ce5c8-1287">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1287">Type</span></span>|<span data-ttu-id="ce5c8-1288">属性</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1288">Attributes</span></span>|<span data-ttu-id="ce5c8-1289">説明</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1289">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="ce5c8-1290">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1290">Object</span></span>|<span data-ttu-id="ce5c8-1291">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1291">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-1292">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1292">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ce5c8-1293">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1293">Object</span></span>|<span data-ttu-id="ce5c8-1294">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1294">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-1295">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1295">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ce5c8-1296">関数</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1296">function</span></span>||<span data-ttu-id="ce5c8-1297">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1297">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ce5c8-1298">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1298">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce5c8-1299">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1299">Requirements</span></span>

|<span data-ttu-id="ce5c8-1300">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1300">Requirement</span></span>|<span data-ttu-id="ce5c8-1301">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1301">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-1302">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-1303">1.3</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1303">1.3</span></span>|
|[<span data-ttu-id="ce5c8-1304">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-1305">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1305">ReadWriteItem</span></span>|
|[<span data-ttu-id="ce5c8-1306">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-1307">作成</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1307">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="ce5c8-1308">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1308">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="ce5c8-p189">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p189">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="ce5c8-1311">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1311">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="ce5c8-1312">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1312">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="ce5c8-p190">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p190">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce5c8-1316">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1316">Parameters</span></span>

|<span data-ttu-id="ce5c8-1317">名前</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1317">Name</span></span>|<span data-ttu-id="ce5c8-1318">型</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1318">Type</span></span>|<span data-ttu-id="ce5c8-1319">属性</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1319">Attributes</span></span>|<span data-ttu-id="ce5c8-1320">説明</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1320">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="ce5c8-1321">String</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1321">String</span></span>||<span data-ttu-id="ce5c8-p191">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-p191">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="ce5c8-1325">Object</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1325">Object</span></span>|<span data-ttu-id="ce5c8-1326">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1326">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-1327">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1327">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ce5c8-1328">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1328">Object</span></span>|<span data-ttu-id="ce5c8-1329">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1329">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-1330">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1330">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="ce5c8-1331">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1331">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="ce5c8-1332">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1332">&lt;optional&gt;</span></span>|<span data-ttu-id="ce5c8-1333">`text` の場合、Outlook on the web とデスクトップ クライアントでは現在のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1333">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="ce5c8-1334">フィールドが HTML エディターの場合、データが HTML の場合でもテキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1334">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="ce5c8-1335">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Outlook on the web では現在のスタイルが適用され、Outlook デスクトップ クライアントでは既定のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1335">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="ce5c8-1336">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1336">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="ce5c8-1337">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1337">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="ce5c8-1338">function</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1338">function</span></span>||<span data-ttu-id="ce5c8-1339">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1339">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce5c8-1340">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1340">Requirements</span></span>

|<span data-ttu-id="ce5c8-1341">要件</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1341">Requirement</span></span>|<span data-ttu-id="ce5c8-1342">値</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1342">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce5c8-1343">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1343">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ce5c8-1344">1.2</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1344">1.2</span></span>|
|[<span data-ttu-id="ce5c8-1345">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1345">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ce5c8-1346">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1346">ReadWriteItem</span></span>|
|[<span data-ttu-id="ce5c8-1347">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1347">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="ce5c8-1348">作成</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1348">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ce5c8-1349">例</span><span class="sxs-lookup"><span data-stu-id="ce5c8-1349">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
