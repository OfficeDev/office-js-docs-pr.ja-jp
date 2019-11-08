---
title: Office. メールボックス-要件セット1.7
description: ''
ms.date: 11/06/2019
localization_priority: Normal
ms.openlocfilehash: 1c0948490c5c0b77252a8605b43f85dd529f2897
ms.sourcegitcommit: 08c0b9ff319c391922fa43d3c2e9783cf6b53b1b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/08/2019
ms.locfileid: "38066215"
---
# <a name="item"></a><span data-ttu-id="f70d6-102">item</span><span class="sxs-lookup"><span data-stu-id="f70d6-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="f70d6-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="f70d6-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="f70d6-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-106">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-106">Requirements</span></span>

|<span data-ttu-id="f70d6-107">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-107">Requirement</span></span>|<span data-ttu-id="f70d6-108">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-110">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-110">1.0</span></span>|
|[<span data-ttu-id="f70d6-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="f70d6-112">Restricted</span></span>|
|[<span data-ttu-id="f70d6-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f70d6-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f70d6-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-115">Members and methods</span></span>

| <span data-ttu-id="f70d6-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-116">Member</span></span> | <span data-ttu-id="f70d6-117">種類</span><span class="sxs-lookup"><span data-stu-id="f70d6-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f70d6-118">attachments</span><span class="sxs-lookup"><span data-stu-id="f70d6-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="f70d6-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-119">Member</span></span> |
| [<span data-ttu-id="f70d6-120">bcc</span><span class="sxs-lookup"><span data-stu-id="f70d6-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="f70d6-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-121">Member</span></span> |
| [<span data-ttu-id="f70d6-122">body</span><span class="sxs-lookup"><span data-stu-id="f70d6-122">body</span></span>](#body-body) | <span data-ttu-id="f70d6-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-123">Member</span></span> |
| [<span data-ttu-id="f70d6-124">cc</span><span class="sxs-lookup"><span data-stu-id="f70d6-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f70d6-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-125">Member</span></span> |
| [<span data-ttu-id="f70d6-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="f70d6-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="f70d6-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-127">Member</span></span> |
| [<span data-ttu-id="f70d6-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="f70d6-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="f70d6-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-129">Member</span></span> |
| [<span data-ttu-id="f70d6-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="f70d6-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="f70d6-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-131">Member</span></span> |
| [<span data-ttu-id="f70d6-132">end</span><span class="sxs-lookup"><span data-stu-id="f70d6-132">end</span></span>](#end-datetime) | <span data-ttu-id="f70d6-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-133">Member</span></span> |
| [<span data-ttu-id="f70d6-134">from</span><span class="sxs-lookup"><span data-stu-id="f70d6-134">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="f70d6-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-135">Member</span></span> |
| [<span data-ttu-id="f70d6-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="f70d6-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="f70d6-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-137">Member</span></span> |
| [<span data-ttu-id="f70d6-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="f70d6-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="f70d6-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-139">Member</span></span> |
| [<span data-ttu-id="f70d6-140">itemId</span><span class="sxs-lookup"><span data-stu-id="f70d6-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="f70d6-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-141">Member</span></span> |
| [<span data-ttu-id="f70d6-142">itemType</span><span class="sxs-lookup"><span data-stu-id="f70d6-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="f70d6-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-143">Member</span></span> |
| [<span data-ttu-id="f70d6-144">location</span><span class="sxs-lookup"><span data-stu-id="f70d6-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="f70d6-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-145">Member</span></span> |
| [<span data-ttu-id="f70d6-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="f70d6-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="f70d6-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-147">Member</span></span> |
| [<span data-ttu-id="f70d6-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="f70d6-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="f70d6-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-149">Member</span></span> |
| [<span data-ttu-id="f70d6-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="f70d6-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f70d6-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-151">Member</span></span> |
| [<span data-ttu-id="f70d6-152">organizer</span><span class="sxs-lookup"><span data-stu-id="f70d6-152">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="f70d6-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-153">Member</span></span> |
| [<span data-ttu-id="f70d6-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="f70d6-154">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="f70d6-155">Member</span><span class="sxs-lookup"><span data-stu-id="f70d6-155">Member</span></span> |
| [<span data-ttu-id="f70d6-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="f70d6-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f70d6-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-157">Member</span></span> |
| [<span data-ttu-id="f70d6-158">sender</span><span class="sxs-lookup"><span data-stu-id="f70d6-158">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="f70d6-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-159">Member</span></span> |
| [<span data-ttu-id="f70d6-160">系列 Id</span><span class="sxs-lookup"><span data-stu-id="f70d6-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="f70d6-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-161">Member</span></span> |
| [<span data-ttu-id="f70d6-162">start</span><span class="sxs-lookup"><span data-stu-id="f70d6-162">start</span></span>](#start-datetime) | <span data-ttu-id="f70d6-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-163">Member</span></span> |
| [<span data-ttu-id="f70d6-164">subject</span><span class="sxs-lookup"><span data-stu-id="f70d6-164">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="f70d6-165">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-165">Member</span></span> |
| [<span data-ttu-id="f70d6-166">to</span><span class="sxs-lookup"><span data-stu-id="f70d6-166">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f70d6-167">メンバー</span><span class="sxs-lookup"><span data-stu-id="f70d6-167">Member</span></span> |
| [<span data-ttu-id="f70d6-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f70d6-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="f70d6-169">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-169">Method</span></span> |
| [<span data-ttu-id="f70d6-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="f70d6-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="f70d6-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-171">Method</span></span> |
| [<span data-ttu-id="f70d6-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f70d6-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="f70d6-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-173">Method</span></span> |
| [<span data-ttu-id="f70d6-174">close</span><span class="sxs-lookup"><span data-stu-id="f70d6-174">close</span></span>](#close) | <span data-ttu-id="f70d6-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-175">Method</span></span> |
| [<span data-ttu-id="f70d6-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="f70d6-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="f70d6-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-177">Method</span></span> |
| [<span data-ttu-id="f70d6-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="f70d6-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="f70d6-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-179">Method</span></span> |
| [<span data-ttu-id="f70d6-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="f70d6-180">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="f70d6-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-181">Method</span></span> |
| [<span data-ttu-id="f70d6-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="f70d6-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="f70d6-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-183">Method</span></span> |
| [<span data-ttu-id="f70d6-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="f70d6-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="f70d6-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-185">Method</span></span> |
| [<span data-ttu-id="f70d6-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="f70d6-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="f70d6-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-187">Method</span></span> |
| [<span data-ttu-id="f70d6-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="f70d6-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="f70d6-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-189">Method</span></span> |
| [<span data-ttu-id="f70d6-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="f70d6-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="f70d6-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-191">Method</span></span> |
| [<span data-ttu-id="f70d6-192">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="f70d6-192">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="f70d6-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-193">Method</span></span> |
| [<span data-ttu-id="f70d6-194">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="f70d6-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="f70d6-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-195">Method</span></span> |
| [<span data-ttu-id="f70d6-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="f70d6-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="f70d6-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-197">Method</span></span> |
| [<span data-ttu-id="f70d6-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f70d6-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="f70d6-199">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-199">Method</span></span> |
| [<span data-ttu-id="f70d6-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="f70d6-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="f70d6-201">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-201">Method</span></span> |
| [<span data-ttu-id="f70d6-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="f70d6-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="f70d6-203">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-203">Method</span></span> |
| [<span data-ttu-id="f70d6-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="f70d6-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="f70d6-205">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="f70d6-206">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-206">Example</span></span>

<span data-ttu-id="f70d6-207">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="f70d6-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="f70d6-208">Members</span><span class="sxs-lookup"><span data-stu-id="f70d6-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-17"></a><span data-ttu-id="f70d6-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="f70d6-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

<span data-ttu-id="f70d6-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-212">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="f70d6-213">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f70d6-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="f70d6-214">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-214">Type</span></span>

*   <span data-ttu-id="f70d6-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="f70d6-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-216">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-216">Requirements</span></span>

|<span data-ttu-id="f70d6-217">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-217">Requirement</span></span>|<span data-ttu-id="f70d6-218">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-219">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-220">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-220">1.0</span></span>|
|[<span data-ttu-id="f70d6-221">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-222">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-223">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-224">読み取り</span><span class="sxs-lookup"><span data-stu-id="f70d6-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f70d6-225">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-225">Example</span></span>

<span data-ttu-id="f70d6-226">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="f70d6-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="f70d6-228">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="f70d6-229">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f70d6-229">Compose mode only.</span></span>

<span data-ttu-id="f70d6-230">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f70d6-230">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f70d6-231">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-231">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f70d6-232">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-232">Get 500 members maximum.</span></span>
- <span data-ttu-id="f70d6-233">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-233">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="f70d6-234">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-234">Type</span></span>

*   [<span data-ttu-id="f70d6-235">受信者</span><span class="sxs-lookup"><span data-stu-id="f70d6-235">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="f70d6-236">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-236">Requirements</span></span>

|<span data-ttu-id="f70d6-237">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-237">Requirement</span></span>|<span data-ttu-id="f70d6-238">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-239">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-240">1.1</span><span class="sxs-lookup"><span data-stu-id="f70d6-240">1.1</span></span>|
|[<span data-ttu-id="f70d6-241">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-242">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-243">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-244">作成</span><span class="sxs-lookup"><span data-stu-id="f70d6-244">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f70d6-245">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-245">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-17"></a><span data-ttu-id="f70d6-246">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-246">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span></span>

<span data-ttu-id="f70d6-247">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-247">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f70d6-248">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-248">Type</span></span>

*   [<span data-ttu-id="f70d6-249">Body</span><span class="sxs-lookup"><span data-stu-id="f70d6-249">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="f70d6-250">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-250">Requirements</span></span>

|<span data-ttu-id="f70d6-251">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-251">Requirement</span></span>|<span data-ttu-id="f70d6-252">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-253">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-254">1.1</span><span class="sxs-lookup"><span data-stu-id="f70d6-254">1.1</span></span>|
|[<span data-ttu-id="f70d6-255">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-256">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-257">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-258">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f70d6-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f70d6-259">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-259">Example</span></span>

<span data-ttu-id="f70d6-260">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-260">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="f70d6-261">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="f70d6-261">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="f70d6-262">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-262">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="f70d6-263">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-263">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="f70d6-264">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-264">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f70d6-265">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-265">Read mode</span></span>

<span data-ttu-id="f70d6-266">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-266">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="f70d6-267">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f70d6-267">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f70d6-268">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-268">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="f70d6-269">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-269">Compose mode</span></span>

<span data-ttu-id="f70d6-270">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-270">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="f70d6-271">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f70d6-271">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f70d6-272">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-272">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f70d6-273">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-273">Get 500 members maximum.</span></span>
- <span data-ttu-id="f70d6-274">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-274">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f70d6-275">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-275">Type</span></span>

*   <span data-ttu-id="f70d6-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-277">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-277">Requirements</span></span>

|<span data-ttu-id="f70d6-278">必要条件</span><span class="sxs-lookup"><span data-stu-id="f70d6-278">Requirement</span></span>|<span data-ttu-id="f70d6-279">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-280">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-281">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-281">1.0</span></span>|
|[<span data-ttu-id="f70d6-282">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-283">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-284">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-285">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f70d6-285">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="f70d6-286">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="f70d6-286">(nullable) conversationId: String</span></span>

<span data-ttu-id="f70d6-287">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-287">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="f70d6-p109">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="f70d6-p110">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="f70d6-292">Type</span><span class="sxs-lookup"><span data-stu-id="f70d6-292">Type</span></span>

*   <span data-ttu-id="f70d6-293">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-293">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-294">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-294">Requirements</span></span>

|<span data-ttu-id="f70d6-295">必要条件</span><span class="sxs-lookup"><span data-stu-id="f70d6-295">Requirement</span></span>|<span data-ttu-id="f70d6-296">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-297">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-297">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-298">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-298">1.0</span></span>|
|[<span data-ttu-id="f70d6-299">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-299">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-300">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-300">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-301">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-301">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-302">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f70d6-302">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f70d6-303">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-303">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="f70d6-304">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="f70d6-304">dateTimeCreated: Date</span></span>

<span data-ttu-id="f70d6-p111">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f70d6-307">種類</span><span class="sxs-lookup"><span data-stu-id="f70d6-307">Type</span></span>

*   <span data-ttu-id="f70d6-308">日付</span><span class="sxs-lookup"><span data-stu-id="f70d6-308">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-309">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-309">Requirements</span></span>

|<span data-ttu-id="f70d6-310">必要条件</span><span class="sxs-lookup"><span data-stu-id="f70d6-310">Requirement</span></span>|<span data-ttu-id="f70d6-311">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-312">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-313">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-313">1.0</span></span>|
|[<span data-ttu-id="f70d6-314">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-315">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-316">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-317">読み取り</span><span class="sxs-lookup"><span data-stu-id="f70d6-317">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f70d6-318">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-318">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="f70d6-319">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="f70d6-319">dateTimeModified: Date</span></span>

<span data-ttu-id="f70d6-p112">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-322">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-322">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="f70d6-323">種類</span><span class="sxs-lookup"><span data-stu-id="f70d6-323">Type</span></span>

*   <span data-ttu-id="f70d6-324">日付</span><span class="sxs-lookup"><span data-stu-id="f70d6-324">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-325">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-325">Requirements</span></span>

|<span data-ttu-id="f70d6-326">必要条件</span><span class="sxs-lookup"><span data-stu-id="f70d6-326">Requirement</span></span>|<span data-ttu-id="f70d6-327">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-328">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-329">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-329">1.0</span></span>|
|[<span data-ttu-id="f70d6-330">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-330">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-331">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-332">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-332">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-333">読み取り</span><span class="sxs-lookup"><span data-stu-id="f70d6-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f70d6-334">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-334">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="f70d6-335">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-335">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="f70d6-336">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-336">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="f70d6-p113">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f70d6-339">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-339">Read mode</span></span>

<span data-ttu-id="f70d6-340">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-340">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="f70d6-341">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-341">Compose mode</span></span>

<span data-ttu-id="f70d6-342">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-342">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="f70d6-343">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-343">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="f70d6-344">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-344">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="f70d6-345">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-345">Type</span></span>

*   <span data-ttu-id="f70d6-346">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-346">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-347">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-347">Requirements</span></span>

|<span data-ttu-id="f70d6-348">必要条件</span><span class="sxs-lookup"><span data-stu-id="f70d6-348">Requirement</span></span>|<span data-ttu-id="f70d6-349">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-349">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-350">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-350">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-351">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-351">1.0</span></span>|
|[<span data-ttu-id="f70d6-352">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-352">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-353">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-354">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-354">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-355">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f70d6-355">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17fromjavascriptapioutlookofficefromviewoutlook-js-17"></a><span data-ttu-id="f70d6-356">from: [emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[from](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-356">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

<span data-ttu-id="f70d6-357">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-357">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="f70d6-p114">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-360">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="f70d6-360">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f70d6-361">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-361">Read mode</span></span>

<span data-ttu-id="f70d6-362">プロパティ`from`は`EmailAddressDetails`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-362">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="f70d6-363">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-363">Compose mode</span></span>

<span data-ttu-id="f70d6-364">プロパティ`from`は、from `From`値を取得するメソッドを提供するオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-364">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f70d6-365">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-365">Type</span></span>

*   <span data-ttu-id="f70d6-366">[電子メールアドレス](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [の](/javascript/api/outlook/office.from?view=outlook-js-1.7)詳細</span><span class="sxs-lookup"><span data-stu-id="f70d6-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-367">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-367">Requirements</span></span>

|<span data-ttu-id="f70d6-368">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-368">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="f70d6-369">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-369">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-370">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-370">1.0</span></span>|<span data-ttu-id="f70d6-371">1.7</span><span class="sxs-lookup"><span data-stu-id="f70d6-371">1.7</span></span>|
|[<span data-ttu-id="f70d6-372">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-372">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-373">ReadItem</span></span>|<span data-ttu-id="f70d6-374">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-374">ReadWriteItem</span></span>|
|[<span data-ttu-id="f70d6-375">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-375">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-376">Read</span><span class="sxs-lookup"><span data-stu-id="f70d6-376">Read</span></span>|<span data-ttu-id="f70d6-377">Compose</span><span class="sxs-lookup"><span data-stu-id="f70d6-377">Compose</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="f70d6-378">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="f70d6-378">internetMessageId: String</span></span>

<span data-ttu-id="f70d6-p115">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f70d6-381">Type</span><span class="sxs-lookup"><span data-stu-id="f70d6-381">Type</span></span>

*   <span data-ttu-id="f70d6-382">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-383">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-383">Requirements</span></span>

|<span data-ttu-id="f70d6-384">必要条件</span><span class="sxs-lookup"><span data-stu-id="f70d6-384">Requirement</span></span>|<span data-ttu-id="f70d6-385">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-386">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-387">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-387">1.0</span></span>|
|[<span data-ttu-id="f70d6-388">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-389">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-390">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-391">読み取り</span><span class="sxs-lookup"><span data-stu-id="f70d6-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f70d6-392">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-392">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="f70d6-393">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="f70d6-393">itemClass: String</span></span>

<span data-ttu-id="f70d6-p116">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="f70d6-p117">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="f70d6-398">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-398">Type</span></span>|<span data-ttu-id="f70d6-399">説明</span><span class="sxs-lookup"><span data-stu-id="f70d6-399">Description</span></span>|<span data-ttu-id="f70d6-400">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="f70d6-400">item class</span></span>|
|---|---|---|
|<span data-ttu-id="f70d6-401">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="f70d6-401">Appointment items</span></span>|<span data-ttu-id="f70d6-402">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="f70d6-402">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="f70d6-403">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="f70d6-403">Message items</span></span>|<span data-ttu-id="f70d6-404">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-404">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="f70d6-405">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-405">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="f70d6-406">Type</span><span class="sxs-lookup"><span data-stu-id="f70d6-406">Type</span></span>

*   <span data-ttu-id="f70d6-407">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-408">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-408">Requirements</span></span>

|<span data-ttu-id="f70d6-409">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-409">Requirement</span></span>|<span data-ttu-id="f70d6-410">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-411">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-412">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-412">1.0</span></span>|
|[<span data-ttu-id="f70d6-413">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-414">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-415">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-416">読み取り</span><span class="sxs-lookup"><span data-stu-id="f70d6-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f70d6-417">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-417">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="f70d6-418">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="f70d6-418">(nullable) itemId: String</span></span>

<span data-ttu-id="f70d6-p118">現在のアイテムの [Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p118">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-421">`itemId` プロパティから返される識別子は、[Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) と同じです。</span><span class="sxs-lookup"><span data-stu-id="f70d6-421">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="f70d6-422">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-422">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="f70d6-423">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-423">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="f70d6-424">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f70d6-424">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="f70d6-p120">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p120">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="f70d6-427">Type</span><span class="sxs-lookup"><span data-stu-id="f70d6-427">Type</span></span>

*   <span data-ttu-id="f70d6-428">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-428">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-429">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-429">Requirements</span></span>

|<span data-ttu-id="f70d6-430">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-430">Requirement</span></span>|<span data-ttu-id="f70d6-431">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-431">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-432">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-432">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-433">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-433">1.0</span></span>|
|[<span data-ttu-id="f70d6-434">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-434">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-435">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-435">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-436">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-436">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-437">読み取り</span><span class="sxs-lookup"><span data-stu-id="f70d6-437">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f70d6-438">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-438">Example</span></span>

<span data-ttu-id="f70d6-p121">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p121">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-17"></a><span data-ttu-id="f70d6-441">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-441">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span></span>

<span data-ttu-id="f70d6-442">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-442">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="f70d6-443">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="f70d6-443">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f70d6-444">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-444">Type</span></span>

*   [<span data-ttu-id="f70d6-445">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="f70d6-445">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="f70d6-446">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-446">Requirements</span></span>

|<span data-ttu-id="f70d6-447">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-447">Requirement</span></span>|<span data-ttu-id="f70d6-448">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-448">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-449">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-449">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-450">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-450">1.0</span></span>|
|[<span data-ttu-id="f70d6-451">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-451">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-452">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-452">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-453">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-453">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-454">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f70d6-454">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f70d6-455">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-455">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-17"></a><span data-ttu-id="f70d6-456">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-456">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

<span data-ttu-id="f70d6-457">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-457">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f70d6-458">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-458">Read mode</span></span>

<span data-ttu-id="f70d6-459">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-459">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="f70d6-460">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-460">Compose mode</span></span>

<span data-ttu-id="f70d6-461">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-461">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f70d6-462">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-462">Type</span></span>

*   <span data-ttu-id="f70d6-463">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-463">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-464">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-464">Requirements</span></span>

|<span data-ttu-id="f70d6-465">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-465">Requirement</span></span>|<span data-ttu-id="f70d6-466">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-467">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-468">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-468">1.0</span></span>|
|[<span data-ttu-id="f70d6-469">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-470">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-471">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-472">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f70d6-472">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="f70d6-473">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="f70d6-473">normalizedSubject: String</span></span>

<span data-ttu-id="f70d6-p122">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p122">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="f70d6-p123">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p123">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="f70d6-478">Type</span><span class="sxs-lookup"><span data-stu-id="f70d6-478">Type</span></span>

*   <span data-ttu-id="f70d6-479">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-479">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-480">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-480">Requirements</span></span>

|<span data-ttu-id="f70d6-481">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-481">Requirement</span></span>|<span data-ttu-id="f70d6-482">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-482">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-483">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-483">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-484">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-484">1.0</span></span>|
|[<span data-ttu-id="f70d6-485">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-485">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-486">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-486">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-487">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-487">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-488">読み取り</span><span class="sxs-lookup"><span data-stu-id="f70d6-488">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f70d6-489">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-489">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-17"></a><span data-ttu-id="f70d6-490">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-490">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span></span>

<span data-ttu-id="f70d6-491">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-491">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f70d6-492">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-492">Type</span></span>

*   [<span data-ttu-id="f70d6-493">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="f70d6-493">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="f70d6-494">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-494">Requirements</span></span>

|<span data-ttu-id="f70d6-495">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-495">Requirement</span></span>|<span data-ttu-id="f70d6-496">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-496">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-497">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-497">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-498">1.3</span><span class="sxs-lookup"><span data-stu-id="f70d6-498">1.3</span></span>|
|[<span data-ttu-id="f70d6-499">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-499">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-500">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-500">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-501">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-501">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-502">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f70d6-502">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f70d6-503">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-503">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="f70d6-504">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-504">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="f70d6-505">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-505">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="f70d6-506">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-506">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f70d6-507">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-507">Read mode</span></span>

<span data-ttu-id="f70d6-508">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-508">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="f70d6-509">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f70d6-509">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f70d6-510">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-510">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="f70d6-511">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-511">Compose mode</span></span>

<span data-ttu-id="f70d6-512">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-512">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="f70d6-513">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f70d6-513">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f70d6-514">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-514">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f70d6-515">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-515">Get 500 members maximum.</span></span>
- <span data-ttu-id="f70d6-516">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-516">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f70d6-517">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-517">Type</span></span>

*   <span data-ttu-id="f70d6-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-519">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-519">Requirements</span></span>

|<span data-ttu-id="f70d6-520">必要条件</span><span class="sxs-lookup"><span data-stu-id="f70d6-520">Requirement</span></span>|<span data-ttu-id="f70d6-521">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-522">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-523">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-523">1.0</span></span>|
|[<span data-ttu-id="f70d6-524">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-525">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-526">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-527">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f70d6-527">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17organizerjavascriptapioutlookofficeorganizerviewoutlook-js-17"></a><span data-ttu-id="f70d6-528">開催者: [emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[開催者](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-528">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

<span data-ttu-id="f70d6-529">指定した会議の開催者の電子メールアドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-529">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f70d6-530">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-530">Read mode</span></span>

<span data-ttu-id="f70d6-531">プロパティ`organizer`は、会議の開催者を表す[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-531">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="f70d6-532">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-532">Compose mode</span></span>

<span data-ttu-id="f70d6-533">プロパティ`organizer`は、開催者の値を取得するためのメソッドを提供する[オーガナイザー](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-533">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="f70d6-534">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-534">Type</span></span>

*   <span data-ttu-id="f70d6-535">[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [開催者](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-535">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-536">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-536">Requirements</span></span>

|<span data-ttu-id="f70d6-537">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-537">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="f70d6-538">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-539">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-539">1.0</span></span>|<span data-ttu-id="f70d6-540">1.7</span><span class="sxs-lookup"><span data-stu-id="f70d6-540">1.7</span></span>|
|[<span data-ttu-id="f70d6-541">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-542">ReadItem</span></span>|<span data-ttu-id="f70d6-543">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-543">ReadWriteItem</span></span>|
|[<span data-ttu-id="f70d6-544">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-544">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-545">Read</span><span class="sxs-lookup"><span data-stu-id="f70d6-545">Read</span></span>|<span data-ttu-id="f70d6-546">Compose</span><span class="sxs-lookup"><span data-stu-id="f70d6-546">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-17"></a><span data-ttu-id="f70d6-547">(nullable) 定期的なスケジュール:[定期的](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)なアイテム</span><span class="sxs-lookup"><span data-stu-id="f70d6-547">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span></span>

<span data-ttu-id="f70d6-548">予定の定期的なパターンを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-548">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="f70d6-549">会議出席依頼の定期的なパターンを取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-549">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="f70d6-550">予定アイテムの読み取りおよび作成モード。</span><span class="sxs-lookup"><span data-stu-id="f70d6-550">Read and compose modes for appointment items.</span></span> <span data-ttu-id="f70d6-551">会議出席依頼アイテムの閲覧モード。</span><span class="sxs-lookup"><span data-stu-id="f70d6-551">Read mode for meeting request items.</span></span>

<span data-ttu-id="f70d6-552">この`recurrence`プロパティは、アイテムが series または series 内のインスタンスの場合、定期的な予定または会議出席依頼に対して[定期的](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-552">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="f70d6-553">`null`は、単一の予定および1つの予定の会議出席依頼に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-553">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="f70d6-554">`undefined`は、会議出席依頼ではないメッセージに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-554">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="f70d6-555">注: 会議出席依頼に`itemClass`は、IPM という値があります。出席依頼。</span><span class="sxs-lookup"><span data-stu-id="f70d6-555">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="f70d6-556">注: 定期的なオブジェクトが`null`の場合は、そのオブジェクトが単一の予定または1つの予定の会議出席依頼であり、データ系列の一部ではないことを示します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-556">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f70d6-557">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-557">Read mode</span></span>

<span data-ttu-id="f70d6-558">この`recurrence`プロパティは、定期的な予定を表す[定期的](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-558">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that represents the appointment recurrence.</span></span> <span data-ttu-id="f70d6-559">これは、予定および会議出席依頼に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-559">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="f70d6-560">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-560">Compose mode</span></span>

<span data-ttu-id="f70d6-561">この`recurrence`プロパティは、予定の繰り返しを管理するためのメソッドを提供する[定期的](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-561">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="f70d6-562">これは予定に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-562">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="f70d6-563">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-563">Type</span></span>

* [<span data-ttu-id="f70d6-564">繰り返さ</span><span class="sxs-lookup"><span data-stu-id="f70d6-564">Recurrence</span></span>](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)

|<span data-ttu-id="f70d6-565">必要条件</span><span class="sxs-lookup"><span data-stu-id="f70d6-565">Requirement</span></span>|<span data-ttu-id="f70d6-566">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-567">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-568">1.7</span><span class="sxs-lookup"><span data-stu-id="f70d6-568">1.7</span></span>|
|[<span data-ttu-id="f70d6-569">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-570">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-571">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-572">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f70d6-572">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="f70d6-573">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-573">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="f70d6-574">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-574">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="f70d6-575">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-575">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f70d6-576">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-576">Read mode</span></span>

<span data-ttu-id="f70d6-577">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-577">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="f70d6-578">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f70d6-578">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f70d6-579">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-579">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="f70d6-580">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-580">Compose mode</span></span>

<span data-ttu-id="f70d6-581">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-581">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="f70d6-582">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f70d6-582">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f70d6-583">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-583">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f70d6-584">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-584">Get 500 members maximum.</span></span>
- <span data-ttu-id="f70d6-585">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-585">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="f70d6-586">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-586">Type</span></span>

*   <span data-ttu-id="f70d6-587">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-587">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-588">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-588">Requirements</span></span>

|<span data-ttu-id="f70d6-589">必要条件</span><span class="sxs-lookup"><span data-stu-id="f70d6-589">Requirement</span></span>|<span data-ttu-id="f70d6-590">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-590">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-591">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-591">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-592">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-592">1.0</span></span>|
|[<span data-ttu-id="f70d6-593">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-593">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-594">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-594">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-595">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-595">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-596">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f70d6-596">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17"></a><span data-ttu-id="f70d6-597">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-597">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span></span>

<span data-ttu-id="f70d6-p134">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p134">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="f70d6-p135">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsfrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p135">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-602">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="f70d6-602">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f70d6-603">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-603">Type</span></span>

*   [<span data-ttu-id="f70d6-604">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f70d6-604">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="f70d6-605">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-605">Requirements</span></span>

|<span data-ttu-id="f70d6-606">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-606">Requirement</span></span>|<span data-ttu-id="f70d6-607">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-608">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-609">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-609">1.0</span></span>|
|[<span data-ttu-id="f70d6-610">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-610">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-611">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-612">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-613">読み取り</span><span class="sxs-lookup"><span data-stu-id="f70d6-613">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f70d6-614">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-614">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="f70d6-615">(nullable) 系列 Id: String</span><span class="sxs-lookup"><span data-stu-id="f70d6-615">(nullable) seriesId: String</span></span>

<span data-ttu-id="f70d6-616">インスタンスが属する系列の id を取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-616">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="f70d6-617">Web 上の Outlook およびデスクトップクライアントでは、 `seriesId`は、このアイテムが属する親 (シリーズ) アイテムの Exchange web サービス (EWS) ID を返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-617">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="f70d6-618">ただし、iOS と Android では、 `seriesId`は親アイテムの REST ID を返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-618">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-619">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="f70d6-619">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="f70d6-620">`seriesId`プロパティが OUTLOOK REST API で使用される outlook id と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-620">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="f70d6-621">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-621">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="f70d6-622">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f70d6-622">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="f70d6-623">この`seriesId`プロパティは`null` 、単一の予定、系列のアイテム、会議出席依頼などの親アイテムを持たないアイテムに`undefined`対して、会議出席依頼以外のアイテムに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-623">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="f70d6-624">Type</span><span class="sxs-lookup"><span data-stu-id="f70d6-624">Type</span></span>

* <span data-ttu-id="f70d6-625">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-625">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-626">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-626">Requirements</span></span>

|<span data-ttu-id="f70d6-627">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-627">Requirement</span></span>|<span data-ttu-id="f70d6-628">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-629">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-630">1.7</span><span class="sxs-lookup"><span data-stu-id="f70d6-630">1.7</span></span>|
|[<span data-ttu-id="f70d6-631">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-631">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-632">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-633">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-633">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-634">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f70d6-634">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f70d6-635">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-635">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="f70d6-636">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-636">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="f70d6-637">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-637">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="f70d6-p138">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p138">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f70d6-640">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-640">Read mode</span></span>

<span data-ttu-id="f70d6-641">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-641">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="f70d6-642">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-642">Compose mode</span></span>

<span data-ttu-id="f70d6-643">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-643">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="f70d6-644">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-644">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="f70d6-645">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-645">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="f70d6-646">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-646">Type</span></span>

*   <span data-ttu-id="f70d6-647">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-647">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-648">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-648">Requirements</span></span>

|<span data-ttu-id="f70d6-649">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-649">Requirement</span></span>|<span data-ttu-id="f70d6-650">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-650">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-651">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-651">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-652">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-652">1.0</span></span>|
|[<span data-ttu-id="f70d6-653">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-653">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-654">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-654">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-655">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-655">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-656">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f70d6-656">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-17"></a><span data-ttu-id="f70d6-657">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-657">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

<span data-ttu-id="f70d6-658">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-658">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="f70d6-659">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-659">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f70d6-660">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-660">Read mode</span></span>

<span data-ttu-id="f70d6-p139">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p139">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="f70d6-663">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="f70d6-663">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="f70d6-664">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-664">Compose mode</span></span>

<span data-ttu-id="f70d6-665">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-665">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="f70d6-666">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-666">Type</span></span>

*   <span data-ttu-id="f70d6-667">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-667">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-668">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-668">Requirements</span></span>

|<span data-ttu-id="f70d6-669">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-669">Requirement</span></span>|<span data-ttu-id="f70d6-670">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-670">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-671">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-671">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-672">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-672">1.0</span></span>|
|[<span data-ttu-id="f70d6-673">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-673">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-674">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-674">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-675">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-675">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-676">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f70d6-676">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="f70d6-677">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-677">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="f70d6-678">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-678">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="f70d6-679">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-679">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f70d6-680">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-680">Read mode</span></span>

<span data-ttu-id="f70d6-681">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-681">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="f70d6-682">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f70d6-682">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f70d6-683">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-683">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="f70d6-684">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f70d6-684">Compose mode</span></span>

<span data-ttu-id="f70d6-685">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-685">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="f70d6-686">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f70d6-686">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f70d6-687">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-687">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f70d6-688">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-688">Get 500 members maximum.</span></span>
- <span data-ttu-id="f70d6-689">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-689">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f70d6-690">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-690">Type</span></span>

*   <span data-ttu-id="f70d6-691">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-691">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-692">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-692">Requirements</span></span>

|<span data-ttu-id="f70d6-693">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-693">Requirement</span></span>|<span data-ttu-id="f70d6-694">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-695">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-696">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-696">1.0</span></span>|
|[<span data-ttu-id="f70d6-697">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-697">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-698">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-698">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-699">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-699">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-700">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f70d6-700">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="f70d6-701">メソッド</span><span class="sxs-lookup"><span data-stu-id="f70d6-701">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="f70d6-702">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f70d6-702">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f70d6-703">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-703">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="f70d6-704">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-704">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="f70d6-705">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-705">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f70d6-706">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f70d6-706">Parameters</span></span>
|<span data-ttu-id="f70d6-707">名前</span><span class="sxs-lookup"><span data-stu-id="f70d6-707">Name</span></span>|<span data-ttu-id="f70d6-708">種類</span><span class="sxs-lookup"><span data-stu-id="f70d6-708">Type</span></span>|<span data-ttu-id="f70d6-709">属性</span><span class="sxs-lookup"><span data-stu-id="f70d6-709">Attributes</span></span>|<span data-ttu-id="f70d6-710">説明</span><span class="sxs-lookup"><span data-stu-id="f70d6-710">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="f70d6-711">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-711">String</span></span>||<span data-ttu-id="f70d6-p143">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p143">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="f70d6-714">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-714">String</span></span>||<span data-ttu-id="f70d6-p144">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p144">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="f70d6-717">Object</span><span class="sxs-lookup"><span data-stu-id="f70d6-717">Object</span></span>|<span data-ttu-id="f70d6-718">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-718">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-719">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f70d6-719">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f70d6-720">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f70d6-720">Object</span></span>|<span data-ttu-id="f70d6-721">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-721">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-722">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-722">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="f70d6-723">Boolean</span><span class="sxs-lookup"><span data-stu-id="f70d6-723">Boolean</span></span>|<span data-ttu-id="f70d6-724">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-724">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-725">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-725">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="f70d6-726">function</span><span class="sxs-lookup"><span data-stu-id="f70d6-726">function</span></span>|<span data-ttu-id="f70d6-727">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-727">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-728">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-728">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f70d6-729">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-729">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f70d6-730">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-730">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f70d6-731">エラー</span><span class="sxs-lookup"><span data-stu-id="f70d6-731">Errors</span></span>

|<span data-ttu-id="f70d6-732">エラー コード</span><span class="sxs-lookup"><span data-stu-id="f70d6-732">Error code</span></span>|<span data-ttu-id="f70d6-733">説明</span><span class="sxs-lookup"><span data-stu-id="f70d6-733">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="f70d6-734">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="f70d6-734">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="f70d6-735">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="f70d6-735">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="f70d6-736">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-736">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f70d6-737">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-737">Requirements</span></span>

|<span data-ttu-id="f70d6-738">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-738">Requirement</span></span>|<span data-ttu-id="f70d6-739">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-739">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-740">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-740">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-741">1.1</span><span class="sxs-lookup"><span data-stu-id="f70d6-741">1.1</span></span>|
|[<span data-ttu-id="f70d6-742">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-742">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-743">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-743">ReadWriteItem</span></span>|
|[<span data-ttu-id="f70d6-744">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-744">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-745">作成</span><span class="sxs-lookup"><span data-stu-id="f70d6-745">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f70d6-746">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-746">Examples</span></span>

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

<span data-ttu-id="f70d6-747">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-747">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="f70d6-748">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f70d6-748">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="f70d6-749">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-749">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="f70d6-750">現在、サポートされて`Office.EventType.AppointmentTimeChanged`いる`Office.EventType.RecipientsChanged`イベントの種類は、、、です。`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="f70d6-750">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="f70d6-751">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f70d6-751">Parameters</span></span>

| <span data-ttu-id="f70d6-752">名前</span><span class="sxs-lookup"><span data-stu-id="f70d6-752">Name</span></span> | <span data-ttu-id="f70d6-753">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-753">Type</span></span> | <span data-ttu-id="f70d6-754">属性</span><span class="sxs-lookup"><span data-stu-id="f70d6-754">Attributes</span></span> | <span data-ttu-id="f70d6-755">説明</span><span class="sxs-lookup"><span data-stu-id="f70d6-755">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="f70d6-756">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="f70d6-756">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="f70d6-757">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="f70d6-757">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="f70d6-758">Function</span><span class="sxs-lookup"><span data-stu-id="f70d6-758">Function</span></span> || <span data-ttu-id="f70d6-p145">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p145">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="f70d6-762">Object</span><span class="sxs-lookup"><span data-stu-id="f70d6-762">Object</span></span> | <span data-ttu-id="f70d6-763">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-763">&lt;optional&gt;</span></span> | <span data-ttu-id="f70d6-764">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f70d6-764">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="f70d6-765">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f70d6-765">Object</span></span> | <span data-ttu-id="f70d6-766">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-766">&lt;optional&gt;</span></span> | <span data-ttu-id="f70d6-767">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-767">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="f70d6-768">function</span><span class="sxs-lookup"><span data-stu-id="f70d6-768">function</span></span>| <span data-ttu-id="f70d6-769">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-769">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-770">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-770">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f70d6-771">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-771">Requirements</span></span>

|<span data-ttu-id="f70d6-772">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-772">Requirement</span></span>| <span data-ttu-id="f70d6-773">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-773">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-774">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-774">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f70d6-775">1.7</span><span class="sxs-lookup"><span data-stu-id="f70d6-775">1.7</span></span> |
|[<span data-ttu-id="f70d6-776">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-776">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f70d6-777">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-777">ReadItem</span></span> |
|[<span data-ttu-id="f70d6-778">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-778">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f70d6-779">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f70d6-779">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="f70d6-780">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-780">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="f70d6-781">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f70d6-781">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f70d6-782">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-782">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="f70d6-p146">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p146">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="f70d6-786">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-786">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="f70d6-787">Office アドインを Outlook on the web で実行している場合、編集中のアイテム以外のアイテムに `addItemAttachmentAsync` メソッドでアイテムを添付できます。ただし、これはサポートされていないため、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-787">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f70d6-788">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f70d6-788">Parameters</span></span>

|<span data-ttu-id="f70d6-789">名前</span><span class="sxs-lookup"><span data-stu-id="f70d6-789">Name</span></span>|<span data-ttu-id="f70d6-790">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-790">Type</span></span>|<span data-ttu-id="f70d6-791">属性</span><span class="sxs-lookup"><span data-stu-id="f70d6-791">Attributes</span></span>|<span data-ttu-id="f70d6-792">説明</span><span class="sxs-lookup"><span data-stu-id="f70d6-792">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="f70d6-793">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-793">String</span></span>||<span data-ttu-id="f70d6-p147">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p147">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="f70d6-796">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-796">String</span></span>||<span data-ttu-id="f70d6-797">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="f70d6-797">The subject of the item to be attached.</span></span> <span data-ttu-id="f70d6-798">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="f70d6-798">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="f70d6-799">Object</span><span class="sxs-lookup"><span data-stu-id="f70d6-799">Object</span></span>|<span data-ttu-id="f70d6-800">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-800">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-801">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f70d6-801">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f70d6-802">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f70d6-802">Object</span></span>|<span data-ttu-id="f70d6-803">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-803">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-804">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-804">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f70d6-805">function</span><span class="sxs-lookup"><span data-stu-id="f70d6-805">function</span></span>|<span data-ttu-id="f70d6-806">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-806">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-807">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-807">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f70d6-808">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-808">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f70d6-809">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-809">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f70d6-810">エラー</span><span class="sxs-lookup"><span data-stu-id="f70d6-810">Errors</span></span>

|<span data-ttu-id="f70d6-811">エラー コード</span><span class="sxs-lookup"><span data-stu-id="f70d6-811">Error code</span></span>|<span data-ttu-id="f70d6-812">説明</span><span class="sxs-lookup"><span data-stu-id="f70d6-812">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="f70d6-813">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-813">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f70d6-814">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-814">Requirements</span></span>

|<span data-ttu-id="f70d6-815">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-815">Requirement</span></span>|<span data-ttu-id="f70d6-816">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-817">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-818">1.1</span><span class="sxs-lookup"><span data-stu-id="f70d6-818">1.1</span></span>|
|[<span data-ttu-id="f70d6-819">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-820">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-820">ReadWriteItem</span></span>|
|[<span data-ttu-id="f70d6-821">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-822">作成</span><span class="sxs-lookup"><span data-stu-id="f70d6-822">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f70d6-823">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-823">Example</span></span>

<span data-ttu-id="f70d6-824">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-824">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="f70d6-825">close()</span><span class="sxs-lookup"><span data-stu-id="f70d6-825">close()</span></span>

<span data-ttu-id="f70d6-826">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-826">Closes the current item that is being composed.</span></span>

<span data-ttu-id="f70d6-p149">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p149">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-829">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-829">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="f70d6-830">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-830">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-831">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-831">Requirements</span></span>

|<span data-ttu-id="f70d6-832">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-832">Requirement</span></span>|<span data-ttu-id="f70d6-833">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-833">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-834">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-834">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-835">1.3</span><span class="sxs-lookup"><span data-stu-id="f70d6-835">1.3</span></span>|
|[<span data-ttu-id="f70d6-836">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-836">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-837">制限あり</span><span class="sxs-lookup"><span data-stu-id="f70d6-837">Restricted</span></span>|
|[<span data-ttu-id="f70d6-838">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-838">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-839">新規作成</span><span class="sxs-lookup"><span data-stu-id="f70d6-839">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="f70d6-840">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="f70d6-840">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="f70d6-841">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-841">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-842">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-842">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f70d6-843">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-843">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f70d6-844">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="f70d6-844">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="f70d6-p150">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p150">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f70d6-848">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f70d6-848">Parameters</span></span>

|<span data-ttu-id="f70d6-849">名前</span><span class="sxs-lookup"><span data-stu-id="f70d6-849">Name</span></span>|<span data-ttu-id="f70d6-850">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-850">Type</span></span>|<span data-ttu-id="f70d6-851">属性</span><span class="sxs-lookup"><span data-stu-id="f70d6-851">Attributes</span></span>|<span data-ttu-id="f70d6-852">説明</span><span class="sxs-lookup"><span data-stu-id="f70d6-852">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="f70d6-853">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="f70d6-853">String &#124; Object</span></span>||<span data-ttu-id="f70d6-p151">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p151">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f70d6-856">**または**</span><span class="sxs-lookup"><span data-stu-id="f70d6-856">**OR**</span></span><br/><span data-ttu-id="f70d6-p152">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p152">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="f70d6-859">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-859">String</span></span>|<span data-ttu-id="f70d6-860">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-860">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-p153">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="f70d6-863">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-863">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="f70d6-864">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-864">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-865">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="f70d6-865">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="f70d6-866">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-866">String</span></span>||<span data-ttu-id="f70d6-p154">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p154">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="f70d6-869">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-869">String</span></span>||<span data-ttu-id="f70d6-870">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="f70d6-870">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="f70d6-871">文字列</span><span class="sxs-lookup"><span data-stu-id="f70d6-871">String</span></span>||<span data-ttu-id="f70d6-p155">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p155">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="f70d6-874">ブール値</span><span class="sxs-lookup"><span data-stu-id="f70d6-874">Boolean</span></span>||<span data-ttu-id="f70d6-p156">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p156">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="f70d6-877">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-877">String</span></span>||<span data-ttu-id="f70d6-p157">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p157">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="f70d6-881">function</span><span class="sxs-lookup"><span data-stu-id="f70d6-881">function</span></span>|<span data-ttu-id="f70d6-882">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-882">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-883">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-883">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f70d6-884">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-884">Requirements</span></span>

|<span data-ttu-id="f70d6-885">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-885">Requirement</span></span>|<span data-ttu-id="f70d6-886">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-886">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-887">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-887">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-888">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-888">1.0</span></span>|
|[<span data-ttu-id="f70d6-889">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-889">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-890">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-890">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-891">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-891">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-892">読み取り</span><span class="sxs-lookup"><span data-stu-id="f70d6-892">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f70d6-893">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-893">Examples</span></span>

<span data-ttu-id="f70d6-894">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-894">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="f70d6-895">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-895">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="f70d6-896">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-896">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f70d6-897">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-897">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f70d6-898">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-898">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f70d6-899">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-899">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="f70d6-900">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="f70d6-900">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="f70d6-901">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-901">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-902">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-902">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f70d6-903">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-903">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f70d6-904">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="f70d6-904">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="f70d6-p158">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p158">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f70d6-908">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f70d6-908">Parameters</span></span>

|<span data-ttu-id="f70d6-909">名前</span><span class="sxs-lookup"><span data-stu-id="f70d6-909">Name</span></span>|<span data-ttu-id="f70d6-910">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-910">Type</span></span>|<span data-ttu-id="f70d6-911">属性</span><span class="sxs-lookup"><span data-stu-id="f70d6-911">Attributes</span></span>|<span data-ttu-id="f70d6-912">説明</span><span class="sxs-lookup"><span data-stu-id="f70d6-912">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="f70d6-913">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="f70d6-913">String &#124; Object</span></span>||<span data-ttu-id="f70d6-p159">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p159">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f70d6-916">**または**</span><span class="sxs-lookup"><span data-stu-id="f70d6-916">**OR**</span></span><br/><span data-ttu-id="f70d6-p160">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p160">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="f70d6-919">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-919">String</span></span>|<span data-ttu-id="f70d6-920">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-920">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-p161">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p161">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="f70d6-923">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-923">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="f70d6-924">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-924">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-925">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="f70d6-925">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="f70d6-926">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-926">String</span></span>||<span data-ttu-id="f70d6-p162">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p162">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="f70d6-929">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-929">String</span></span>||<span data-ttu-id="f70d6-930">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="f70d6-930">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="f70d6-931">文字列</span><span class="sxs-lookup"><span data-stu-id="f70d6-931">String</span></span>||<span data-ttu-id="f70d6-p163">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p163">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="f70d6-934">ブール値</span><span class="sxs-lookup"><span data-stu-id="f70d6-934">Boolean</span></span>||<span data-ttu-id="f70d6-p164">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p164">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="f70d6-937">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-937">String</span></span>||<span data-ttu-id="f70d6-p165">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p165">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="f70d6-941">function</span><span class="sxs-lookup"><span data-stu-id="f70d6-941">function</span></span>|<span data-ttu-id="f70d6-942">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-942">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-943">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-943">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f70d6-944">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-944">Requirements</span></span>

|<span data-ttu-id="f70d6-945">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-945">Requirement</span></span>|<span data-ttu-id="f70d6-946">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-947">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-948">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-948">1.0</span></span>|
|[<span data-ttu-id="f70d6-949">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-949">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-950">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-950">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-951">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-951">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-952">読み取り</span><span class="sxs-lookup"><span data-stu-id="f70d6-952">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f70d6-953">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-953">Examples</span></span>

<span data-ttu-id="f70d6-954">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-954">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="f70d6-955">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-955">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="f70d6-956">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-956">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f70d6-957">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-957">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f70d6-958">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-958">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f70d6-959">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-959">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="f70d6-960">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="f70d6-960">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="f70d6-961">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-961">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-962">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-962">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-963">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-963">Requirements</span></span>

|<span data-ttu-id="f70d6-964">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-964">Requirement</span></span>|<span data-ttu-id="f70d6-965">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-966">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-966">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-967">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-967">1.0</span></span>|
|[<span data-ttu-id="f70d6-968">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-968">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-969">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-970">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-970">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-971">読み取り</span><span class="sxs-lookup"><span data-stu-id="f70d6-971">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f70d6-972">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f70d6-972">Returns:</span></span>

<span data-ttu-id="f70d6-973">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-973">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="f70d6-974">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-974">Example</span></span>

<span data-ttu-id="f70d6-975">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="f70d6-975">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="f70d6-976">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="f70d6-976">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="f70d6-977">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-977">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-978">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-978">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f70d6-979">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f70d6-979">Parameters</span></span>

|<span data-ttu-id="f70d6-980">名前</span><span class="sxs-lookup"><span data-stu-id="f70d6-980">Name</span></span>|<span data-ttu-id="f70d6-981">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-981">Type</span></span>|<span data-ttu-id="f70d6-982">説明</span><span class="sxs-lookup"><span data-stu-id="f70d6-982">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="f70d6-983">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="f70d6-983">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.7)|<span data-ttu-id="f70d6-984">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="f70d6-984">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f70d6-985">Requirements</span><span class="sxs-lookup"><span data-stu-id="f70d6-985">Requirements</span></span>

|<span data-ttu-id="f70d6-986">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-986">Requirement</span></span>|<span data-ttu-id="f70d6-987">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-987">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-988">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-988">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-989">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-989">1.0</span></span>|
|[<span data-ttu-id="f70d6-990">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-990">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-991">制限あり</span><span class="sxs-lookup"><span data-stu-id="f70d6-991">Restricted</span></span>|
|[<span data-ttu-id="f70d6-992">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-992">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-993">読み取り</span><span class="sxs-lookup"><span data-stu-id="f70d6-993">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f70d6-994">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f70d6-994">Returns:</span></span>

<span data-ttu-id="f70d6-995">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-995">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="f70d6-996">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-996">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="f70d6-997">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-997">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="f70d6-998">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="f70d6-998">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="f70d6-999">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="f70d6-999">Value of `entityType`</span></span>|<span data-ttu-id="f70d6-1000">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="f70d6-1000">Type of objects in returned array</span></span>|<span data-ttu-id="f70d6-1001">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-1001">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="f70d6-1002">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-1002">String</span></span>|<span data-ttu-id="f70d6-1003">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="f70d6-1003">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="f70d6-1004">連絡先</span><span class="sxs-lookup"><span data-stu-id="f70d6-1004">Contact</span></span>|<span data-ttu-id="f70d6-1005">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f70d6-1005">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="f70d6-1006">文字列</span><span class="sxs-lookup"><span data-stu-id="f70d6-1006">String</span></span>|<span data-ttu-id="f70d6-1007">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f70d6-1007">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="f70d6-1008">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="f70d6-1008">MeetingSuggestion</span></span>|<span data-ttu-id="f70d6-1009">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f70d6-1009">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="f70d6-1010">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="f70d6-1010">PhoneNumber</span></span>|<span data-ttu-id="f70d6-1011">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="f70d6-1011">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="f70d6-1012">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="f70d6-1012">TaskSuggestion</span></span>|<span data-ttu-id="f70d6-1013">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f70d6-1013">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="f70d6-1014">文字列</span><span class="sxs-lookup"><span data-stu-id="f70d6-1014">String</span></span>|<span data-ttu-id="f70d6-1015">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="f70d6-1015">**Restricted**</span></span>|

<span data-ttu-id="f70d6-1016">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="f70d6-1016">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

##### <a name="example"></a><span data-ttu-id="f70d6-1017">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-1017">Example</span></span>

<span data-ttu-id="f70d6-1018">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1018">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="f70d6-1019">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="f70d6-1019">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="f70d6-1020">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1020">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-1021">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1021">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f70d6-1022">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1022">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f70d6-1023">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f70d6-1023">Parameters</span></span>

|<span data-ttu-id="f70d6-1024">名前</span><span class="sxs-lookup"><span data-stu-id="f70d6-1024">Name</span></span>|<span data-ttu-id="f70d6-1025">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-1025">Type</span></span>|<span data-ttu-id="f70d6-1026">説明</span><span class="sxs-lookup"><span data-stu-id="f70d6-1026">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="f70d6-1027">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-1027">String</span></span>|<span data-ttu-id="f70d6-1028">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1028">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f70d6-1029">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1029">Requirements</span></span>

|<span data-ttu-id="f70d6-1030">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1030">Requirement</span></span>|<span data-ttu-id="f70d6-1031">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-1031">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-1032">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-1032">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-1033">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-1033">1.0</span></span>|
|[<span data-ttu-id="f70d6-1034">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-1034">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-1035">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-1035">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-1036">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-1036">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-1037">読み取り</span><span class="sxs-lookup"><span data-stu-id="f70d6-1037">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f70d6-1038">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f70d6-1038">Returns:</span></span>

<span data-ttu-id="f70d6-p167">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p167">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="f70d6-1041">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="f70d6-1041">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="f70d6-1042">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f70d6-1042">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="f70d6-1043">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1043">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-1044">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1044">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f70d6-p168">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p168">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f70d6-1048">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1048">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f70d6-1049">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1049">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f70d6-p169">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-1053">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1053">Requirements</span></span>

|<span data-ttu-id="f70d6-1054">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1054">Requirement</span></span>|<span data-ttu-id="f70d6-1055">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-1056">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-1057">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-1057">1.0</span></span>|
|[<span data-ttu-id="f70d6-1058">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-1058">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-1059">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-1059">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-1060">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-1060">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-1061">読み取り</span><span class="sxs-lookup"><span data-stu-id="f70d6-1061">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f70d6-1062">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f70d6-1062">Returns:</span></span>

<span data-ttu-id="f70d6-p170">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="f70d6-1065">型: Object</span><span class="sxs-lookup"><span data-stu-id="f70d6-1065">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="f70d6-1066">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-1066">Example</span></span>

<span data-ttu-id="f70d6-1067">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1067">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="f70d6-1068">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="f70d6-1068">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="f70d6-1069">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1069">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-1070">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1070">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f70d6-1071">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1071">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="f70d6-p171">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f70d6-1074">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f70d6-1074">Parameters</span></span>

|<span data-ttu-id="f70d6-1075">名前</span><span class="sxs-lookup"><span data-stu-id="f70d6-1075">Name</span></span>|<span data-ttu-id="f70d6-1076">種類</span><span class="sxs-lookup"><span data-stu-id="f70d6-1076">Type</span></span>|<span data-ttu-id="f70d6-1077">説明</span><span class="sxs-lookup"><span data-stu-id="f70d6-1077">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="f70d6-1078">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-1078">String</span></span>|<span data-ttu-id="f70d6-1079">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1079">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f70d6-1080">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1080">Requirements</span></span>

|<span data-ttu-id="f70d6-1081">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1081">Requirement</span></span>|<span data-ttu-id="f70d6-1082">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-1083">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-1083">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-1084">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-1084">1.0</span></span>|
|[<span data-ttu-id="f70d6-1085">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-1085">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-1086">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-1087">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-1087">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-1088">読み取り</span><span class="sxs-lookup"><span data-stu-id="f70d6-1088">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f70d6-1089">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f70d6-1089">Returns:</span></span>

<span data-ttu-id="f70d6-1090">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1090">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="f70d6-1091">型: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="f70d6-1091">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="f70d6-1092">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-1092">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="f70d6-1093">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="f70d6-1093">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="f70d6-1094">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1094">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="f70d6-1095">選択されていないが、カーソルが本文または件名にある場合、メソッドは選択されたデータに対して空の文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1095">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data.</span></span> <span data-ttu-id="f70d6-1096">本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1096">If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-1097">Outlook on the web で、テキストが選択されていないのにカーソルが本文内にある場合、メソッドでは文字列 "null" を返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1097">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="f70d6-1098">このような状況を確認するには、このセクションで後述する例を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1098">To check for this situation, see the example later in this section.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f70d6-1099">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f70d6-1099">Parameters</span></span>

|<span data-ttu-id="f70d6-1100">名前</span><span class="sxs-lookup"><span data-stu-id="f70d6-1100">Name</span></span>|<span data-ttu-id="f70d6-1101">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-1101">Type</span></span>|<span data-ttu-id="f70d6-1102">属性</span><span class="sxs-lookup"><span data-stu-id="f70d6-1102">Attributes</span></span>|<span data-ttu-id="f70d6-1103">説明</span><span class="sxs-lookup"><span data-stu-id="f70d6-1103">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="f70d6-1104">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f70d6-1104">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="f70d6-p174">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p174">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="f70d6-1108">Object</span><span class="sxs-lookup"><span data-stu-id="f70d6-1108">Object</span></span>|<span data-ttu-id="f70d6-1109">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-1109">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-1110">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1110">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f70d6-1111">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f70d6-1111">Object</span></span>|<span data-ttu-id="f70d6-1112">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-1112">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-1113">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1113">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f70d6-1114">function</span><span class="sxs-lookup"><span data-stu-id="f70d6-1114">function</span></span>||<span data-ttu-id="f70d6-1115">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1115">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f70d6-1116">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1116">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="f70d6-1117">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1117">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f70d6-1118">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1118">Requirements</span></span>

|<span data-ttu-id="f70d6-1119">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1119">Requirement</span></span>|<span data-ttu-id="f70d6-1120">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-1120">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-1121">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-1121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-1122">1.2</span><span class="sxs-lookup"><span data-stu-id="f70d6-1122">1.2</span></span>|
|[<span data-ttu-id="f70d6-1123">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-1123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-1124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-1124">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-1125">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-1125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-1126">作成</span><span class="sxs-lookup"><span data-stu-id="f70d6-1126">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="f70d6-1127">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f70d6-1127">Returns:</span></span>

<span data-ttu-id="f70d6-1128">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1128">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="f70d6-1129">型:String</span><span class="sxs-lookup"><span data-stu-id="f70d6-1129">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="f70d6-1130">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-1130">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  // Handle where Outlook on the web erroneously returns "null" instead of empty string.
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookWebApp'
      && asyncResult.value.endPosition === asyncResult.value.startPosition) {
    text = "";
  }

  console.log("Selected text in " + prop + ": " + text);
}
```

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="f70d6-1131">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="f70d6-1131">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="f70d6-1132">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1132">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="f70d6-1133">強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1133">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-1134">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1134">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-1135">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1135">Requirements</span></span>

|<span data-ttu-id="f70d6-1136">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1136">Requirement</span></span>|<span data-ttu-id="f70d6-1137">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-1138">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-1139">1.6</span><span class="sxs-lookup"><span data-stu-id="f70d6-1139">1.6</span></span>|
|[<span data-ttu-id="f70d6-1140">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-1141">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-1141">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-1142">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-1143">読み取り</span><span class="sxs-lookup"><span data-stu-id="f70d6-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f70d6-1144">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f70d6-1144">Returns:</span></span>

<span data-ttu-id="f70d6-1145">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="f70d6-1145">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="f70d6-1146">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-1146">Example</span></span>

<span data-ttu-id="f70d6-1147">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1147">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="f70d6-1148">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f70d6-1148">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="f70d6-p177">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p177">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-1151">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1151">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f70d6-p178">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p178">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f70d6-1155">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1155">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f70d6-1156">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1156">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f70d6-p179">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70d6-1160">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1160">Requirements</span></span>

|<span data-ttu-id="f70d6-1161">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1161">Requirement</span></span>|<span data-ttu-id="f70d6-1162">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-1162">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-1163">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-1163">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-1164">1.6</span><span class="sxs-lookup"><span data-stu-id="f70d6-1164">1.6</span></span>|
|[<span data-ttu-id="f70d6-1165">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-1165">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-1166">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-1166">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-1167">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-1167">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-1168">読み取り</span><span class="sxs-lookup"><span data-stu-id="f70d6-1168">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f70d6-1169">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f70d6-1169">Returns:</span></span>

<span data-ttu-id="f70d6-p180">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="f70d6-1172">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-1172">Example</span></span>

<span data-ttu-id="f70d6-1173">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1173">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="f70d6-1174">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f70d6-1174">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="f70d6-1175">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1175">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="f70d6-p181">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p181">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f70d6-1179">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f70d6-1179">Parameters</span></span>

|<span data-ttu-id="f70d6-1180">名前</span><span class="sxs-lookup"><span data-stu-id="f70d6-1180">Name</span></span>|<span data-ttu-id="f70d6-1181">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-1181">Type</span></span>|<span data-ttu-id="f70d6-1182">属性</span><span class="sxs-lookup"><span data-stu-id="f70d6-1182">Attributes</span></span>|<span data-ttu-id="f70d6-1183">説明</span><span class="sxs-lookup"><span data-stu-id="f70d6-1183">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="f70d6-1184">function</span><span class="sxs-lookup"><span data-stu-id="f70d6-1184">function</span></span>||<span data-ttu-id="f70d6-1185">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1185">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f70d6-1186">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1186">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="f70d6-1187">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1187">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="f70d6-1188">Object</span><span class="sxs-lookup"><span data-stu-id="f70d6-1188">Object</span></span>|<span data-ttu-id="f70d6-1189">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-1189">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-1190">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1190">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="f70d6-1191">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1191">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f70d6-1192">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1192">Requirements</span></span>

|<span data-ttu-id="f70d6-1193">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1193">Requirement</span></span>|<span data-ttu-id="f70d6-1194">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-1194">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-1195">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-1195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-1196">1.0</span><span class="sxs-lookup"><span data-stu-id="f70d6-1196">1.0</span></span>|
|[<span data-ttu-id="f70d6-1197">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-1197">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-1198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-1198">ReadItem</span></span>|
|[<span data-ttu-id="f70d6-1199">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-1199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-1200">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f70d6-1200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f70d6-1201">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-1201">Example</span></span>

<span data-ttu-id="f70d6-p184">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p184">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="f70d6-1205">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f70d6-1205">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="f70d6-1206">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1206">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="f70d6-1207">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1207">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="f70d6-1208">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1208">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="f70d6-1209">Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1209">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="f70d6-1210">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1210">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f70d6-1211">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f70d6-1211">Parameters</span></span>

|<span data-ttu-id="f70d6-1212">名前</span><span class="sxs-lookup"><span data-stu-id="f70d6-1212">Name</span></span>|<span data-ttu-id="f70d6-1213">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-1213">Type</span></span>|<span data-ttu-id="f70d6-1214">属性</span><span class="sxs-lookup"><span data-stu-id="f70d6-1214">Attributes</span></span>|<span data-ttu-id="f70d6-1215">説明</span><span class="sxs-lookup"><span data-stu-id="f70d6-1215">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="f70d6-1216">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-1216">String</span></span>||<span data-ttu-id="f70d6-1217">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1217">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="f70d6-1218">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f70d6-1218">Object</span></span>|<span data-ttu-id="f70d6-1219">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-1219">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-1220">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1220">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f70d6-1221">Object</span><span class="sxs-lookup"><span data-stu-id="f70d6-1221">Object</span></span>|<span data-ttu-id="f70d6-1222">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-1222">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-1223">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1223">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f70d6-1224">function</span><span class="sxs-lookup"><span data-stu-id="f70d6-1224">function</span></span>|<span data-ttu-id="f70d6-1225">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-1225">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-1226">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1226">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f70d6-1227">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1227">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f70d6-1228">エラー</span><span class="sxs-lookup"><span data-stu-id="f70d6-1228">Errors</span></span>

|<span data-ttu-id="f70d6-1229">エラー コード</span><span class="sxs-lookup"><span data-stu-id="f70d6-1229">Error code</span></span>|<span data-ttu-id="f70d6-1230">説明</span><span class="sxs-lookup"><span data-stu-id="f70d6-1230">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="f70d6-1231">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1231">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f70d6-1232">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1232">Requirements</span></span>

|<span data-ttu-id="f70d6-1233">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1233">Requirement</span></span>|<span data-ttu-id="f70d6-1234">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-1234">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-1235">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-1235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-1236">1.1</span><span class="sxs-lookup"><span data-stu-id="f70d6-1236">1.1</span></span>|
|[<span data-ttu-id="f70d6-1237">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-1237">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-1238">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-1238">ReadWriteItem</span></span>|
|[<span data-ttu-id="f70d6-1239">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-1239">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-1240">作成</span><span class="sxs-lookup"><span data-stu-id="f70d6-1240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f70d6-1241">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-1241">Example</span></span>

<span data-ttu-id="f70d6-1242">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1242">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="f70d6-1243">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f70d6-1243">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="f70d6-1244">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1244">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="f70d6-1245">現在、サポートされて`Office.EventType.AppointmentTimeChanged`いる`Office.EventType.RecipientsChanged`イベントの種類は、、、です。`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="f70d6-1245">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="f70d6-1246">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f70d6-1246">Parameters</span></span>

| <span data-ttu-id="f70d6-1247">名前</span><span class="sxs-lookup"><span data-stu-id="f70d6-1247">Name</span></span> | <span data-ttu-id="f70d6-1248">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-1248">Type</span></span> | <span data-ttu-id="f70d6-1249">属性</span><span class="sxs-lookup"><span data-stu-id="f70d6-1249">Attributes</span></span> | <span data-ttu-id="f70d6-1250">説明</span><span class="sxs-lookup"><span data-stu-id="f70d6-1250">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="f70d6-1251">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="f70d6-1251">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="f70d6-1252">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1252">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="f70d6-1253">Object</span><span class="sxs-lookup"><span data-stu-id="f70d6-1253">Object</span></span> | <span data-ttu-id="f70d6-1254">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-1254">&lt;optional&gt;</span></span> | <span data-ttu-id="f70d6-1255">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1255">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="f70d6-1256">Object</span><span class="sxs-lookup"><span data-stu-id="f70d6-1256">Object</span></span> | <span data-ttu-id="f70d6-1257">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-1257">&lt;optional&gt;</span></span> | <span data-ttu-id="f70d6-1258">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1258">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="f70d6-1259">function</span><span class="sxs-lookup"><span data-stu-id="f70d6-1259">function</span></span>| <span data-ttu-id="f70d6-1260">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-1260">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-1261">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1261">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f70d6-1262">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1262">Requirements</span></span>

|<span data-ttu-id="f70d6-1263">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1263">Requirement</span></span>| <span data-ttu-id="f70d6-1264">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-1264">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-1265">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-1265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f70d6-1266">1.7</span><span class="sxs-lookup"><span data-stu-id="f70d6-1266">1.7</span></span> |
|[<span data-ttu-id="f70d6-1267">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-1267">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f70d6-1268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-1268">ReadItem</span></span> |
|[<span data-ttu-id="f70d6-1269">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-1269">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f70d6-1270">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f70d6-1270">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="f70d6-1271">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-1271">Example</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="f70d6-1272">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="f70d6-1272">saveAsync([options], callback)</span></span>

<span data-ttu-id="f70d6-1273">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1273">Asynchronously saves an item.</span></span>

<span data-ttu-id="f70d6-1274">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1274">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="f70d6-1275">Outlook on the web またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1275">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="f70d6-1276">キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1276">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-1277">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1277">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="f70d6-1278">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1278">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="f70d6-p188">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p188">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="f70d6-1282">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1282">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="f70d6-1283">Outlook on Mac では、会議の保存はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1283">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="f70d6-1284">`saveAsync` メソッドは、作成モードの会議から呼び出されると失敗します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1284">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="f70d6-1285">回避策については、「[Office JS API を使用して Outlook for Mac で会議を下書きとして保存できない](https://support.microsoft.com/help/4505745)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1285">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="f70d6-1286">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1286">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f70d6-1287">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f70d6-1287">Parameters</span></span>

|<span data-ttu-id="f70d6-1288">名前</span><span class="sxs-lookup"><span data-stu-id="f70d6-1288">Name</span></span>|<span data-ttu-id="f70d6-1289">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-1289">Type</span></span>|<span data-ttu-id="f70d6-1290">属性</span><span class="sxs-lookup"><span data-stu-id="f70d6-1290">Attributes</span></span>|<span data-ttu-id="f70d6-1291">説明</span><span class="sxs-lookup"><span data-stu-id="f70d6-1291">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="f70d6-1292">Object</span><span class="sxs-lookup"><span data-stu-id="f70d6-1292">Object</span></span>|<span data-ttu-id="f70d6-1293">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-1293">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-1294">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1294">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f70d6-1295">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f70d6-1295">Object</span></span>|<span data-ttu-id="f70d6-1296">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-1296">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-1297">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1297">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f70d6-1298">関数</span><span class="sxs-lookup"><span data-stu-id="f70d6-1298">function</span></span>||<span data-ttu-id="f70d6-1299">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1299">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f70d6-1300">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1300">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f70d6-1301">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1301">Requirements</span></span>

|<span data-ttu-id="f70d6-1302">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1302">Requirement</span></span>|<span data-ttu-id="f70d6-1303">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-1303">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-1304">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-1304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-1305">1.3</span><span class="sxs-lookup"><span data-stu-id="f70d6-1305">1.3</span></span>|
|[<span data-ttu-id="f70d6-1306">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-1306">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-1307">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-1307">ReadWriteItem</span></span>|
|[<span data-ttu-id="f70d6-1308">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-1308">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-1309">作成</span><span class="sxs-lookup"><span data-stu-id="f70d6-1309">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f70d6-1310">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-1310">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="f70d6-p190">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p190">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="f70d6-1313">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="f70d6-1313">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="f70d6-1314">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1314">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="f70d6-p191">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p191">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f70d6-1318">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f70d6-1318">Parameters</span></span>

|<span data-ttu-id="f70d6-1319">名前</span><span class="sxs-lookup"><span data-stu-id="f70d6-1319">Name</span></span>|<span data-ttu-id="f70d6-1320">型</span><span class="sxs-lookup"><span data-stu-id="f70d6-1320">Type</span></span>|<span data-ttu-id="f70d6-1321">属性</span><span class="sxs-lookup"><span data-stu-id="f70d6-1321">Attributes</span></span>|<span data-ttu-id="f70d6-1322">説明</span><span class="sxs-lookup"><span data-stu-id="f70d6-1322">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="f70d6-1323">String</span><span class="sxs-lookup"><span data-stu-id="f70d6-1323">String</span></span>||<span data-ttu-id="f70d6-p192">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-p192">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="f70d6-1327">Object</span><span class="sxs-lookup"><span data-stu-id="f70d6-1327">Object</span></span>|<span data-ttu-id="f70d6-1328">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-1328">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-1329">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1329">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f70d6-1330">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f70d6-1330">Object</span></span>|<span data-ttu-id="f70d6-1331">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-1331">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-1332">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1332">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="f70d6-1333">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f70d6-1333">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="f70d6-1334">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f70d6-1334">&lt;optional&gt;</span></span>|<span data-ttu-id="f70d6-1335">`text` の場合、Outlook on the web とデスクトップ クライアントでは現在のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1335">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="f70d6-1336">フィールドが HTML エディターの場合、データが HTML の場合でもテキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1336">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="f70d6-1337">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Outlook on the web では現在のスタイルが適用され、Outlook デスクトップ クライアントでは既定のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1337">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="f70d6-1338">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1338">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="f70d6-1339">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1339">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="f70d6-1340">function</span><span class="sxs-lookup"><span data-stu-id="f70d6-1340">function</span></span>||<span data-ttu-id="f70d6-1341">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f70d6-1341">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f70d6-1342">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1342">Requirements</span></span>

|<span data-ttu-id="f70d6-1343">要件</span><span class="sxs-lookup"><span data-stu-id="f70d6-1343">Requirement</span></span>|<span data-ttu-id="f70d6-1344">値</span><span class="sxs-lookup"><span data-stu-id="f70d6-1344">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70d6-1345">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f70d6-1345">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f70d6-1346">1.2</span><span class="sxs-lookup"><span data-stu-id="f70d6-1346">1.2</span></span>|
|[<span data-ttu-id="f70d6-1347">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f70d6-1347">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f70d6-1348">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f70d6-1348">ReadWriteItem</span></span>|
|[<span data-ttu-id="f70d6-1349">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f70d6-1349">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f70d6-1350">作成</span><span class="sxs-lookup"><span data-stu-id="f70d6-1350">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f70d6-1351">例</span><span class="sxs-lookup"><span data-stu-id="f70d6-1351">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
