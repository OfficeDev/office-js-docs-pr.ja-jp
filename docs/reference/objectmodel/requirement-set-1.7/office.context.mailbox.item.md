---
title: Office. メールボックス-要件セット1.7
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 94fcb090fb7e676ed0c30caef342e1d6cfc6e9f9
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/14/2019
ms.locfileid: "30600292"
---
# <a name="item"></a><span data-ttu-id="1e558-102">item</span><span class="sxs-lookup"><span data-stu-id="1e558-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="1e558-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="1e558-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="1e558-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="1e558-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-106">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-106">Requirements</span></span>

|<span data-ttu-id="1e558-107">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-107">Requirement</span></span>|<span data-ttu-id="1e558-108">値</span><span class="sxs-lookup"><span data-stu-id="1e558-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-110">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-110">1.0</span></span>|
|[<span data-ttu-id="1e558-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="1e558-112">Restricted</span></span>|
|[<span data-ttu-id="1e558-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e558-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1e558-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="1e558-115">Members and methods</span></span>

| <span data-ttu-id="1e558-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="1e558-116">Member</span></span> | <span data-ttu-id="1e558-117">種類</span><span class="sxs-lookup"><span data-stu-id="1e558-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1e558-118">attachments</span><span class="sxs-lookup"><span data-stu-id="1e558-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="1e558-119">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-119">Member</span></span> |
| [<span data-ttu-id="1e558-120">bcc</span><span class="sxs-lookup"><span data-stu-id="1e558-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="1e558-121">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-121">Member</span></span> |
| [<span data-ttu-id="1e558-122">body</span><span class="sxs-lookup"><span data-stu-id="1e558-122">body</span></span>](#body-body) | <span data-ttu-id="1e558-123">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-123">Member</span></span> |
| [<span data-ttu-id="1e558-124">cc</span><span class="sxs-lookup"><span data-stu-id="1e558-124">cc</span></span>](#cc-arrayemailaddressdetails) | <span data-ttu-id="1e558-125">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-125">Member</span></span> |
| [<span data-ttu-id="1e558-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="1e558-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="1e558-127">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-127">Member</span></span> |
| [<span data-ttu-id="1e558-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="1e558-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="1e558-129">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-129">Member</span></span> |
| [<span data-ttu-id="1e558-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="1e558-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="1e558-131">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-131">Member</span></span> |
| [<span data-ttu-id="1e558-132">end</span><span class="sxs-lookup"><span data-stu-id="1e558-132">end</span></span>](#end-datetime) | <span data-ttu-id="1e558-133">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-133">Member</span></span> |
| [<span data-ttu-id="1e558-134">from</span><span class="sxs-lookup"><span data-stu-id="1e558-134">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="1e558-135">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-135">Member</span></span> |
| [<span data-ttu-id="1e558-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="1e558-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="1e558-137">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-137">Member</span></span> |
| [<span data-ttu-id="1e558-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="1e558-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="1e558-139">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-139">Member</span></span> |
| [<span data-ttu-id="1e558-140">itemId</span><span class="sxs-lookup"><span data-stu-id="1e558-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="1e558-141">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-141">Member</span></span> |
| [<span data-ttu-id="1e558-142">itemType</span><span class="sxs-lookup"><span data-stu-id="1e558-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="1e558-143">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-143">Member</span></span> |
| [<span data-ttu-id="1e558-144">location</span><span class="sxs-lookup"><span data-stu-id="1e558-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="1e558-145">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-145">Member</span></span> |
| [<span data-ttu-id="1e558-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="1e558-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="1e558-147">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-147">Member</span></span> |
| [<span data-ttu-id="1e558-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="1e558-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="1e558-149">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-149">Member</span></span> |
| [<span data-ttu-id="1e558-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="1e558-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetails) | <span data-ttu-id="1e558-151">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-151">Member</span></span> |
| [<span data-ttu-id="1e558-152">organizer</span><span class="sxs-lookup"><span data-stu-id="1e558-152">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="1e558-153">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-153">Member</span></span> |
| [<span data-ttu-id="1e558-154">繰り返さ</span><span class="sxs-lookup"><span data-stu-id="1e558-154">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="1e558-155">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-155">Member</span></span> |
| [<span data-ttu-id="1e558-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="1e558-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetails) | <span data-ttu-id="1e558-157">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-157">Member</span></span> |
| [<span data-ttu-id="1e558-158">sender</span><span class="sxs-lookup"><span data-stu-id="1e558-158">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="1e558-159">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-159">Member</span></span> |
| [<span data-ttu-id="1e558-160">系列 id</span><span class="sxs-lookup"><span data-stu-id="1e558-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="1e558-161">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-161">Member</span></span> |
| [<span data-ttu-id="1e558-162">start</span><span class="sxs-lookup"><span data-stu-id="1e558-162">start</span></span>](#start-datetime) | <span data-ttu-id="1e558-163">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-163">Member</span></span> |
| [<span data-ttu-id="1e558-164">subject</span><span class="sxs-lookup"><span data-stu-id="1e558-164">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="1e558-165">Member</span><span class="sxs-lookup"><span data-stu-id="1e558-165">Member</span></span> |
| [<span data-ttu-id="1e558-166">to</span><span class="sxs-lookup"><span data-stu-id="1e558-166">to</span></span>](#to-arrayemailaddressdetails) | <span data-ttu-id="1e558-167">メンバー</span><span class="sxs-lookup"><span data-stu-id="1e558-167">Member</span></span> |
| [<span data-ttu-id="1e558-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1e558-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="1e558-169">Method</span><span class="sxs-lookup"><span data-stu-id="1e558-169">Method</span></span> |
| [<span data-ttu-id="1e558-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="1e558-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="1e558-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="1e558-171">Method</span></span> |
| [<span data-ttu-id="1e558-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1e558-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="1e558-173">Method</span><span class="sxs-lookup"><span data-stu-id="1e558-173">Method</span></span> |
| [<span data-ttu-id="1e558-174">close</span><span class="sxs-lookup"><span data-stu-id="1e558-174">close</span></span>](#close) | <span data-ttu-id="1e558-175">Method</span><span class="sxs-lookup"><span data-stu-id="1e558-175">Method</span></span> |
| [<span data-ttu-id="1e558-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="1e558-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="1e558-177">Method</span><span class="sxs-lookup"><span data-stu-id="1e558-177">Method</span></span> |
| [<span data-ttu-id="1e558-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="1e558-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="1e558-179">Method</span><span class="sxs-lookup"><span data-stu-id="1e558-179">Method</span></span> |
| [<span data-ttu-id="1e558-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="1e558-180">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="1e558-181">Method</span><span class="sxs-lookup"><span data-stu-id="1e558-181">Method</span></span> |
| [<span data-ttu-id="1e558-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="1e558-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontact) | <span data-ttu-id="1e558-183">Method</span><span class="sxs-lookup"><span data-stu-id="1e558-183">Method</span></span> |
| [<span data-ttu-id="1e558-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="1e558-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontact) | <span data-ttu-id="1e558-185">Method</span><span class="sxs-lookup"><span data-stu-id="1e558-185">Method</span></span> |
| [<span data-ttu-id="1e558-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="1e558-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="1e558-187">Method</span><span class="sxs-lookup"><span data-stu-id="1e558-187">Method</span></span> |
| [<span data-ttu-id="1e558-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="1e558-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="1e558-189">Method</span><span class="sxs-lookup"><span data-stu-id="1e558-189">Method</span></span> |
| [<span data-ttu-id="1e558-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="1e558-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="1e558-191">Method</span><span class="sxs-lookup"><span data-stu-id="1e558-191">Method</span></span> |
| [<span data-ttu-id="1e558-192">office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="1e558-192">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="1e558-193">Method</span><span class="sxs-lookup"><span data-stu-id="1e558-193">Method</span></span> |
| [<span data-ttu-id="1e558-194">office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="1e558-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="1e558-195">Method</span><span class="sxs-lookup"><span data-stu-id="1e558-195">Method</span></span> |
| [<span data-ttu-id="1e558-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="1e558-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="1e558-197">Method</span><span class="sxs-lookup"><span data-stu-id="1e558-197">Method</span></span> |
| [<span data-ttu-id="1e558-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1e558-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="1e558-199">メソッド</span><span class="sxs-lookup"><span data-stu-id="1e558-199">Method</span></span> |
| [<span data-ttu-id="1e558-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="1e558-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="1e558-201">メソッド</span><span class="sxs-lookup"><span data-stu-id="1e558-201">Method</span></span> |
| [<span data-ttu-id="1e558-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="1e558-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="1e558-203">Method</span><span class="sxs-lookup"><span data-stu-id="1e558-203">Method</span></span> |
| [<span data-ttu-id="1e558-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="1e558-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="1e558-205">Method</span><span class="sxs-lookup"><span data-stu-id="1e558-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="1e558-206">例</span><span class="sxs-lookup"><span data-stu-id="1e558-206">Example</span></span>

<span data-ttu-id="1e558-207">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1e558-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="1e558-208">メンバー</span><span class="sxs-lookup"><span data-stu-id="1e558-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="1e558-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="1e558-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="1e558-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1e558-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1e558-212">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="1e558-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="1e558-213">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1e558-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="1e558-214">型</span><span class="sxs-lookup"><span data-stu-id="1e558-214">Type</span></span>

*   <span data-ttu-id="1e558-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="1e558-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-216">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-216">Requirements</span></span>

|<span data-ttu-id="1e558-217">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-217">Requirement</span></span>|<span data-ttu-id="1e558-218">値</span><span class="sxs-lookup"><span data-stu-id="1e558-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-219">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-220">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-220">1.0</span></span>|
|[<span data-ttu-id="1e558-221">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-221">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-222">ReadItem</span></span>|
|[<span data-ttu-id="1e558-223">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-223">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-224">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e558-225">例</span><span class="sxs-lookup"><span data-stu-id="1e558-225">Example</span></span>

<span data-ttu-id="1e558-226">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="1e558-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="1e558-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1e558-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="1e558-228">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="1e558-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="1e558-229">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1e558-229">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1e558-230">型</span><span class="sxs-lookup"><span data-stu-id="1e558-230">Type</span></span>

*   [<span data-ttu-id="1e558-231">受信者</span><span class="sxs-lookup"><span data-stu-id="1e558-231">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="1e558-232">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-232">Requirements</span></span>

|<span data-ttu-id="1e558-233">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-233">Requirement</span></span>|<span data-ttu-id="1e558-234">値</span><span class="sxs-lookup"><span data-stu-id="1e558-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-235">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-236">1.1</span><span class="sxs-lookup"><span data-stu-id="1e558-236">1.1</span></span>|
|[<span data-ttu-id="1e558-237">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-237">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-238">ReadItem</span></span>|
|[<span data-ttu-id="1e558-239">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-239">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-240">作成</span><span class="sxs-lookup"><span data-stu-id="1e558-240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1e558-241">例</span><span class="sxs-lookup"><span data-stu-id="1e558-241">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="1e558-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="1e558-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="1e558-243">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="1e558-243">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1e558-244">型</span><span class="sxs-lookup"><span data-stu-id="1e558-244">Type</span></span>

*   [<span data-ttu-id="1e558-245">Body</span><span class="sxs-lookup"><span data-stu-id="1e558-245">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="1e558-246">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-246">Requirements</span></span>

|<span data-ttu-id="1e558-247">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-247">Requirement</span></span>|<span data-ttu-id="1e558-248">値</span><span class="sxs-lookup"><span data-stu-id="1e558-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-249">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-250">1.1</span><span class="sxs-lookup"><span data-stu-id="1e558-250">1.1</span></span>|
|[<span data-ttu-id="1e558-251">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-252">ReadItem</span></span>|
|[<span data-ttu-id="1e558-253">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-254">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e558-254">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e558-255">例</span><span class="sxs-lookup"><span data-stu-id="1e558-255">Example</span></span>

<span data-ttu-id="1e558-256">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="1e558-256">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="1e558-257">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="1e558-257">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="1e558-258">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1e558-258">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="1e558-259">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1e558-259">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="1e558-260">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="1e558-260">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1e558-261">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1e558-261">Read mode</span></span>

<span data-ttu-id="1e558-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="1e558-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="1e558-264">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1e558-264">Compose mode</span></span>

<span data-ttu-id="1e558-265">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-265">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1e558-266">型</span><span class="sxs-lookup"><span data-stu-id="1e558-266">Type</span></span>

*   <span data-ttu-id="1e558-267">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1e558-267">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-268">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-268">Requirements</span></span>

|<span data-ttu-id="1e558-269">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-269">Requirement</span></span>|<span data-ttu-id="1e558-270">値</span><span class="sxs-lookup"><span data-stu-id="1e558-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-271">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-272">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-272">1.0</span></span>|
|[<span data-ttu-id="1e558-273">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-274">ReadItem</span></span>|
|[<span data-ttu-id="1e558-275">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-276">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e558-276">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="1e558-277">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="1e558-277">(nullable) conversationId :String</span></span>

<span data-ttu-id="1e558-278">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="1e558-278">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="1e558-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="1e558-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="1e558-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="1e558-283">Type</span><span class="sxs-lookup"><span data-stu-id="1e558-283">Type</span></span>

*   <span data-ttu-id="1e558-284">文字列</span><span class="sxs-lookup"><span data-stu-id="1e558-284">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-285">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-285">Requirements</span></span>

|<span data-ttu-id="1e558-286">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-286">Requirement</span></span>|<span data-ttu-id="1e558-287">値</span><span class="sxs-lookup"><span data-stu-id="1e558-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-288">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-289">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-289">1.0</span></span>|
|[<span data-ttu-id="1e558-290">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-291">ReadItem</span></span>|
|[<span data-ttu-id="1e558-292">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-293">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e558-293">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e558-294">例</span><span class="sxs-lookup"><span data-stu-id="1e558-294">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="1e558-295">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="1e558-295">dateTimeCreated :Date</span></span>

<span data-ttu-id="1e558-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1e558-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1e558-298">型</span><span class="sxs-lookup"><span data-stu-id="1e558-298">Type</span></span>

*   <span data-ttu-id="1e558-299">日付</span><span class="sxs-lookup"><span data-stu-id="1e558-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-300">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-300">Requirements</span></span>

|<span data-ttu-id="1e558-301">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-301">Requirement</span></span>|<span data-ttu-id="1e558-302">値</span><span class="sxs-lookup"><span data-stu-id="1e558-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-303">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-304">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-304">1.0</span></span>|
|[<span data-ttu-id="1e558-305">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-306">ReadItem</span></span>|
|[<span data-ttu-id="1e558-307">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-308">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e558-309">例</span><span class="sxs-lookup"><span data-stu-id="1e558-309">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="1e558-310">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="1e558-310">dateTimeModified :Date</span></span>

<span data-ttu-id="1e558-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1e558-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1e558-313">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1e558-313">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="1e558-314">型</span><span class="sxs-lookup"><span data-stu-id="1e558-314">Type</span></span>

*   <span data-ttu-id="1e558-315">日付</span><span class="sxs-lookup"><span data-stu-id="1e558-315">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-316">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-316">Requirements</span></span>

|<span data-ttu-id="1e558-317">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-317">Requirement</span></span>|<span data-ttu-id="1e558-318">値</span><span class="sxs-lookup"><span data-stu-id="1e558-318">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-319">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-320">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-320">1.0</span></span>|
|[<span data-ttu-id="1e558-321">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-322">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-322">ReadItem</span></span>|
|[<span data-ttu-id="1e558-323">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-324">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-324">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e558-325">例</span><span class="sxs-lookup"><span data-stu-id="1e558-325">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="1e558-326">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="1e558-326">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="1e558-327">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1e558-327">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="1e558-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="1e558-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1e558-330">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1e558-330">Read mode</span></span>

<span data-ttu-id="1e558-331">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-331">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="1e558-332">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1e558-332">Compose mode</span></span>

<span data-ttu-id="1e558-333">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-333">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="1e558-334">[`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1e558-334">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="1e558-335">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="1e558-335">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="1e558-336">型</span><span class="sxs-lookup"><span data-stu-id="1e558-336">Type</span></span>

*   <span data-ttu-id="1e558-337">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="1e558-337">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-338">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-338">Requirements</span></span>

|<span data-ttu-id="1e558-339">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-339">Requirement</span></span>|<span data-ttu-id="1e558-340">値</span><span class="sxs-lookup"><span data-stu-id="1e558-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-341">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-342">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-342">1.0</span></span>|
|[<span data-ttu-id="1e558-343">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-344">ReadItem</span></span>|
|[<span data-ttu-id="1e558-345">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-346">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e558-346">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="1e558-347">from:[emailaddressdetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[from](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="1e558-347">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="1e558-348">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="1e558-348">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="1e558-p112">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="1e558-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1e558-351">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="1e558-351">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1e558-352">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1e558-352">Read mode</span></span>

<span data-ttu-id="1e558-353">プロパティ`from`は`EmailAddressDetails`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-353">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="1e558-354">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1e558-354">Compose mode</span></span>

<span data-ttu-id="1e558-355">プロパティ`from`は、from `From`値を取得するメソッドを提供するオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-355">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1e558-356">型</span><span class="sxs-lookup"><span data-stu-id="1e558-356">Type</span></span>

*   <span data-ttu-id="1e558-357">[電子メールアドレス](/javascript/api/outlook_1_7/office.emailaddressdetails) | [の](/javascript/api/outlook_1_7/office.from)詳細</span><span class="sxs-lookup"><span data-stu-id="1e558-357">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-358">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-358">Requirements</span></span>

|<span data-ttu-id="1e558-359">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-359">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="1e558-360">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-361">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-361">1.0</span></span>|<span data-ttu-id="1e558-362">1.7</span><span class="sxs-lookup"><span data-stu-id="1e558-362">1.7</span></span>|
|[<span data-ttu-id="1e558-363">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-364">ReadItem</span></span>|<span data-ttu-id="1e558-365">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1e558-365">ReadWriteItem</span></span>|
|[<span data-ttu-id="1e558-366">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-367">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-367">Read</span></span>|<span data-ttu-id="1e558-368">作成</span><span class="sxs-lookup"><span data-stu-id="1e558-368">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="1e558-369">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="1e558-369">internetMessageId :String</span></span>

<span data-ttu-id="1e558-p113">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1e558-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1e558-372">Type</span><span class="sxs-lookup"><span data-stu-id="1e558-372">Type</span></span>

*   <span data-ttu-id="1e558-373">文字列</span><span class="sxs-lookup"><span data-stu-id="1e558-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-374">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-374">Requirements</span></span>

|<span data-ttu-id="1e558-375">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-375">Requirement</span></span>|<span data-ttu-id="1e558-376">値</span><span class="sxs-lookup"><span data-stu-id="1e558-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-377">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-378">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-378">1.0</span></span>|
|[<span data-ttu-id="1e558-379">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-379">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-380">ReadItem</span></span>|
|[<span data-ttu-id="1e558-381">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-381">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-382">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e558-383">例</span><span class="sxs-lookup"><span data-stu-id="1e558-383">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="1e558-384">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="1e558-384">itemClass :String</span></span>

<span data-ttu-id="1e558-p114">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1e558-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="1e558-p115">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="1e558-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="1e558-389">型</span><span class="sxs-lookup"><span data-stu-id="1e558-389">Type</span></span>|<span data-ttu-id="1e558-390">説明</span><span class="sxs-lookup"><span data-stu-id="1e558-390">Description</span></span>|<span data-ttu-id="1e558-391">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="1e558-391">item class</span></span>|
|---|---|---|
|<span data-ttu-id="1e558-392">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="1e558-392">Appointment items</span></span>|<span data-ttu-id="1e558-393">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="1e558-393">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="1e558-394">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="1e558-394">Message items</span></span>|<span data-ttu-id="1e558-395">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="1e558-395">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="1e558-396">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="1e558-396">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="1e558-397">型</span><span class="sxs-lookup"><span data-stu-id="1e558-397">Type</span></span>

*   <span data-ttu-id="1e558-398">文字列</span><span class="sxs-lookup"><span data-stu-id="1e558-398">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-399">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-399">Requirements</span></span>

|<span data-ttu-id="1e558-400">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-400">Requirement</span></span>|<span data-ttu-id="1e558-401">値</span><span class="sxs-lookup"><span data-stu-id="1e558-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-402">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-403">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-403">1.0</span></span>|
|[<span data-ttu-id="1e558-404">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-404">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-405">ReadItem</span></span>|
|[<span data-ttu-id="1e558-406">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-406">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-407">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-407">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e558-408">例</span><span class="sxs-lookup"><span data-stu-id="1e558-408">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="1e558-409">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="1e558-409">(nullable) itemId :String</span></span>

<span data-ttu-id="1e558-p116">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1e558-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1e558-412">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="1e558-412">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="1e558-413">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="1e558-413">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="1e558-414">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1e558-414">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="1e558-415">詳細は、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1e558-415">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="1e558-p118">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="1e558-418">Type</span><span class="sxs-lookup"><span data-stu-id="1e558-418">Type</span></span>

*   <span data-ttu-id="1e558-419">文字列</span><span class="sxs-lookup"><span data-stu-id="1e558-419">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-420">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-420">Requirements</span></span>

|<span data-ttu-id="1e558-421">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-421">Requirement</span></span>|<span data-ttu-id="1e558-422">値</span><span class="sxs-lookup"><span data-stu-id="1e558-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-423">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-424">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-424">1.0</span></span>|
|[<span data-ttu-id="1e558-425">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-426">ReadItem</span></span>|
|[<span data-ttu-id="1e558-427">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-428">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e558-429">例</span><span class="sxs-lookup"><span data-stu-id="1e558-429">Example</span></span>

<span data-ttu-id="1e558-p119">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="1e558-432">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="1e558-432">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="1e558-433">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="1e558-433">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="1e558-434">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="1e558-434">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1e558-435">型</span><span class="sxs-lookup"><span data-stu-id="1e558-435">Type</span></span>

*   [<span data-ttu-id="1e558-436">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="1e558-436">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="1e558-437">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-437">Requirements</span></span>

|<span data-ttu-id="1e558-438">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-438">Requirement</span></span>|<span data-ttu-id="1e558-439">値</span><span class="sxs-lookup"><span data-stu-id="1e558-439">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-440">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-441">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-441">1.0</span></span>|
|[<span data-ttu-id="1e558-442">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-442">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-443">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-443">ReadItem</span></span>|
|[<span data-ttu-id="1e558-444">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-444">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-445">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e558-445">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e558-446">例</span><span class="sxs-lookup"><span data-stu-id="1e558-446">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="1e558-447">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="1e558-447">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="1e558-448">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1e558-448">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1e558-449">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1e558-449">Read mode</span></span>

<span data-ttu-id="1e558-450">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-450">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="1e558-451">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1e558-451">Compose mode</span></span>

<span data-ttu-id="1e558-452">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-452">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1e558-453">型</span><span class="sxs-lookup"><span data-stu-id="1e558-453">Type</span></span>

*   <span data-ttu-id="1e558-454">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="1e558-454">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-455">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-455">Requirements</span></span>

|<span data-ttu-id="1e558-456">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-456">Requirement</span></span>|<span data-ttu-id="1e558-457">値</span><span class="sxs-lookup"><span data-stu-id="1e558-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-458">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-459">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-459">1.0</span></span>|
|[<span data-ttu-id="1e558-460">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-460">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-461">ReadItem</span></span>|
|[<span data-ttu-id="1e558-462">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-462">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-463">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e558-463">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="1e558-464">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="1e558-464">normalizedSubject :String</span></span>

<span data-ttu-id="1e558-p120">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1e558-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="1e558-p121">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="1e558-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="1e558-469">Type</span><span class="sxs-lookup"><span data-stu-id="1e558-469">Type</span></span>

*   <span data-ttu-id="1e558-470">文字列</span><span class="sxs-lookup"><span data-stu-id="1e558-470">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-471">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-471">Requirements</span></span>

|<span data-ttu-id="1e558-472">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-472">Requirement</span></span>|<span data-ttu-id="1e558-473">値</span><span class="sxs-lookup"><span data-stu-id="1e558-473">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-474">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-474">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-475">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-475">1.0</span></span>|
|[<span data-ttu-id="1e558-476">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-476">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-477">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-477">ReadItem</span></span>|
|[<span data-ttu-id="1e558-478">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-478">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-479">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-479">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e558-480">例</span><span class="sxs-lookup"><span data-stu-id="1e558-480">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="1e558-481">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="1e558-481">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="1e558-482">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="1e558-482">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1e558-483">型</span><span class="sxs-lookup"><span data-stu-id="1e558-483">Type</span></span>

*   [<span data-ttu-id="1e558-484">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="1e558-484">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="1e558-485">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-485">Requirements</span></span>

|<span data-ttu-id="1e558-486">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-486">Requirement</span></span>|<span data-ttu-id="1e558-487">値</span><span class="sxs-lookup"><span data-stu-id="1e558-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-488">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-489">1.3</span><span class="sxs-lookup"><span data-stu-id="1e558-489">1.3</span></span>|
|[<span data-ttu-id="1e558-490">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-491">ReadItem</span></span>|
|[<span data-ttu-id="1e558-492">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-493">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e558-493">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e558-494">例</span><span class="sxs-lookup"><span data-stu-id="1e558-494">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="1e558-495">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1e558-495">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="1e558-496">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1e558-496">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="1e558-497">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="1e558-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1e558-498">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1e558-498">Read mode</span></span>

<span data-ttu-id="1e558-499">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-499">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="1e558-500">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1e558-500">Compose mode</span></span>

<span data-ttu-id="1e558-501">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-501">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1e558-502">型</span><span class="sxs-lookup"><span data-stu-id="1e558-502">Type</span></span>

*   <span data-ttu-id="1e558-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1e558-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-504">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-504">Requirements</span></span>

|<span data-ttu-id="1e558-505">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-505">Requirement</span></span>|<span data-ttu-id="1e558-506">値</span><span class="sxs-lookup"><span data-stu-id="1e558-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-507">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-508">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-508">1.0</span></span>|
|[<span data-ttu-id="1e558-509">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-510">ReadItem</span></span>|
|[<span data-ttu-id="1e558-511">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-512">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e558-512">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="1e558-513">開催者:[emailaddressdetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[開催者](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="1e558-513">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="1e558-514">指定した会議の開催者の電子メールアドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="1e558-514">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1e558-515">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1e558-515">Read mode</span></span>

<span data-ttu-id="1e558-516">プロパティ`organizer`は、会議の開催者を表す[emailaddressdetails](/javascript/api/outlook_1_7/office.emailaddressdetails)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-516">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="1e558-517">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1e558-517">Compose mode</span></span>

<span data-ttu-id="1e558-518">プロパティ`organizer`は、開催者の値を取得するためのメソッドを提供する[オーガナイザー](/javascript/api/outlook_1_7/office.organizer)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-518">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="1e558-519">型</span><span class="sxs-lookup"><span data-stu-id="1e558-519">Type</span></span>

*   <span data-ttu-id="1e558-520">[emailaddressdetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [開催者](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="1e558-520">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-521">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-521">Requirements</span></span>

|<span data-ttu-id="1e558-522">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-522">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="1e558-523">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-524">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-524">1.0</span></span>|<span data-ttu-id="1e558-525">1.7</span><span class="sxs-lookup"><span data-stu-id="1e558-525">1.7</span></span>|
|[<span data-ttu-id="1e558-526">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-526">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-527">ReadItem</span></span>|<span data-ttu-id="1e558-528">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1e558-528">ReadWriteItem</span></span>|
|[<span data-ttu-id="1e558-529">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-529">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-530">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-530">Read</span></span>|<span data-ttu-id="1e558-531">作成</span><span class="sxs-lookup"><span data-stu-id="1e558-531">Compose</span></span>|

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="1e558-532">(nullable) 定期的なスケジュール:[定期的](/javascript/api/outlook_1_7/office.recurrence)なアイテム</span><span class="sxs-lookup"><span data-stu-id="1e558-532">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="1e558-533">予定の定期的なパターンを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1e558-533">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="1e558-534">会議出席依頼の定期的なパターンを取得します。</span><span class="sxs-lookup"><span data-stu-id="1e558-534">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="1e558-535">予定アイテムの読み取りおよび作成モード。</span><span class="sxs-lookup"><span data-stu-id="1e558-535">Read and compose modes for appointment items.</span></span> <span data-ttu-id="1e558-536">会議出席依頼アイテムの閲覧モード。</span><span class="sxs-lookup"><span data-stu-id="1e558-536">Read mode for meeting request items.</span></span>

<span data-ttu-id="1e558-537">この`recurrence`プロパティは、アイテムが series または series 内のインスタンスの場合、定期的な予定または会議出席依頼に対して[定期的](/javascript/api/outlook_1_7/office.recurrence)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-537">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="1e558-538">`null`は、単一の予定および1つの予定の会議出席依頼に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-538">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="1e558-539">`undefined`は、会議出席依頼ではないメッセージに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-539">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="1e558-540">注: 会議出席依頼に`itemClass`は、IPM という値があります。出席依頼。</span><span class="sxs-lookup"><span data-stu-id="1e558-540">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="1e558-541">注: 定期的なオブジェクトが`null`の場合は、そのオブジェクトが単一の予定または1つの予定の会議出席依頼であり、データ系列の一部ではないことを示します。</span><span class="sxs-lookup"><span data-stu-id="1e558-541">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1e558-542">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1e558-542">Read mode</span></span>

<span data-ttu-id="1e558-543">この`recurrence`プロパティは、定期的な予定を表す[定期的](/javascript/api/outlook_1_7/office.recurrence)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-543">The `recurrence` property returns a [Recurrence](/javascript/api/outlook_1_7/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="1e558-544">これは、予定および会議出席依頼に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="1e558-544">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="1e558-545">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1e558-545">Compose mode</span></span>

<span data-ttu-id="1e558-546">この`recurrence`プロパティは、予定の繰り返しを管理するためのメソッドを提供する[定期的](/javascript/api/outlook_1_7/office.recurrence)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-546">The `recurrence` property returns a [Recurrence](/javascript/api/outlook_1_7/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="1e558-547">これは予定に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="1e558-547">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="1e558-548">型</span><span class="sxs-lookup"><span data-stu-id="1e558-548">Type</span></span>

* [<span data-ttu-id="1e558-549">Recurrence</span><span class="sxs-lookup"><span data-stu-id="1e558-549">Recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="1e558-550">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-550">Requirement</span></span>|<span data-ttu-id="1e558-551">値</span><span class="sxs-lookup"><span data-stu-id="1e558-551">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-552">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-552">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-553">1.7</span><span class="sxs-lookup"><span data-stu-id="1e558-553">1.7</span></span>|
|[<span data-ttu-id="1e558-554">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-554">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-555">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-555">ReadItem</span></span>|
|[<span data-ttu-id="1e558-556">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-556">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-557">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e558-557">Compose or Read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="1e558-558">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1e558-558">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="1e558-559">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1e558-559">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="1e558-560">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="1e558-560">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1e558-561">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1e558-561">Read mode</span></span>

<span data-ttu-id="1e558-562">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-562">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="1e558-563">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1e558-563">Compose mode</span></span>

<span data-ttu-id="1e558-564">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-564">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="1e558-565">型</span><span class="sxs-lookup"><span data-stu-id="1e558-565">Type</span></span>

*   <span data-ttu-id="1e558-566">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1e558-566">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-567">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-567">Requirements</span></span>

|<span data-ttu-id="1e558-568">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-568">Requirement</span></span>|<span data-ttu-id="1e558-569">値</span><span class="sxs-lookup"><span data-stu-id="1e558-569">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-570">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-571">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-571">1.0</span></span>|
|[<span data-ttu-id="1e558-572">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-572">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-573">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-573">ReadItem</span></span>|
|[<span data-ttu-id="1e558-574">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-574">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-575">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e558-575">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="1e558-576">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1e558-576">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="1e558-p128">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1e558-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="1e558-p129">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="1e558-p129">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1e558-581">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="1e558-581">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1e558-582">型</span><span class="sxs-lookup"><span data-stu-id="1e558-582">Type</span></span>

*   [<span data-ttu-id="1e558-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1e558-583">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1e558-584">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-584">Requirements</span></span>

|<span data-ttu-id="1e558-585">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-585">Requirement</span></span>|<span data-ttu-id="1e558-586">値</span><span class="sxs-lookup"><span data-stu-id="1e558-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-587">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-588">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-588">1.0</span></span>|
|[<span data-ttu-id="1e558-589">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-589">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-590">ReadItem</span></span>|
|[<span data-ttu-id="1e558-591">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-591">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-592">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e558-593">例</span><span class="sxs-lookup"><span data-stu-id="1e558-593">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="1e558-594">(nullable) 系列 id: String</span><span class="sxs-lookup"><span data-stu-id="1e558-594">(nullable) seriesId :String</span></span>

<span data-ttu-id="1e558-595">インスタンスが属する系列の id を取得します。</span><span class="sxs-lookup"><span data-stu-id="1e558-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="1e558-596">OWA および Outlook で、は`seriesId` 、このアイテムが属する親 (シリーズ) アイテムの Exchange Web サービス (EWS) ID を返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-596">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="1e558-597">ただし、iOS と Android では、 `seriesId`は親アイテムの REST ID を返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="1e558-598">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="1e558-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="1e558-599">`seriesId`プロパティが outlook REST API で使用される outlook id と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="1e558-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="1e558-600">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1e558-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="1e558-601">詳細は、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1e558-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="1e558-602">この`seriesId`プロパティは`null` 、単一の予定、系列のアイテム、会議出席依頼などの親アイテムを持たないアイテムに`undefined`対して、会議出席依頼以外のアイテムに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="1e558-603">Type</span><span class="sxs-lookup"><span data-stu-id="1e558-603">Type</span></span>

* <span data-ttu-id="1e558-604">文字列</span><span class="sxs-lookup"><span data-stu-id="1e558-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-605">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-605">Requirements</span></span>

|<span data-ttu-id="1e558-606">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-606">Requirement</span></span>|<span data-ttu-id="1e558-607">値</span><span class="sxs-lookup"><span data-stu-id="1e558-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-608">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-609">1.7</span><span class="sxs-lookup"><span data-stu-id="1e558-609">1.7</span></span>|
|[<span data-ttu-id="1e558-610">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-611">ReadItem</span></span>|
|[<span data-ttu-id="1e558-612">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-613">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e558-613">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e558-614">例</span><span class="sxs-lookup"><span data-stu-id="1e558-614">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="1e558-615">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="1e558-615">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="1e558-616">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1e558-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="1e558-p132">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="1e558-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1e558-619">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1e558-619">Read mode</span></span>

<span data-ttu-id="1e558-620">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-620">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="1e558-621">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1e558-621">Compose mode</span></span>

<span data-ttu-id="1e558-622">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="1e558-623">[`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1e558-623">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="1e558-624">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="1e558-624">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="1e558-625">型</span><span class="sxs-lookup"><span data-stu-id="1e558-625">Type</span></span>

*   <span data-ttu-id="1e558-626">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="1e558-626">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-627">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-627">Requirements</span></span>

|<span data-ttu-id="1e558-628">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-628">Requirement</span></span>|<span data-ttu-id="1e558-629">値</span><span class="sxs-lookup"><span data-stu-id="1e558-629">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-630">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-630">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-631">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-631">1.0</span></span>|
|[<span data-ttu-id="1e558-632">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-632">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-633">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-633">ReadItem</span></span>|
|[<span data-ttu-id="1e558-634">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-634">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-635">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e558-635">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="1e558-636">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="1e558-636">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="1e558-637">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1e558-637">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="1e558-638">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1e558-638">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1e558-639">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1e558-639">Read mode</span></span>

<span data-ttu-id="1e558-p133">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="1e558-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="1e558-642">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1e558-642">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="1e558-643">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1e558-643">Compose mode</span></span>

<span data-ttu-id="1e558-644">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="1e558-645">型</span><span class="sxs-lookup"><span data-stu-id="1e558-645">Type</span></span>

*   <span data-ttu-id="1e558-646">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="1e558-646">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-647">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-647">Requirements</span></span>

|<span data-ttu-id="1e558-648">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-648">Requirement</span></span>|<span data-ttu-id="1e558-649">値</span><span class="sxs-lookup"><span data-stu-id="1e558-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-650">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-651">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-651">1.0</span></span>|
|[<span data-ttu-id="1e558-652">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-653">ReadItem</span></span>|
|[<span data-ttu-id="1e558-654">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-655">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e558-655">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="1e558-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1e558-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="1e558-657">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1e558-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="1e558-658">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="1e558-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1e558-659">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1e558-659">Read mode</span></span>

<span data-ttu-id="1e558-p135">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="1e558-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="1e558-662">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1e558-662">Compose mode</span></span>

<span data-ttu-id="1e558-663">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-663">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1e558-664">型</span><span class="sxs-lookup"><span data-stu-id="1e558-664">Type</span></span>

*   <span data-ttu-id="1e558-665">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1e558-665">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-666">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-666">Requirements</span></span>

|<span data-ttu-id="1e558-667">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-667">Requirement</span></span>|<span data-ttu-id="1e558-668">値</span><span class="sxs-lookup"><span data-stu-id="1e558-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-669">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-670">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-670">1.0</span></span>|
|[<span data-ttu-id="1e558-671">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-672">ReadItem</span></span>|
|[<span data-ttu-id="1e558-673">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-674">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e558-674">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="1e558-675">メソッド</span><span class="sxs-lookup"><span data-stu-id="1e558-675">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="1e558-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1e558-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1e558-677">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="1e558-677">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="1e558-678">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="1e558-678">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="1e558-679">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="1e558-679">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e558-680">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1e558-680">Parameters</span></span>
|<span data-ttu-id="1e558-681">名前</span><span class="sxs-lookup"><span data-stu-id="1e558-681">Name</span></span>|<span data-ttu-id="1e558-682">種類</span><span class="sxs-lookup"><span data-stu-id="1e558-682">Type</span></span>|<span data-ttu-id="1e558-683">属性</span><span class="sxs-lookup"><span data-stu-id="1e558-683">Attributes</span></span>|<span data-ttu-id="1e558-684">説明</span><span class="sxs-lookup"><span data-stu-id="1e558-684">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="1e558-685">String</span><span class="sxs-lookup"><span data-stu-id="1e558-685">String</span></span>||<span data-ttu-id="1e558-p136">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="1e558-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="1e558-688">String</span><span class="sxs-lookup"><span data-stu-id="1e558-688">String</span></span>||<span data-ttu-id="1e558-p137">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="1e558-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="1e558-691">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1e558-691">Object</span></span>|<span data-ttu-id="1e558-692">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-692">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-693">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="1e558-693">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1e558-694">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1e558-694">Object</span></span>|<span data-ttu-id="1e558-695">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-695">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-696">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1e558-696">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="1e558-697">ブール値</span><span class="sxs-lookup"><span data-stu-id="1e558-697">Boolean</span></span>|<span data-ttu-id="1e558-698">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-698">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-699">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="1e558-699">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="1e558-700">関数</span><span class="sxs-lookup"><span data-stu-id="1e558-700">function</span></span>|<span data-ttu-id="1e558-701">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-701">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-702">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="1e558-702">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1e558-703">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-703">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1e558-704">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="1e558-704">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1e558-705">エラー</span><span class="sxs-lookup"><span data-stu-id="1e558-705">Errors</span></span>

|<span data-ttu-id="1e558-706">エラー コード</span><span class="sxs-lookup"><span data-stu-id="1e558-706">Error code</span></span>|<span data-ttu-id="1e558-707">説明</span><span class="sxs-lookup"><span data-stu-id="1e558-707">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="1e558-708">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="1e558-708">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="1e558-709">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="1e558-709">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="1e558-710">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="1e558-710">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e558-711">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-711">Requirements</span></span>

|<span data-ttu-id="1e558-712">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-712">Requirement</span></span>|<span data-ttu-id="1e558-713">値</span><span class="sxs-lookup"><span data-stu-id="1e558-713">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-714">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-714">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-715">1.1</span><span class="sxs-lookup"><span data-stu-id="1e558-715">1.1</span></span>|
|[<span data-ttu-id="1e558-716">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-716">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-717">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1e558-717">ReadWriteItem</span></span>|
|[<span data-ttu-id="1e558-718">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-718">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-719">作成</span><span class="sxs-lookup"><span data-stu-id="1e558-719">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="1e558-720">例</span><span class="sxs-lookup"><span data-stu-id="1e558-720">Examples</span></span>

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

<span data-ttu-id="1e558-721">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="1e558-721">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="1e558-722">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1e558-722">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="1e558-723">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="1e558-723">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="1e558-724">現在、サポートされて`Office.EventType.AppointmentTimeChanged`いる`Office.EventType.RecipientsChanged`イベントの種類は、、、です。`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="1e558-724">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e558-725">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1e558-725">Parameters</span></span>

| <span data-ttu-id="1e558-726">名前</span><span class="sxs-lookup"><span data-stu-id="1e558-726">Name</span></span> | <span data-ttu-id="1e558-727">種類</span><span class="sxs-lookup"><span data-stu-id="1e558-727">Type</span></span> | <span data-ttu-id="1e558-728">属性</span><span class="sxs-lookup"><span data-stu-id="1e558-728">Attributes</span></span> | <span data-ttu-id="1e558-729">説明</span><span class="sxs-lookup"><span data-stu-id="1e558-729">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="1e558-730">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="1e558-730">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="1e558-731">ハンドラを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="1e558-731">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="1e558-732">関数</span><span class="sxs-lookup"><span data-stu-id="1e558-732">Function</span></span> || <span data-ttu-id="1e558-p138">イベントを処理する関数。この関数は、オブジェクト リテラルである単一パラメータを受け入れる必要があります。パラメータの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメータと一致します。</span><span class="sxs-lookup"><span data-stu-id="1e558-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="1e558-736">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1e558-736">Object</span></span> | <span data-ttu-id="1e558-737">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-737">&lt;optional&gt;</span></span> | <span data-ttu-id="1e558-738">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="1e558-738">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="1e558-739">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1e558-739">Object</span></span> | <span data-ttu-id="1e558-740">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-740">&lt;optional&gt;</span></span> | <span data-ttu-id="1e558-741">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1e558-741">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="1e558-742">関数</span><span class="sxs-lookup"><span data-stu-id="1e558-742">function</span></span>| <span data-ttu-id="1e558-743">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-743">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-744">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="1e558-744">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e558-745">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-745">Requirements</span></span>

|<span data-ttu-id="1e558-746">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-746">Requirement</span></span>| <span data-ttu-id="1e558-747">値</span><span class="sxs-lookup"><span data-stu-id="1e558-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-748">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e558-749">1.7</span><span class="sxs-lookup"><span data-stu-id="1e558-749">1.7</span></span> |
|[<span data-ttu-id="1e558-750">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e558-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-751">ReadItem</span></span> |
|[<span data-ttu-id="1e558-752">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1e558-753">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e558-753">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="1e558-754">例</span><span class="sxs-lookup"><span data-stu-id="1e558-754">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="1e558-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1e558-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1e558-756">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="1e558-756">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="1e558-p139">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="1e558-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="1e558-760">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="1e558-760">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="1e558-761">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="1e558-761">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e558-762">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1e558-762">Parameters</span></span>

|<span data-ttu-id="1e558-763">名前</span><span class="sxs-lookup"><span data-stu-id="1e558-763">Name</span></span>|<span data-ttu-id="1e558-764">種類</span><span class="sxs-lookup"><span data-stu-id="1e558-764">Type</span></span>|<span data-ttu-id="1e558-765">属性</span><span class="sxs-lookup"><span data-stu-id="1e558-765">Attributes</span></span>|<span data-ttu-id="1e558-766">説明</span><span class="sxs-lookup"><span data-stu-id="1e558-766">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="1e558-767">String</span><span class="sxs-lookup"><span data-stu-id="1e558-767">String</span></span>||<span data-ttu-id="1e558-p140">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="1e558-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="1e558-770">String</span><span class="sxs-lookup"><span data-stu-id="1e558-770">String</span></span>||<span data-ttu-id="1e558-771">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="1e558-771">The subject of the item to be attached.</span></span> <span data-ttu-id="1e558-772">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="1e558-772">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="1e558-773">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1e558-773">Object</span></span>|<span data-ttu-id="1e558-774">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-774">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-775">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="1e558-775">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1e558-776">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1e558-776">Object</span></span>|<span data-ttu-id="1e558-777">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-777">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-778">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1e558-778">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1e558-779">関数</span><span class="sxs-lookup"><span data-stu-id="1e558-779">function</span></span>|<span data-ttu-id="1e558-780">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-780">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-781">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="1e558-781">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1e558-782">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-782">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1e558-783">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="1e558-783">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1e558-784">エラー</span><span class="sxs-lookup"><span data-stu-id="1e558-784">Errors</span></span>

|<span data-ttu-id="1e558-785">エラー コード</span><span class="sxs-lookup"><span data-stu-id="1e558-785">Error code</span></span>|<span data-ttu-id="1e558-786">説明</span><span class="sxs-lookup"><span data-stu-id="1e558-786">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="1e558-787">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="1e558-787">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e558-788">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-788">Requirements</span></span>

|<span data-ttu-id="1e558-789">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-789">Requirement</span></span>|<span data-ttu-id="1e558-790">値</span><span class="sxs-lookup"><span data-stu-id="1e558-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-791">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-792">1.1</span><span class="sxs-lookup"><span data-stu-id="1e558-792">1.1</span></span>|
|[<span data-ttu-id="1e558-793">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-793">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-794">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1e558-794">ReadWriteItem</span></span>|
|[<span data-ttu-id="1e558-795">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-795">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-796">作成</span><span class="sxs-lookup"><span data-stu-id="1e558-796">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1e558-797">例</span><span class="sxs-lookup"><span data-stu-id="1e558-797">Example</span></span>

<span data-ttu-id="1e558-798">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-798">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="1e558-799">close()</span><span class="sxs-lookup"><span data-stu-id="1e558-799">close()</span></span>

<span data-ttu-id="1e558-800">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="1e558-800">Closes the current item that is being composed.</span></span>

<span data-ttu-id="1e558-p142">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="1e558-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="1e558-803">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-803">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="1e558-804">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="1e558-804">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-805">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-805">Requirements</span></span>

|<span data-ttu-id="1e558-806">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-806">Requirement</span></span>|<span data-ttu-id="1e558-807">値</span><span class="sxs-lookup"><span data-stu-id="1e558-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-808">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-809">1.3</span><span class="sxs-lookup"><span data-stu-id="1e558-809">1.3</span></span>|
|[<span data-ttu-id="1e558-810">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-810">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-811">制限あり</span><span class="sxs-lookup"><span data-stu-id="1e558-811">Restricted</span></span>|
|[<span data-ttu-id="1e558-812">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-812">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-813">新規作成</span><span class="sxs-lookup"><span data-stu-id="1e558-813">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="1e558-814">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="1e558-814">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="1e558-815">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-815">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1e558-816">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1e558-816">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1e558-817">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-817">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1e558-818">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="1e558-818">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="1e558-p143">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="1e558-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e558-822">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1e558-822">Parameters</span></span>

|<span data-ttu-id="1e558-823">名前</span><span class="sxs-lookup"><span data-stu-id="1e558-823">Name</span></span>|<span data-ttu-id="1e558-824">種類</span><span class="sxs-lookup"><span data-stu-id="1e558-824">Type</span></span>|<span data-ttu-id="1e558-825">属性</span><span class="sxs-lookup"><span data-stu-id="1e558-825">Attributes</span></span>|<span data-ttu-id="1e558-826">説明</span><span class="sxs-lookup"><span data-stu-id="1e558-826">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="1e558-827">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="1e558-827">String &#124; Object</span></span>||<span data-ttu-id="1e558-828">回答フォームの本文を表すテキストと HTML が含まれる文字列。</span><span class="sxs-lookup"><span data-stu-id="1e558-828">A string that contains text and HTML and that represents the body of the reply form.</span></span> <span data-ttu-id="1e558-829">文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="1e558-829">The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1e558-830">**または**</span><span class="sxs-lookup"><span data-stu-id="1e558-830">**OR**</span></span><br/><span data-ttu-id="1e558-p145">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="1e558-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="1e558-833">String</span><span class="sxs-lookup"><span data-stu-id="1e558-833">String</span></span>|<span data-ttu-id="1e558-834">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-834">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="1e558-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="1e558-837">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-837">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="1e558-838">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-838">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-839">添付ファイルまたは添付アイテムである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="1e558-839">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="1e558-840">String</span><span class="sxs-lookup"><span data-stu-id="1e558-840">String</span></span>||<span data-ttu-id="1e558-p147">添付ファイルの種類を示します。添付ファイルの場合は`file`、添付項目の場合は`item`でなければなりません。</span><span class="sxs-lookup"><span data-stu-id="1e558-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="1e558-843">String</span><span class="sxs-lookup"><span data-stu-id="1e558-843">String</span></span>||<span data-ttu-id="1e558-844">添付ファイル名を含む文字列で、255 文字以内で入力が可能です。</span><span class="sxs-lookup"><span data-stu-id="1e558-844">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="1e558-845">文字列</span><span class="sxs-lookup"><span data-stu-id="1e558-845">String</span></span>||<span data-ttu-id="1e558-p148">`type`が`file`に設定されている場合にのみ使用されます。ファイルの場所の URIです。</span><span class="sxs-lookup"><span data-stu-id="1e558-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="1e558-848">ブール値</span><span class="sxs-lookup"><span data-stu-id="1e558-848">Boolean</span></span>||<span data-ttu-id="1e558-p149">`type`が`file`に設定されている場合にのみ使用されます。`true`の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="1e558-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="1e558-851">String</span><span class="sxs-lookup"><span data-stu-id="1e558-851">String</span></span>||<span data-ttu-id="1e558-p150">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="1e558-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="1e558-855">関数</span><span class="sxs-lookup"><span data-stu-id="1e558-855">function</span></span>|<span data-ttu-id="1e558-856">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-856">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-857">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-857">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e558-858">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-858">Requirements</span></span>

|<span data-ttu-id="1e558-859">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-859">Requirement</span></span>|<span data-ttu-id="1e558-860">値</span><span class="sxs-lookup"><span data-stu-id="1e558-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-861">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-862">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-862">1.0</span></span>|
|[<span data-ttu-id="1e558-863">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-864">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-864">ReadItem</span></span>|
|[<span data-ttu-id="1e558-865">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-866">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-866">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1e558-867">例</span><span class="sxs-lookup"><span data-stu-id="1e558-867">Examples</span></span>

<span data-ttu-id="1e558-868">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="1e558-868">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="1e558-869">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="1e558-869">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="1e558-870">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="1e558-870">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1e558-871">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="1e558-871">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1e558-872">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="1e558-872">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1e558-873">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="1e558-873">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="1e558-874">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="1e558-874">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="1e558-875">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-875">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1e558-876">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1e558-876">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1e558-877">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-877">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1e558-878">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="1e558-878">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="1e558-p151">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="1e558-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e558-882">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1e558-882">Parameters</span></span>

|<span data-ttu-id="1e558-883">名前</span><span class="sxs-lookup"><span data-stu-id="1e558-883">Name</span></span>|<span data-ttu-id="1e558-884">種類</span><span class="sxs-lookup"><span data-stu-id="1e558-884">Type</span></span>|<span data-ttu-id="1e558-885">属性</span><span class="sxs-lookup"><span data-stu-id="1e558-885">Attributes</span></span>|<span data-ttu-id="1e558-886">説明</span><span class="sxs-lookup"><span data-stu-id="1e558-886">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="1e558-887">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="1e558-887">String &#124; Object</span></span>||<span data-ttu-id="1e558-888">回答フォームの本文を表すテキストと HTML が含まれる文字列。</span><span class="sxs-lookup"><span data-stu-id="1e558-888">A string that contains text and HTML and that represents the body of the reply form.</span></span> <span data-ttu-id="1e558-889">文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="1e558-889">The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1e558-890">**または**</span><span class="sxs-lookup"><span data-stu-id="1e558-890">**OR**</span></span><br/><span data-ttu-id="1e558-p153">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="1e558-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="1e558-893">String</span><span class="sxs-lookup"><span data-stu-id="1e558-893">String</span></span>|<span data-ttu-id="1e558-894">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-894">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-p154">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="1e558-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="1e558-897">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-897">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="1e558-898">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-898">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-899">添付ファイルまたは添付アイテムである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="1e558-899">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="1e558-900">String</span><span class="sxs-lookup"><span data-stu-id="1e558-900">String</span></span>||<span data-ttu-id="1e558-p155">添付ファイルの種類を示します。添付ファイルの場合は`file`、添付項目の場合は`item`でなければなりません。</span><span class="sxs-lookup"><span data-stu-id="1e558-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="1e558-903">String</span><span class="sxs-lookup"><span data-stu-id="1e558-903">String</span></span>||<span data-ttu-id="1e558-904">添付ファイル名を含む文字列で、255 文字以内で入力が可能です。</span><span class="sxs-lookup"><span data-stu-id="1e558-904">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="1e558-905">文字列</span><span class="sxs-lookup"><span data-stu-id="1e558-905">String</span></span>||<span data-ttu-id="1e558-p156">`type`が`file`に設定されている場合にのみ使用されます。ファイルの場所の URIです。</span><span class="sxs-lookup"><span data-stu-id="1e558-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="1e558-908">ブール値</span><span class="sxs-lookup"><span data-stu-id="1e558-908">Boolean</span></span>||<span data-ttu-id="1e558-p157">`type`が`file`に設定されている場合にのみ使用されます。`true`の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="1e558-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="1e558-911">String</span><span class="sxs-lookup"><span data-stu-id="1e558-911">String</span></span>||<span data-ttu-id="1e558-p158">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="1e558-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="1e558-915">関数</span><span class="sxs-lookup"><span data-stu-id="1e558-915">function</span></span>|<span data-ttu-id="1e558-916">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-916">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-917">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-917">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e558-918">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-918">Requirements</span></span>

|<span data-ttu-id="1e558-919">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-919">Requirement</span></span>|<span data-ttu-id="1e558-920">値</span><span class="sxs-lookup"><span data-stu-id="1e558-920">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-921">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-921">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-922">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-922">1.0</span></span>|
|[<span data-ttu-id="1e558-923">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-923">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-924">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-924">ReadItem</span></span>|
|[<span data-ttu-id="1e558-925">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-925">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-926">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-926">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1e558-927">例</span><span class="sxs-lookup"><span data-stu-id="1e558-927">Examples</span></span>

<span data-ttu-id="1e558-928">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="1e558-928">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="1e558-929">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="1e558-929">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="1e558-930">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="1e558-930">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1e558-931">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="1e558-931">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1e558-932">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="1e558-932">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1e558-933">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="1e558-933">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="1e558-934">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="1e558-934">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="1e558-935">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="1e558-935">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1e558-936">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1e558-936">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-937">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-937">Requirements</span></span>

|<span data-ttu-id="1e558-938">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-938">Requirement</span></span>|<span data-ttu-id="1e558-939">値</span><span class="sxs-lookup"><span data-stu-id="1e558-939">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-940">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-940">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-941">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-941">1.0</span></span>|
|[<span data-ttu-id="1e558-942">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-942">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-943">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-943">ReadItem</span></span>|
|[<span data-ttu-id="1e558-944">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-944">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-945">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-945">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1e558-946">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1e558-946">Returns:</span></span>

<span data-ttu-id="1e558-947">型:[Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="1e558-947">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="1e558-948">例</span><span class="sxs-lookup"><span data-stu-id="1e558-948">Example</span></span>

<span data-ttu-id="1e558-949">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="1e558-949">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="1e558-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="1e558-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="1e558-951">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="1e558-951">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1e558-952">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1e558-952">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e558-953">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1e558-953">Parameters</span></span>

|<span data-ttu-id="1e558-954">名前</span><span class="sxs-lookup"><span data-stu-id="1e558-954">Name</span></span>|<span data-ttu-id="1e558-955">型</span><span class="sxs-lookup"><span data-stu-id="1e558-955">Type</span></span>|<span data-ttu-id="1e558-956">説明</span><span class="sxs-lookup"><span data-stu-id="1e558-956">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="1e558-957">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="1e558-957">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="1e558-958">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="1e558-958">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e558-959">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-959">Requirements</span></span>

|<span data-ttu-id="1e558-960">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-960">Requirement</span></span>|<span data-ttu-id="1e558-961">値</span><span class="sxs-lookup"><span data-stu-id="1e558-961">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-962">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-962">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-963">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-963">1.0</span></span>|
|[<span data-ttu-id="1e558-964">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-964">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-965">制限あり</span><span class="sxs-lookup"><span data-stu-id="1e558-965">Restricted</span></span>|
|[<span data-ttu-id="1e558-966">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-966">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-967">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-967">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1e558-968">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1e558-968">Returns:</span></span>

<span data-ttu-id="1e558-969">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-969">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="1e558-970">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-970">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="1e558-971">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="1e558-971">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="1e558-972">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="1e558-972">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="1e558-973">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="1e558-973">Value of `entityType`</span></span>|<span data-ttu-id="1e558-974">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="1e558-974">Type of objects in returned array</span></span>|<span data-ttu-id="1e558-975">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="1e558-975">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="1e558-976">String</span><span class="sxs-lookup"><span data-stu-id="1e558-976">String</span></span>|<span data-ttu-id="1e558-977">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="1e558-977">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="1e558-978">Contact</span><span class="sxs-lookup"><span data-stu-id="1e558-978">Contact</span></span>|<span data-ttu-id="1e558-979">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1e558-979">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="1e558-980">String</span><span class="sxs-lookup"><span data-stu-id="1e558-980">String</span></span>|<span data-ttu-id="1e558-981">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1e558-981">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="1e558-982">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="1e558-982">MeetingSuggestion</span></span>|<span data-ttu-id="1e558-983">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1e558-983">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="1e558-984">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="1e558-984">PhoneNumber</span></span>|<span data-ttu-id="1e558-985">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="1e558-985">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="1e558-986">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="1e558-986">TaskSuggestion</span></span>|<span data-ttu-id="1e558-987">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1e558-987">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="1e558-988">String</span><span class="sxs-lookup"><span data-stu-id="1e558-988">String</span></span>|<span data-ttu-id="1e558-989">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="1e558-989">**Restricted**</span></span>|

<span data-ttu-id="1e558-990">型:Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="1e558-990">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="1e558-991">例</span><span class="sxs-lookup"><span data-stu-id="1e558-991">Example</span></span>

<span data-ttu-id="1e558-992">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="1e558-992">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="1e558-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="1e558-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="1e558-994">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-994">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1e558-995">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1e558-995">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1e558-996">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-996">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e558-997">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1e558-997">Parameters</span></span>

|<span data-ttu-id="1e558-998">名前</span><span class="sxs-lookup"><span data-stu-id="1e558-998">Name</span></span>|<span data-ttu-id="1e558-999">型</span><span class="sxs-lookup"><span data-stu-id="1e558-999">Type</span></span>|<span data-ttu-id="1e558-1000">説明</span><span class="sxs-lookup"><span data-stu-id="1e558-1000">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="1e558-1001">String</span><span class="sxs-lookup"><span data-stu-id="1e558-1001">String</span></span>|<span data-ttu-id="1e558-1002">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="1e558-1002">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e558-1003">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-1003">Requirements</span></span>

|<span data-ttu-id="1e558-1004">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-1004">Requirement</span></span>|<span data-ttu-id="1e558-1005">値</span><span class="sxs-lookup"><span data-stu-id="1e558-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-1006">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-1006">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-1007">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-1007">1.0</span></span>|
|[<span data-ttu-id="1e558-1008">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-1008">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-1009">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-1009">ReadItem</span></span>|
|[<span data-ttu-id="1e558-1010">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-1010">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-1011">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-1011">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1e558-1012">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="1e558-1012">Returns:</span></span>

<span data-ttu-id="1e558-p160">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="1e558-1015">型:Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="1e558-1015">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="1e558-1016">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1e558-1016">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="1e558-1017">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-1017">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1e558-1018">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1e558-1018">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1e558-p161">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="1e558-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1e558-1022">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="1e558-1022">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1e558-1023">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="1e558-1023">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="1e558-p162">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="1e558-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-1027">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-1027">Requirements</span></span>

|<span data-ttu-id="1e558-1028">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-1028">Requirement</span></span>|<span data-ttu-id="1e558-1029">値</span><span class="sxs-lookup"><span data-stu-id="1e558-1029">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-1030">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-1030">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-1031">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-1031">1.0</span></span>|
|[<span data-ttu-id="1e558-1032">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-1032">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-1033">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-1033">ReadItem</span></span>|
|[<span data-ttu-id="1e558-1034">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-1034">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-1035">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-1035">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1e558-1036">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="1e558-1036">Returns:</span></span>

<span data-ttu-id="1e558-p163">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="1e558-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="1e558-1039">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="1e558-1039">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1e558-1040">Object</span><span class="sxs-lookup"><span data-stu-id="1e558-1040">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1e558-1041">例</span><span class="sxs-lookup"><span data-stu-id="1e558-1041">Example</span></span>

<span data-ttu-id="1e558-1042">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1e558-1042">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="1e558-1043">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="1e558-1043">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="1e558-1044">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-1044">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1e558-1045">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1e558-1045">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1e558-1046">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="1e558-1046">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="1e558-p164">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="1e558-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e558-1049">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1e558-1049">Parameters</span></span>

|<span data-ttu-id="1e558-1050">名前</span><span class="sxs-lookup"><span data-stu-id="1e558-1050">Name</span></span>|<span data-ttu-id="1e558-1051">型</span><span class="sxs-lookup"><span data-stu-id="1e558-1051">Type</span></span>|<span data-ttu-id="1e558-1052">説明</span><span class="sxs-lookup"><span data-stu-id="1e558-1052">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="1e558-1053">String</span><span class="sxs-lookup"><span data-stu-id="1e558-1053">String</span></span>|<span data-ttu-id="1e558-1054">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="1e558-1054">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e558-1055">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-1055">Requirements</span></span>

|<span data-ttu-id="1e558-1056">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-1056">Requirement</span></span>|<span data-ttu-id="1e558-1057">値</span><span class="sxs-lookup"><span data-stu-id="1e558-1057">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-1058">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-1058">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-1059">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-1059">1.0</span></span>|
|[<span data-ttu-id="1e558-1060">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-1060">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-1061">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-1061">ReadItem</span></span>|
|[<span data-ttu-id="1e558-1062">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-1062">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-1063">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-1063">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1e558-1064">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="1e558-1064">Returns:</span></span>

<span data-ttu-id="1e558-1065">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="1e558-1065">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="1e558-1066">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="1e558-1066">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1e558-1067">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="1e558-1067">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1e558-1068">例</span><span class="sxs-lookup"><span data-stu-id="1e558-1068">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="1e558-1069">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="1e558-1069">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="1e558-1070">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-1070">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="1e558-p165">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e558-1073">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1e558-1073">Parameters</span></span>

|<span data-ttu-id="1e558-1074">名前</span><span class="sxs-lookup"><span data-stu-id="1e558-1074">Name</span></span>|<span data-ttu-id="1e558-1075">種類</span><span class="sxs-lookup"><span data-stu-id="1e558-1075">Type</span></span>|<span data-ttu-id="1e558-1076">属性</span><span class="sxs-lookup"><span data-stu-id="1e558-1076">Attributes</span></span>|<span data-ttu-id="1e558-1077">説明</span><span class="sxs-lookup"><span data-stu-id="1e558-1077">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="1e558-1078">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1e558-1078">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="1e558-p166">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="1e558-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="1e558-1082">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1e558-1082">Object</span></span>|<span data-ttu-id="1e558-1083">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-1083">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-1084">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="1e558-1084">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1e558-1085">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1e558-1085">Object</span></span>|<span data-ttu-id="1e558-1086">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-1086">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-1087">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1e558-1087">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1e558-1088">関数</span><span class="sxs-lookup"><span data-stu-id="1e558-1088">function</span></span>||<span data-ttu-id="1e558-1089">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="1e558-1089">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1e558-1090">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="1e558-1090">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="1e558-1091">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="1e558-1091">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e558-1092">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-1092">Requirements</span></span>

|<span data-ttu-id="1e558-1093">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-1093">Requirement</span></span>|<span data-ttu-id="1e558-1094">値</span><span class="sxs-lookup"><span data-stu-id="1e558-1094">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-1095">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-1095">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-1096">1.2</span><span class="sxs-lookup"><span data-stu-id="1e558-1096">1.2</span></span>|
|[<span data-ttu-id="1e558-1097">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-1097">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-1098">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1e558-1098">ReadWriteItem</span></span>|
|[<span data-ttu-id="1e558-1099">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-1099">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-1100">作成</span><span class="sxs-lookup"><span data-stu-id="1e558-1100">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="1e558-1101">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1e558-1101">Returns:</span></span>

<span data-ttu-id="1e558-1102">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="1e558-1102">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="1e558-1103">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="1e558-1103">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1e558-1104">String</span><span class="sxs-lookup"><span data-stu-id="1e558-1104">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1e558-1105">例</span><span class="sxs-lookup"><span data-stu-id="1e558-1105">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="1e558-1106">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="1e558-1106">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="1e558-1107">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="1e558-1107">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="1e558-1108">強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-1108">Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="1e558-1109">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1e558-1109">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-1110">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-1110">Requirements</span></span>

|<span data-ttu-id="1e558-1111">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-1111">Requirement</span></span>|<span data-ttu-id="1e558-1112">値</span><span class="sxs-lookup"><span data-stu-id="1e558-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-1113">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-1114">1.6</span><span class="sxs-lookup"><span data-stu-id="1e558-1114">1.6</span></span>|
|[<span data-ttu-id="1e558-1115">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-1115">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-1116">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-1116">ReadItem</span></span>|
|[<span data-ttu-id="1e558-1117">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-1117">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-1118">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-1118">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1e558-1119">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="1e558-1119">Returns:</span></span>

<span data-ttu-id="1e558-1120">型:[Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="1e558-1120">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="1e558-1121">例</span><span class="sxs-lookup"><span data-stu-id="1e558-1121">Example</span></span>

<span data-ttu-id="1e558-1122">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="1e558-1122">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="1e558-1123">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1e558-1123">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="1e558-p169">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="1e558-1126">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1e558-1126">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1e558-p170">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="1e558-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1e558-1130">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="1e558-1130">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1e558-1131">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="1e558-1131">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="1e558-p171">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="1e558-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e558-1135">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-1135">Requirements</span></span>

|<span data-ttu-id="1e558-1136">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-1136">Requirement</span></span>|<span data-ttu-id="1e558-1137">値</span><span class="sxs-lookup"><span data-stu-id="1e558-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-1138">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-1139">1.6</span><span class="sxs-lookup"><span data-stu-id="1e558-1139">1.6</span></span>|
|[<span data-ttu-id="1e558-1140">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-1140">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-1141">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-1141">ReadItem</span></span>|
|[<span data-ttu-id="1e558-1142">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-1142">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-1143">読み取り</span><span class="sxs-lookup"><span data-stu-id="1e558-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1e558-1144">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="1e558-1144">Returns:</span></span>

<span data-ttu-id="1e558-p172">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="1e558-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="1e558-1147">例</span><span class="sxs-lookup"><span data-stu-id="1e558-1147">Example</span></span>

<span data-ttu-id="1e558-1148">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1e558-1148">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="1e558-1149">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1e558-1149">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="1e558-1150">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="1e558-1150">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="1e558-p173">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="1e558-p173">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e558-1154">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1e558-1154">Parameters</span></span>

|<span data-ttu-id="1e558-1155">名前</span><span class="sxs-lookup"><span data-stu-id="1e558-1155">Name</span></span>|<span data-ttu-id="1e558-1156">種類</span><span class="sxs-lookup"><span data-stu-id="1e558-1156">Type</span></span>|<span data-ttu-id="1e558-1157">属性</span><span class="sxs-lookup"><span data-stu-id="1e558-1157">Attributes</span></span>|<span data-ttu-id="1e558-1158">説明</span><span class="sxs-lookup"><span data-stu-id="1e558-1158">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="1e558-1159">関数</span><span class="sxs-lookup"><span data-stu-id="1e558-1159">function</span></span>||<span data-ttu-id="1e558-1160">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="1e558-1160">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1e558-1161">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-1161">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="1e558-1162">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="1e558-1162">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="1e558-1163">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1e558-1163">Object</span></span>|<span data-ttu-id="1e558-1164">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-1164">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-1165">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="1e558-1165">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="1e558-1166">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="1e558-1166">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e558-1167">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-1167">Requirements</span></span>

|<span data-ttu-id="1e558-1168">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-1168">Requirement</span></span>|<span data-ttu-id="1e558-1169">値</span><span class="sxs-lookup"><span data-stu-id="1e558-1169">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-1170">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-1170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-1171">1.0以降</span><span class="sxs-lookup"><span data-stu-id="1e558-1171">1.0</span></span>|
|[<span data-ttu-id="1e558-1172">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-1172">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-1173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-1173">ReadItem</span></span>|
|[<span data-ttu-id="1e558-1174">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-1174">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-1175">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e558-1175">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e558-1176">例</span><span class="sxs-lookup"><span data-stu-id="1e558-1176">Example</span></span>

<span data-ttu-id="1e558-p176">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="1e558-p176">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="1e558-1180">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1e558-1180">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="1e558-1181">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="1e558-1181">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="1e558-p177">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="1e558-p177">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e558-1186">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1e558-1186">Parameters</span></span>

|<span data-ttu-id="1e558-1187">名前</span><span class="sxs-lookup"><span data-stu-id="1e558-1187">Name</span></span>|<span data-ttu-id="1e558-1188">種類</span><span class="sxs-lookup"><span data-stu-id="1e558-1188">Type</span></span>|<span data-ttu-id="1e558-1189">属性</span><span class="sxs-lookup"><span data-stu-id="1e558-1189">Attributes</span></span>|<span data-ttu-id="1e558-1190">説明</span><span class="sxs-lookup"><span data-stu-id="1e558-1190">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="1e558-1191">String</span><span class="sxs-lookup"><span data-stu-id="1e558-1191">String</span></span>||<span data-ttu-id="1e558-1192">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="1e558-1192">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="1e558-1193">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1e558-1193">Object</span></span>|<span data-ttu-id="1e558-1194">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-1194">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-1195">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="1e558-1195">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1e558-1196">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1e558-1196">Object</span></span>|<span data-ttu-id="1e558-1197">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-1197">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-1198">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1e558-1198">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1e558-1199">function</span><span class="sxs-lookup"><span data-stu-id="1e558-1199">function</span></span>|<span data-ttu-id="1e558-1200">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-1200">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-1201">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="1e558-1201">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1e558-1202">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="1e558-1202">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1e558-1203">エラー</span><span class="sxs-lookup"><span data-stu-id="1e558-1203">Errors</span></span>

|<span data-ttu-id="1e558-1204">エラー コード</span><span class="sxs-lookup"><span data-stu-id="1e558-1204">Error code</span></span>|<span data-ttu-id="1e558-1205">説明</span><span class="sxs-lookup"><span data-stu-id="1e558-1205">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="1e558-1206">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="1e558-1206">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e558-1207">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-1207">Requirements</span></span>

|<span data-ttu-id="1e558-1208">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-1208">Requirement</span></span>|<span data-ttu-id="1e558-1209">値</span><span class="sxs-lookup"><span data-stu-id="1e558-1209">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-1210">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-1210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-1211">1.1</span><span class="sxs-lookup"><span data-stu-id="1e558-1211">1.1</span></span>|
|[<span data-ttu-id="1e558-1212">最小のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-1212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-1213">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1e558-1213">ReadWriteItem</span></span>|
|[<span data-ttu-id="1e558-1214">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-1214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-1215">作成</span><span class="sxs-lookup"><span data-stu-id="1e558-1215">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1e558-1216">例</span><span class="sxs-lookup"><span data-stu-id="1e558-1216">Example</span></span>

<span data-ttu-id="1e558-1217">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="1e558-1217">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="1e558-1218">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1e558-1218">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="1e558-1219">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="1e558-1219">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="1e558-1220">現在、サポートされて`Office.EventType.AppointmentTimeChanged`いる`Office.EventType.RecipientsChanged`イベントの種類は、、、です。`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="1e558-1220">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e558-1221">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1e558-1221">Parameters</span></span>

| <span data-ttu-id="1e558-1222">名前</span><span class="sxs-lookup"><span data-stu-id="1e558-1222">Name</span></span> | <span data-ttu-id="1e558-1223">種類</span><span class="sxs-lookup"><span data-stu-id="1e558-1223">Type</span></span> | <span data-ttu-id="1e558-1224">属性</span><span class="sxs-lookup"><span data-stu-id="1e558-1224">Attributes</span></span> | <span data-ttu-id="1e558-1225">説明</span><span class="sxs-lookup"><span data-stu-id="1e558-1225">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="1e558-1226">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="1e558-1226">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="1e558-1227">ハンドラを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="1e558-1227">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="1e558-1228">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1e558-1228">Object</span></span> | <span data-ttu-id="1e558-1229">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="1e558-1230">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="1e558-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="1e558-1231">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1e558-1231">Object</span></span> | <span data-ttu-id="1e558-1232">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="1e558-1233">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1e558-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="1e558-1234">関数</span><span class="sxs-lookup"><span data-stu-id="1e558-1234">function</span></span>| <span data-ttu-id="1e558-1235">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-1236">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="1e558-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e558-1237">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-1237">Requirements</span></span>

|<span data-ttu-id="1e558-1238">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-1238">Requirement</span></span>| <span data-ttu-id="1e558-1239">値</span><span class="sxs-lookup"><span data-stu-id="1e558-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-1240">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-1240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e558-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="1e558-1241">1.7</span></span> |
|[<span data-ttu-id="1e558-1242">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-1242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e558-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1e558-1243">ReadItem</span></span> |
|[<span data-ttu-id="1e558-1244">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-1244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1e558-1245">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e558-1245">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="1e558-1246">例</span><span class="sxs-lookup"><span data-stu-id="1e558-1246">Example</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="1e558-1247">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="1e558-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="1e558-1248">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="1e558-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="1e558-p178">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-p178">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="1e558-1252">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="1e558-1252">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="1e558-1253">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="1e558-p180">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="1e558-1257">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="1e558-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="1e558-1258">Mac Outlook では、新規作成モードの会議で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="1e558-1258">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="1e558-1259">Mac Outlook では、会議で `saveAsync` を呼び出すとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-1259">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="1e558-1260">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-1260">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e558-1261">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1e558-1261">Parameters</span></span>

|<span data-ttu-id="1e558-1262">名前</span><span class="sxs-lookup"><span data-stu-id="1e558-1262">Name</span></span>|<span data-ttu-id="1e558-1263">種類</span><span class="sxs-lookup"><span data-stu-id="1e558-1263">Type</span></span>|<span data-ttu-id="1e558-1264">属性</span><span class="sxs-lookup"><span data-stu-id="1e558-1264">Attributes</span></span>|<span data-ttu-id="1e558-1265">説明</span><span class="sxs-lookup"><span data-stu-id="1e558-1265">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="1e558-1266">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1e558-1266">Object</span></span>|<span data-ttu-id="1e558-1267">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-1268">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="1e558-1268">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1e558-1269">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1e558-1269">Object</span></span>|<span data-ttu-id="1e558-1270">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-1271">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1e558-1271">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1e558-1272">関数</span><span class="sxs-lookup"><span data-stu-id="1e558-1272">function</span></span>||<span data-ttu-id="1e558-1273">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="1e558-1273">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1e558-1274">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-1274">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e558-1275">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-1275">Requirements</span></span>

|<span data-ttu-id="1e558-1276">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-1276">Requirement</span></span>|<span data-ttu-id="1e558-1277">値</span><span class="sxs-lookup"><span data-stu-id="1e558-1277">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-1278">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-1278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-1279">1.3</span><span class="sxs-lookup"><span data-stu-id="1e558-1279">1.3</span></span>|
|[<span data-ttu-id="1e558-1280">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-1280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-1281">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1e558-1281">ReadWriteItem</span></span>|
|[<span data-ttu-id="1e558-1282">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-1282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-1283">作成</span><span class="sxs-lookup"><span data-stu-id="1e558-1283">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="1e558-1284">例</span><span class="sxs-lookup"><span data-stu-id="1e558-1284">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="1e558-p182">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="1e558-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="1e558-1287">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="1e558-1287">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="1e558-1288">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="1e558-1288">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="1e558-p183">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="1e558-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1e558-1292">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1e558-1292">Parameters</span></span>

|<span data-ttu-id="1e558-1293">名前</span><span class="sxs-lookup"><span data-stu-id="1e558-1293">Name</span></span>|<span data-ttu-id="1e558-1294">種類</span><span class="sxs-lookup"><span data-stu-id="1e558-1294">Type</span></span>|<span data-ttu-id="1e558-1295">属性</span><span class="sxs-lookup"><span data-stu-id="1e558-1295">Attributes</span></span>|<span data-ttu-id="1e558-1296">説明</span><span class="sxs-lookup"><span data-stu-id="1e558-1296">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="1e558-1297">String</span><span class="sxs-lookup"><span data-stu-id="1e558-1297">String</span></span>||<span data-ttu-id="1e558-p184">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="1e558-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="1e558-1301">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1e558-1301">Object</span></span>|<span data-ttu-id="1e558-1302">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-1303">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="1e558-1303">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1e558-1304">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1e558-1304">Object</span></span>|<span data-ttu-id="1e558-1305">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-1305">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-1306">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1e558-1306">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="1e558-1307">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1e558-1307">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="1e558-1308">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1e558-1308">&lt;optional&gt;</span></span>|<span data-ttu-id="1e558-p185">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-p185">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="1e558-p186">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-p186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="1e558-1313">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="1e558-1313">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="1e558-1314">関数</span><span class="sxs-lookup"><span data-stu-id="1e558-1314">function</span></span>||<span data-ttu-id="1e558-1315">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="1e558-1315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e558-1316">要件</span><span class="sxs-lookup"><span data-stu-id="1e558-1316">Requirements</span></span>

|<span data-ttu-id="1e558-1317">必要条件</span><span class="sxs-lookup"><span data-stu-id="1e558-1317">Requirement</span></span>|<span data-ttu-id="1e558-1318">値</span><span class="sxs-lookup"><span data-stu-id="1e558-1318">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e558-1319">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e558-1319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1e558-1320">1.2</span><span class="sxs-lookup"><span data-stu-id="1e558-1320">1.2</span></span>|
|[<span data-ttu-id="1e558-1321">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e558-1321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1e558-1322">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1e558-1322">ReadWriteItem</span></span>|
|[<span data-ttu-id="1e558-1323">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e558-1323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1e558-1324">作成</span><span class="sxs-lookup"><span data-stu-id="1e558-1324">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1e558-1325">例</span><span class="sxs-lookup"><span data-stu-id="1e558-1325">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
