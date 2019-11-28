---
title: Office. メールボックス-要件セット1.6
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 46dc6148ea150e9e2ab1b245ead006a2ad377d88
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629701"
---
# <a name="item"></a><span data-ttu-id="f0b3f-102">item</span><span class="sxs-lookup"><span data-stu-id="f0b3f-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="f0b3f-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="f0b3f-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="f0b3f-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-106">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-106">Requirements</span></span>

|<span data-ttu-id="f0b3f-107">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-107">Requirement</span></span>| <span data-ttu-id="f0b3f-108">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-110">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-110">1.0</span></span>|
|[<span data-ttu-id="f0b3f-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="f0b3f-112">Restricted</span></span>|
|[<span data-ttu-id="f0b3f-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0b3f-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f0b3f-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-115">Members and methods</span></span>

| <span data-ttu-id="f0b3f-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-116">Member</span></span> | <span data-ttu-id="f0b3f-117">種類</span><span class="sxs-lookup"><span data-stu-id="f0b3f-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f0b3f-118">attachments</span><span class="sxs-lookup"><span data-stu-id="f0b3f-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="f0b3f-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-119">Member</span></span> |
| [<span data-ttu-id="f0b3f-120">bcc</span><span class="sxs-lookup"><span data-stu-id="f0b3f-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="f0b3f-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-121">Member</span></span> |
| [<span data-ttu-id="f0b3f-122">body</span><span class="sxs-lookup"><span data-stu-id="f0b3f-122">body</span></span>](#body-body) | <span data-ttu-id="f0b3f-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-123">Member</span></span> |
| [<span data-ttu-id="f0b3f-124">cc</span><span class="sxs-lookup"><span data-stu-id="f0b3f-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f0b3f-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-125">Member</span></span> |
| [<span data-ttu-id="f0b3f-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="f0b3f-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="f0b3f-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-127">Member</span></span> |
| [<span data-ttu-id="f0b3f-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="f0b3f-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="f0b3f-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-129">Member</span></span> |
| [<span data-ttu-id="f0b3f-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="f0b3f-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="f0b3f-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-131">Member</span></span> |
| [<span data-ttu-id="f0b3f-132">end</span><span class="sxs-lookup"><span data-stu-id="f0b3f-132">end</span></span>](#end-datetime) | <span data-ttu-id="f0b3f-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-133">Member</span></span> |
| [<span data-ttu-id="f0b3f-134">from</span><span class="sxs-lookup"><span data-stu-id="f0b3f-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="f0b3f-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-135">Member</span></span> |
| [<span data-ttu-id="f0b3f-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="f0b3f-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="f0b3f-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-137">Member</span></span> |
| [<span data-ttu-id="f0b3f-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="f0b3f-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="f0b3f-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-139">Member</span></span> |
| [<span data-ttu-id="f0b3f-140">itemId</span><span class="sxs-lookup"><span data-stu-id="f0b3f-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="f0b3f-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-141">Member</span></span> |
| [<span data-ttu-id="f0b3f-142">itemType</span><span class="sxs-lookup"><span data-stu-id="f0b3f-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="f0b3f-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-143">Member</span></span> |
| [<span data-ttu-id="f0b3f-144">location</span><span class="sxs-lookup"><span data-stu-id="f0b3f-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="f0b3f-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-145">Member</span></span> |
| [<span data-ttu-id="f0b3f-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="f0b3f-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="f0b3f-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-147">Member</span></span> |
| [<span data-ttu-id="f0b3f-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="f0b3f-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="f0b3f-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-149">Member</span></span> |
| [<span data-ttu-id="f0b3f-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="f0b3f-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f0b3f-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-151">Member</span></span> |
| [<span data-ttu-id="f0b3f-152">organizer</span><span class="sxs-lookup"><span data-stu-id="f0b3f-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="f0b3f-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-153">Member</span></span> |
| [<span data-ttu-id="f0b3f-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="f0b3f-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f0b3f-155">Member</span><span class="sxs-lookup"><span data-stu-id="f0b3f-155">Member</span></span> |
| [<span data-ttu-id="f0b3f-156">sender</span><span class="sxs-lookup"><span data-stu-id="f0b3f-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="f0b3f-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-157">Member</span></span> |
| [<span data-ttu-id="f0b3f-158">start</span><span class="sxs-lookup"><span data-stu-id="f0b3f-158">start</span></span>](#start-datetime) | <span data-ttu-id="f0b3f-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-159">Member</span></span> |
| [<span data-ttu-id="f0b3f-160">subject</span><span class="sxs-lookup"><span data-stu-id="f0b3f-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="f0b3f-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-161">Member</span></span> |
| [<span data-ttu-id="f0b3f-162">to</span><span class="sxs-lookup"><span data-stu-id="f0b3f-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f0b3f-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-163">Member</span></span> |
| [<span data-ttu-id="f0b3f-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f0b3f-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="f0b3f-165">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-165">Method</span></span> |
| [<span data-ttu-id="f0b3f-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f0b3f-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="f0b3f-167">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-167">Method</span></span> |
| [<span data-ttu-id="f0b3f-168">close</span><span class="sxs-lookup"><span data-stu-id="f0b3f-168">close</span></span>](#close) | <span data-ttu-id="f0b3f-169">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-169">Method</span></span> |
| [<span data-ttu-id="f0b3f-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="f0b3f-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="f0b3f-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-171">Method</span></span> |
| [<span data-ttu-id="f0b3f-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="f0b3f-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="f0b3f-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-173">Method</span></span> |
| [<span data-ttu-id="f0b3f-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="f0b3f-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="f0b3f-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-175">Method</span></span> |
| [<span data-ttu-id="f0b3f-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="f0b3f-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="f0b3f-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-177">Method</span></span> |
| [<span data-ttu-id="f0b3f-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="f0b3f-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="f0b3f-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-179">Method</span></span> |
| [<span data-ttu-id="f0b3f-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="f0b3f-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="f0b3f-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-181">Method</span></span> |
| [<span data-ttu-id="f0b3f-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="f0b3f-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="f0b3f-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-183">Method</span></span> |
| [<span data-ttu-id="f0b3f-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="f0b3f-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="f0b3f-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-185">Method</span></span> |
| [<span data-ttu-id="f0b3f-186">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="f0b3f-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="f0b3f-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-187">Method</span></span> |
| [<span data-ttu-id="f0b3f-188">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="f0b3f-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="f0b3f-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-189">Method</span></span> |
| [<span data-ttu-id="f0b3f-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="f0b3f-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="f0b3f-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-191">Method</span></span> |
| [<span data-ttu-id="f0b3f-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f0b3f-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="f0b3f-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-193">Method</span></span> |
| [<span data-ttu-id="f0b3f-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="f0b3f-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="f0b3f-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-195">Method</span></span> |
| [<span data-ttu-id="f0b3f-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="f0b3f-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="f0b3f-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="f0b3f-198">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-198">Example</span></span>

<span data-ttu-id="f0b3f-199">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="f0b3f-200">Members</span><span class="sxs-lookup"><span data-stu-id="f0b3f-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-16"></a><span data-ttu-id="f0b3f-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="f0b3f-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

<span data-ttu-id="f0b3f-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f0b3f-204">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="f0b3f-205">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="f0b3f-206">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-206">Type</span></span>

*   <span data-ttu-id="f0b3f-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="f0b3f-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-208">Requirements</span></span>

|<span data-ttu-id="f0b3f-209">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-209">Requirement</span></span>| <span data-ttu-id="f0b3f-210">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-211">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-212">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-212">1.0</span></span>|
|[<span data-ttu-id="f0b3f-213">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-214">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-215">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-216">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0b3f-217">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-217">Example</span></span>

<span data-ttu-id="f0b3f-218">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="f0b3f-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f0b3f-220">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="f0b3f-221">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-221">Compose mode only.</span></span>

<span data-ttu-id="f0b3f-222">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-222">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f0b3f-223">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-223">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f0b3f-224">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-224">Get 500 members maximum.</span></span>
- <span data-ttu-id="f0b3f-225">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-225">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="f0b3f-226">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-226">Type</span></span>

*   [<span data-ttu-id="f0b3f-227">受信者</span><span class="sxs-lookup"><span data-stu-id="f0b3f-227">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="f0b3f-228">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-228">Requirements</span></span>

|<span data-ttu-id="f0b3f-229">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-229">Requirement</span></span>| <span data-ttu-id="f0b3f-230">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-231">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-231">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-232">1.1</span><span class="sxs-lookup"><span data-stu-id="f0b3f-232">1.1</span></span>|
|[<span data-ttu-id="f0b3f-233">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-233">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-234">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-234">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-235">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-235">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-236">作成</span><span class="sxs-lookup"><span data-stu-id="f0b3f-236">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f0b3f-237">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-237">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-16"></a><span data-ttu-id="f0b3f-238">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-238">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f0b3f-239">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-239">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f0b3f-240">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-240">Type</span></span>

*   [<span data-ttu-id="f0b3f-241">Body</span><span class="sxs-lookup"><span data-stu-id="f0b3f-241">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="f0b3f-242">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-242">Requirements</span></span>

|<span data-ttu-id="f0b3f-243">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-243">Requirement</span></span>| <span data-ttu-id="f0b3f-244">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-245">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-246">1.1</span><span class="sxs-lookup"><span data-stu-id="f0b3f-246">1.1</span></span>|
|[<span data-ttu-id="f0b3f-247">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-247">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-248">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-249">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-249">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-250">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0b3f-250">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0b3f-251">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-251">Example</span></span>

<span data-ttu-id="f0b3f-252">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-252">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="f0b3f-253">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-253">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="f0b3f-254">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-254">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f0b3f-255">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-255">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="f0b3f-256">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-256">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f0b3f-257">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-257">Read mode</span></span>

<span data-ttu-id="f0b3f-258">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-258">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="f0b3f-259">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-259">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f0b3f-260">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-260">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="f0b3f-261">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-261">Compose mode</span></span>

<span data-ttu-id="f0b3f-262">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-262">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="f0b3f-263">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-263">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f0b3f-264">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-264">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f0b3f-265">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-265">Get 500 members maximum.</span></span>
- <span data-ttu-id="f0b3f-266">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-266">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f0b3f-267">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-267">Type</span></span>

*   <span data-ttu-id="f0b3f-268">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-268">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-269">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-269">Requirements</span></span>

|<span data-ttu-id="f0b3f-270">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-270">Requirement</span></span>| <span data-ttu-id="f0b3f-271">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-271">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-272">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-272">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-273">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-273">1.0</span></span>|
|[<span data-ttu-id="f0b3f-274">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-274">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-275">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-275">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-276">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-276">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-277">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0b3f-277">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="f0b3f-278">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-278">(nullable) conversationId: String</span></span>

<span data-ttu-id="f0b3f-279">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-279">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="f0b3f-p109">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="f0b3f-p110">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="f0b3f-284">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-284">Type</span></span>

*   <span data-ttu-id="f0b3f-285">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-285">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-286">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-286">Requirements</span></span>

|<span data-ttu-id="f0b3f-287">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-287">Requirement</span></span>| <span data-ttu-id="f0b3f-288">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-289">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-289">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-290">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-290">1.0</span></span>|
|[<span data-ttu-id="f0b3f-291">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-291">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-292">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-292">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-293">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-293">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-294">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0b3f-294">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0b3f-295">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-295">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="f0b3f-296">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="f0b3f-296">dateTimeCreated: Date</span></span>

<span data-ttu-id="f0b3f-p111">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f0b3f-299">種類</span><span class="sxs-lookup"><span data-stu-id="f0b3f-299">Type</span></span>

*   <span data-ttu-id="f0b3f-300">日付</span><span class="sxs-lookup"><span data-stu-id="f0b3f-300">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-301">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-301">Requirements</span></span>

|<span data-ttu-id="f0b3f-302">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-302">Requirement</span></span>| <span data-ttu-id="f0b3f-303">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-304">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-305">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-305">1.0</span></span>|
|[<span data-ttu-id="f0b3f-306">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-306">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-307">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-308">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-308">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-309">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0b3f-310">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-310">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="f0b3f-311">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="f0b3f-311">dateTimeModified: Date</span></span>

<span data-ttu-id="f0b3f-p112">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f0b3f-314">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-314">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="f0b3f-315">種類</span><span class="sxs-lookup"><span data-stu-id="f0b3f-315">Type</span></span>

*   <span data-ttu-id="f0b3f-316">日付</span><span class="sxs-lookup"><span data-stu-id="f0b3f-316">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-317">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-317">Requirements</span></span>

|<span data-ttu-id="f0b3f-318">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-318">Requirement</span></span>| <span data-ttu-id="f0b3f-319">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-320">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-321">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-321">1.0</span></span>|
|[<span data-ttu-id="f0b3f-322">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-323">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-324">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-325">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-325">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0b3f-326">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-326">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="f0b3f-327">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-327">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f0b3f-328">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-328">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="f0b3f-p113">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f0b3f-331">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-331">Read mode</span></span>

<span data-ttu-id="f0b3f-332">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-332">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="f0b3f-333">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-333">Compose mode</span></span>

<span data-ttu-id="f0b3f-334">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-334">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="f0b3f-335">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-335">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="f0b3f-336">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-336">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="f0b3f-337">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-337">Type</span></span>

*   <span data-ttu-id="f0b3f-338">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-338">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-339">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-339">Requirements</span></span>

|<span data-ttu-id="f0b3f-340">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-340">Requirement</span></span>| <span data-ttu-id="f0b3f-341">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-342">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-342">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-343">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-343">1.0</span></span>|
|[<span data-ttu-id="f0b3f-344">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-344">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-345">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-346">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-346">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-347">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0b3f-347">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="f0b3f-348">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-348">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f0b3f-p114">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p114">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="f0b3f-p115">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p115">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f0b3f-353">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-353">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f0b3f-354">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-354">Type</span></span>

*   [<span data-ttu-id="f0b3f-355">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f0b3f-355">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="example"></a><span data-ttu-id="f0b3f-356">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-356">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="f0b3f-357">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-357">Requirements</span></span>

|<span data-ttu-id="f0b3f-358">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-358">Requirement</span></span>| <span data-ttu-id="f0b3f-359">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-360">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-361">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-361">1.0</span></span>|
|[<span data-ttu-id="f0b3f-362">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-362">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-363">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-364">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-364">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-365">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-365">Read</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="f0b3f-366">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-366">internetMessageId: String</span></span>

<span data-ttu-id="f0b3f-p116">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f0b3f-369">Type</span><span class="sxs-lookup"><span data-stu-id="f0b3f-369">Type</span></span>

*   <span data-ttu-id="f0b3f-370">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-370">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-371">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-371">Requirements</span></span>

|<span data-ttu-id="f0b3f-372">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-372">Requirement</span></span>| <span data-ttu-id="f0b3f-373">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-373">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-374">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-374">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-375">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-375">1.0</span></span>|
|[<span data-ttu-id="f0b3f-376">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-376">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-377">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-378">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-378">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-379">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-379">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0b3f-380">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-380">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="f0b3f-381">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-381">itemClass: String</span></span>

<span data-ttu-id="f0b3f-p117">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="f0b3f-p118">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="f0b3f-386">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-386">Type</span></span> | <span data-ttu-id="f0b3f-387">説明</span><span class="sxs-lookup"><span data-stu-id="f0b3f-387">Description</span></span> | <span data-ttu-id="f0b3f-388">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="f0b3f-388">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="f0b3f-389">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="f0b3f-389">Appointment items</span></span> | <span data-ttu-id="f0b3f-390">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-390">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="f0b3f-391">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="f0b3f-391">Message items</span></span> | <span data-ttu-id="f0b3f-392">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-392">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="f0b3f-393">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-393">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="f0b3f-394">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-394">Type</span></span>

*   <span data-ttu-id="f0b3f-395">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-396">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-396">Requirements</span></span>

|<span data-ttu-id="f0b3f-397">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-397">Requirement</span></span>| <span data-ttu-id="f0b3f-398">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-399">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-399">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-400">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-400">1.0</span></span>|
|[<span data-ttu-id="f0b3f-401">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-401">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-402">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-403">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-403">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-404">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0b3f-405">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-405">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="f0b3f-406">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-406">(nullable) itemId: String</span></span>

<span data-ttu-id="f0b3f-p119">現在のアイテムの [Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p119">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f0b3f-409">`itemId` プロパティから返される識別子は、[Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) と同じです。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-409">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="f0b3f-410">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-410">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="f0b3f-411">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-411">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="f0b3f-412">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-412">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="f0b3f-p121">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="f0b3f-415">Type</span><span class="sxs-lookup"><span data-stu-id="f0b3f-415">Type</span></span>

*   <span data-ttu-id="f0b3f-416">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-416">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-417">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-417">Requirements</span></span>

|<span data-ttu-id="f0b3f-418">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-418">Requirement</span></span>| <span data-ttu-id="f0b3f-419">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-420">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-420">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-421">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-421">1.0</span></span>|
|[<span data-ttu-id="f0b3f-422">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-422">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-423">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-424">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-424">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-425">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-425">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0b3f-426">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-426">Example</span></span>

<span data-ttu-id="f0b3f-p122">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-16"></a><span data-ttu-id="f0b3f-429">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-429">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f0b3f-430">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-430">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="f0b3f-431">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-431">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f0b3f-432">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-432">Type</span></span>

*   [<span data-ttu-id="f0b3f-433">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="f0b3f-433">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="f0b3f-434">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-434">Requirements</span></span>

|<span data-ttu-id="f0b3f-435">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-435">Requirement</span></span>| <span data-ttu-id="f0b3f-436">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-436">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-437">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-437">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-438">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-438">1.0</span></span>|
|[<span data-ttu-id="f0b3f-439">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-439">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-440">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-440">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-441">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-441">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-442">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0b3f-442">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0b3f-443">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-443">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-16"></a><span data-ttu-id="f0b3f-444">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-444">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f0b3f-445">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-445">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f0b3f-446">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-446">Read mode</span></span>

<span data-ttu-id="f0b3f-447">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-447">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="f0b3f-448">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-448">Compose mode</span></span>

<span data-ttu-id="f0b3f-449">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-449">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f0b3f-450">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-450">Type</span></span>

*   <span data-ttu-id="f0b3f-451">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-451">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-452">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-452">Requirements</span></span>

|<span data-ttu-id="f0b3f-453">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-453">Requirement</span></span>| <span data-ttu-id="f0b3f-454">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-455">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-455">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-456">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-456">1.0</span></span>|
|[<span data-ttu-id="f0b3f-457">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-457">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-458">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-459">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-459">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-460">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0b3f-460">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="f0b3f-461">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-461">normalizedSubject: String</span></span>

<span data-ttu-id="f0b3f-p123">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="f0b3f-p124">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="f0b3f-466">Type</span><span class="sxs-lookup"><span data-stu-id="f0b3f-466">Type</span></span>

*   <span data-ttu-id="f0b3f-467">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-467">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-468">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-468">Requirements</span></span>

|<span data-ttu-id="f0b3f-469">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-469">Requirement</span></span>| <span data-ttu-id="f0b3f-470">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-470">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-471">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-471">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-472">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-472">1.0</span></span>|
|[<span data-ttu-id="f0b3f-473">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-473">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-474">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-474">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-475">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-475">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-476">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-476">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0b3f-477">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-477">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-16"></a><span data-ttu-id="f0b3f-478">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-478">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f0b3f-479">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-479">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f0b3f-480">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-480">Type</span></span>

*   [<span data-ttu-id="f0b3f-481">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="f0b3f-481">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="f0b3f-482">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-482">Requirements</span></span>

|<span data-ttu-id="f0b3f-483">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-483">Requirement</span></span>| <span data-ttu-id="f0b3f-484">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-484">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-485">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-485">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-486">1.3</span><span class="sxs-lookup"><span data-stu-id="f0b3f-486">1.3</span></span>|
|[<span data-ttu-id="f0b3f-487">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-487">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-488">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-488">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-489">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-489">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-490">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0b3f-490">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0b3f-491">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-491">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="f0b3f-492">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-492">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f0b3f-493">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-493">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="f0b3f-494">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-494">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f0b3f-495">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-495">Read mode</span></span>

<span data-ttu-id="f0b3f-496">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-496">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="f0b3f-497">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-497">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f0b3f-498">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-498">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="f0b3f-499">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-499">Compose mode</span></span>

<span data-ttu-id="f0b3f-500">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-500">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="f0b3f-501">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-501">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f0b3f-502">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-502">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f0b3f-503">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-503">Get 500 members maximum.</span></span>
- <span data-ttu-id="f0b3f-504">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-504">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f0b3f-505">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-505">Type</span></span>

*   <span data-ttu-id="f0b3f-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-507">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-507">Requirements</span></span>

|<span data-ttu-id="f0b3f-508">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-508">Requirement</span></span>| <span data-ttu-id="f0b3f-509">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-509">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-510">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-510">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-511">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-511">1.0</span></span>|
|[<span data-ttu-id="f0b3f-512">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-512">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-513">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-514">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-514">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-515">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0b3f-515">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="f0b3f-516">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-516">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f0b3f-p128">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f0b3f-519">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-519">Type</span></span>

*   [<span data-ttu-id="f0b3f-520">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f0b3f-520">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="f0b3f-521">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-521">Requirements</span></span>

|<span data-ttu-id="f0b3f-522">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-522">Requirement</span></span>| <span data-ttu-id="f0b3f-523">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-524">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-525">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-525">1.0</span></span>|
|[<span data-ttu-id="f0b3f-526">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-527">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-528">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-529">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-529">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0b3f-530">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-530">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="f0b3f-531">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-531">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f0b3f-532">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-532">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="f0b3f-533">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-533">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f0b3f-534">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-534">Read mode</span></span>

<span data-ttu-id="f0b3f-535">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-535">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="f0b3f-536">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-536">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f0b3f-537">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-537">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="f0b3f-538">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-538">Compose mode</span></span>

<span data-ttu-id="f0b3f-539">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-539">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="f0b3f-540">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-540">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f0b3f-541">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-541">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f0b3f-542">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-542">Get 500 members maximum.</span></span>
- <span data-ttu-id="f0b3f-543">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-543">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="f0b3f-544">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-544">Type</span></span>

*   <span data-ttu-id="f0b3f-545">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-545">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-546">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-546">Requirements</span></span>

|<span data-ttu-id="f0b3f-547">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-547">Requirement</span></span>| <span data-ttu-id="f0b3f-548">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-549">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-550">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-550">1.0</span></span>|
|[<span data-ttu-id="f0b3f-551">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-552">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-552">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-553">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-554">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0b3f-554">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="f0b3f-555">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-555">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f0b3f-p132">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="f0b3f-p133">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f0b3f-560">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-560">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f0b3f-561">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-561">Type</span></span>

*   [<span data-ttu-id="f0b3f-562">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f0b3f-562">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="f0b3f-563">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-563">Requirements</span></span>

|<span data-ttu-id="f0b3f-564">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-564">Requirement</span></span>| <span data-ttu-id="f0b3f-565">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-565">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-566">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-566">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-567">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-567">1.0</span></span>|
|[<span data-ttu-id="f0b3f-568">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-568">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-569">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-569">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-570">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-570">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-571">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-571">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0b3f-572">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-572">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="f0b3f-573">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-573">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f0b3f-574">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-574">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="f0b3f-p134">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f0b3f-577">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-577">Read mode</span></span>

<span data-ttu-id="f0b3f-578">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-578">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="f0b3f-579">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-579">Compose mode</span></span>

<span data-ttu-id="f0b3f-580">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-580">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="f0b3f-581">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-581">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="f0b3f-582">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-582">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="f0b3f-583">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-583">Type</span></span>

*   <span data-ttu-id="f0b3f-584">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-584">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-585">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-585">Requirements</span></span>

|<span data-ttu-id="f0b3f-586">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-586">Requirement</span></span>| <span data-ttu-id="f0b3f-587">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-587">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-588">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-589">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-589">1.0</span></span>|
|[<span data-ttu-id="f0b3f-590">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-590">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-591">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-591">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-592">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-592">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-593">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0b3f-593">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-16"></a><span data-ttu-id="f0b3f-594">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-594">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f0b3f-595">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-595">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="f0b3f-596">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-596">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f0b3f-597">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-597">Read mode</span></span>

<span data-ttu-id="f0b3f-p135">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="f0b3f-600">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-600">Compose mode</span></span>

<span data-ttu-id="f0b3f-601">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-601">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="f0b3f-602">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-602">Type</span></span>

*   <span data-ttu-id="f0b3f-603">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-603">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-604">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-604">Requirements</span></span>

|<span data-ttu-id="f0b3f-605">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-605">Requirement</span></span>| <span data-ttu-id="f0b3f-606">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-607">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-608">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-608">1.0</span></span>|
|[<span data-ttu-id="f0b3f-609">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-610">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-610">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-611">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-611">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-612">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0b3f-612">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="f0b3f-613">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-613">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="f0b3f-614">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-614">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="f0b3f-615">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-615">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f0b3f-616">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-616">Read mode</span></span>

<span data-ttu-id="f0b3f-617">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-617">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="f0b3f-618">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-618">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f0b3f-619">ただし、Windows および Mac では、500メンバーの最大値を取得するように設定できます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-619">However, on Windows and Mac, you can set up to get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="f0b3f-620">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-620">Compose mode</span></span>

<span data-ttu-id="f0b3f-621">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-621">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="f0b3f-622">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-622">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f0b3f-623">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-623">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f0b3f-624">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-624">Get 500 members maximum.</span></span>
- <span data-ttu-id="f0b3f-625">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-625">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f0b3f-626">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-626">Type</span></span>

*   <span data-ttu-id="f0b3f-627">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-627">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-628">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-628">Requirements</span></span>

|<span data-ttu-id="f0b3f-629">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-629">Requirement</span></span>| <span data-ttu-id="f0b3f-630">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-630">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-631">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-631">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-632">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-632">1.0</span></span>|
|[<span data-ttu-id="f0b3f-633">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-633">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-634">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-634">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-635">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-635">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-636">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0b3f-636">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="f0b3f-637">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0b3f-637">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="f0b3f-638">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f0b3f-638">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f0b3f-639">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-639">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="f0b3f-640">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-640">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="f0b3f-641">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-641">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0b3f-642">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f0b3f-642">Parameters</span></span>

|<span data-ttu-id="f0b3f-643">名前</span><span class="sxs-lookup"><span data-stu-id="f0b3f-643">Name</span></span>| <span data-ttu-id="f0b3f-644">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-644">Type</span></span>| <span data-ttu-id="f0b3f-645">属性</span><span class="sxs-lookup"><span data-stu-id="f0b3f-645">Attributes</span></span>| <span data-ttu-id="f0b3f-646">説明</span><span class="sxs-lookup"><span data-stu-id="f0b3f-646">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="f0b3f-647">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-647">String</span></span>||<span data-ttu-id="f0b3f-p139">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="f0b3f-650">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-650">String</span></span>||<span data-ttu-id="f0b3f-p140">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="f0b3f-653">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f0b3f-653">Object</span></span>| <span data-ttu-id="f0b3f-654">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-654">&lt;optional&gt;</span></span>|<span data-ttu-id="f0b3f-655">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-655">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="f0b3f-656">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f0b3f-656">Object</span></span> | <span data-ttu-id="f0b3f-657">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-657">&lt;optional&gt;</span></span> | <span data-ttu-id="f0b3f-658">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-658">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="f0b3f-659">Boolean</span><span class="sxs-lookup"><span data-stu-id="f0b3f-659">Boolean</span></span> | <span data-ttu-id="f0b3f-660">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-660">&lt;optional&gt;</span></span> | <span data-ttu-id="f0b3f-661">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-661">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="f0b3f-662">function</span><span class="sxs-lookup"><span data-stu-id="f0b3f-662">function</span></span>| <span data-ttu-id="f0b3f-663">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-663">&lt;optional&gt;</span></span>|<span data-ttu-id="f0b3f-664">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-664">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f0b3f-665">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-665">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f0b3f-666">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-666">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f0b3f-667">エラー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-667">Errors</span></span>

| <span data-ttu-id="f0b3f-668">エラー コード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-668">Error code</span></span> | <span data-ttu-id="f0b3f-669">説明</span><span class="sxs-lookup"><span data-stu-id="f0b3f-669">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="f0b3f-670">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-670">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="f0b3f-671">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-671">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="f0b3f-672">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-672">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f0b3f-673">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-673">Requirements</span></span>

|<span data-ttu-id="f0b3f-674">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-674">Requirement</span></span>| <span data-ttu-id="f0b3f-675">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-675">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-676">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-676">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-677">1.1</span><span class="sxs-lookup"><span data-stu-id="f0b3f-677">1.1</span></span>|
|[<span data-ttu-id="f0b3f-678">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-678">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-679">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-679">ReadWriteItem</span></span>|
|[<span data-ttu-id="f0b3f-680">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-680">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-681">作成</span><span class="sxs-lookup"><span data-stu-id="f0b3f-681">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f0b3f-682">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-682">Examples</span></span>

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

<span data-ttu-id="f0b3f-683">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-683">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="f0b3f-684">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f0b3f-684">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f0b3f-685">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-685">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="f0b3f-p141">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="f0b3f-689">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-689">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="f0b3f-690">Office アドインを Outlook on the web で実行している場合、編集中のアイテム以外のアイテムに `addItemAttachmentAsync` メソッドでアイテムを添付できます。ただし、これはサポートされていないため、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-690">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0b3f-691">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f0b3f-691">Parameters</span></span>

|<span data-ttu-id="f0b3f-692">名前</span><span class="sxs-lookup"><span data-stu-id="f0b3f-692">Name</span></span>| <span data-ttu-id="f0b3f-693">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-693">Type</span></span>| <span data-ttu-id="f0b3f-694">属性</span><span class="sxs-lookup"><span data-stu-id="f0b3f-694">Attributes</span></span>| <span data-ttu-id="f0b3f-695">説明</span><span class="sxs-lookup"><span data-stu-id="f0b3f-695">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="f0b3f-696">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-696">String</span></span>||<span data-ttu-id="f0b3f-p142">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="f0b3f-699">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-699">String</span></span>||<span data-ttu-id="f0b3f-700">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-700">The subject of the item to be attached.</span></span> <span data-ttu-id="f0b3f-701">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-701">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="f0b3f-702">Object</span><span class="sxs-lookup"><span data-stu-id="f0b3f-702">Object</span></span>| <span data-ttu-id="f0b3f-703">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-703">&lt;optional&gt;</span></span>|<span data-ttu-id="f0b3f-704">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-704">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f0b3f-705">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f0b3f-705">Object</span></span>| <span data-ttu-id="f0b3f-706">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-706">&lt;optional&gt;</span></span>|<span data-ttu-id="f0b3f-707">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-707">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f0b3f-708">function</span><span class="sxs-lookup"><span data-stu-id="f0b3f-708">function</span></span>| <span data-ttu-id="f0b3f-709">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-709">&lt;optional&gt;</span></span>|<span data-ttu-id="f0b3f-710">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-710">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f0b3f-711">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-711">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f0b3f-712">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-712">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f0b3f-713">エラー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-713">Errors</span></span>

| <span data-ttu-id="f0b3f-714">エラー コード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-714">Error code</span></span> | <span data-ttu-id="f0b3f-715">説明</span><span class="sxs-lookup"><span data-stu-id="f0b3f-715">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="f0b3f-716">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-716">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f0b3f-717">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-717">Requirements</span></span>

|<span data-ttu-id="f0b3f-718">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-718">Requirement</span></span>| <span data-ttu-id="f0b3f-719">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-719">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-720">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-720">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-721">1.1</span><span class="sxs-lookup"><span data-stu-id="f0b3f-721">1.1</span></span>|
|[<span data-ttu-id="f0b3f-722">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-722">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-723">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-723">ReadWriteItem</span></span>|
|[<span data-ttu-id="f0b3f-724">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-724">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-725">作成</span><span class="sxs-lookup"><span data-stu-id="f0b3f-725">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f0b3f-726">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-726">Example</span></span>

<span data-ttu-id="f0b3f-727">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-727">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="f0b3f-728">close()</span><span class="sxs-lookup"><span data-stu-id="f0b3f-728">close()</span></span>

<span data-ttu-id="f0b3f-729">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-729">Closes the current item that is being composed.</span></span>

<span data-ttu-id="f0b3f-p144">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="f0b3f-732">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-732">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="f0b3f-733">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-733">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-734">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-734">Requirements</span></span>

|<span data-ttu-id="f0b3f-735">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-735">Requirement</span></span>| <span data-ttu-id="f0b3f-736">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-736">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-737">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-737">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-738">1.3</span><span class="sxs-lookup"><span data-stu-id="f0b3f-738">1.3</span></span>|
|[<span data-ttu-id="f0b3f-739">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-739">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-740">制限あり</span><span class="sxs-lookup"><span data-stu-id="f0b3f-740">Restricted</span></span>|
|[<span data-ttu-id="f0b3f-741">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-741">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-742">新規作成</span><span class="sxs-lookup"><span data-stu-id="f0b3f-742">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="f0b3f-743">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="f0b3f-743">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="f0b3f-744">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-744">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f0b3f-745">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-745">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f0b3f-746">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-746">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f0b3f-747">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-747">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="f0b3f-p145">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0b3f-751">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f0b3f-751">Parameters</span></span>

| <span data-ttu-id="f0b3f-752">名前</span><span class="sxs-lookup"><span data-stu-id="f0b3f-752">Name</span></span> | <span data-ttu-id="f0b3f-753">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-753">Type</span></span> | <span data-ttu-id="f0b3f-754">属性</span><span class="sxs-lookup"><span data-stu-id="f0b3f-754">Attributes</span></span> | <span data-ttu-id="f0b3f-755">説明</span><span class="sxs-lookup"><span data-stu-id="f0b3f-755">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="f0b3f-756">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="f0b3f-756">String &#124; Object</span></span>| |<span data-ttu-id="f0b3f-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f0b3f-759">**または**</span><span class="sxs-lookup"><span data-stu-id="f0b3f-759">**OR**</span></span><br/><span data-ttu-id="f0b3f-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="f0b3f-762">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-762">String</span></span> | <span data-ttu-id="f0b3f-763">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-763">&lt;optional&gt;</span></span> | <span data-ttu-id="f0b3f-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="f0b3f-766">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-766">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="f0b3f-767">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-767">&lt;optional&gt;</span></span> | <span data-ttu-id="f0b3f-768">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-768">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="f0b3f-769">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-769">String</span></span> | | <span data-ttu-id="f0b3f-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="f0b3f-772">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-772">String</span></span> | | <span data-ttu-id="f0b3f-773">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-773">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="f0b3f-774">文字列</span><span class="sxs-lookup"><span data-stu-id="f0b3f-774">String</span></span> | | <span data-ttu-id="f0b3f-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="f0b3f-777">ブール値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-777">Boolean</span></span> | | <span data-ttu-id="f0b3f-p151">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="f0b3f-780">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-780">String</span></span> | | <span data-ttu-id="f0b3f-p152">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="f0b3f-784">function</span><span class="sxs-lookup"><span data-stu-id="f0b3f-784">function</span></span> | <span data-ttu-id="f0b3f-785">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-785">&lt;optional&gt;</span></span> | <span data-ttu-id="f0b3f-786">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-786">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f0b3f-787">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-787">Requirements</span></span>

|<span data-ttu-id="f0b3f-788">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-788">Requirement</span></span>| <span data-ttu-id="f0b3f-789">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-789">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-790">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-790">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-791">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-791">1.0</span></span>|
|[<span data-ttu-id="f0b3f-792">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-792">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-793">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-793">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-794">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-794">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-795">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-795">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f0b3f-796">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-796">Examples</span></span>

<span data-ttu-id="f0b3f-797">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-797">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="f0b3f-798">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-798">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="f0b3f-799">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-799">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f0b3f-800">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-800">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f0b3f-801">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-801">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f0b3f-802">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-802">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="f0b3f-803">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="f0b3f-803">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="f0b3f-804">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-804">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f0b3f-805">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-805">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f0b3f-806">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-806">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f0b3f-807">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-807">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="f0b3f-p153">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0b3f-811">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f0b3f-811">Parameters</span></span>

| <span data-ttu-id="f0b3f-812">名前</span><span class="sxs-lookup"><span data-stu-id="f0b3f-812">Name</span></span> | <span data-ttu-id="f0b3f-813">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-813">Type</span></span> | <span data-ttu-id="f0b3f-814">属性</span><span class="sxs-lookup"><span data-stu-id="f0b3f-814">Attributes</span></span> | <span data-ttu-id="f0b3f-815">説明</span><span class="sxs-lookup"><span data-stu-id="f0b3f-815">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="f0b3f-816">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="f0b3f-816">String &#124; Object</span></span>| | <span data-ttu-id="f0b3f-p154">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f0b3f-819">**または**</span><span class="sxs-lookup"><span data-stu-id="f0b3f-819">**OR**</span></span><br/><span data-ttu-id="f0b3f-p155">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="f0b3f-822">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-822">String</span></span> | <span data-ttu-id="f0b3f-823">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-823">&lt;optional&gt;</span></span> | <span data-ttu-id="f0b3f-p156">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="f0b3f-826">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-826">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="f0b3f-827">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-827">&lt;optional&gt;</span></span> | <span data-ttu-id="f0b3f-828">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-828">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="f0b3f-829">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-829">String</span></span> | | <span data-ttu-id="f0b3f-p157">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="f0b3f-832">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-832">String</span></span> | | <span data-ttu-id="f0b3f-833">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-833">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="f0b3f-834">文字列</span><span class="sxs-lookup"><span data-stu-id="f0b3f-834">String</span></span> | | <span data-ttu-id="f0b3f-p158">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="f0b3f-837">ブール値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-837">Boolean</span></span> | | <span data-ttu-id="f0b3f-p159">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="f0b3f-840">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-840">String</span></span> | | <span data-ttu-id="f0b3f-p160">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="f0b3f-844">function</span><span class="sxs-lookup"><span data-stu-id="f0b3f-844">function</span></span> | <span data-ttu-id="f0b3f-845">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-845">&lt;optional&gt;</span></span> | <span data-ttu-id="f0b3f-846">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-846">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f0b3f-847">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-847">Requirements</span></span>

|<span data-ttu-id="f0b3f-848">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-848">Requirement</span></span>| <span data-ttu-id="f0b3f-849">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-849">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-850">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-850">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-851">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-851">1.0</span></span>|
|[<span data-ttu-id="f0b3f-852">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-852">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-853">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-853">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-854">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-854">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-855">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-855">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f0b3f-856">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-856">Examples</span></span>

<span data-ttu-id="f0b3f-857">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-857">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="f0b3f-858">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-858">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="f0b3f-859">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-859">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f0b3f-860">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-860">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f0b3f-861">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-861">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f0b3f-862">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-862">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="f0b3f-863">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="f0b3f-863">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="f0b3f-864">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-864">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="f0b3f-865">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-865">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-866">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-866">Requirements</span></span>

|<span data-ttu-id="f0b3f-867">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-867">Requirement</span></span>| <span data-ttu-id="f0b3f-868">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-868">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-869">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-869">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-870">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-870">1.0</span></span>|
|[<span data-ttu-id="f0b3f-871">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-871">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-872">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-872">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-873">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-873">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-874">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-874">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f0b3f-875">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f0b3f-875">Returns:</span></span>

<span data-ttu-id="f0b3f-876">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-876">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="f0b3f-877">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-877">Example</span></span>

<span data-ttu-id="f0b3f-878">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-878">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="f0b3f-879">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="f0b3f-879">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="f0b3f-880">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-880">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="f0b3f-881">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-881">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0b3f-882">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f0b3f-882">Parameters</span></span>

|<span data-ttu-id="f0b3f-883">名前</span><span class="sxs-lookup"><span data-stu-id="f0b3f-883">Name</span></span>| <span data-ttu-id="f0b3f-884">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-884">Type</span></span>| <span data-ttu-id="f0b3f-885">説明</span><span class="sxs-lookup"><span data-stu-id="f0b3f-885">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="f0b3f-886">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="f0b3f-886">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.6)|<span data-ttu-id="f0b3f-887">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-887">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0b3f-888">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-888">Requirements</span></span>

|<span data-ttu-id="f0b3f-889">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-889">Requirement</span></span>| <span data-ttu-id="f0b3f-890">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-890">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-891">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-891">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-892">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-892">1.0</span></span>|
|[<span data-ttu-id="f0b3f-893">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-893">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-894">制限あり</span><span class="sxs-lookup"><span data-stu-id="f0b3f-894">Restricted</span></span>|
|[<span data-ttu-id="f0b3f-895">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-895">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-896">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-896">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f0b3f-897">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f0b3f-897">Returns:</span></span>

<span data-ttu-id="f0b3f-898">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-898">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="f0b3f-899">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-899">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="f0b3f-900">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-900">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="f0b3f-901">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-901">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="f0b3f-902">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-902">Value of `entityType`</span></span> | <span data-ttu-id="f0b3f-903">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-903">Type of objects in returned array</span></span> | <span data-ttu-id="f0b3f-904">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-904">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="f0b3f-905">文字列</span><span class="sxs-lookup"><span data-stu-id="f0b3f-905">String</span></span> | <span data-ttu-id="f0b3f-906">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="f0b3f-906">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="f0b3f-907">連絡先</span><span class="sxs-lookup"><span data-stu-id="f0b3f-907">Contact</span></span> | <span data-ttu-id="f0b3f-908">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f0b3f-908">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="f0b3f-909">文字列</span><span class="sxs-lookup"><span data-stu-id="f0b3f-909">String</span></span> | <span data-ttu-id="f0b3f-910">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f0b3f-910">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="f0b3f-911">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="f0b3f-911">MeetingSuggestion</span></span> | <span data-ttu-id="f0b3f-912">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f0b3f-912">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="f0b3f-913">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="f0b3f-913">PhoneNumber</span></span> | <span data-ttu-id="f0b3f-914">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="f0b3f-914">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="f0b3f-915">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="f0b3f-915">TaskSuggestion</span></span> | <span data-ttu-id="f0b3f-916">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f0b3f-916">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="f0b3f-917">文字列</span><span class="sxs-lookup"><span data-stu-id="f0b3f-917">String</span></span> | <span data-ttu-id="f0b3f-918">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="f0b3f-918">**Restricted**</span></span> |

<span data-ttu-id="f0b3f-919">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="f0b3f-919">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

##### <a name="example"></a><span data-ttu-id="f0b3f-920">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-920">Example</span></span>

<span data-ttu-id="f0b3f-921">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-921">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="f0b3f-922">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="f0b3f-922">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="f0b3f-923">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-923">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f0b3f-924">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-924">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f0b3f-925">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-925">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0b3f-926">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f0b3f-926">Parameters</span></span>

|<span data-ttu-id="f0b3f-927">名前</span><span class="sxs-lookup"><span data-stu-id="f0b3f-927">Name</span></span>| <span data-ttu-id="f0b3f-928">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-928">Type</span></span>| <span data-ttu-id="f0b3f-929">説明</span><span class="sxs-lookup"><span data-stu-id="f0b3f-929">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="f0b3f-930">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-930">String</span></span>|<span data-ttu-id="f0b3f-931">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-931">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0b3f-932">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-932">Requirements</span></span>

|<span data-ttu-id="f0b3f-933">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-933">Requirement</span></span>| <span data-ttu-id="f0b3f-934">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-935">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-936">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-936">1.0</span></span>|
|[<span data-ttu-id="f0b3f-937">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-937">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-938">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-939">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-939">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-940">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-940">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f0b3f-941">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f0b3f-941">Returns:</span></span>

<span data-ttu-id="f0b3f-p162">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="f0b3f-944">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="f0b3f-944">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="f0b3f-945">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f0b3f-945">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="f0b3f-946">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-946">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f0b3f-947">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-947">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f0b3f-p163">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f0b3f-951">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-951">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f0b3f-952">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-952">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f0b3f-p164">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-956">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-956">Requirements</span></span>

|<span data-ttu-id="f0b3f-957">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-957">Requirement</span></span>| <span data-ttu-id="f0b3f-958">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-958">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-959">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-959">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-960">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-960">1.0</span></span>|
|[<span data-ttu-id="f0b3f-961">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-961">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-962">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-962">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-963">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-963">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-964">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-964">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f0b3f-965">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f0b3f-965">Returns:</span></span>

<span data-ttu-id="f0b3f-p165">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="f0b3f-968">型: Object</span><span class="sxs-lookup"><span data-stu-id="f0b3f-968">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="f0b3f-969">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-969">Example</span></span>

<span data-ttu-id="f0b3f-970">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-970">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="f0b3f-971">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="f0b3f-971">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="f0b3f-972">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-972">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f0b3f-973">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-973">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f0b3f-974">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-974">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="f0b3f-p166">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0b3f-977">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f0b3f-977">Parameters</span></span>

|<span data-ttu-id="f0b3f-978">名前</span><span class="sxs-lookup"><span data-stu-id="f0b3f-978">Name</span></span>| <span data-ttu-id="f0b3f-979">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-979">Type</span></span>| <span data-ttu-id="f0b3f-980">説明</span><span class="sxs-lookup"><span data-stu-id="f0b3f-980">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="f0b3f-981">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-981">String</span></span>|<span data-ttu-id="f0b3f-982">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-982">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0b3f-983">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-983">Requirements</span></span>

|<span data-ttu-id="f0b3f-984">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-984">Requirement</span></span>| <span data-ttu-id="f0b3f-985">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-986">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-987">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-987">1.0</span></span>|
|[<span data-ttu-id="f0b3f-988">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-988">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-989">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-989">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-990">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-990">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-991">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-991">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f0b3f-992">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f0b3f-992">Returns:</span></span>

<span data-ttu-id="f0b3f-993">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-993">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="f0b3f-994">型: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="f0b3f-994">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="f0b3f-995">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-995">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="f0b3f-996">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="f0b3f-996">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="f0b3f-997">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-997">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="f0b3f-p167">選択されていない状態でカーソルが本文または件名にある場合、メソッドは選択されたデータに対し空の文字列を返します。本文または件名以外のフィールドが選択されている場合には、メソッドは`InvalidSelection`エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p167">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0b3f-1000">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1000">Parameters</span></span>

|<span data-ttu-id="f0b3f-1001">名前</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1001">Name</span></span>| <span data-ttu-id="f0b3f-1002">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1002">Type</span></span>| <span data-ttu-id="f0b3f-1003">属性</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1003">Attributes</span></span>| <span data-ttu-id="f0b3f-1004">説明</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1004">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="f0b3f-1005">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1005">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="f0b3f-p168">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p168">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="f0b3f-1009">Object</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1009">Object</span></span>| <span data-ttu-id="f0b3f-1010">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1010">&lt;optional&gt;</span></span>|<span data-ttu-id="f0b3f-1011">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1011">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f0b3f-1012">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1012">Object</span></span>| <span data-ttu-id="f0b3f-1013">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1013">&lt;optional&gt;</span></span>|<span data-ttu-id="f0b3f-1014">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1014">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f0b3f-1015">function</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1015">function</span></span>||<span data-ttu-id="f0b3f-1016">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f0b3f-1017">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1017">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="f0b3f-1018">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1018">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0b3f-1019">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1019">Requirements</span></span>

|<span data-ttu-id="f0b3f-1020">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1020">Requirement</span></span>| <span data-ttu-id="f0b3f-1021">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1021">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-1022">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1022">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-1023">1.2</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1023">1.2</span></span>|
|[<span data-ttu-id="f0b3f-1024">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1024">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-1025">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1025">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-1026">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1026">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-1027">作成</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1027">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="f0b3f-1028">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1028">Returns:</span></span>

<span data-ttu-id="f0b3f-1029">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1029">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="f0b3f-1030">型:String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1030">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="f0b3f-1031">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1031">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="f0b3f-1032">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1032">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="f0b3f-1033">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1033">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="f0b3f-1034">強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1034">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="f0b3f-1035">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1035">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-1036">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1036">Requirements</span></span>

|<span data-ttu-id="f0b3f-1037">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1037">Requirement</span></span>| <span data-ttu-id="f0b3f-1038">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1038">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-1039">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1039">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-1040">1.6</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1040">1.6</span></span> |
|[<span data-ttu-id="f0b3f-1041">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1041">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-1042">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1042">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-1043">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1043">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-1044">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1044">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f0b3f-1045">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1045">Returns:</span></span>

<span data-ttu-id="f0b3f-1046">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1046">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="f0b3f-1047">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1047">Example</span></span>

<span data-ttu-id="f0b3f-1048">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1048">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="f0b3f-1049">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1049">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="f0b3f-p171">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p171">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="f0b3f-1052">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1052">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f0b3f-p172">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p172">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f0b3f-1056">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1056">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f0b3f-1057">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1057">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f0b3f-p173">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p173">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0b3f-1061">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1061">Requirements</span></span>

|<span data-ttu-id="f0b3f-1062">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1062">Requirement</span></span>| <span data-ttu-id="f0b3f-1063">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1063">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-1064">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1064">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-1065">1.6</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1065">1.6</span></span> |
|[<span data-ttu-id="f0b3f-1066">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1066">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-1067">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1067">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-1068">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1068">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-1069">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1069">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f0b3f-1070">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1070">Returns:</span></span>

<span data-ttu-id="f0b3f-p174">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p174">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="f0b3f-1073">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1073">Example</span></span>

<span data-ttu-id="f0b3f-1074">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1074">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="f0b3f-1075">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1075">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="f0b3f-1076">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1076">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="f0b3f-p175">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p175">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0b3f-1080">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1080">Parameters</span></span>

|<span data-ttu-id="f0b3f-1081">名前</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1081">Name</span></span>| <span data-ttu-id="f0b3f-1082">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1082">Type</span></span>| <span data-ttu-id="f0b3f-1083">属性</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1083">Attributes</span></span>| <span data-ttu-id="f0b3f-1084">説明</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1084">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="f0b3f-1085">function</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1085">function</span></span>||<span data-ttu-id="f0b3f-1086">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1086">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f0b3f-1087">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1087">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="f0b3f-1088">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1088">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="f0b3f-1089">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1089">Object</span></span>| <span data-ttu-id="f0b3f-1090">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1090">&lt;optional&gt;</span></span>|<span data-ttu-id="f0b3f-1091">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1091">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="f0b3f-1092">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1092">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0b3f-1093">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1093">Requirements</span></span>

|<span data-ttu-id="f0b3f-1094">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1094">Requirement</span></span>| <span data-ttu-id="f0b3f-1095">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1095">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-1096">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1096">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-1097">1.0</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1097">1.0</span></span>|
|[<span data-ttu-id="f0b3f-1098">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1098">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-1099">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1099">ReadItem</span></span>|
|[<span data-ttu-id="f0b3f-1100">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1100">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-1101">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1101">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0b3f-1102">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1102">Example</span></span>

<span data-ttu-id="f0b3f-p178">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p178">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="f0b3f-1106">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1106">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="f0b3f-1107">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1107">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="f0b3f-1108">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1108">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="f0b3f-1109">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1109">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="f0b3f-1110">Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1110">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="f0b3f-1111">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1111">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0b3f-1112">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1112">Parameters</span></span>

|<span data-ttu-id="f0b3f-1113">名前</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1113">Name</span></span>| <span data-ttu-id="f0b3f-1114">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1114">Type</span></span>| <span data-ttu-id="f0b3f-1115">属性</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1115">Attributes</span></span>| <span data-ttu-id="f0b3f-1116">説明</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1116">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="f0b3f-1117">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1117">String</span></span>||<span data-ttu-id="f0b3f-1118">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1118">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="f0b3f-1119">Object</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1119">Object</span></span>| <span data-ttu-id="f0b3f-1120">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1120">&lt;optional&gt;</span></span>|<span data-ttu-id="f0b3f-1121">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1121">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f0b3f-1122">Object</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1122">Object</span></span>| <span data-ttu-id="f0b3f-1123">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1123">&lt;optional&gt;</span></span>|<span data-ttu-id="f0b3f-1124">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1124">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f0b3f-1125">function</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1125">function</span></span>| <span data-ttu-id="f0b3f-1126">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1126">&lt;optional&gt;</span></span>|<span data-ttu-id="f0b3f-1127">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1127">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f0b3f-1128">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1128">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f0b3f-1129">エラー</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1129">Errors</span></span>

| <span data-ttu-id="f0b3f-1130">エラー コード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1130">Error code</span></span> | <span data-ttu-id="f0b3f-1131">説明</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1131">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="f0b3f-1132">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1132">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f0b3f-1133">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1133">Requirements</span></span>

|<span data-ttu-id="f0b3f-1134">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1134">Requirement</span></span>| <span data-ttu-id="f0b3f-1135">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1135">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-1136">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1136">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-1137">1.1</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1137">1.1</span></span>|
|[<span data-ttu-id="f0b3f-1138">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1138">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-1139">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1139">ReadWriteItem</span></span>|
|[<span data-ttu-id="f0b3f-1140">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1140">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-1141">作成</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1141">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f0b3f-1142">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1142">Example</span></span>

<span data-ttu-id="f0b3f-1143">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1143">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="f0b3f-1144">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1144">saveAsync([options], callback)</span></span>

<span data-ttu-id="f0b3f-1145">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1145">Asynchronously saves an item.</span></span>

<span data-ttu-id="f0b3f-1146">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1146">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="f0b3f-1147">Outlook on the web またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1147">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="f0b3f-1148">キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1148">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="f0b3f-1149">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1149">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="f0b3f-1150">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1150">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="f0b3f-p182">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p182">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="f0b3f-1154">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1154">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="f0b3f-1155">Outlook on Mac では、会議の保存はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1155">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="f0b3f-1156">`saveAsync` メソッドは、作成モードの会議から呼び出されると失敗します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1156">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="f0b3f-1157">回避策については、「[Office JS API を使用して Outlook for Mac で会議を下書きとして保存できない](https://support.microsoft.com/help/4505745)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1157">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="f0b3f-1158">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1158">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0b3f-1159">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1159">Parameters</span></span>

|<span data-ttu-id="f0b3f-1160">名前</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1160">Name</span></span>| <span data-ttu-id="f0b3f-1161">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1161">Type</span></span>| <span data-ttu-id="f0b3f-1162">属性</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1162">Attributes</span></span>| <span data-ttu-id="f0b3f-1163">説明</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1163">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="f0b3f-1164">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1164">Object</span></span>| <span data-ttu-id="f0b3f-1165">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1165">&lt;optional&gt;</span></span>|<span data-ttu-id="f0b3f-1166">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1166">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f0b3f-1167">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1167">Object</span></span>| <span data-ttu-id="f0b3f-1168">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1168">&lt;optional&gt;</span></span>|<span data-ttu-id="f0b3f-1169">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1169">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f0b3f-1170">関数</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1170">function</span></span>||<span data-ttu-id="f0b3f-1171">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1171">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f0b3f-1172">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1172">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0b3f-1173">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1173">Requirements</span></span>

|<span data-ttu-id="f0b3f-1174">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1174">Requirement</span></span>| <span data-ttu-id="f0b3f-1175">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1175">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-1176">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-1177">1.3</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1177">1.3</span></span>|
|[<span data-ttu-id="f0b3f-1178">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1178">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-1179">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1179">ReadWriteItem</span></span>|
|[<span data-ttu-id="f0b3f-1180">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-1181">作成</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1181">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f0b3f-1182">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1182">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="f0b3f-p184">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p184">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="f0b3f-1185">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1185">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="f0b3f-1186">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1186">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="f0b3f-p185">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p185">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0b3f-1190">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1190">Parameters</span></span>

|<span data-ttu-id="f0b3f-1191">名前</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1191">Name</span></span>| <span data-ttu-id="f0b3f-1192">型</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1192">Type</span></span>| <span data-ttu-id="f0b3f-1193">属性</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1193">Attributes</span></span>| <span data-ttu-id="f0b3f-1194">説明</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1194">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="f0b3f-1195">String</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1195">String</span></span>||<span data-ttu-id="f0b3f-p186">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-p186">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="f0b3f-1199">Object</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1199">Object</span></span>| <span data-ttu-id="f0b3f-1200">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1200">&lt;optional&gt;</span></span>|<span data-ttu-id="f0b3f-1201">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1201">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f0b3f-1202">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1202">Object</span></span>| <span data-ttu-id="f0b3f-1203">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1203">&lt;optional&gt;</span></span>|<span data-ttu-id="f0b3f-1204">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1204">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="f0b3f-1205">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1205">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="f0b3f-1206">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1206">&lt;optional&gt;</span></span>|<span data-ttu-id="f0b3f-1207">`text` の場合、Outlook on the web とデスクトップ クライアントでは現在のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1207">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="f0b3f-1208">フィールドが HTML エディターの場合、データが HTML の場合でもテキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1208">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="f0b3f-1209">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Outlook on the web では現在のスタイルが適用され、Outlook デスクトップ クライアントでは既定のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1209">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="f0b3f-1210">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1210">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="f0b3f-1211">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1211">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="f0b3f-1212">function</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1212">function</span></span>||<span data-ttu-id="f0b3f-1213">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1213">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f0b3f-1214">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1214">Requirements</span></span>

|<span data-ttu-id="f0b3f-1215">要件</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1215">Requirement</span></span>| <span data-ttu-id="f0b3f-1216">値</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1216">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0b3f-1217">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0b3f-1218">1.2</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1218">1.2</span></span>|
|[<span data-ttu-id="f0b3f-1219">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1219">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0b3f-1220">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1220">ReadWriteItem</span></span>|
|[<span data-ttu-id="f0b3f-1221">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1221">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0b3f-1222">作成</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1222">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f0b3f-1223">例</span><span class="sxs-lookup"><span data-stu-id="f0b3f-1223">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
