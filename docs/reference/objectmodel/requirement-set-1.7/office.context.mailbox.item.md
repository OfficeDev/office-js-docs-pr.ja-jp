---
title: Office.context.mailbox.item - 1.7 を設定する要件
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: 6ac795d426cf80071d7b83d5e10714f4d3a6036b
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701891"
---
# <a name="item"></a><span data-ttu-id="f72be-102">item</span><span class="sxs-lookup"><span data-stu-id="f72be-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="f72be-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="f72be-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="f72be-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="f72be-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-106">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-106">Requirements</span></span>

|<span data-ttu-id="f72be-107">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-107">Requirement</span></span>|<span data-ttu-id="f72be-108">値</span><span class="sxs-lookup"><span data-stu-id="f72be-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-110">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-110">1.0</span></span>|
|[<span data-ttu-id="f72be-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="f72be-112">Restricted</span></span>|
|[<span data-ttu-id="f72be-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-114">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f72be-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-115">Members and methods</span></span>

| <span data-ttu-id="f72be-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-116">Member</span></span> | <span data-ttu-id="f72be-117">種類</span><span class="sxs-lookup"><span data-stu-id="f72be-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f72be-118">attachments</span><span class="sxs-lookup"><span data-stu-id="f72be-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails) | <span data-ttu-id="f72be-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-119">Member</span></span> |
| [<span data-ttu-id="f72be-120">bcc</span><span class="sxs-lookup"><span data-stu-id="f72be-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="f72be-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-121">Member</span></span> |
| [<span data-ttu-id="f72be-122">body</span><span class="sxs-lookup"><span data-stu-id="f72be-122">body</span></span>](#body-bodyjavascriptapioutlook17officebody) | <span data-ttu-id="f72be-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-123">Member</span></span> |
| [<span data-ttu-id="f72be-124">cc</span><span class="sxs-lookup"><span data-stu-id="f72be-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="f72be-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-125">Member</span></span> |
| [<span data-ttu-id="f72be-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="f72be-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="f72be-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-127">Member</span></span> |
| [<span data-ttu-id="f72be-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="f72be-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="f72be-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-129">Member</span></span> |
| [<span data-ttu-id="f72be-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="f72be-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="f72be-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-131">Member</span></span> |
| [<span data-ttu-id="f72be-132">end</span><span class="sxs-lookup"><span data-stu-id="f72be-132">end</span></span>](#end-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="f72be-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-133">Member</span></span> |
| [<span data-ttu-id="f72be-134">from</span><span class="sxs-lookup"><span data-stu-id="f72be-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) | <span data-ttu-id="f72be-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-135">Member</span></span> |
| [<span data-ttu-id="f72be-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="f72be-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="f72be-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-137">Member</span></span> |
| [<span data-ttu-id="f72be-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="f72be-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="f72be-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-139">Member</span></span> |
| [<span data-ttu-id="f72be-140">itemId</span><span class="sxs-lookup"><span data-stu-id="f72be-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="f72be-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-141">Member</span></span> |
| [<span data-ttu-id="f72be-142">itemType</span><span class="sxs-lookup"><span data-stu-id="f72be-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) | <span data-ttu-id="f72be-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-143">Member</span></span> |
| [<span data-ttu-id="f72be-144">location</span><span class="sxs-lookup"><span data-stu-id="f72be-144">location</span></span>](#location-stringlocationjavascriptapioutlook17officelocation) | <span data-ttu-id="f72be-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-145">Member</span></span> |
| [<span data-ttu-id="f72be-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="f72be-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="f72be-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-147">Member</span></span> |
| [<span data-ttu-id="f72be-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="f72be-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages) | <span data-ttu-id="f72be-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-149">Member</span></span> |
| [<span data-ttu-id="f72be-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="f72be-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="f72be-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-151">Member</span></span> |
| [<span data-ttu-id="f72be-152">organizer</span><span class="sxs-lookup"><span data-stu-id="f72be-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) | <span data-ttu-id="f72be-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-153">Member</span></span> |
| [<span data-ttu-id="f72be-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="f72be-154">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) | <span data-ttu-id="f72be-155">Member</span><span class="sxs-lookup"><span data-stu-id="f72be-155">Member</span></span> |
| [<span data-ttu-id="f72be-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="f72be-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="f72be-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-157">Member</span></span> |
| [<span data-ttu-id="f72be-158">sender</span><span class="sxs-lookup"><span data-stu-id="f72be-158">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) | <span data-ttu-id="f72be-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-159">Member</span></span> |
| [<span data-ttu-id="f72be-160">seriesId</span><span class="sxs-lookup"><span data-stu-id="f72be-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="f72be-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-161">Member</span></span> |
| [<span data-ttu-id="f72be-162">start</span><span class="sxs-lookup"><span data-stu-id="f72be-162">start</span></span>](#start-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="f72be-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-163">Member</span></span> |
| [<span data-ttu-id="f72be-164">subject</span><span class="sxs-lookup"><span data-stu-id="f72be-164">subject</span></span>](#subject-stringsubjectjavascriptapioutlook17officesubject) | <span data-ttu-id="f72be-165">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-165">Member</span></span> |
| [<span data-ttu-id="f72be-166">to</span><span class="sxs-lookup"><span data-stu-id="f72be-166">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="f72be-167">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-167">Member</span></span> |
| [<span data-ttu-id="f72be-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f72be-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="f72be-169">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-169">Method</span></span> |
| [<span data-ttu-id="f72be-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="f72be-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="f72be-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-171">Method</span></span> |
| [<span data-ttu-id="f72be-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f72be-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="f72be-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-173">Method</span></span> |
| [<span data-ttu-id="f72be-174">close</span><span class="sxs-lookup"><span data-stu-id="f72be-174">close</span></span>](#close) | <span data-ttu-id="f72be-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-175">Method</span></span> |
| [<span data-ttu-id="f72be-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="f72be-176">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="f72be-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-177">Method</span></span> |
| [<span data-ttu-id="f72be-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="f72be-178">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="f72be-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-179">Method</span></span> |
| [<span data-ttu-id="f72be-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="f72be-180">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="f72be-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-181">Method</span></span> |
| [<span data-ttu-id="f72be-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="f72be-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="f72be-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-183">Method</span></span> |
| [<span data-ttu-id="f72be-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="f72be-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="f72be-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-185">Method</span></span> |
| [<span data-ttu-id="f72be-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="f72be-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="f72be-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-187">Method</span></span> |
| [<span data-ttu-id="f72be-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="f72be-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="f72be-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-189">Method</span></span> |
| [<span data-ttu-id="f72be-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="f72be-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="f72be-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-191">Method</span></span> |
| [<span data-ttu-id="f72be-192">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="f72be-192">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="f72be-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-193">Method</span></span> |
| [<span data-ttu-id="f72be-194">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="f72be-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="f72be-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-195">Method</span></span> |
| [<span data-ttu-id="f72be-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="f72be-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="f72be-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-197">Method</span></span> |
| [<span data-ttu-id="f72be-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f72be-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="f72be-199">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-199">Method</span></span> |
| [<span data-ttu-id="f72be-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="f72be-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="f72be-201">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-201">Method</span></span> |
| [<span data-ttu-id="f72be-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="f72be-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="f72be-203">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-203">Method</span></span> |
| [<span data-ttu-id="f72be-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="f72be-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="f72be-205">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="f72be-206">例</span><span class="sxs-lookup"><span data-stu-id="f72be-206">Example</span></span>

<span data-ttu-id="f72be-207">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="f72be-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
}
```

### <a name="members"></a><span data-ttu-id="f72be-208">メンバー</span><span class="sxs-lookup"><span data-stu-id="f72be-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="f72be-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f72be-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="f72be-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f72be-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f72be-212">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="f72be-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="f72be-213">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f72be-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-214">種類:</span><span class="sxs-lookup"><span data-stu-id="f72be-214">Type:</span></span>

*   <span data-ttu-id="f72be-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f72be-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-216">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-216">Requirements</span></span>

|<span data-ttu-id="f72be-217">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-217">Requirement</span></span>|<span data-ttu-id="f72be-218">値</span><span class="sxs-lookup"><span data-stu-id="f72be-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-219">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-220">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-220">1.0</span></span>|
|[<span data-ttu-id="f72be-221">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-221">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-222">ReadItem</span></span>|
|[<span data-ttu-id="f72be-223">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-223">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-224">読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-225">例</span><span class="sxs-lookup"><span data-stu-id="f72be-225">Example</span></span>

<span data-ttu-id="f72be-226">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="f72be-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
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

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="f72be-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f72be-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="f72be-228">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="f72be-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="f72be-229">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f72be-229">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-230">種類:</span><span class="sxs-lookup"><span data-stu-id="f72be-230">Type:</span></span>

*   [<span data-ttu-id="f72be-231">Recipients</span><span class="sxs-lookup"><span data-stu-id="f72be-231">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="f72be-232">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-232">Requirements</span></span>

|<span data-ttu-id="f72be-233">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-233">Requirement</span></span>|<span data-ttu-id="f72be-234">値</span><span class="sxs-lookup"><span data-stu-id="f72be-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-235">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-236">1.1</span><span class="sxs-lookup"><span data-stu-id="f72be-236">1.1</span></span>|
|[<span data-ttu-id="f72be-237">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-237">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-238">ReadItem</span></span>|
|[<span data-ttu-id="f72be-239">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-239">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-240">作成</span><span class="sxs-lookup"><span data-stu-id="f72be-240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-241">例</span><span class="sxs-lookup"><span data-stu-id="f72be-241">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="f72be-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="f72be-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="f72be-243">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="f72be-243">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-244">型:</span><span class="sxs-lookup"><span data-stu-id="f72be-244">Type:</span></span>

*   [<span data-ttu-id="f72be-245">Body</span><span class="sxs-lookup"><span data-stu-id="f72be-245">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="f72be-246">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-246">Requirements</span></span>

|<span data-ttu-id="f72be-247">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-247">Requirement</span></span>|<span data-ttu-id="f72be-248">値</span><span class="sxs-lookup"><span data-stu-id="f72be-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-249">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-250">1.1</span><span class="sxs-lookup"><span data-stu-id="f72be-250">1.1</span></span>|
|[<span data-ttu-id="f72be-251">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-252">ReadItem</span></span>|
|[<span data-ttu-id="f72be-253">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-254">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-254">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="f72be-255">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f72be-255">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="f72be-256">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f72be-256">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="f72be-257">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f72be-257">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f72be-258">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f72be-258">Read mode</span></span>

<span data-ttu-id="f72be-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f72be-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f72be-261">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f72be-261">Compose mode</span></span>

<span data-ttu-id="f72be-262">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-262">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-263">種類:</span><span class="sxs-lookup"><span data-stu-id="f72be-263">Type:</span></span>

*   <span data-ttu-id="f72be-264">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f72be-264">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-265">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-265">Requirements</span></span>

|<span data-ttu-id="f72be-266">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-266">Requirement</span></span>|<span data-ttu-id="f72be-267">値</span><span class="sxs-lookup"><span data-stu-id="f72be-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-268">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-269">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-269">1.0</span></span>|
|[<span data-ttu-id="f72be-270">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-270">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-271">ReadItem</span></span>|
|[<span data-ttu-id="f72be-272">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-272">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-273">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="f72be-273">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-274">例</span><span class="sxs-lookup"><span data-stu-id="f72be-274">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="f72be-275">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="f72be-275">(nullable) conversationId :String</span></span>

<span data-ttu-id="f72be-276">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="f72be-276">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="f72be-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="f72be-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="f72be-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-281">型:</span><span class="sxs-lookup"><span data-stu-id="f72be-281">Type:</span></span>

*   <span data-ttu-id="f72be-282">String</span><span class="sxs-lookup"><span data-stu-id="f72be-282">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-283">Requirements</span><span class="sxs-lookup"><span data-stu-id="f72be-283">Requirements</span></span>

|<span data-ttu-id="f72be-284">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-284">Requirement</span></span>|<span data-ttu-id="f72be-285">値</span><span class="sxs-lookup"><span data-stu-id="f72be-285">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-286">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-286">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-287">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-287">1.0</span></span>|
|[<span data-ttu-id="f72be-288">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-288">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-289">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-289">ReadItem</span></span>|
|[<span data-ttu-id="f72be-290">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-290">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-291">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-291">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="f72be-292">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="f72be-292">dateTimeCreated :Date</span></span>

<span data-ttu-id="f72be-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f72be-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-295">種類:</span><span class="sxs-lookup"><span data-stu-id="f72be-295">Type:</span></span>

*   <span data-ttu-id="f72be-296">日付</span><span class="sxs-lookup"><span data-stu-id="f72be-296">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-297">Requirements</span><span class="sxs-lookup"><span data-stu-id="f72be-297">Requirements</span></span>

|<span data-ttu-id="f72be-298">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-298">Requirement</span></span>|<span data-ttu-id="f72be-299">値</span><span class="sxs-lookup"><span data-stu-id="f72be-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-300">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-300">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-301">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-301">1.0</span></span>|
|[<span data-ttu-id="f72be-302">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-302">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-303">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-303">ReadItem</span></span>|
|[<span data-ttu-id="f72be-304">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-304">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-305">読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-305">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-306">例</span><span class="sxs-lookup"><span data-stu-id="f72be-306">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="f72be-307">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="f72be-307">dateTimeModified :Date</span></span>

<span data-ttu-id="f72be-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f72be-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f72be-310">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f72be-310">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-311">種類:</span><span class="sxs-lookup"><span data-stu-id="f72be-311">Type:</span></span>

*   <span data-ttu-id="f72be-312">日付</span><span class="sxs-lookup"><span data-stu-id="f72be-312">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-313">Requirements</span><span class="sxs-lookup"><span data-stu-id="f72be-313">Requirements</span></span>

|<span data-ttu-id="f72be-314">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-314">Requirement</span></span>|<span data-ttu-id="f72be-315">値</span><span class="sxs-lookup"><span data-stu-id="f72be-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-316">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-317">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-317">1.0</span></span>|
|[<span data-ttu-id="f72be-318">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-318">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-319">ReadItem</span></span>|
|[<span data-ttu-id="f72be-320">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-320">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-321">読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-321">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-322">例</span><span class="sxs-lookup"><span data-stu-id="f72be-322">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="f72be-323">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="f72be-323">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="f72be-324">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f72be-324">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="f72be-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="f72be-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f72be-327">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f72be-327">Read mode</span></span>

<span data-ttu-id="f72be-328">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-328">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f72be-329">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f72be-329">Compose mode</span></span>

<span data-ttu-id="f72be-330">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-330">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="f72be-331">[`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f72be-331">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-332">型:</span><span class="sxs-lookup"><span data-stu-id="f72be-332">Type:</span></span>

*   <span data-ttu-id="f72be-333">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="f72be-333">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-334">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-334">Requirements</span></span>

|<span data-ttu-id="f72be-335">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-335">Requirement</span></span>|<span data-ttu-id="f72be-336">値</span><span class="sxs-lookup"><span data-stu-id="f72be-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-337">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-338">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-338">1.0</span></span>|
|[<span data-ttu-id="f72be-339">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-339">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-340">ReadItem</span></span>|
|[<span data-ttu-id="f72be-341">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-341">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-342">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-342">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-343">例</span><span class="sxs-lookup"><span data-stu-id="f72be-343">Example</span></span>

<span data-ttu-id="f72be-344">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="f72be-344">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="f72be-345">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="f72be-345">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="f72be-346">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="f72be-346">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="f72be-p112">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="f72be-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f72be-349">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="f72be-349">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f72be-350">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f72be-350">Read mode</span></span>

<span data-ttu-id="f72be-351">`from` プロパティは `EmailAddressDetails` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-351">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="f72be-352">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f72be-352">Compose mode</span></span>

<span data-ttu-id="f72be-353">`from` プロパティは From 値を取得するメソッドを提供する `From` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-353">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f72be-354">型:</span><span class="sxs-lookup"><span data-stu-id="f72be-354">Type:</span></span>

*   <span data-ttu-id="f72be-355">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="f72be-355">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-356">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-356">Requirements</span></span>

|<span data-ttu-id="f72be-357">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-357">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="f72be-358">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-358">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-359">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-359">1.0</span></span>|<span data-ttu-id="f72be-360">1.7</span><span class="sxs-lookup"><span data-stu-id="f72be-360">1.7</span></span>|
|[<span data-ttu-id="f72be-361">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-362">ReadItem</span></span>|<span data-ttu-id="f72be-363">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f72be-363">ReadWriteItem</span></span>|
|[<span data-ttu-id="f72be-364">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-365">Read</span><span class="sxs-lookup"><span data-stu-id="f72be-365">Read</span></span>|<span data-ttu-id="f72be-366">Compose</span><span class="sxs-lookup"><span data-stu-id="f72be-366">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="f72be-367">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="f72be-367">internetMessageId :String</span></span>

<span data-ttu-id="f72be-p113">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f72be-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-370">型:</span><span class="sxs-lookup"><span data-stu-id="f72be-370">Type:</span></span>

*   <span data-ttu-id="f72be-371">String</span><span class="sxs-lookup"><span data-stu-id="f72be-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-372">Requirements</span><span class="sxs-lookup"><span data-stu-id="f72be-372">Requirements</span></span>

|<span data-ttu-id="f72be-373">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-373">Requirement</span></span>|<span data-ttu-id="f72be-374">値</span><span class="sxs-lookup"><span data-stu-id="f72be-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-375">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-376">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-376">1.0</span></span>|
|[<span data-ttu-id="f72be-377">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-378">ReadItem</span></span>|
|[<span data-ttu-id="f72be-379">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-380">読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-381">例</span><span class="sxs-lookup"><span data-stu-id="f72be-381">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="f72be-382">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="f72be-382">itemClass :String</span></span>

<span data-ttu-id="f72be-p114">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f72be-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="f72be-p115">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="f72be-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="f72be-387">型</span><span class="sxs-lookup"><span data-stu-id="f72be-387">Type</span></span>|<span data-ttu-id="f72be-388">説明</span><span class="sxs-lookup"><span data-stu-id="f72be-388">Description</span></span>|<span data-ttu-id="f72be-389">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="f72be-389">item class</span></span>|
|---|---|---|
|<span data-ttu-id="f72be-390">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="f72be-390">Appointment items</span></span>|<span data-ttu-id="f72be-391">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="f72be-391">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="f72be-392">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="f72be-392">Message items</span></span>|<span data-ttu-id="f72be-393">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="f72be-393">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="f72be-394">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="f72be-394">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-395">型:</span><span class="sxs-lookup"><span data-stu-id="f72be-395">Type:</span></span>

*   <span data-ttu-id="f72be-396">String</span><span class="sxs-lookup"><span data-stu-id="f72be-396">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-397">Requirements</span><span class="sxs-lookup"><span data-stu-id="f72be-397">Requirements</span></span>

|<span data-ttu-id="f72be-398">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-398">Requirement</span></span>|<span data-ttu-id="f72be-399">値</span><span class="sxs-lookup"><span data-stu-id="f72be-399">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-400">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-401">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-401">1.0</span></span>|
|[<span data-ttu-id="f72be-402">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-402">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-403">ReadItem</span></span>|
|[<span data-ttu-id="f72be-404">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-404">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-405">読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-405">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-406">例</span><span class="sxs-lookup"><span data-stu-id="f72be-406">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="f72be-407">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="f72be-407">(nullable) itemId :String</span></span>

<span data-ttu-id="f72be-p116">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f72be-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f72be-410">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="f72be-410">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="f72be-411">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="f72be-411">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="f72be-412">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f72be-412">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="f72be-413">詳細は、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f72be-413">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="f72be-p118">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-416">型:</span><span class="sxs-lookup"><span data-stu-id="f72be-416">Type:</span></span>

*   <span data-ttu-id="f72be-417">String</span><span class="sxs-lookup"><span data-stu-id="f72be-417">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-418">Requirements</span><span class="sxs-lookup"><span data-stu-id="f72be-418">Requirements</span></span>

|<span data-ttu-id="f72be-419">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-419">Requirement</span></span>|<span data-ttu-id="f72be-420">値</span><span class="sxs-lookup"><span data-stu-id="f72be-420">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-421">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-421">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-422">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-422">1.0</span></span>|
|[<span data-ttu-id="f72be-423">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-423">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-424">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-424">ReadItem</span></span>|
|[<span data-ttu-id="f72be-425">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-425">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-426">読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-426">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-427">例</span><span class="sxs-lookup"><span data-stu-id="f72be-427">Example</span></span>

<span data-ttu-id="f72be-p119">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="f72be-430">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="f72be-430">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="f72be-431">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="f72be-431">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="f72be-432">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="f72be-432">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-433">種類:</span><span class="sxs-lookup"><span data-stu-id="f72be-433">Type:</span></span>

*   [<span data-ttu-id="f72be-434">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="f72be-434">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="f72be-435">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-435">Requirements</span></span>

|<span data-ttu-id="f72be-436">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-436">Requirement</span></span>|<span data-ttu-id="f72be-437">値</span><span class="sxs-lookup"><span data-stu-id="f72be-437">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-438">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-438">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-439">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-439">1.0</span></span>|
|[<span data-ttu-id="f72be-440">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-440">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-441">ReadItem</span></span>|
|[<span data-ttu-id="f72be-442">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-442">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-443">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-443">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-444">例</span><span class="sxs-lookup"><span data-stu-id="f72be-444">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="f72be-445">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="f72be-445">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="f72be-446">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f72be-446">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f72be-447">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f72be-447">Read mode</span></span>

<span data-ttu-id="f72be-448">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-448">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f72be-449">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f72be-449">Compose mode</span></span>

<span data-ttu-id="f72be-450">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-450">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-451">種類:</span><span class="sxs-lookup"><span data-stu-id="f72be-451">Type:</span></span>

*   <span data-ttu-id="f72be-452">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="f72be-452">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-453">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-453">Requirements</span></span>

|<span data-ttu-id="f72be-454">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-454">Requirement</span></span>|<span data-ttu-id="f72be-455">値</span><span class="sxs-lookup"><span data-stu-id="f72be-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-456">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-457">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-457">1.0</span></span>|
|[<span data-ttu-id="f72be-458">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-458">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-459">ReadItem</span></span>|
|[<span data-ttu-id="f72be-460">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-460">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-461">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-461">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-462">例</span><span class="sxs-lookup"><span data-stu-id="f72be-462">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="f72be-463">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="f72be-463">normalizedSubject :String</span></span>

<span data-ttu-id="f72be-p120">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f72be-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="f72be-p121">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="f72be-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-468">型:</span><span class="sxs-lookup"><span data-stu-id="f72be-468">Type:</span></span>

*   <span data-ttu-id="f72be-469">String</span><span class="sxs-lookup"><span data-stu-id="f72be-469">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-470">Requirements</span><span class="sxs-lookup"><span data-stu-id="f72be-470">Requirements</span></span>

|<span data-ttu-id="f72be-471">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-471">Requirement</span></span>|<span data-ttu-id="f72be-472">値</span><span class="sxs-lookup"><span data-stu-id="f72be-472">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-473">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-473">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-474">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-474">1.0</span></span>|
|[<span data-ttu-id="f72be-475">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-475">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-476">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-476">ReadItem</span></span>|
|[<span data-ttu-id="f72be-477">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-477">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-478">読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-478">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-479">例</span><span class="sxs-lookup"><span data-stu-id="f72be-479">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="f72be-480">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="f72be-480">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="f72be-481">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="f72be-481">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-482">型:</span><span class="sxs-lookup"><span data-stu-id="f72be-482">Type:</span></span>

*   [<span data-ttu-id="f72be-483">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="f72be-483">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="f72be-484">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-484">Requirements</span></span>

|<span data-ttu-id="f72be-485">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-485">Requirement</span></span>|<span data-ttu-id="f72be-486">値</span><span class="sxs-lookup"><span data-stu-id="f72be-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-487">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-487">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-488">1.3</span><span class="sxs-lookup"><span data-stu-id="f72be-488">1.3</span></span>|
|[<span data-ttu-id="f72be-489">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-490">ReadItem</span></span>|
|[<span data-ttu-id="f72be-491">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-492">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-492">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="f72be-493">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f72be-493">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="f72be-494">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f72be-494">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="f72be-495">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f72be-495">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f72be-496">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f72be-496">Read mode</span></span>

<span data-ttu-id="f72be-497">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-497">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f72be-498">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f72be-498">Compose mode</span></span>

<span data-ttu-id="f72be-499">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-499">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-500">種類:</span><span class="sxs-lookup"><span data-stu-id="f72be-500">Type:</span></span>

*   <span data-ttu-id="f72be-501">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f72be-501">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-502">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-502">Requirements</span></span>

|<span data-ttu-id="f72be-503">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-503">Requirement</span></span>|<span data-ttu-id="f72be-504">値</span><span class="sxs-lookup"><span data-stu-id="f72be-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-505">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-506">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-506">1.0</span></span>|
|[<span data-ttu-id="f72be-507">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-507">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-508">ReadItem</span></span>|
|[<span data-ttu-id="f72be-509">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-509">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-510">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-510">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-511">例</span><span class="sxs-lookup"><span data-stu-id="f72be-511">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="f72be-512">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="f72be-512">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="f72be-513">指定の会議の開催者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="f72be-513">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f72be-514">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f72be-514">Read mode</span></span>

<span data-ttu-id="f72be-515">`organizer` プロパティは、会議開催者を表す [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-515">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f72be-516">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f72be-516">Compose mode</span></span>

<span data-ttu-id="f72be-517">`organizer` プロパティは Organizer 値を取得するメソッドを提供する [Organizer](/javascript/api/outlook_1_7/office.organizer) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-517">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-518">種類:</span><span class="sxs-lookup"><span data-stu-id="f72be-518">Type:</span></span>

*   <span data-ttu-id="f72be-519">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="f72be-519">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-520">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-520">Requirements</span></span>

|<span data-ttu-id="f72be-521">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-521">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="f72be-522">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-523">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-523">1.0</span></span>|<span data-ttu-id="f72be-524">1.7</span><span class="sxs-lookup"><span data-stu-id="f72be-524">1.7</span></span>|
|[<span data-ttu-id="f72be-525">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-525">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-526">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-526">ReadItem</span></span>|<span data-ttu-id="f72be-527">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f72be-527">ReadWriteItem</span></span>|
|[<span data-ttu-id="f72be-528">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-528">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-529">Read</span><span class="sxs-lookup"><span data-stu-id="f72be-529">Read</span></span>|<span data-ttu-id="f72be-530">Compose</span><span class="sxs-lookup"><span data-stu-id="f72be-530">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-531">例</span><span class="sxs-lookup"><span data-stu-id="f72be-531">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="f72be-532">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="f72be-532">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="f72be-533">予定の定期的なパターンを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f72be-533">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="f72be-534">会議出席依頼の定期的なパターンを取得します。</span><span class="sxs-lookup"><span data-stu-id="f72be-534">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="f72be-535">予定アイテムの閲覧モードと新規作成モード。</span><span class="sxs-lookup"><span data-stu-id="f72be-535">Read and compose modes for appointment items.</span></span> <span data-ttu-id="f72be-536">会議出席依頼アイテムの閲覧モード。</span><span class="sxs-lookup"><span data-stu-id="f72be-536">Read mode for meeting request items.</span></span>

<span data-ttu-id="f72be-537">`recurrence` プロパティは、アイテムがシリーズか、シリーズに含まれるインスタンスの場合、定期的な予定または会議出席依頼に対して [recurrence](/javascript/api/outlook_1_7/office.recurrence) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-537">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="f72be-538">`null` は、単発の予定および単発の予定の会議出席依頼に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-538">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="f72be-539">`undefined` は、会議出席依頼ではないメッセージに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-539">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="f72be-540">注: 会議出席依頼の `itemClass` 値は IPM.Schedule.Meeting.Request です。</span><span class="sxs-lookup"><span data-stu-id="f72be-540">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="f72be-541">注: recurrence オブジェクトが `null` の場合、オブジェクトがシリーズの一部ではなく、1 つの単発の予定または 1 つの単発の予定の会議出席依頼であることを示します。</span><span class="sxs-lookup"><span data-stu-id="f72be-541">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-542">型:</span><span class="sxs-lookup"><span data-stu-id="f72be-542">Type:</span></span>

* [<span data-ttu-id="f72be-543">Recurrence</span><span class="sxs-lookup"><span data-stu-id="f72be-543">Recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="f72be-544">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-544">Requirement</span></span>|<span data-ttu-id="f72be-545">値</span><span class="sxs-lookup"><span data-stu-id="f72be-545">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-546">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-547">1.7</span><span class="sxs-lookup"><span data-stu-id="f72be-547">1.7</span></span>|
|[<span data-ttu-id="f72be-548">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-548">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-549">ReadItem</span></span>|
|[<span data-ttu-id="f72be-550">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-550">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-551">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="f72be-551">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="f72be-552">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f72be-552">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="f72be-553">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f72be-553">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="f72be-554">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f72be-554">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f72be-555">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f72be-555">Read mode</span></span>

<span data-ttu-id="f72be-556">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-556">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f72be-557">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f72be-557">Compose mode</span></span>

<span data-ttu-id="f72be-558">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-558">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-559">種類:</span><span class="sxs-lookup"><span data-stu-id="f72be-559">Type:</span></span>

*   <span data-ttu-id="f72be-560">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f72be-560">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-561">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-561">Requirements</span></span>

|<span data-ttu-id="f72be-562">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-562">Requirement</span></span>|<span data-ttu-id="f72be-563">値</span><span class="sxs-lookup"><span data-stu-id="f72be-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-564">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-565">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-565">1.0</span></span>|
|[<span data-ttu-id="f72be-566">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-567">ReadItem</span></span>|
|[<span data-ttu-id="f72be-568">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-569">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-570">例</span><span class="sxs-lookup"><span data-stu-id="f72be-570">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="f72be-571">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f72be-571">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="f72be-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f72be-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="f72be-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="f72be-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f72be-576">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="f72be-576">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-577">型:</span><span class="sxs-lookup"><span data-stu-id="f72be-577">Type:</span></span>

*   [<span data-ttu-id="f72be-578">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f72be-578">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f72be-579">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-579">Requirements</span></span>

|<span data-ttu-id="f72be-580">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-580">Requirement</span></span>|<span data-ttu-id="f72be-581">値</span><span class="sxs-lookup"><span data-stu-id="f72be-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-582">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-583">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-583">1.0</span></span>|
|[<span data-ttu-id="f72be-584">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-584">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-585">ReadItem</span></span>|
|[<span data-ttu-id="f72be-586">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-586">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-587">読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-587">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-588">例</span><span class="sxs-lookup"><span data-stu-id="f72be-588">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="f72be-589">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="f72be-589">(nullable) seriesId :String</span></span>

<span data-ttu-id="f72be-590">あるインスタンスが属するシリーズの ID を取得します。</span><span class="sxs-lookup"><span data-stu-id="f72be-590">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="f72be-591">OWA と Outlook では、`seriesId` はこのアイテムが属する親 (シリーズ) アイテムの Exchange Web Services (EWS) ID を返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-591">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="f72be-592">ただし、iOS と Android の場合、`seriesId` は親アイテムの REST ID を返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-592">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="f72be-593">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="f72be-593">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="f72be-594">`seriesId` プロパティは、Outlook REST API で使用される Outlook ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="f72be-594">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="f72be-595">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f72be-595">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="f72be-596">詳細については、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f72be-596">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="f72be-597">`seriesId` プロパティは、単発の予定、シリーズ アイテム、会議出席依頼など、親アイテムを持たないアイテムに対して `null` を返し、会議出席依頼ではないその他のアイテムに対して `undefined` を返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-597">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-598">型:</span><span class="sxs-lookup"><span data-stu-id="f72be-598">Type:</span></span>

* <span data-ttu-id="f72be-599">String</span><span class="sxs-lookup"><span data-stu-id="f72be-599">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-600">Requirements</span><span class="sxs-lookup"><span data-stu-id="f72be-600">Requirements</span></span>

|<span data-ttu-id="f72be-601">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-601">Requirement</span></span>|<span data-ttu-id="f72be-602">値</span><span class="sxs-lookup"><span data-stu-id="f72be-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-603">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-603">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-604">1.7</span><span class="sxs-lookup"><span data-stu-id="f72be-604">1.7</span></span>|
|[<span data-ttu-id="f72be-605">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-605">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-606">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-606">ReadItem</span></span>|
|[<span data-ttu-id="f72be-607">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-607">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-608">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-608">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-609">例</span><span class="sxs-lookup"><span data-stu-id="f72be-609">Example</span></span>

```js
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="f72be-610">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="f72be-610">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="f72be-611">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f72be-611">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="f72be-p130">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="f72be-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f72be-614">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f72be-614">Read mode</span></span>

<span data-ttu-id="f72be-615">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-615">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f72be-616">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f72be-616">Compose mode</span></span>

<span data-ttu-id="f72be-617">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-617">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="f72be-618">[`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f72be-618">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-619">種類:</span><span class="sxs-lookup"><span data-stu-id="f72be-619">Type:</span></span>

*   <span data-ttu-id="f72be-620">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="f72be-620">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-621">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-621">Requirements</span></span>

|<span data-ttu-id="f72be-622">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-622">Requirement</span></span>|<span data-ttu-id="f72be-623">値</span><span class="sxs-lookup"><span data-stu-id="f72be-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-624">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-625">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-625">1.0</span></span>|
|[<span data-ttu-id="f72be-626">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-627">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-627">ReadItem</span></span>|
|[<span data-ttu-id="f72be-628">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-629">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-629">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-630">例</span><span class="sxs-lookup"><span data-stu-id="f72be-630">Example</span></span>

<span data-ttu-id="f72be-631">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="f72be-631">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="f72be-632">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f72be-632">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="f72be-633">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f72be-633">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="f72be-634">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f72be-634">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f72be-635">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f72be-635">Read mode</span></span>

<span data-ttu-id="f72be-p131">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="f72be-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="f72be-638">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f72be-638">Compose mode</span></span>

<span data-ttu-id="f72be-639">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-639">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f72be-640">型:</span><span class="sxs-lookup"><span data-stu-id="f72be-640">Type:</span></span>

*   <span data-ttu-id="f72be-641">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f72be-641">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-642">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-642">Requirements</span></span>

|<span data-ttu-id="f72be-643">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-643">Requirement</span></span>|<span data-ttu-id="f72be-644">値</span><span class="sxs-lookup"><span data-stu-id="f72be-644">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-645">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-645">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-646">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-646">1.0</span></span>|
|[<span data-ttu-id="f72be-647">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-647">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-648">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-648">ReadItem</span></span>|
|[<span data-ttu-id="f72be-649">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-649">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-650">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-650">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="f72be-651">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f72be-651">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="f72be-652">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f72be-652">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="f72be-653">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f72be-653">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f72be-654">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f72be-654">Read mode</span></span>

<span data-ttu-id="f72be-p133">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f72be-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f72be-657">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f72be-657">Compose mode</span></span>

<span data-ttu-id="f72be-658">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-658">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="f72be-659">種類:</span><span class="sxs-lookup"><span data-stu-id="f72be-659">Type:</span></span>

*   <span data-ttu-id="f72be-660">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f72be-660">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-661">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-661">Requirements</span></span>

|<span data-ttu-id="f72be-662">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-662">Requirement</span></span>|<span data-ttu-id="f72be-663">値</span><span class="sxs-lookup"><span data-stu-id="f72be-663">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-664">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-664">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-665">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-665">1.0</span></span>|
|[<span data-ttu-id="f72be-666">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-666">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-667">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-667">ReadItem</span></span>|
|[<span data-ttu-id="f72be-668">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-668">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-669">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-669">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-670">例</span><span class="sxs-lookup"><span data-stu-id="f72be-670">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="f72be-671">メソッド</span><span class="sxs-lookup"><span data-stu-id="f72be-671">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="f72be-672">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f72be-672">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f72be-673">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="f72be-673">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="f72be-674">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="f72be-674">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="f72be-675">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="f72be-675">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f72be-676">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f72be-676">Parameters:</span></span>
|<span data-ttu-id="f72be-677">名前</span><span class="sxs-lookup"><span data-stu-id="f72be-677">Name</span></span>|<span data-ttu-id="f72be-678">型</span><span class="sxs-lookup"><span data-stu-id="f72be-678">Type</span></span>|<span data-ttu-id="f72be-679">属性</span><span class="sxs-lookup"><span data-stu-id="f72be-679">Attributes</span></span>|<span data-ttu-id="f72be-680">説明</span><span class="sxs-lookup"><span data-stu-id="f72be-680">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="f72be-681">String</span><span class="sxs-lookup"><span data-stu-id="f72be-681">String</span></span>||<span data-ttu-id="f72be-p134">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="f72be-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="f72be-684">String</span><span class="sxs-lookup"><span data-stu-id="f72be-684">String</span></span>||<span data-ttu-id="f72be-p135">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="f72be-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="f72be-687">Object</span><span class="sxs-lookup"><span data-stu-id="f72be-687">Object</span></span>|<span data-ttu-id="f72be-688">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-688">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-689">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f72be-689">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f72be-690">Object</span><span class="sxs-lookup"><span data-stu-id="f72be-690">Object</span></span>|<span data-ttu-id="f72be-691">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-691">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-692">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f72be-692">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="f72be-693">Boolean</span><span class="sxs-lookup"><span data-stu-id="f72be-693">Boolean</span></span>|<span data-ttu-id="f72be-694">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-694">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-695">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="f72be-695">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="f72be-696">function</span><span class="sxs-lookup"><span data-stu-id="f72be-696">function</span></span>|<span data-ttu-id="f72be-697">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-697">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-698">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-698">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f72be-699">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-699">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f72be-700">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="f72be-700">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f72be-701">エラー</span><span class="sxs-lookup"><span data-stu-id="f72be-701">Errors</span></span>

|<span data-ttu-id="f72be-702">エラー コード</span><span class="sxs-lookup"><span data-stu-id="f72be-702">Error code</span></span>|<span data-ttu-id="f72be-703">説明</span><span class="sxs-lookup"><span data-stu-id="f72be-703">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="f72be-704">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="f72be-704">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="f72be-705">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="f72be-705">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="f72be-706">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="f72be-706">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f72be-707">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-707">Requirements</span></span>

|<span data-ttu-id="f72be-708">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-708">Requirement</span></span>|<span data-ttu-id="f72be-709">値</span><span class="sxs-lookup"><span data-stu-id="f72be-709">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-710">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-710">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-711">1.1</span><span class="sxs-lookup"><span data-stu-id="f72be-711">1.1</span></span>|
|[<span data-ttu-id="f72be-712">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-712">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-713">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f72be-713">ReadWriteItem</span></span>|
|[<span data-ttu-id="f72be-714">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-714">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-715">作成</span><span class="sxs-lookup"><span data-stu-id="f72be-715">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f72be-716">例</span><span class="sxs-lookup"><span data-stu-id="f72be-716">Examples</span></span>

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

<span data-ttu-id="f72be-717">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="f72be-717">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="f72be-718">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f72be-718">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="f72be-719">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="f72be-719">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="f72be-720">現在、サポートされているイベントの種類は `Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged`、`Office.EventType.RecurrenceChanged` です。</span><span class="sxs-lookup"><span data-stu-id="f72be-720">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="f72be-721">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f72be-721">Parameters:</span></span>

| <span data-ttu-id="f72be-722">名前</span><span class="sxs-lookup"><span data-stu-id="f72be-722">Name</span></span> | <span data-ttu-id="f72be-723">型</span><span class="sxs-lookup"><span data-stu-id="f72be-723">Type</span></span> | <span data-ttu-id="f72be-724">属性</span><span class="sxs-lookup"><span data-stu-id="f72be-724">Attributes</span></span> | <span data-ttu-id="f72be-725">説明</span><span class="sxs-lookup"><span data-stu-id="f72be-725">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="f72be-726">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="f72be-726">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="f72be-727">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="f72be-727">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="f72be-728">Function</span><span class="sxs-lookup"><span data-stu-id="f72be-728">Function</span></span> || <span data-ttu-id="f72be-p136">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="f72be-p136">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="f72be-732">Object</span><span class="sxs-lookup"><span data-stu-id="f72be-732">Object</span></span> | <span data-ttu-id="f72be-733">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-733">&lt;optional&gt;</span></span> | <span data-ttu-id="f72be-734">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f72be-734">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="f72be-735">Object</span><span class="sxs-lookup"><span data-stu-id="f72be-735">Object</span></span> | <span data-ttu-id="f72be-736">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-736">&lt;optional&gt;</span></span> | <span data-ttu-id="f72be-737">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f72be-737">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="f72be-738">function</span><span class="sxs-lookup"><span data-stu-id="f72be-738">function</span></span>| <span data-ttu-id="f72be-739">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-739">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-740">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-740">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f72be-741">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-741">Requirements</span></span>

|<span data-ttu-id="f72be-742">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-742">Requirement</span></span>| <span data-ttu-id="f72be-743">値</span><span class="sxs-lookup"><span data-stu-id="f72be-743">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-744">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-744">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f72be-745">1.7</span><span class="sxs-lookup"><span data-stu-id="f72be-745">1.7</span></span> |
|[<span data-ttu-id="f72be-746">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-746">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f72be-747">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-747">ReadItem</span></span> |
|[<span data-ttu-id="f72be-748">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-748">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f72be-749">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-749">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="f72be-750">例</span><span class="sxs-lookup"><span data-stu-id="f72be-750">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="f72be-751">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f72be-751">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f72be-752">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="f72be-752">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="f72be-p137">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="f72be-p137">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="f72be-756">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="f72be-756">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="f72be-757">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="f72be-757">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f72be-758">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f72be-758">Parameters:</span></span>

|<span data-ttu-id="f72be-759">名前</span><span class="sxs-lookup"><span data-stu-id="f72be-759">Name</span></span>|<span data-ttu-id="f72be-760">型</span><span class="sxs-lookup"><span data-stu-id="f72be-760">Type</span></span>|<span data-ttu-id="f72be-761">属性</span><span class="sxs-lookup"><span data-stu-id="f72be-761">Attributes</span></span>|<span data-ttu-id="f72be-762">説明</span><span class="sxs-lookup"><span data-stu-id="f72be-762">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="f72be-763">String</span><span class="sxs-lookup"><span data-stu-id="f72be-763">String</span></span>||<span data-ttu-id="f72be-p138">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="f72be-p138">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="f72be-766">String</span><span class="sxs-lookup"><span data-stu-id="f72be-766">String</span></span>||<span data-ttu-id="f72be-p139">添付するアイテムの件名。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="f72be-p139">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="f72be-769">Object</span><span class="sxs-lookup"><span data-stu-id="f72be-769">Object</span></span>|<span data-ttu-id="f72be-770">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-770">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-771">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f72be-771">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f72be-772">Object</span><span class="sxs-lookup"><span data-stu-id="f72be-772">Object</span></span>|<span data-ttu-id="f72be-773">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-773">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-774">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f72be-774">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f72be-775">function</span><span class="sxs-lookup"><span data-stu-id="f72be-775">function</span></span>|<span data-ttu-id="f72be-776">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-776">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-777">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-777">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f72be-778">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-778">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f72be-779">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="f72be-779">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f72be-780">エラー</span><span class="sxs-lookup"><span data-stu-id="f72be-780">Errors</span></span>

|<span data-ttu-id="f72be-781">エラー コード</span><span class="sxs-lookup"><span data-stu-id="f72be-781">Error code</span></span>|<span data-ttu-id="f72be-782">説明</span><span class="sxs-lookup"><span data-stu-id="f72be-782">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="f72be-783">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="f72be-783">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f72be-784">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-784">Requirements</span></span>

|<span data-ttu-id="f72be-785">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-785">Requirement</span></span>|<span data-ttu-id="f72be-786">値</span><span class="sxs-lookup"><span data-stu-id="f72be-786">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-787">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-787">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-788">1.1</span><span class="sxs-lookup"><span data-stu-id="f72be-788">1.1</span></span>|
|[<span data-ttu-id="f72be-789">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-789">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-790">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f72be-790">ReadWriteItem</span></span>|
|[<span data-ttu-id="f72be-791">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-791">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-792">作成</span><span class="sxs-lookup"><span data-stu-id="f72be-792">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-793">例</span><span class="sxs-lookup"><span data-stu-id="f72be-793">Example</span></span>

<span data-ttu-id="f72be-794">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-794">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
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

####  <a name="close"></a><span data-ttu-id="f72be-795">close()</span><span class="sxs-lookup"><span data-stu-id="f72be-795">close()</span></span>

<span data-ttu-id="f72be-796">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="f72be-796">Closes the current item that is being composed.</span></span>

<span data-ttu-id="f72be-p140">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="f72be-p140">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="f72be-799">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-799">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="f72be-800">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="f72be-800">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-801">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-801">Requirements</span></span>

|<span data-ttu-id="f72be-802">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-802">Requirement</span></span>|<span data-ttu-id="f72be-803">値</span><span class="sxs-lookup"><span data-stu-id="f72be-803">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-804">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-804">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-805">1.3</span><span class="sxs-lookup"><span data-stu-id="f72be-805">1.3</span></span>|
|[<span data-ttu-id="f72be-806">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-806">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-807">制限あり</span><span class="sxs-lookup"><span data-stu-id="f72be-807">Restricted</span></span>|
|[<span data-ttu-id="f72be-808">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-808">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-809">作成</span><span class="sxs-lookup"><span data-stu-id="f72be-809">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="f72be-810">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="f72be-810">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="f72be-811">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-811">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f72be-812">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f72be-812">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f72be-813">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-813">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f72be-814">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="f72be-814">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="f72be-p141">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="f72be-p141">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f72be-818">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f72be-818">Parameters:</span></span>

|<span data-ttu-id="f72be-819">名前</span><span class="sxs-lookup"><span data-stu-id="f72be-819">Name</span></span>|<span data-ttu-id="f72be-820">型</span><span class="sxs-lookup"><span data-stu-id="f72be-820">Type</span></span>|<span data-ttu-id="f72be-821">属性</span><span class="sxs-lookup"><span data-stu-id="f72be-821">Attributes</span></span>|<span data-ttu-id="f72be-822">説明</span><span class="sxs-lookup"><span data-stu-id="f72be-822">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="f72be-823">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="f72be-823">String &#124; Object</span></span>||<span data-ttu-id="f72be-p142">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f72be-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f72be-826">**または**</span><span class="sxs-lookup"><span data-stu-id="f72be-826">**OR**</span></span><br/><span data-ttu-id="f72be-p143">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="f72be-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="f72be-829">String</span><span class="sxs-lookup"><span data-stu-id="f72be-829">String</span></span>|<span data-ttu-id="f72be-830">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-830">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-p144">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f72be-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="f72be-833">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-833">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="f72be-834">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-834">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-835">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="f72be-835">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="f72be-836">String</span><span class="sxs-lookup"><span data-stu-id="f72be-836">String</span></span>||<span data-ttu-id="f72be-p145">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="f72be-p145">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="f72be-839">String</span><span class="sxs-lookup"><span data-stu-id="f72be-839">String</span></span>||<span data-ttu-id="f72be-840">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="f72be-840">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="f72be-841">String</span><span class="sxs-lookup"><span data-stu-id="f72be-841">String</span></span>||<span data-ttu-id="f72be-p146">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="f72be-p146">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="f72be-844">ブール値</span><span class="sxs-lookup"><span data-stu-id="f72be-844">Boolean</span></span>||<span data-ttu-id="f72be-p147">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="f72be-p147">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="f72be-847">String</span><span class="sxs-lookup"><span data-stu-id="f72be-847">String</span></span>||<span data-ttu-id="f72be-p148">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="f72be-p148">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="f72be-851">function</span><span class="sxs-lookup"><span data-stu-id="f72be-851">function</span></span>|<span data-ttu-id="f72be-852">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-852">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-853">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-853">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f72be-854">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-854">Requirements</span></span>

|<span data-ttu-id="f72be-855">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-855">Requirement</span></span>|<span data-ttu-id="f72be-856">値</span><span class="sxs-lookup"><span data-stu-id="f72be-856">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-857">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-857">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-858">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-858">1.0</span></span>|
|[<span data-ttu-id="f72be-859">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-859">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-860">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-860">ReadItem</span></span>|
|[<span data-ttu-id="f72be-861">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-861">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-862">読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-862">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f72be-863">例</span><span class="sxs-lookup"><span data-stu-id="f72be-863">Examples</span></span>

<span data-ttu-id="f72be-864">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="f72be-864">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="f72be-865">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="f72be-865">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="f72be-866">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="f72be-866">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f72be-867">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="f72be-867">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f72be-868">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="f72be-868">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f72be-869">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="f72be-869">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="f72be-870">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="f72be-870">displayReplyForm(formData)</span></span>

<span data-ttu-id="f72be-871">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-871">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f72be-872">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f72be-872">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f72be-873">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-873">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f72be-874">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="f72be-874">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="f72be-p149">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="f72be-p149">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f72be-878">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f72be-878">Parameters:</span></span>

|<span data-ttu-id="f72be-879">名前</span><span class="sxs-lookup"><span data-stu-id="f72be-879">Name</span></span>|<span data-ttu-id="f72be-880">型</span><span class="sxs-lookup"><span data-stu-id="f72be-880">Type</span></span>|<span data-ttu-id="f72be-881">属性</span><span class="sxs-lookup"><span data-stu-id="f72be-881">Attributes</span></span>|<span data-ttu-id="f72be-882">説明</span><span class="sxs-lookup"><span data-stu-id="f72be-882">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="f72be-883">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="f72be-883">String &#124; Object</span></span>||<span data-ttu-id="f72be-p150">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f72be-p150">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f72be-886">**または**</span><span class="sxs-lookup"><span data-stu-id="f72be-886">**OR**</span></span><br/><span data-ttu-id="f72be-p151">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="f72be-p151">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="f72be-889">String</span><span class="sxs-lookup"><span data-stu-id="f72be-889">String</span></span>|<span data-ttu-id="f72be-890">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-890">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-p152">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f72be-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="f72be-893">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-893">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="f72be-894">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-894">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-895">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="f72be-895">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="f72be-896">String</span><span class="sxs-lookup"><span data-stu-id="f72be-896">String</span></span>||<span data-ttu-id="f72be-p153">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="f72be-p153">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="f72be-899">String</span><span class="sxs-lookup"><span data-stu-id="f72be-899">String</span></span>||<span data-ttu-id="f72be-900">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="f72be-900">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="f72be-901">String</span><span class="sxs-lookup"><span data-stu-id="f72be-901">String</span></span>||<span data-ttu-id="f72be-p154">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="f72be-p154">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="f72be-904">ブール値</span><span class="sxs-lookup"><span data-stu-id="f72be-904">Boolean</span></span>||<span data-ttu-id="f72be-p155">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="f72be-p155">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="f72be-907">String</span><span class="sxs-lookup"><span data-stu-id="f72be-907">String</span></span>||<span data-ttu-id="f72be-p156">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="f72be-p156">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="f72be-911">function</span><span class="sxs-lookup"><span data-stu-id="f72be-911">function</span></span>|<span data-ttu-id="f72be-912">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-912">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-913">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-913">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f72be-914">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-914">Requirements</span></span>

|<span data-ttu-id="f72be-915">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-915">Requirement</span></span>|<span data-ttu-id="f72be-916">値</span><span class="sxs-lookup"><span data-stu-id="f72be-916">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-917">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-917">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-918">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-918">1.0</span></span>|
|[<span data-ttu-id="f72be-919">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-919">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-920">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-920">ReadItem</span></span>|
|[<span data-ttu-id="f72be-921">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-921">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-922">読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-922">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f72be-923">例</span><span class="sxs-lookup"><span data-stu-id="f72be-923">Examples</span></span>

<span data-ttu-id="f72be-924">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="f72be-924">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="f72be-925">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="f72be-925">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="f72be-926">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="f72be-926">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f72be-927">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="f72be-927">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f72be-928">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="f72be-928">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f72be-929">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="f72be-929">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="f72be-930">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="f72be-930">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="f72be-931">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="f72be-931">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="f72be-932">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f72be-932">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-933">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-933">Requirements</span></span>

|<span data-ttu-id="f72be-934">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-934">Requirement</span></span>|<span data-ttu-id="f72be-935">値</span><span class="sxs-lookup"><span data-stu-id="f72be-935">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-936">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-936">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-937">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-937">1.0</span></span>|
|[<span data-ttu-id="f72be-938">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-938">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-939">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-939">ReadItem</span></span>|
|[<span data-ttu-id="f72be-940">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-940">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-941">読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-941">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f72be-942">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f72be-942">Returns:</span></span>

<span data-ttu-id="f72be-943">型:[Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="f72be-943">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="f72be-944">例</span><span class="sxs-lookup"><span data-stu-id="f72be-944">Example</span></span>

<span data-ttu-id="f72be-945">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="f72be-945">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="f72be-946">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f72be-946">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f72be-947">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="f72be-947">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="f72be-948">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f72be-948">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f72be-949">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f72be-949">Parameters:</span></span>

|<span data-ttu-id="f72be-950">名前</span><span class="sxs-lookup"><span data-stu-id="f72be-950">Name</span></span>|<span data-ttu-id="f72be-951">種類</span><span class="sxs-lookup"><span data-stu-id="f72be-951">Type</span></span>|<span data-ttu-id="f72be-952">説明</span><span class="sxs-lookup"><span data-stu-id="f72be-952">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="f72be-953">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="f72be-953">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="f72be-954">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="f72be-954">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f72be-955">Requirements</span><span class="sxs-lookup"><span data-stu-id="f72be-955">Requirements</span></span>

|<span data-ttu-id="f72be-956">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-956">Requirement</span></span>|<span data-ttu-id="f72be-957">値</span><span class="sxs-lookup"><span data-stu-id="f72be-957">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-958">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-958">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-959">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-959">1.0</span></span>|
|[<span data-ttu-id="f72be-960">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-960">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-961">制限あり</span><span class="sxs-lookup"><span data-stu-id="f72be-961">Restricted</span></span>|
|[<span data-ttu-id="f72be-962">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-962">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-963">読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-963">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f72be-964">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f72be-964">Returns:</span></span>

<span data-ttu-id="f72be-965">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-965">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="f72be-966">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-966">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="f72be-967">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="f72be-967">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="f72be-968">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="f72be-968">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="f72be-969">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="f72be-969">Value of `entityType`</span></span>|<span data-ttu-id="f72be-970">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="f72be-970">Type of objects in returned array</span></span>|<span data-ttu-id="f72be-971">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="f72be-971">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="f72be-972">String</span><span class="sxs-lookup"><span data-stu-id="f72be-972">String</span></span>|<span data-ttu-id="f72be-973">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="f72be-973">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="f72be-974">連絡先</span><span class="sxs-lookup"><span data-stu-id="f72be-974">Contact</span></span>|<span data-ttu-id="f72be-975">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f72be-975">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="f72be-976">文字列</span><span class="sxs-lookup"><span data-stu-id="f72be-976">String</span></span>|<span data-ttu-id="f72be-977">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f72be-977">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="f72be-978">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="f72be-978">MeetingSuggestion</span></span>|<span data-ttu-id="f72be-979">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f72be-979">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="f72be-980">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="f72be-980">PhoneNumber</span></span>|<span data-ttu-id="f72be-981">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="f72be-981">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="f72be-982">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="f72be-982">TaskSuggestion</span></span>|<span data-ttu-id="f72be-983">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f72be-983">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="f72be-984">String</span><span class="sxs-lookup"><span data-stu-id="f72be-984">String</span></span>|<span data-ttu-id="f72be-985">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="f72be-985">**Restricted**</span></span>|

<span data-ttu-id="f72be-986">型:Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f72be-986">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="f72be-987">例</span><span class="sxs-lookup"><span data-stu-id="f72be-987">Example</span></span>

<span data-ttu-id="f72be-988">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="f72be-988">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="f72be-989">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f72be-989">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f72be-990">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-990">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f72be-991">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f72be-991">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f72be-992">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-992">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f72be-993">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f72be-993">Parameters:</span></span>

|<span data-ttu-id="f72be-994">名前</span><span class="sxs-lookup"><span data-stu-id="f72be-994">Name</span></span>|<span data-ttu-id="f72be-995">種類</span><span class="sxs-lookup"><span data-stu-id="f72be-995">Type</span></span>|<span data-ttu-id="f72be-996">説明</span><span class="sxs-lookup"><span data-stu-id="f72be-996">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="f72be-997">String</span><span class="sxs-lookup"><span data-stu-id="f72be-997">String</span></span>|<span data-ttu-id="f72be-998">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="f72be-998">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f72be-999">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-999">Requirements</span></span>

|<span data-ttu-id="f72be-1000">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1000">Requirement</span></span>|<span data-ttu-id="f72be-1001">値</span><span class="sxs-lookup"><span data-stu-id="f72be-1001">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-1002">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-1002">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-1003">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-1003">1.0</span></span>|
|[<span data-ttu-id="f72be-1004">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-1004">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-1005">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-1005">ReadItem</span></span>|
|[<span data-ttu-id="f72be-1006">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-1006">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-1007">読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-1007">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f72be-1008">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f72be-1008">Returns:</span></span>

<span data-ttu-id="f72be-p158">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-p158">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="f72be-1011">型:Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f72be-1011">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="f72be-1012">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f72be-1012">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="f72be-1013">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-1013">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f72be-1014">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f72be-1014">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f72be-p159">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="f72be-p159">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f72be-1018">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="f72be-1018">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f72be-1019">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="f72be-1019">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f72be-p160">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="f72be-p160">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-1023">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1023">Requirements</span></span>

|<span data-ttu-id="f72be-1024">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1024">Requirement</span></span>|<span data-ttu-id="f72be-1025">値</span><span class="sxs-lookup"><span data-stu-id="f72be-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-1026">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-1027">1.0</span></span>|
|[<span data-ttu-id="f72be-1028">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-1029">ReadItem</span></span>|
|[<span data-ttu-id="f72be-1030">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-1031">読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-1031">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f72be-1032">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f72be-1032">Returns:</span></span>

<span data-ttu-id="f72be-p161">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="f72be-p161">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="f72be-1035">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="f72be-1035">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f72be-1036">Object</span><span class="sxs-lookup"><span data-stu-id="f72be-1036">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f72be-1037">例</span><span class="sxs-lookup"><span data-stu-id="f72be-1037">Example</span></span>

<span data-ttu-id="f72be-1038">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="f72be-1038">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="f72be-1039">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="f72be-1039">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="f72be-1040">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-1040">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f72be-1041">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f72be-1041">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f72be-1042">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="f72be-1042">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="f72be-p162">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="f72be-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f72be-1045">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f72be-1045">Parameters:</span></span>

|<span data-ttu-id="f72be-1046">名前</span><span class="sxs-lookup"><span data-stu-id="f72be-1046">Name</span></span>|<span data-ttu-id="f72be-1047">種類</span><span class="sxs-lookup"><span data-stu-id="f72be-1047">Type</span></span>|<span data-ttu-id="f72be-1048">説明</span><span class="sxs-lookup"><span data-stu-id="f72be-1048">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="f72be-1049">String</span><span class="sxs-lookup"><span data-stu-id="f72be-1049">String</span></span>|<span data-ttu-id="f72be-1050">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="f72be-1050">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f72be-1051">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1051">Requirements</span></span>

|<span data-ttu-id="f72be-1052">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1052">Requirement</span></span>|<span data-ttu-id="f72be-1053">値</span><span class="sxs-lookup"><span data-stu-id="f72be-1053">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-1054">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-1054">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-1055">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-1055">1.0</span></span>|
|[<span data-ttu-id="f72be-1056">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-1056">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-1057">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-1057">ReadItem</span></span>|
|[<span data-ttu-id="f72be-1058">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-1058">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-1059">読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-1059">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f72be-1060">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f72be-1060">Returns:</span></span>

<span data-ttu-id="f72be-1061">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="f72be-1061">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="f72be-1062">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="f72be-1062">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f72be-1063">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="f72be-1063">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f72be-1064">例</span><span class="sxs-lookup"><span data-stu-id="f72be-1064">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="f72be-1065">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="f72be-1065">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="f72be-1066">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-1066">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="f72be-p163">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-p163">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f72be-1069">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f72be-1069">Parameters:</span></span>

|<span data-ttu-id="f72be-1070">名前</span><span class="sxs-lookup"><span data-stu-id="f72be-1070">Name</span></span>|<span data-ttu-id="f72be-1071">型</span><span class="sxs-lookup"><span data-stu-id="f72be-1071">Type</span></span>|<span data-ttu-id="f72be-1072">属性</span><span class="sxs-lookup"><span data-stu-id="f72be-1072">Attributes</span></span>|<span data-ttu-id="f72be-1073">説明</span><span class="sxs-lookup"><span data-stu-id="f72be-1073">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="f72be-1074">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f72be-1074">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="f72be-p164">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="f72be-p164">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="f72be-1078">Object</span><span class="sxs-lookup"><span data-stu-id="f72be-1078">Object</span></span>|<span data-ttu-id="f72be-1079">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-1079">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-1080">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f72be-1080">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f72be-1081">Object</span><span class="sxs-lookup"><span data-stu-id="f72be-1081">Object</span></span>|<span data-ttu-id="f72be-1082">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-1082">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-1083">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1083">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f72be-1084">function</span><span class="sxs-lookup"><span data-stu-id="f72be-1084">function</span></span>||<span data-ttu-id="f72be-1085">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1085">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f72be-1086">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="f72be-1086">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="f72be-1087">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="f72be-1087">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f72be-1088">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1088">Requirements</span></span>

|<span data-ttu-id="f72be-1089">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1089">Requirement</span></span>|<span data-ttu-id="f72be-1090">値</span><span class="sxs-lookup"><span data-stu-id="f72be-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-1091">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-1092">1.2</span><span class="sxs-lookup"><span data-stu-id="f72be-1092">1.2</span></span>|
|[<span data-ttu-id="f72be-1093">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-1093">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-1094">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f72be-1094">ReadWriteItem</span></span>|
|[<span data-ttu-id="f72be-1095">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-1095">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-1096">作成</span><span class="sxs-lookup"><span data-stu-id="f72be-1096">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="f72be-1097">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f72be-1097">Returns:</span></span>

<span data-ttu-id="f72be-1098">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="f72be-1098">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="f72be-1099">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="f72be-1099">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f72be-1100">String</span><span class="sxs-lookup"><span data-stu-id="f72be-1100">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f72be-1101">例</span><span class="sxs-lookup"><span data-stu-id="f72be-1101">Example</span></span>

```js
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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="f72be-1102">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="f72be-1102">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="f72be-p166">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-p166">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="f72be-1105">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f72be-1105">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-1106">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1106">Requirements</span></span>

|<span data-ttu-id="f72be-1107">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1107">Requirement</span></span>|<span data-ttu-id="f72be-1108">値</span><span class="sxs-lookup"><span data-stu-id="f72be-1108">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-1109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-1109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-1110">1.6</span><span class="sxs-lookup"><span data-stu-id="f72be-1110">1.6</span></span>|
|[<span data-ttu-id="f72be-1111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-1111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-1112">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-1112">ReadItem</span></span>|
|[<span data-ttu-id="f72be-1113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-1113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-1114">読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-1114">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f72be-1115">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f72be-1115">Returns:</span></span>

<span data-ttu-id="f72be-1116">型:[Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="f72be-1116">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="f72be-1117">例</span><span class="sxs-lookup"><span data-stu-id="f72be-1117">Example</span></span>

<span data-ttu-id="f72be-1118">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="f72be-1118">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="f72be-1119">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f72be-1119">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="f72be-p167">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-p167">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="f72be-1122">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f72be-1122">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f72be-p168">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="f72be-p168">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f72be-1126">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="f72be-1126">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f72be-1127">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="f72be-1127">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f72be-p169">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="f72be-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f72be-1131">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1131">Requirements</span></span>

|<span data-ttu-id="f72be-1132">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1132">Requirement</span></span>|<span data-ttu-id="f72be-1133">値</span><span class="sxs-lookup"><span data-stu-id="f72be-1133">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-1134">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-1134">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-1135">1.6</span><span class="sxs-lookup"><span data-stu-id="f72be-1135">1.6</span></span>|
|[<span data-ttu-id="f72be-1136">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-1136">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-1137">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-1137">ReadItem</span></span>|
|[<span data-ttu-id="f72be-1138">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-1138">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-1139">読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-1139">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f72be-1140">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f72be-1140">Returns:</span></span>

<span data-ttu-id="f72be-p170">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="f72be-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="f72be-1143">例</span><span class="sxs-lookup"><span data-stu-id="f72be-1143">Example</span></span>

<span data-ttu-id="f72be-1144">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="f72be-1144">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="f72be-1145">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f72be-1145">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="f72be-1146">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1146">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="f72be-p171">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="f72be-p171">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f72be-1150">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f72be-1150">Parameters:</span></span>

|<span data-ttu-id="f72be-1151">名前</span><span class="sxs-lookup"><span data-stu-id="f72be-1151">Name</span></span>|<span data-ttu-id="f72be-1152">型</span><span class="sxs-lookup"><span data-stu-id="f72be-1152">Type</span></span>|<span data-ttu-id="f72be-1153">属性</span><span class="sxs-lookup"><span data-stu-id="f72be-1153">Attributes</span></span>|<span data-ttu-id="f72be-1154">説明</span><span class="sxs-lookup"><span data-stu-id="f72be-1154">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="f72be-1155">function</span><span class="sxs-lookup"><span data-stu-id="f72be-1155">function</span></span>||<span data-ttu-id="f72be-1156">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1156">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f72be-1157">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1157">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="f72be-1158">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1158">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="f72be-1159">Object</span><span class="sxs-lookup"><span data-stu-id="f72be-1159">Object</span></span>|<span data-ttu-id="f72be-1160">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-1161">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1161">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="f72be-1162">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1162">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f72be-1163">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1163">Requirements</span></span>

|<span data-ttu-id="f72be-1164">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1164">Requirement</span></span>|<span data-ttu-id="f72be-1165">値</span><span class="sxs-lookup"><span data-stu-id="f72be-1165">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-1166">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-1166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-1167">1.0</span><span class="sxs-lookup"><span data-stu-id="f72be-1167">1.0</span></span>|
|[<span data-ttu-id="f72be-1168">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-1168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-1169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-1169">ReadItem</span></span>|
|[<span data-ttu-id="f72be-1170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-1170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-1171">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-1171">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-1172">例</span><span class="sxs-lookup"><span data-stu-id="f72be-1172">Example</span></span>

<span data-ttu-id="f72be-p174">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="f72be-p174">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="f72be-1176">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f72be-1176">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="f72be-1177">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="f72be-1177">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="f72be-p175">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="f72be-p175">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f72be-1182">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f72be-1182">Parameters:</span></span>

|<span data-ttu-id="f72be-1183">名前</span><span class="sxs-lookup"><span data-stu-id="f72be-1183">Name</span></span>|<span data-ttu-id="f72be-1184">型</span><span class="sxs-lookup"><span data-stu-id="f72be-1184">Type</span></span>|<span data-ttu-id="f72be-1185">属性</span><span class="sxs-lookup"><span data-stu-id="f72be-1185">Attributes</span></span>|<span data-ttu-id="f72be-1186">説明</span><span class="sxs-lookup"><span data-stu-id="f72be-1186">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="f72be-1187">String</span><span class="sxs-lookup"><span data-stu-id="f72be-1187">String</span></span>||<span data-ttu-id="f72be-1188">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="f72be-1188">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="f72be-1189">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f72be-1189">Object</span></span>|<span data-ttu-id="f72be-1190">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-1190">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-1191">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f72be-1191">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f72be-1192">Object</span><span class="sxs-lookup"><span data-stu-id="f72be-1192">Object</span></span>|<span data-ttu-id="f72be-1193">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-1193">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-1194">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1194">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f72be-1195">function</span><span class="sxs-lookup"><span data-stu-id="f72be-1195">function</span></span>|<span data-ttu-id="f72be-1196">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-1196">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-1197">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1197">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f72be-1198">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1198">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f72be-1199">エラー</span><span class="sxs-lookup"><span data-stu-id="f72be-1199">Errors</span></span>

|<span data-ttu-id="f72be-1200">エラー コード</span><span class="sxs-lookup"><span data-stu-id="f72be-1200">Error code</span></span>|<span data-ttu-id="f72be-1201">説明</span><span class="sxs-lookup"><span data-stu-id="f72be-1201">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="f72be-1202">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="f72be-1202">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f72be-1203">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1203">Requirements</span></span>

|<span data-ttu-id="f72be-1204">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1204">Requirement</span></span>|<span data-ttu-id="f72be-1205">値</span><span class="sxs-lookup"><span data-stu-id="f72be-1205">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-1206">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-1206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-1207">1.1</span><span class="sxs-lookup"><span data-stu-id="f72be-1207">1.1</span></span>|
|[<span data-ttu-id="f72be-1208">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-1208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-1209">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f72be-1209">ReadWriteItem</span></span>|
|[<span data-ttu-id="f72be-1210">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-1210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-1211">作成</span><span class="sxs-lookup"><span data-stu-id="f72be-1211">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-1212">例</span><span class="sxs-lookup"><span data-stu-id="f72be-1212">Example</span></span>

<span data-ttu-id="f72be-1213">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="f72be-1213">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="f72be-1214">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f72be-1214">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="f72be-1215">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="f72be-1215">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="f72be-1216">現在、サポートされているイベントの種類は `Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged`、`Office.EventType.RecurrenceChanged` です。</span><span class="sxs-lookup"><span data-stu-id="f72be-1216">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="f72be-1217">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f72be-1217">Parameters:</span></span>

| <span data-ttu-id="f72be-1218">名前</span><span class="sxs-lookup"><span data-stu-id="f72be-1218">Name</span></span> | <span data-ttu-id="f72be-1219">型</span><span class="sxs-lookup"><span data-stu-id="f72be-1219">Type</span></span> | <span data-ttu-id="f72be-1220">属性</span><span class="sxs-lookup"><span data-stu-id="f72be-1220">Attributes</span></span> | <span data-ttu-id="f72be-1221">説明</span><span class="sxs-lookup"><span data-stu-id="f72be-1221">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="f72be-1222">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="f72be-1222">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="f72be-1223">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="f72be-1223">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="f72be-1224">Object</span><span class="sxs-lookup"><span data-stu-id="f72be-1224">Object</span></span> | <span data-ttu-id="f72be-1225">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-1225">&lt;optional&gt;</span></span> | <span data-ttu-id="f72be-1226">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f72be-1226">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="f72be-1227">Object</span><span class="sxs-lookup"><span data-stu-id="f72be-1227">Object</span></span> | <span data-ttu-id="f72be-1228">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-1228">&lt;optional&gt;</span></span> | <span data-ttu-id="f72be-1229">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1229">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="f72be-1230">function</span><span class="sxs-lookup"><span data-stu-id="f72be-1230">function</span></span>| <span data-ttu-id="f72be-1231">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-1231">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-1232">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1232">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f72be-1233">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1233">Requirements</span></span>

|<span data-ttu-id="f72be-1234">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1234">Requirement</span></span>| <span data-ttu-id="f72be-1235">値</span><span class="sxs-lookup"><span data-stu-id="f72be-1235">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-1236">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-1236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f72be-1237">1.7</span><span class="sxs-lookup"><span data-stu-id="f72be-1237">1.7</span></span> |
|[<span data-ttu-id="f72be-1238">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-1238">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f72be-1239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f72be-1239">ReadItem</span></span> |
|[<span data-ttu-id="f72be-1240">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-1240">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f72be-1241">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f72be-1241">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="f72be-1242">例</span><span class="sxs-lookup"><span data-stu-id="f72be-1242">Example</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="f72be-1243">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="f72be-1243">saveAsync([options], callback)</span></span>

<span data-ttu-id="f72be-1244">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="f72be-1244">Asynchronously saves an item.</span></span>

<span data-ttu-id="f72be-p176">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-p176">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="f72be-1248">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="f72be-1248">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="f72be-1249">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1249">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="f72be-p178">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-p178">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="f72be-1253">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="f72be-1253">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="f72be-1254">Mac Outlook では、新規作成モードの会議で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="f72be-1254">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="f72be-1255">Mac Outlook では、会議で `saveAsync` を呼び出すとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1255">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="f72be-1256">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1256">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f72be-1257">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f72be-1257">Parameters:</span></span>

|<span data-ttu-id="f72be-1258">名前</span><span class="sxs-lookup"><span data-stu-id="f72be-1258">Name</span></span>|<span data-ttu-id="f72be-1259">型</span><span class="sxs-lookup"><span data-stu-id="f72be-1259">Type</span></span>|<span data-ttu-id="f72be-1260">属性</span><span class="sxs-lookup"><span data-stu-id="f72be-1260">Attributes</span></span>|<span data-ttu-id="f72be-1261">説明</span><span class="sxs-lookup"><span data-stu-id="f72be-1261">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="f72be-1262">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f72be-1262">Object</span></span>|<span data-ttu-id="f72be-1263">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-1263">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-1264">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f72be-1264">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f72be-1265">Object</span><span class="sxs-lookup"><span data-stu-id="f72be-1265">Object</span></span>|<span data-ttu-id="f72be-1266">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-1266">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-1267">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1267">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f72be-1268">function</span><span class="sxs-lookup"><span data-stu-id="f72be-1268">function</span></span>||<span data-ttu-id="f72be-1269">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1269">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f72be-1270">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1270">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f72be-1271">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1271">Requirements</span></span>

|<span data-ttu-id="f72be-1272">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1272">Requirement</span></span>|<span data-ttu-id="f72be-1273">値</span><span class="sxs-lookup"><span data-stu-id="f72be-1273">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-1274">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-1274">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-1275">1.3</span><span class="sxs-lookup"><span data-stu-id="f72be-1275">1.3</span></span>|
|[<span data-ttu-id="f72be-1276">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-1276">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-1277">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f72be-1277">ReadWriteItem</span></span>|
|[<span data-ttu-id="f72be-1278">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-1278">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-1279">新規作成</span><span class="sxs-lookup"><span data-stu-id="f72be-1279">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f72be-1280">例</span><span class="sxs-lookup"><span data-stu-id="f72be-1280">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="f72be-p180">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="f72be-p180">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="f72be-1283">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="f72be-1283">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="f72be-1284">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="f72be-1284">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="f72be-p181">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="f72be-p181">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f72be-1288">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f72be-1288">Parameters:</span></span>

|<span data-ttu-id="f72be-1289">名前</span><span class="sxs-lookup"><span data-stu-id="f72be-1289">Name</span></span>|<span data-ttu-id="f72be-1290">型</span><span class="sxs-lookup"><span data-stu-id="f72be-1290">Type</span></span>|<span data-ttu-id="f72be-1291">属性</span><span class="sxs-lookup"><span data-stu-id="f72be-1291">Attributes</span></span>|<span data-ttu-id="f72be-1292">説明</span><span class="sxs-lookup"><span data-stu-id="f72be-1292">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="f72be-1293">String</span><span class="sxs-lookup"><span data-stu-id="f72be-1293">String</span></span>||<span data-ttu-id="f72be-p182">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="f72be-p182">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="f72be-1297">Object</span><span class="sxs-lookup"><span data-stu-id="f72be-1297">Object</span></span>|<span data-ttu-id="f72be-1298">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-1298">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-1299">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f72be-1299">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f72be-1300">Object</span><span class="sxs-lookup"><span data-stu-id="f72be-1300">Object</span></span>|<span data-ttu-id="f72be-1301">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-1301">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-1302">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1302">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="f72be-1303">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f72be-1303">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="f72be-1304">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f72be-1304">&lt;optional&gt;</span></span>|<span data-ttu-id="f72be-p183">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-p183">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="f72be-p184">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-p184">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="f72be-1309">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1309">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="f72be-1310">function</span><span class="sxs-lookup"><span data-stu-id="f72be-1310">function</span></span>||<span data-ttu-id="f72be-1311">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f72be-1311">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f72be-1312">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1312">Requirements</span></span>

|<span data-ttu-id="f72be-1313">要件</span><span class="sxs-lookup"><span data-stu-id="f72be-1313">Requirement</span></span>|<span data-ttu-id="f72be-1314">値</span><span class="sxs-lookup"><span data-stu-id="f72be-1314">Value</span></span>|
|---|---|
|[<span data-ttu-id="f72be-1315">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f72be-1315">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f72be-1316">1.2</span><span class="sxs-lookup"><span data-stu-id="f72be-1316">1.2</span></span>|
|[<span data-ttu-id="f72be-1317">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f72be-1317">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f72be-1318">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f72be-1318">ReadWriteItem</span></span>|
|[<span data-ttu-id="f72be-1319">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f72be-1319">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f72be-1320">作成</span><span class="sxs-lookup"><span data-stu-id="f72be-1320">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f72be-1321">例</span><span class="sxs-lookup"><span data-stu-id="f72be-1321">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
