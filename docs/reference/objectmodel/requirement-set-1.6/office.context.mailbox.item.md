---
title: Office.context.mailbox.item の要件は、1.6 を設定
description: ''
ms.date: 12/18/2018
localization_priority: Normal
ms.openlocfilehash: 23f27a2949ddcdaa17ffe3f4711002d47d699458
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29387318"
---
# <a name="item"></a><span data-ttu-id="1ca84-102">item</span><span class="sxs-lookup"><span data-stu-id="1ca84-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="1ca84-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="1ca84-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="1ca84-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-106">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-106">Requirements</span></span>

|<span data-ttu-id="1ca84-107">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-107">Requirement</span></span>| <span data-ttu-id="1ca84-108">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-110">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-110">1.0</span></span>|
|[<span data-ttu-id="1ca84-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="1ca84-112">Restricted</span></span>|
|[<span data-ttu-id="1ca84-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-114">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1ca84-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-115">Members and methods</span></span>

| <span data-ttu-id="1ca84-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-116">Member</span></span> | <span data-ttu-id="1ca84-117">種類</span><span class="sxs-lookup"><span data-stu-id="1ca84-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1ca84-118">attachments</span><span class="sxs-lookup"><span data-stu-id="1ca84-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | <span data-ttu-id="1ca84-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-119">Member</span></span> |
| [<span data-ttu-id="1ca84-120">bcc</span><span class="sxs-lookup"><span data-stu-id="1ca84-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="1ca84-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-121">Member</span></span> |
| [<span data-ttu-id="1ca84-122">body</span><span class="sxs-lookup"><span data-stu-id="1ca84-122">body</span></span>](#body-bodyjavascriptapioutlook16officebody) | <span data-ttu-id="1ca84-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-123">Member</span></span> |
| [<span data-ttu-id="1ca84-124">cc</span><span class="sxs-lookup"><span data-stu-id="1ca84-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="1ca84-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-125">Member</span></span> |
| [<span data-ttu-id="1ca84-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="1ca84-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="1ca84-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-127">Member</span></span> |
| [<span data-ttu-id="1ca84-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="1ca84-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="1ca84-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-129">Member</span></span> |
| [<span data-ttu-id="1ca84-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="1ca84-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="1ca84-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-131">Member</span></span> |
| [<span data-ttu-id="1ca84-132">end</span><span class="sxs-lookup"><span data-stu-id="1ca84-132">end</span></span>](#end-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="1ca84-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-133">Member</span></span> |
| [<span data-ttu-id="1ca84-134">from</span><span class="sxs-lookup"><span data-stu-id="1ca84-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="1ca84-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-135">Member</span></span> |
| [<span data-ttu-id="1ca84-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="1ca84-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="1ca84-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-137">Member</span></span> |
| [<span data-ttu-id="1ca84-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="1ca84-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="1ca84-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-139">Member</span></span> |
| [<span data-ttu-id="1ca84-140">itemId</span><span class="sxs-lookup"><span data-stu-id="1ca84-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="1ca84-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-141">Member</span></span> |
| [<span data-ttu-id="1ca84-142">itemType</span><span class="sxs-lookup"><span data-stu-id="1ca84-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | <span data-ttu-id="1ca84-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-143">Member</span></span> |
| [<span data-ttu-id="1ca84-144">location</span><span class="sxs-lookup"><span data-stu-id="1ca84-144">location</span></span>](#location-stringlocationjavascriptapioutlook16officelocation) | <span data-ttu-id="1ca84-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-145">Member</span></span> |
| [<span data-ttu-id="1ca84-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="1ca84-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="1ca84-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-147">Member</span></span> |
| [<span data-ttu-id="1ca84-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="1ca84-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | <span data-ttu-id="1ca84-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-149">Member</span></span> |
| [<span data-ttu-id="1ca84-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="1ca84-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="1ca84-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-151">Member</span></span> |
| [<span data-ttu-id="1ca84-152">organizer</span><span class="sxs-lookup"><span data-stu-id="1ca84-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="1ca84-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-153">Member</span></span> |
| [<span data-ttu-id="1ca84-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="1ca84-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="1ca84-155">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-155">Member</span></span> |
| [<span data-ttu-id="1ca84-156">sender</span><span class="sxs-lookup"><span data-stu-id="1ca84-156">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="1ca84-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-157">Member</span></span> |
| [<span data-ttu-id="1ca84-158">start</span><span class="sxs-lookup"><span data-stu-id="1ca84-158">start</span></span>](#start-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="1ca84-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-159">Member</span></span> |
| [<span data-ttu-id="1ca84-160">subject</span><span class="sxs-lookup"><span data-stu-id="1ca84-160">subject</span></span>](#subject-stringsubjectjavascriptapioutlook16officesubject) | <span data-ttu-id="1ca84-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-161">Member</span></span> |
| [<span data-ttu-id="1ca84-162">to</span><span class="sxs-lookup"><span data-stu-id="1ca84-162">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="1ca84-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-163">Member</span></span> |
| [<span data-ttu-id="1ca84-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1ca84-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="1ca84-165">メソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-165">Method</span></span> |
| [<span data-ttu-id="1ca84-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1ca84-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="1ca84-167">メソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-167">Method</span></span> |
| [<span data-ttu-id="1ca84-168">close</span><span class="sxs-lookup"><span data-stu-id="1ca84-168">close</span></span>](#close) | <span data-ttu-id="1ca84-169">メソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-169">Method</span></span> |
| [<span data-ttu-id="1ca84-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="1ca84-170">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="1ca84-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-171">Method</span></span> |
| [<span data-ttu-id="1ca84-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="1ca84-172">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="1ca84-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-173">Method</span></span> |
| [<span data-ttu-id="1ca84-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="1ca84-174">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="1ca84-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-175">Method</span></span> |
| [<span data-ttu-id="1ca84-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="1ca84-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="1ca84-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-177">Method</span></span> |
| [<span data-ttu-id="1ca84-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="1ca84-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="1ca84-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-179">Method</span></span> |
| [<span data-ttu-id="1ca84-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="1ca84-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="1ca84-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-181">Method</span></span> |
| [<span data-ttu-id="1ca84-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="1ca84-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="1ca84-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-183">Method</span></span> |
| [<span data-ttu-id="1ca84-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="1ca84-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="1ca84-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-185">Method</span></span> |
| [<span data-ttu-id="1ca84-186">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="1ca84-186">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="1ca84-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-187">Method</span></span> |
| [<span data-ttu-id="1ca84-188">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="1ca84-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="1ca84-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-189">Method</span></span> |
| [<span data-ttu-id="1ca84-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="1ca84-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="1ca84-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-191">Method</span></span> |
| [<span data-ttu-id="1ca84-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1ca84-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="1ca84-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-193">Method</span></span> |
| [<span data-ttu-id="1ca84-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="1ca84-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="1ca84-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-195">Method</span></span> |
| [<span data-ttu-id="1ca84-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="1ca84-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="1ca84-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="1ca84-198">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-198">Example</span></span>

<span data-ttu-id="1ca84-199">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1ca84-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="1ca84-200">メンバー</span><span class="sxs-lookup"><span data-stu-id="1ca84-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="1ca84-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="1ca84-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="1ca84-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1ca84-204">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="1ca84-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="1ca84-205">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1ca84-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-206">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-206">Type:</span></span>

*   <span data-ttu-id="1ca84-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="1ca84-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-208">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-208">Requirements</span></span>

|<span data-ttu-id="1ca84-209">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-209">Requirement</span></span>| <span data-ttu-id="1ca84-210">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-211">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-212">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-212">1.0</span></span>|
|[<span data-ttu-id="1ca84-213">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-213">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-214">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-215">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-215">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-216">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-217">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-217">Example</span></span>

<span data-ttu-id="1ca84-218">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="1ca84-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1ca84-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="1ca84-220">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="1ca84-221">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1ca84-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-222">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-222">Type:</span></span>

*   [<span data-ttu-id="1ca84-223">Recipients</span><span class="sxs-lookup"><span data-stu-id="1ca84-223">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="1ca84-224">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-224">Requirements</span></span>

|<span data-ttu-id="1ca84-225">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-225">Requirement</span></span>| <span data-ttu-id="1ca84-226">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-227">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-228">1.1</span><span class="sxs-lookup"><span data-stu-id="1ca84-228">1.1</span></span>|
|[<span data-ttu-id="1ca84-229">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-229">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-230">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-231">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-231">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-232">作成</span><span class="sxs-lookup"><span data-stu-id="1ca84-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-233">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-233">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="1ca84-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="1ca84-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="1ca84-235">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-236">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-236">Type:</span></span>

*   [<span data-ttu-id="1ca84-237">Body</span><span class="sxs-lookup"><span data-stu-id="1ca84-237">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="1ca84-238">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-238">Requirements</span></span>

|<span data-ttu-id="1ca84-239">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-239">Requirement</span></span>| <span data-ttu-id="1ca84-240">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-241">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-242">1.1</span><span class="sxs-lookup"><span data-stu-id="1ca84-242">1.1</span></span>|
|[<span data-ttu-id="1ca84-243">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-243">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-244">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-245">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-246">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-246">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="1ca84-247">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1ca84-247">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="1ca84-248">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-248">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="1ca84-249">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="1ca84-249">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1ca84-250">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1ca84-250">Read mode</span></span>

<span data-ttu-id="1ca84-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1ca84-253">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1ca84-253">Compose mode</span></span>

<span data-ttu-id="1ca84-254">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-254">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-255">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-255">Type:</span></span>

*   <span data-ttu-id="1ca84-256">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1ca84-256">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-257">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-257">Requirements</span></span>

|<span data-ttu-id="1ca84-258">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-258">Requirement</span></span>| <span data-ttu-id="1ca84-259">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-260">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-260">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-261">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-261">1.0</span></span>|
|[<span data-ttu-id="1ca84-262">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-262">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-263">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-263">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-264">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-264">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-265">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-265">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-266">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-266">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="1ca84-267">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="1ca84-267">(nullable) conversationId :String</span></span>

<span data-ttu-id="1ca84-268">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-268">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="1ca84-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="1ca84-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-273">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-273">Type:</span></span>

*   <span data-ttu-id="1ca84-274">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-275">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-275">Requirements</span></span>

|<span data-ttu-id="1ca84-276">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-276">Requirement</span></span>| <span data-ttu-id="1ca84-277">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-278">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-279">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-279">1.0</span></span>|
|[<span data-ttu-id="1ca84-280">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-281">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-282">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-283">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-283">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="1ca84-284">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="1ca84-284">dateTimeCreated :Date</span></span>

<span data-ttu-id="1ca84-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-287">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-287">Type:</span></span>

*   <span data-ttu-id="1ca84-288">日付</span><span class="sxs-lookup"><span data-stu-id="1ca84-288">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-289">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-289">Requirements</span></span>

|<span data-ttu-id="1ca84-290">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-290">Requirement</span></span>| <span data-ttu-id="1ca84-291">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-291">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-292">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-292">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-293">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-293">1.0</span></span>|
|[<span data-ttu-id="1ca84-294">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-294">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-295">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-295">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-296">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-296">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-297">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-297">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-298">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-298">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="1ca84-299">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="1ca84-299">dateTimeModified :Date</span></span>

<span data-ttu-id="1ca84-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1ca84-302">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1ca84-302">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-303">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-303">Type:</span></span>

*   <span data-ttu-id="1ca84-304">日付</span><span class="sxs-lookup"><span data-stu-id="1ca84-304">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-305">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-305">Requirements</span></span>

|<span data-ttu-id="1ca84-306">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-306">Requirement</span></span>| <span data-ttu-id="1ca84-307">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-308">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-309">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-309">1.0</span></span>|
|[<span data-ttu-id="1ca84-310">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-310">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-311">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-312">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-312">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-313">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-314">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-314">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="1ca84-315">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="1ca84-315">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="1ca84-316">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-316">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="1ca84-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1ca84-319">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1ca84-319">Read mode</span></span>

<span data-ttu-id="1ca84-320">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-320">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1ca84-321">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1ca84-321">Compose mode</span></span>

<span data-ttu-id="1ca84-322">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-322">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="1ca84-323">[`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1ca84-323">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-324">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-324">Type:</span></span>

*   <span data-ttu-id="1ca84-325">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="1ca84-325">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-326">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-326">Requirements</span></span>

|<span data-ttu-id="1ca84-327">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-327">Requirement</span></span>| <span data-ttu-id="1ca84-328">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-329">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-330">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-330">1.0</span></span>|
|[<span data-ttu-id="1ca84-331">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-331">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-332">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-333">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-333">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-334">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-334">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-335">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-335">Example</span></span>

<span data-ttu-id="1ca84-336">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-336">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="1ca84-337">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1ca84-337">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="1ca84-p112">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="1ca84-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1ca84-342">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="1ca84-342">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-343">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-343">Type:</span></span>

*   [<span data-ttu-id="1ca84-344">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1ca84-344">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1ca84-345">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-345">Requirements</span></span>

|<span data-ttu-id="1ca84-346">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-346">Requirement</span></span>| <span data-ttu-id="1ca84-347">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-347">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-348">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-349">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-349">1.0</span></span>|
|[<span data-ttu-id="1ca84-350">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-350">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-351">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-352">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-352">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-353">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-353">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="1ca84-354">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="1ca84-354">internetMessageId :String</span></span>

<span data-ttu-id="1ca84-p114">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-357">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-357">Type:</span></span>

*   <span data-ttu-id="1ca84-358">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-358">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-359">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-359">Requirements</span></span>

|<span data-ttu-id="1ca84-360">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-360">Requirement</span></span>| <span data-ttu-id="1ca84-361">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-362">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-363">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-363">1.0</span></span>|
|[<span data-ttu-id="1ca84-364">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-365">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-366">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-367">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-367">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-368">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-368">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="1ca84-369">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="1ca84-369">itemClass :String</span></span>

<span data-ttu-id="1ca84-p115">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="1ca84-p116">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="1ca84-374">型</span><span class="sxs-lookup"><span data-stu-id="1ca84-374">Type</span></span> | <span data-ttu-id="1ca84-375">説明</span><span class="sxs-lookup"><span data-stu-id="1ca84-375">Description</span></span> | <span data-ttu-id="1ca84-376">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="1ca84-376">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="1ca84-377">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="1ca84-377">Appointment items</span></span> | <span data-ttu-id="1ca84-378">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="1ca84-378">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="1ca84-379">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="1ca84-379">Message items</span></span> | <span data-ttu-id="1ca84-380">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-380">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="1ca84-381">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-381">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-382">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-382">Type:</span></span>

*   <span data-ttu-id="1ca84-383">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-383">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-384">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-384">Requirements</span></span>

|<span data-ttu-id="1ca84-385">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-385">Requirement</span></span>| <span data-ttu-id="1ca84-386">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-386">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-387">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-387">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-388">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-388">1.0</span></span>|
|[<span data-ttu-id="1ca84-389">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-389">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-390">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-390">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-391">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-391">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-392">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-392">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-393">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-393">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="1ca84-394">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="1ca84-394">(nullable) itemId :String</span></span>

<span data-ttu-id="1ca84-p117">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1ca84-397">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="1ca84-397">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="1ca84-398">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="1ca84-398">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="1ca84-399">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1ca84-399">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="1ca84-400">詳細は、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1ca84-400">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="1ca84-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-403">種類:</span><span class="sxs-lookup"><span data-stu-id="1ca84-403">Type:</span></span>

*   <span data-ttu-id="1ca84-404">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-404">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-405">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-405">Requirements</span></span>

|<span data-ttu-id="1ca84-406">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-406">Requirement</span></span>| <span data-ttu-id="1ca84-407">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-408">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-409">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-409">1.0</span></span>|
|[<span data-ttu-id="1ca84-410">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-411">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-412">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-413">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-413">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-414">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-414">Example</span></span>

<span data-ttu-id="1ca84-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="1ca84-417">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="1ca84-417">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="1ca84-418">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-418">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="1ca84-419">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="1ca84-419">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-420">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-420">Type:</span></span>

*   [<span data-ttu-id="1ca84-421">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="1ca84-421">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="1ca84-422">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-422">Requirements</span></span>

|<span data-ttu-id="1ca84-423">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-423">Requirement</span></span>| <span data-ttu-id="1ca84-424">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-424">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-425">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-425">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-426">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-426">1.0</span></span>|
|[<span data-ttu-id="1ca84-427">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-427">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-428">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-429">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-429">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-430">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-430">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-431">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-431">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="1ca84-432">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="1ca84-432">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="1ca84-433">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-433">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1ca84-434">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1ca84-434">Read mode</span></span>

<span data-ttu-id="1ca84-435">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-435">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1ca84-436">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1ca84-436">Compose mode</span></span>

<span data-ttu-id="1ca84-437">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-437">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-438">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-438">Type:</span></span>

*   <span data-ttu-id="1ca84-439">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="1ca84-439">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-440">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-440">Requirements</span></span>

|<span data-ttu-id="1ca84-441">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-441">Requirement</span></span>| <span data-ttu-id="1ca84-442">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-443">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-443">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-444">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-444">1.0</span></span>|
|[<span data-ttu-id="1ca84-445">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-446">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-447">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-448">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-449">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-449">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="1ca84-450">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="1ca84-450">normalizedSubject :String</span></span>

<span data-ttu-id="1ca84-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="1ca84-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-455">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-455">Type:</span></span>

*   <span data-ttu-id="1ca84-456">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-456">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-457">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-457">Requirements</span></span>

|<span data-ttu-id="1ca84-458">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-458">Requirement</span></span>| <span data-ttu-id="1ca84-459">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-460">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-461">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-461">1.0</span></span>|
|[<span data-ttu-id="1ca84-462">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-463">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-464">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-465">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-466">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-466">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="1ca84-467">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="1ca84-467">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="1ca84-468">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-468">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-469">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-469">Type:</span></span>

*   [<span data-ttu-id="1ca84-470">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="1ca84-470">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="1ca84-471">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-471">Requirements</span></span>

|<span data-ttu-id="1ca84-472">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-472">Requirement</span></span>| <span data-ttu-id="1ca84-473">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-473">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-474">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-474">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-475">1.3</span><span class="sxs-lookup"><span data-stu-id="1ca84-475">1.3</span></span>|
|[<span data-ttu-id="1ca84-476">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-476">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-477">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-477">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-478">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-478">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-479">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-479">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="1ca84-480">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1ca84-480">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="1ca84-481">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-481">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="1ca84-482">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="1ca84-482">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1ca84-483">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1ca84-483">Read mode</span></span>

<span data-ttu-id="1ca84-484">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-484">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1ca84-485">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1ca84-485">Compose mode</span></span>

<span data-ttu-id="1ca84-486">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-486">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-487">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-487">Type:</span></span>

*   <span data-ttu-id="1ca84-488">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1ca84-488">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-489">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-489">Requirements</span></span>

|<span data-ttu-id="1ca84-490">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-490">Requirement</span></span>| <span data-ttu-id="1ca84-491">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-491">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-492">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-492">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-493">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-493">1.0</span></span>|
|[<span data-ttu-id="1ca84-494">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-494">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-495">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-495">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-496">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-496">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-497">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-497">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-498">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-498">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="1ca84-499">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1ca84-499">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="1ca84-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-502">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-502">Type:</span></span>

*   [<span data-ttu-id="1ca84-503">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1ca84-503">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1ca84-504">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-504">Requirements</span></span>

|<span data-ttu-id="1ca84-505">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-505">Requirement</span></span>| <span data-ttu-id="1ca84-506">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-507">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-508">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-508">1.0</span></span>|
|[<span data-ttu-id="1ca84-509">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-510">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-511">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-512">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-512">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-513">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-513">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="1ca84-514">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1ca84-514">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="1ca84-515">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-515">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="1ca84-516">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="1ca84-516">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1ca84-517">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1ca84-517">Read mode</span></span>

<span data-ttu-id="1ca84-518">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-518">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1ca84-519">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1ca84-519">Compose mode</span></span>

<span data-ttu-id="1ca84-520">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-520">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-521">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-521">Type:</span></span>

*   <span data-ttu-id="1ca84-522">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1ca84-522">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-523">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-523">Requirements</span></span>

|<span data-ttu-id="1ca84-524">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-524">Requirement</span></span>| <span data-ttu-id="1ca84-525">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-525">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-526">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-526">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-527">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-527">1.0</span></span>|
|[<span data-ttu-id="1ca84-528">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-528">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-529">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-529">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-530">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-530">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-531">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-531">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-532">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-532">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="1ca84-533">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1ca84-533">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="1ca84-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="1ca84-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1ca84-538">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="1ca84-538">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-539">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-539">Type:</span></span>

*   [<span data-ttu-id="1ca84-540">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1ca84-540">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1ca84-541">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-541">Requirements</span></span>

|<span data-ttu-id="1ca84-542">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-542">Requirement</span></span>| <span data-ttu-id="1ca84-543">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-543">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-544">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-544">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-545">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-545">1.0</span></span>|
|[<span data-ttu-id="1ca84-546">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-546">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-547">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-547">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-548">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-548">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-549">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-549">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-550">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-550">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="1ca84-551">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="1ca84-551">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="1ca84-552">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-552">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="1ca84-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1ca84-555">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1ca84-555">Read mode</span></span>

<span data-ttu-id="1ca84-556">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-556">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1ca84-557">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1ca84-557">Compose mode</span></span>

<span data-ttu-id="1ca84-558">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-558">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="1ca84-559">[`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1ca84-559">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-560">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-560">Type:</span></span>

*   <span data-ttu-id="1ca84-561">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="1ca84-561">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-562">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-562">Requirements</span></span>

|<span data-ttu-id="1ca84-563">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-563">Requirement</span></span>| <span data-ttu-id="1ca84-564">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-564">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-565">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-565">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-566">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-566">1.0</span></span>|
|[<span data-ttu-id="1ca84-567">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-567">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-568">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-568">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-569">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-569">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-570">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-570">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-571">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-571">Example</span></span>

<span data-ttu-id="1ca84-572">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-572">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="1ca84-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="1ca84-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="1ca84-574">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="1ca84-575">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1ca84-576">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1ca84-576">Read mode</span></span>

<span data-ttu-id="1ca84-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="1ca84-579">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1ca84-579">Compose mode</span></span>

<span data-ttu-id="1ca84-580">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1ca84-581">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-581">Type:</span></span>

*   <span data-ttu-id="1ca84-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="1ca84-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-583">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-583">Requirements</span></span>

|<span data-ttu-id="1ca84-584">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-584">Requirement</span></span>| <span data-ttu-id="1ca84-585">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-586">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-587">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-587">1.0</span></span>|
|[<span data-ttu-id="1ca84-588">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-588">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-589">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-590">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-590">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-591">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-591">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="1ca84-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1ca84-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="1ca84-593">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="1ca84-594">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="1ca84-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1ca84-595">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1ca84-595">Read mode</span></span>

<span data-ttu-id="1ca84-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1ca84-598">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1ca84-598">Compose mode</span></span>

<span data-ttu-id="1ca84-599">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="1ca84-600">型:</span><span class="sxs-lookup"><span data-stu-id="1ca84-600">Type:</span></span>

*   <span data-ttu-id="1ca84-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1ca84-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-602">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-602">Requirements</span></span>

|<span data-ttu-id="1ca84-603">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-603">Requirement</span></span>| <span data-ttu-id="1ca84-604">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-605">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-606">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-606">1.0</span></span>|
|[<span data-ttu-id="1ca84-607">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-607">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-608">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-609">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-609">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-610">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-610">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-611">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-611">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="1ca84-612">メソッド</span><span class="sxs-lookup"><span data-stu-id="1ca84-612">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="1ca84-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1ca84-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1ca84-614">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-614">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="1ca84-615">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-615">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="1ca84-616">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-616">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1ca84-617">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="1ca84-617">Parameters:</span></span>

|<span data-ttu-id="1ca84-618">名前</span><span class="sxs-lookup"><span data-stu-id="1ca84-618">Name</span></span>| <span data-ttu-id="1ca84-619">型</span><span class="sxs-lookup"><span data-stu-id="1ca84-619">Type</span></span>| <span data-ttu-id="1ca84-620">属性</span><span class="sxs-lookup"><span data-stu-id="1ca84-620">Attributes</span></span>| <span data-ttu-id="1ca84-621">説明</span><span class="sxs-lookup"><span data-stu-id="1ca84-621">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="1ca84-622">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-622">String</span></span>||<span data-ttu-id="1ca84-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="1ca84-625">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-625">String</span></span>||<span data-ttu-id="1ca84-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="1ca84-628">Object</span><span class="sxs-lookup"><span data-stu-id="1ca84-628">Object</span></span>| <span data-ttu-id="1ca84-629">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-629">&lt;optional&gt;</span></span>|<span data-ttu-id="1ca84-630">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="1ca84-630">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="1ca84-631">Object</span><span class="sxs-lookup"><span data-stu-id="1ca84-631">Object</span></span> | <span data-ttu-id="1ca84-632">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-632">&lt;optional&gt;</span></span> | <span data-ttu-id="1ca84-633">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-633">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="1ca84-634">Boolean</span><span class="sxs-lookup"><span data-stu-id="1ca84-634">Boolean</span></span> | <span data-ttu-id="1ca84-635">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-635">&lt;optional&gt;</span></span> | <span data-ttu-id="1ca84-636">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-636">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="1ca84-637">function</span><span class="sxs-lookup"><span data-stu-id="1ca84-637">function</span></span>| <span data-ttu-id="1ca84-638">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-638">&lt;optional&gt;</span></span>|<span data-ttu-id="1ca84-639">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-639">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1ca84-640">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-640">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1ca84-641">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-641">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1ca84-642">エラー</span><span class="sxs-lookup"><span data-stu-id="1ca84-642">Errors</span></span>

| <span data-ttu-id="1ca84-643">エラー コード</span><span class="sxs-lookup"><span data-stu-id="1ca84-643">Error code</span></span> | <span data-ttu-id="1ca84-644">説明</span><span class="sxs-lookup"><span data-stu-id="1ca84-644">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="1ca84-645">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="1ca84-645">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="1ca84-646">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="1ca84-646">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="1ca84-647">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-647">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1ca84-648">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-648">Requirements</span></span>

|<span data-ttu-id="1ca84-649">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-649">Requirement</span></span>| <span data-ttu-id="1ca84-650">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-650">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-651">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-651">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-652">1.1</span><span class="sxs-lookup"><span data-stu-id="1ca84-652">1.1</span></span>|
|[<span data-ttu-id="1ca84-653">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-653">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-654">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-654">ReadWriteItem</span></span>|
|[<span data-ttu-id="1ca84-655">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-655">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-656">作成</span><span class="sxs-lookup"><span data-stu-id="1ca84-656">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="1ca84-657">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-657">Examples</span></span>

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

<span data-ttu-id="1ca84-658">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-658">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="1ca84-659">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1ca84-659">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1ca84-660">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-660">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="1ca84-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="1ca84-664">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-664">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="1ca84-665">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="1ca84-665">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1ca84-666">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="1ca84-666">Parameters:</span></span>

|<span data-ttu-id="1ca84-667">名前</span><span class="sxs-lookup"><span data-stu-id="1ca84-667">Name</span></span>| <span data-ttu-id="1ca84-668">型</span><span class="sxs-lookup"><span data-stu-id="1ca84-668">Type</span></span>| <span data-ttu-id="1ca84-669">属性</span><span class="sxs-lookup"><span data-stu-id="1ca84-669">Attributes</span></span>| <span data-ttu-id="1ca84-670">説明</span><span class="sxs-lookup"><span data-stu-id="1ca84-670">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="1ca84-671">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-671">String</span></span>||<span data-ttu-id="1ca84-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="1ca84-674">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-674">String</span></span>||<span data-ttu-id="1ca84-p136">添付するアイテムの件名。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="1ca84-677">Object</span><span class="sxs-lookup"><span data-stu-id="1ca84-677">Object</span></span>| <span data-ttu-id="1ca84-678">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-678">&lt;optional&gt;</span></span>|<span data-ttu-id="1ca84-679">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="1ca84-679">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1ca84-680">Object</span><span class="sxs-lookup"><span data-stu-id="1ca84-680">Object</span></span>| <span data-ttu-id="1ca84-681">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-681">&lt;optional&gt;</span></span>|<span data-ttu-id="1ca84-682">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-682">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1ca84-683">function</span><span class="sxs-lookup"><span data-stu-id="1ca84-683">function</span></span>| <span data-ttu-id="1ca84-684">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-684">&lt;optional&gt;</span></span>|<span data-ttu-id="1ca84-685">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-685">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1ca84-686">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-686">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1ca84-687">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-687">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1ca84-688">エラー</span><span class="sxs-lookup"><span data-stu-id="1ca84-688">Errors</span></span>

| <span data-ttu-id="1ca84-689">エラー コード</span><span class="sxs-lookup"><span data-stu-id="1ca84-689">Error code</span></span> | <span data-ttu-id="1ca84-690">説明</span><span class="sxs-lookup"><span data-stu-id="1ca84-690">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="1ca84-691">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-691">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1ca84-692">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-692">Requirements</span></span>

|<span data-ttu-id="1ca84-693">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-693">Requirement</span></span>| <span data-ttu-id="1ca84-694">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-695">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-696">1.1</span><span class="sxs-lookup"><span data-stu-id="1ca84-696">1.1</span></span>|
|[<span data-ttu-id="1ca84-697">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-697">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-698">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-698">ReadWriteItem</span></span>|
|[<span data-ttu-id="1ca84-699">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-699">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-700">作成</span><span class="sxs-lookup"><span data-stu-id="1ca84-700">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-701">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-701">Example</span></span>

<span data-ttu-id="1ca84-702">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-702">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="1ca84-703">close()</span><span class="sxs-lookup"><span data-stu-id="1ca84-703">close()</span></span>

<span data-ttu-id="1ca84-704">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-704">Closes the current item that is being composed.</span></span>

<span data-ttu-id="1ca84-p137">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="1ca84-707">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-707">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="1ca84-708">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="1ca84-708">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-709">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-709">Requirements</span></span>

|<span data-ttu-id="1ca84-710">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-710">Requirement</span></span>| <span data-ttu-id="1ca84-711">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-711">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-712">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-712">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-713">1.3</span><span class="sxs-lookup"><span data-stu-id="1ca84-713">1.3</span></span>|
|[<span data-ttu-id="1ca84-714">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-714">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-715">制限あり</span><span class="sxs-lookup"><span data-stu-id="1ca84-715">Restricted</span></span>|
|[<span data-ttu-id="1ca84-716">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-716">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-717">作成</span><span class="sxs-lookup"><span data-stu-id="1ca84-717">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="1ca84-718">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="1ca84-718">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="1ca84-719">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-719">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1ca84-720">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1ca84-720">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1ca84-721">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-721">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1ca84-722">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="1ca84-722">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="1ca84-p138">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1ca84-726">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="1ca84-726">Parameters:</span></span>

| <span data-ttu-id="1ca84-727">名前</span><span class="sxs-lookup"><span data-stu-id="1ca84-727">Name</span></span> | <span data-ttu-id="1ca84-728">型</span><span class="sxs-lookup"><span data-stu-id="1ca84-728">Type</span></span> | <span data-ttu-id="1ca84-729">属性</span><span class="sxs-lookup"><span data-stu-id="1ca84-729">Attributes</span></span> | <span data-ttu-id="1ca84-730">説明</span><span class="sxs-lookup"><span data-stu-id="1ca84-730">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="1ca84-731">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="1ca84-731">String &#124; Object</span></span>| |<span data-ttu-id="1ca84-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1ca84-734">**または**</span><span class="sxs-lookup"><span data-stu-id="1ca84-734">**OR**</span></span><br/><span data-ttu-id="1ca84-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="1ca84-737">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-737">String</span></span> | <span data-ttu-id="1ca84-738">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-738">&lt;optional&gt;</span></span> | <span data-ttu-id="1ca84-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="1ca84-741">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-741">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="1ca84-742">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-742">&lt;optional&gt;</span></span> | <span data-ttu-id="1ca84-743">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="1ca84-743">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="1ca84-744">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-744">String</span></span> | | <span data-ttu-id="1ca84-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="1ca84-747">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-747">String</span></span> | | <span data-ttu-id="1ca84-748">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="1ca84-748">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="1ca84-749">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-749">String</span></span> | | <span data-ttu-id="1ca84-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="1ca84-752">Boolean</span><span class="sxs-lookup"><span data-stu-id="1ca84-752">Boolean</span></span> | | <span data-ttu-id="1ca84-p144">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="1ca84-755">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-755">String</span></span> | | <span data-ttu-id="1ca84-p145">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="1ca84-759">function</span><span class="sxs-lookup"><span data-stu-id="1ca84-759">function</span></span> | <span data-ttu-id="1ca84-760">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-760">&lt;optional&gt;</span></span> | <span data-ttu-id="1ca84-761">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-761">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1ca84-762">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-762">Requirements</span></span>

|<span data-ttu-id="1ca84-763">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-763">Requirement</span></span>| <span data-ttu-id="1ca84-764">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-764">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-765">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-765">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-766">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-766">1.0</span></span>|
|[<span data-ttu-id="1ca84-767">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-767">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-768">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-768">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-769">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-769">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-770">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-770">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1ca84-771">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-771">Examples</span></span>

<span data-ttu-id="1ca84-772">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-772">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="1ca84-773">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-773">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="1ca84-774">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-774">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1ca84-775">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-775">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1ca84-776">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-776">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1ca84-777">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-777">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="1ca84-778">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="1ca84-778">displayReplyForm(formData)</span></span>

<span data-ttu-id="1ca84-779">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-779">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1ca84-780">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1ca84-780">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1ca84-781">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-781">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1ca84-782">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="1ca84-782">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="1ca84-p146">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1ca84-786">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="1ca84-786">Parameters:</span></span>

| <span data-ttu-id="1ca84-787">名前</span><span class="sxs-lookup"><span data-stu-id="1ca84-787">Name</span></span> | <span data-ttu-id="1ca84-788">型</span><span class="sxs-lookup"><span data-stu-id="1ca84-788">Type</span></span> | <span data-ttu-id="1ca84-789">属性</span><span class="sxs-lookup"><span data-stu-id="1ca84-789">Attributes</span></span> | <span data-ttu-id="1ca84-790">説明</span><span class="sxs-lookup"><span data-stu-id="1ca84-790">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="1ca84-791">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="1ca84-791">String &#124; Object</span></span>| | <span data-ttu-id="1ca84-p147">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1ca84-794">**または**</span><span class="sxs-lookup"><span data-stu-id="1ca84-794">**OR**</span></span><br/><span data-ttu-id="1ca84-p148">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="1ca84-797">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-797">String</span></span> | <span data-ttu-id="1ca84-798">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-798">&lt;optional&gt;</span></span> | <span data-ttu-id="1ca84-p149">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="1ca84-801">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-801">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="1ca84-802">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-802">&lt;optional&gt;</span></span> | <span data-ttu-id="1ca84-803">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="1ca84-803">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="1ca84-804">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-804">String</span></span> | | <span data-ttu-id="1ca84-p150">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="1ca84-807">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-807">String</span></span> | | <span data-ttu-id="1ca84-808">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="1ca84-808">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="1ca84-809">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-809">String</span></span> | | <span data-ttu-id="1ca84-p151">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="1ca84-812">Boolean</span><span class="sxs-lookup"><span data-stu-id="1ca84-812">Boolean</span></span> | | <span data-ttu-id="1ca84-p152">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="1ca84-815">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-815">String</span></span> | | <span data-ttu-id="1ca84-p153">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="1ca84-819">function</span><span class="sxs-lookup"><span data-stu-id="1ca84-819">function</span></span> | <span data-ttu-id="1ca84-820">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-820">&lt;optional&gt;</span></span> | <span data-ttu-id="1ca84-821">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-821">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1ca84-822">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-822">Requirements</span></span>

|<span data-ttu-id="1ca84-823">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-823">Requirement</span></span>| <span data-ttu-id="1ca84-824">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-824">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-825">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-825">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-826">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-826">1.0</span></span>|
|[<span data-ttu-id="1ca84-827">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-827">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-828">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-828">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-829">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-829">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-830">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-830">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1ca84-831">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-831">Examples</span></span>

<span data-ttu-id="1ca84-832">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-832">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="1ca84-833">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-833">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="1ca84-834">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-834">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1ca84-835">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-835">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1ca84-836">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-836">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1ca84-837">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-837">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="1ca84-838">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="1ca84-838">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="1ca84-839">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-839">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1ca84-840">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1ca84-840">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-841">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-841">Requirements</span></span>

|<span data-ttu-id="1ca84-842">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-842">Requirement</span></span>| <span data-ttu-id="1ca84-843">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-843">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-844">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-844">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-845">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-845">1.0</span></span>|
|[<span data-ttu-id="1ca84-846">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-846">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-847">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-847">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-848">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-848">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-849">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-849">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1ca84-850">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1ca84-850">Returns:</span></span>

<span data-ttu-id="1ca84-851">型:[Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="1ca84-851">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="1ca84-852">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-852">Example</span></span>

<span data-ttu-id="1ca84-853">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="1ca84-853">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="1ca84-854">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="1ca84-854">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="1ca84-855">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-855">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1ca84-856">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1ca84-856">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1ca84-857">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="1ca84-857">Parameters:</span></span>

|<span data-ttu-id="1ca84-858">名前</span><span class="sxs-lookup"><span data-stu-id="1ca84-858">Name</span></span>| <span data-ttu-id="1ca84-859">型</span><span class="sxs-lookup"><span data-stu-id="1ca84-859">Type</span></span>| <span data-ttu-id="1ca84-860">説明</span><span class="sxs-lookup"><span data-stu-id="1ca84-860">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="1ca84-861">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="1ca84-861">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="1ca84-862">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="1ca84-862">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1ca84-863">Requirements</span><span class="sxs-lookup"><span data-stu-id="1ca84-863">Requirements</span></span>

|<span data-ttu-id="1ca84-864">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-864">Requirement</span></span>| <span data-ttu-id="1ca84-865">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-866">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-867">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-867">1.0</span></span>|
|[<span data-ttu-id="1ca84-868">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-869">制限あり</span><span class="sxs-lookup"><span data-stu-id="1ca84-869">Restricted</span></span>|
|[<span data-ttu-id="1ca84-870">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-871">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1ca84-872">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1ca84-872">Returns:</span></span>

<span data-ttu-id="1ca84-873">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-873">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="1ca84-874">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-874">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="1ca84-875">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="1ca84-875">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="1ca84-876">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="1ca84-876">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="1ca84-877">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="1ca84-877">Value of `entityType`</span></span> | <span data-ttu-id="1ca84-878">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="1ca84-878">Type of objects in returned array</span></span> | <span data-ttu-id="1ca84-879">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-879">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="1ca84-880">文字列</span><span class="sxs-lookup"><span data-stu-id="1ca84-880">String</span></span> | <span data-ttu-id="1ca84-881">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="1ca84-881">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="1ca84-882">連絡先</span><span class="sxs-lookup"><span data-stu-id="1ca84-882">Contact</span></span> | <span data-ttu-id="1ca84-883">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1ca84-883">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="1ca84-884">文字列</span><span class="sxs-lookup"><span data-stu-id="1ca84-884">String</span></span> | <span data-ttu-id="1ca84-885">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1ca84-885">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="1ca84-886">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="1ca84-886">MeetingSuggestion</span></span> | <span data-ttu-id="1ca84-887">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1ca84-887">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="1ca84-888">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="1ca84-888">PhoneNumber</span></span> | <span data-ttu-id="1ca84-889">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="1ca84-889">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="1ca84-890">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="1ca84-890">TaskSuggestion</span></span> | <span data-ttu-id="1ca84-891">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1ca84-891">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="1ca84-892">文字列</span><span class="sxs-lookup"><span data-stu-id="1ca84-892">String</span></span> | <span data-ttu-id="1ca84-893">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="1ca84-893">**Restricted**</span></span> |

<span data-ttu-id="1ca84-894">型:Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="1ca84-894">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="1ca84-895">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-895">Example</span></span>

<span data-ttu-id="1ca84-896">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-896">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="1ca84-897">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="1ca84-897">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="1ca84-898">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-898">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1ca84-899">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1ca84-899">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1ca84-900">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-900">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1ca84-901">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="1ca84-901">Parameters:</span></span>

|<span data-ttu-id="1ca84-902">名前</span><span class="sxs-lookup"><span data-stu-id="1ca84-902">Name</span></span>| <span data-ttu-id="1ca84-903">型</span><span class="sxs-lookup"><span data-stu-id="1ca84-903">Type</span></span>| <span data-ttu-id="1ca84-904">説明</span><span class="sxs-lookup"><span data-stu-id="1ca84-904">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="1ca84-905">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-905">String</span></span>|<span data-ttu-id="1ca84-906">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="1ca84-906">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1ca84-907">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-907">Requirements</span></span>

|<span data-ttu-id="1ca84-908">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-908">Requirement</span></span>| <span data-ttu-id="1ca84-909">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-909">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-910">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-910">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-911">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-911">1.0</span></span>|
|[<span data-ttu-id="1ca84-912">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-912">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-913">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-913">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-914">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-914">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-915">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-915">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1ca84-916">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1ca84-916">Returns:</span></span>

<span data-ttu-id="1ca84-p155">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="1ca84-919">型:Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="1ca84-919">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="1ca84-920">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1ca84-920">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="1ca84-921">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-921">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1ca84-922">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1ca84-922">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1ca84-p156">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1ca84-926">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="1ca84-926">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1ca84-927">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="1ca84-927">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="1ca84-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-931">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-931">Requirements</span></span>

|<span data-ttu-id="1ca84-932">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-932">Requirement</span></span>| <span data-ttu-id="1ca84-933">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-933">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-934">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-934">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-935">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-935">1.0</span></span>|
|[<span data-ttu-id="1ca84-936">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-936">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-937">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-937">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-938">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-938">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-939">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-939">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1ca84-940">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1ca84-940">Returns:</span></span>

<span data-ttu-id="1ca84-p158">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="1ca84-943">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="1ca84-943">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1ca84-944">Object</span><span class="sxs-lookup"><span data-stu-id="1ca84-944">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1ca84-945">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-945">Example</span></span>

<span data-ttu-id="1ca84-946">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1ca84-946">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="1ca84-947">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="1ca84-947">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="1ca84-948">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-948">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1ca84-949">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1ca84-949">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1ca84-950">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="1ca84-950">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="1ca84-p159">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1ca84-953">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="1ca84-953">Parameters:</span></span>

|<span data-ttu-id="1ca84-954">名前</span><span class="sxs-lookup"><span data-stu-id="1ca84-954">Name</span></span>| <span data-ttu-id="1ca84-955">型</span><span class="sxs-lookup"><span data-stu-id="1ca84-955">Type</span></span>| <span data-ttu-id="1ca84-956">説明</span><span class="sxs-lookup"><span data-stu-id="1ca84-956">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="1ca84-957">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-957">String</span></span>|<span data-ttu-id="1ca84-958">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="1ca84-958">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1ca84-959">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-959">Requirements</span></span>

|<span data-ttu-id="1ca84-960">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-960">Requirement</span></span>| <span data-ttu-id="1ca84-961">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-961">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-962">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-962">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-963">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-963">1.0</span></span>|
|[<span data-ttu-id="1ca84-964">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-964">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-965">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-965">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-966">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-966">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-967">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-967">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1ca84-968">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1ca84-968">Returns:</span></span>

<span data-ttu-id="1ca84-969">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="1ca84-969">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="1ca84-970">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="1ca84-970">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1ca84-971">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="1ca84-971">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1ca84-972">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-972">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="1ca84-973">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="1ca84-973">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="1ca84-974">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-974">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="1ca84-p160">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1ca84-977">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="1ca84-977">Parameters:</span></span>

|<span data-ttu-id="1ca84-978">名前</span><span class="sxs-lookup"><span data-stu-id="1ca84-978">Name</span></span>| <span data-ttu-id="1ca84-979">型</span><span class="sxs-lookup"><span data-stu-id="1ca84-979">Type</span></span>| <span data-ttu-id="1ca84-980">属性</span><span class="sxs-lookup"><span data-stu-id="1ca84-980">Attributes</span></span>| <span data-ttu-id="1ca84-981">説明</span><span class="sxs-lookup"><span data-stu-id="1ca84-981">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="1ca84-982">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1ca84-982">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="1ca84-p161">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="1ca84-986">Object</span><span class="sxs-lookup"><span data-stu-id="1ca84-986">Object</span></span>| <span data-ttu-id="1ca84-987">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-987">&lt;optional&gt;</span></span>|<span data-ttu-id="1ca84-988">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="1ca84-988">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1ca84-989">Object</span><span class="sxs-lookup"><span data-stu-id="1ca84-989">Object</span></span>| <span data-ttu-id="1ca84-990">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-990">&lt;optional&gt;</span></span>|<span data-ttu-id="1ca84-991">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-991">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1ca84-992">function</span><span class="sxs-lookup"><span data-stu-id="1ca84-992">function</span></span>||<span data-ttu-id="1ca84-993">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-993">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1ca84-994">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-994">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="1ca84-995">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="1ca84-995">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1ca84-996">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-996">Requirements</span></span>

|<span data-ttu-id="1ca84-997">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-997">Requirement</span></span>| <span data-ttu-id="1ca84-998">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-998">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-999">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-999">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-1000">1.2</span><span class="sxs-lookup"><span data-stu-id="1ca84-1000">1.2</span></span>|
|[<span data-ttu-id="1ca84-1001">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-1001">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-1002">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-1002">ReadWriteItem</span></span>|
|[<span data-ttu-id="1ca84-1003">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-1003">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-1004">作成</span><span class="sxs-lookup"><span data-stu-id="1ca84-1004">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="1ca84-1005">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1ca84-1005">Returns:</span></span>

<span data-ttu-id="1ca84-1006">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1006">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="1ca84-1007">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="1ca84-1007">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1ca84-1008">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-1008">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1ca84-1009">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-1009">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="1ca84-1010">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="1ca84-1010">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="1ca84-p163">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p163">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="1ca84-1013">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1013">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-1014">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-1014">Requirements</span></span>

|<span data-ttu-id="1ca84-1015">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-1015">Requirement</span></span>| <span data-ttu-id="1ca84-1016">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-1016">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-1017">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-1017">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-1018">1.6</span><span class="sxs-lookup"><span data-stu-id="1ca84-1018">1.6</span></span> |
|[<span data-ttu-id="1ca84-1019">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-1019">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-1020">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-1020">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-1021">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-1021">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-1022">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-1022">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1ca84-1023">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1ca84-1023">Returns:</span></span>

<span data-ttu-id="1ca84-1024">型:[Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="1ca84-1024">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="1ca84-1025">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-1025">Example</span></span>

<span data-ttu-id="1ca84-1026">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1026">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="1ca84-1027">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1ca84-1027">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="1ca84-p164">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="1ca84-1030">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1030">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1ca84-p165">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1ca84-1034">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1034">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1ca84-1035">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1035">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="1ca84-p166">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ca84-1039">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-1039">Requirements</span></span>

|<span data-ttu-id="1ca84-1040">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-1040">Requirement</span></span>| <span data-ttu-id="1ca84-1041">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-1041">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-1042">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-1042">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-1043">1.6</span><span class="sxs-lookup"><span data-stu-id="1ca84-1043">1.6</span></span> |
|[<span data-ttu-id="1ca84-1044">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-1044">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-1045">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-1045">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-1046">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-1046">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-1047">読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-1047">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1ca84-1048">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1ca84-1048">Returns:</span></span>

<span data-ttu-id="1ca84-p167">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="1ca84-1051">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-1051">Example</span></span>

<span data-ttu-id="1ca84-1052">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1052">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="1ca84-1053">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1ca84-1053">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="1ca84-1054">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1054">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="1ca84-p168">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1ca84-1058">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="1ca84-1058">Parameters:</span></span>

|<span data-ttu-id="1ca84-1059">名前</span><span class="sxs-lookup"><span data-stu-id="1ca84-1059">Name</span></span>| <span data-ttu-id="1ca84-1060">型</span><span class="sxs-lookup"><span data-stu-id="1ca84-1060">Type</span></span>| <span data-ttu-id="1ca84-1061">属性</span><span class="sxs-lookup"><span data-stu-id="1ca84-1061">Attributes</span></span>| <span data-ttu-id="1ca84-1062">説明</span><span class="sxs-lookup"><span data-stu-id="1ca84-1062">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1ca84-1063">function</span><span class="sxs-lookup"><span data-stu-id="1ca84-1063">function</span></span>||<span data-ttu-id="1ca84-1064">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1064">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1ca84-1065">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1065">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="1ca84-1066">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1066">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="1ca84-1067">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1ca84-1067">Object</span></span>| <span data-ttu-id="1ca84-1068">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-1068">&lt;optional&gt;</span></span>|<span data-ttu-id="1ca84-1069">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1069">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="1ca84-1070">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1070">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1ca84-1071">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-1071">Requirements</span></span>

|<span data-ttu-id="1ca84-1072">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-1072">Requirement</span></span>| <span data-ttu-id="1ca84-1073">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-1073">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-1074">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-1074">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-1075">1.0</span><span class="sxs-lookup"><span data-stu-id="1ca84-1075">1.0</span></span>|
|[<span data-ttu-id="1ca84-1076">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-1076">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-1077">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-1077">ReadItem</span></span>|
|[<span data-ttu-id="1ca84-1078">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-1078">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-1079">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="1ca84-1079">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-1080">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-1080">Example</span></span>

<span data-ttu-id="1ca84-p171">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="1ca84-1084">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1ca84-1084">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="1ca84-1085">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1085">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="1ca84-p172">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1ca84-1090">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="1ca84-1090">Parameters:</span></span>

|<span data-ttu-id="1ca84-1091">名前</span><span class="sxs-lookup"><span data-stu-id="1ca84-1091">Name</span></span>| <span data-ttu-id="1ca84-1092">型</span><span class="sxs-lookup"><span data-stu-id="1ca84-1092">Type</span></span>| <span data-ttu-id="1ca84-1093">属性</span><span class="sxs-lookup"><span data-stu-id="1ca84-1093">Attributes</span></span>| <span data-ttu-id="1ca84-1094">説明</span><span class="sxs-lookup"><span data-stu-id="1ca84-1094">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="1ca84-1095">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-1095">String</span></span>||<span data-ttu-id="1ca84-1096">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1096">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="1ca84-1097">Object</span><span class="sxs-lookup"><span data-stu-id="1ca84-1097">Object</span></span>| <span data-ttu-id="1ca84-1098">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="1ca84-1099">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1099">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1ca84-1100">Object</span><span class="sxs-lookup"><span data-stu-id="1ca84-1100">Object</span></span>| <span data-ttu-id="1ca84-1101">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-1101">&lt;optional&gt;</span></span>|<span data-ttu-id="1ca84-1102">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1102">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1ca84-1103">function</span><span class="sxs-lookup"><span data-stu-id="1ca84-1103">function</span></span>| <span data-ttu-id="1ca84-1104">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-1104">&lt;optional&gt;</span></span>|<span data-ttu-id="1ca84-1105">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1105">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1ca84-1106">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1106">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1ca84-1107">エラー</span><span class="sxs-lookup"><span data-stu-id="1ca84-1107">Errors</span></span>

| <span data-ttu-id="1ca84-1108">エラー コード</span><span class="sxs-lookup"><span data-stu-id="1ca84-1108">Error code</span></span> | <span data-ttu-id="1ca84-1109">説明</span><span class="sxs-lookup"><span data-stu-id="1ca84-1109">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="1ca84-1110">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1110">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1ca84-1111">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-1111">Requirements</span></span>

|<span data-ttu-id="1ca84-1112">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-1112">Requirement</span></span>| <span data-ttu-id="1ca84-1113">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-1114">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-1114">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-1115">1.1</span><span class="sxs-lookup"><span data-stu-id="1ca84-1115">1.1</span></span>|
|[<span data-ttu-id="1ca84-1116">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-1116">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-1117">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-1117">ReadWriteItem</span></span>|
|[<span data-ttu-id="1ca84-1118">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-1118">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-1119">作成</span><span class="sxs-lookup"><span data-stu-id="1ca84-1119">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-1120">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-1120">Example</span></span>

<span data-ttu-id="1ca84-1121">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1121">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="1ca84-1122">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="1ca84-1122">saveAsync([options], callback)</span></span>

<span data-ttu-id="1ca84-1123">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1123">Asynchronously saves an item.</span></span>

<span data-ttu-id="1ca84-p173">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p173">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="1ca84-1127">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1127">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="1ca84-1128">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1128">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="1ca84-p175">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="1ca84-1132">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1132">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="1ca84-1133">Mac Outlook では、新規作成モードの会議で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1133">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="1ca84-1134">Mac Outlook では、会議で `saveAsync` を呼び出すとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1134">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="1ca84-1135">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1ca84-1136">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="1ca84-1136">Parameters:</span></span>

|<span data-ttu-id="1ca84-1137">名前</span><span class="sxs-lookup"><span data-stu-id="1ca84-1137">Name</span></span>| <span data-ttu-id="1ca84-1138">型</span><span class="sxs-lookup"><span data-stu-id="1ca84-1138">Type</span></span>| <span data-ttu-id="1ca84-1139">属性</span><span class="sxs-lookup"><span data-stu-id="1ca84-1139">Attributes</span></span>| <span data-ttu-id="1ca84-1140">説明</span><span class="sxs-lookup"><span data-stu-id="1ca84-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="1ca84-1141">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1ca84-1141">Object</span></span>| <span data-ttu-id="1ca84-1142">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="1ca84-1143">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1ca84-1144">Object</span><span class="sxs-lookup"><span data-stu-id="1ca84-1144">Object</span></span>| <span data-ttu-id="1ca84-1145">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="1ca84-1146">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1ca84-1147">function</span><span class="sxs-lookup"><span data-stu-id="1ca84-1147">function</span></span>||<span data-ttu-id="1ca84-1148">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1ca84-1149">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1149">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1ca84-1150">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-1150">Requirements</span></span>

|<span data-ttu-id="1ca84-1151">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-1151">Requirement</span></span>| <span data-ttu-id="1ca84-1152">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-1153">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-1153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="1ca84-1154">1.3</span></span>|
|[<span data-ttu-id="1ca84-1155">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-1155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="1ca84-1157">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-1157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-1158">作成</span><span class="sxs-lookup"><span data-stu-id="1ca84-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="1ca84-1159">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-1159">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="1ca84-p177">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="1ca84-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="1ca84-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="1ca84-1163">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="1ca84-p178">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1ca84-1167">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="1ca84-1167">Parameters:</span></span>

|<span data-ttu-id="1ca84-1168">名前</span><span class="sxs-lookup"><span data-stu-id="1ca84-1168">Name</span></span>| <span data-ttu-id="1ca84-1169">型</span><span class="sxs-lookup"><span data-stu-id="1ca84-1169">Type</span></span>| <span data-ttu-id="1ca84-1170">属性</span><span class="sxs-lookup"><span data-stu-id="1ca84-1170">Attributes</span></span>| <span data-ttu-id="1ca84-1171">説明</span><span class="sxs-lookup"><span data-stu-id="1ca84-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="1ca84-1172">String</span><span class="sxs-lookup"><span data-stu-id="1ca84-1172">String</span></span>||<span data-ttu-id="1ca84-p179">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="1ca84-1176">Object</span><span class="sxs-lookup"><span data-stu-id="1ca84-1176">Object</span></span>| <span data-ttu-id="1ca84-1177">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="1ca84-1178">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1ca84-1179">Object</span><span class="sxs-lookup"><span data-stu-id="1ca84-1179">Object</span></span>| <span data-ttu-id="1ca84-1180">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="1ca84-1181">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="1ca84-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1ca84-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="1ca84-1183">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1ca84-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="1ca84-p180">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p180">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="1ca84-p181">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-p181">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="1ca84-1188">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="1ca84-1189">function</span><span class="sxs-lookup"><span data-stu-id="1ca84-1189">function</span></span>||<span data-ttu-id="1ca84-1190">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1ca84-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1ca84-1191">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-1191">Requirements</span></span>

|<span data-ttu-id="1ca84-1192">要件</span><span class="sxs-lookup"><span data-stu-id="1ca84-1192">Requirement</span></span>| <span data-ttu-id="1ca84-1193">値</span><span class="sxs-lookup"><span data-stu-id="1ca84-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ca84-1194">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1ca84-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1ca84-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="1ca84-1195">1.2</span></span>|
|[<span data-ttu-id="1ca84-1196">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1ca84-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1ca84-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1ca84-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="1ca84-1198">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1ca84-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1ca84-1199">作成</span><span class="sxs-lookup"><span data-stu-id="1ca84-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1ca84-1200">例</span><span class="sxs-lookup"><span data-stu-id="1ca84-1200">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
