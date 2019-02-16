---
title: Office. メールボックス-要件セット1.6
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: b6e51438695f95ed1060bc28ae93f98fb0994b45
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068211"
---
# <a name="item"></a><span data-ttu-id="c6c99-102">item</span><span class="sxs-lookup"><span data-stu-id="c6c99-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="c6c99-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="c6c99-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="c6c99-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-106">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-106">Requirements</span></span>

|<span data-ttu-id="c6c99-107">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-107">Requirement</span></span>| <span data-ttu-id="c6c99-108">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-110">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-110">1.0</span></span>|
|[<span data-ttu-id="c6c99-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="c6c99-112">Restricted</span></span>|
|[<span data-ttu-id="c6c99-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-114">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c6c99-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-115">Members and methods</span></span>

| <span data-ttu-id="c6c99-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-116">Member</span></span> | <span data-ttu-id="c6c99-117">種類</span><span class="sxs-lookup"><span data-stu-id="c6c99-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c6c99-118">attachments</span><span class="sxs-lookup"><span data-stu-id="c6c99-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | <span data-ttu-id="c6c99-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-119">Member</span></span> |
| [<span data-ttu-id="c6c99-120">bcc</span><span class="sxs-lookup"><span data-stu-id="c6c99-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="c6c99-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-121">Member</span></span> |
| [<span data-ttu-id="c6c99-122">body</span><span class="sxs-lookup"><span data-stu-id="c6c99-122">body</span></span>](#body-bodyjavascriptapioutlook16officebody) | <span data-ttu-id="c6c99-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-123">Member</span></span> |
| [<span data-ttu-id="c6c99-124">cc</span><span class="sxs-lookup"><span data-stu-id="c6c99-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="c6c99-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-125">Member</span></span> |
| [<span data-ttu-id="c6c99-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="c6c99-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="c6c99-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-127">Member</span></span> |
| [<span data-ttu-id="c6c99-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="c6c99-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="c6c99-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-129">Member</span></span> |
| [<span data-ttu-id="c6c99-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="c6c99-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="c6c99-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-131">Member</span></span> |
| [<span data-ttu-id="c6c99-132">end</span><span class="sxs-lookup"><span data-stu-id="c6c99-132">end</span></span>](#end-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="c6c99-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-133">Member</span></span> |
| [<span data-ttu-id="c6c99-134">from</span><span class="sxs-lookup"><span data-stu-id="c6c99-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="c6c99-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-135">Member</span></span> |
| [<span data-ttu-id="c6c99-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="c6c99-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="c6c99-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-137">Member</span></span> |
| [<span data-ttu-id="c6c99-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="c6c99-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="c6c99-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-139">Member</span></span> |
| [<span data-ttu-id="c6c99-140">itemId</span><span class="sxs-lookup"><span data-stu-id="c6c99-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="c6c99-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-141">Member</span></span> |
| [<span data-ttu-id="c6c99-142">itemType</span><span class="sxs-lookup"><span data-stu-id="c6c99-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | <span data-ttu-id="c6c99-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-143">Member</span></span> |
| [<span data-ttu-id="c6c99-144">location</span><span class="sxs-lookup"><span data-stu-id="c6c99-144">location</span></span>](#location-stringlocationjavascriptapioutlook16officelocation) | <span data-ttu-id="c6c99-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-145">Member</span></span> |
| [<span data-ttu-id="c6c99-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="c6c99-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="c6c99-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-147">Member</span></span> |
| [<span data-ttu-id="c6c99-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="c6c99-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | <span data-ttu-id="c6c99-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-149">Member</span></span> |
| [<span data-ttu-id="c6c99-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="c6c99-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="c6c99-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-151">Member</span></span> |
| [<span data-ttu-id="c6c99-152">organizer</span><span class="sxs-lookup"><span data-stu-id="c6c99-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="c6c99-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-153">Member</span></span> |
| [<span data-ttu-id="c6c99-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="c6c99-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="c6c99-155">Member</span><span class="sxs-lookup"><span data-stu-id="c6c99-155">Member</span></span> |
| [<span data-ttu-id="c6c99-156">sender</span><span class="sxs-lookup"><span data-stu-id="c6c99-156">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="c6c99-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-157">Member</span></span> |
| [<span data-ttu-id="c6c99-158">start</span><span class="sxs-lookup"><span data-stu-id="c6c99-158">start</span></span>](#start-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="c6c99-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-159">Member</span></span> |
| [<span data-ttu-id="c6c99-160">subject</span><span class="sxs-lookup"><span data-stu-id="c6c99-160">subject</span></span>](#subject-stringsubjectjavascriptapioutlook16officesubject) | <span data-ttu-id="c6c99-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-161">Member</span></span> |
| [<span data-ttu-id="c6c99-162">to</span><span class="sxs-lookup"><span data-stu-id="c6c99-162">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="c6c99-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-163">Member</span></span> |
| [<span data-ttu-id="c6c99-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c6c99-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="c6c99-165">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-165">Method</span></span> |
| [<span data-ttu-id="c6c99-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c6c99-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="c6c99-167">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-167">Method</span></span> |
| [<span data-ttu-id="c6c99-168">close</span><span class="sxs-lookup"><span data-stu-id="c6c99-168">close</span></span>](#close) | <span data-ttu-id="c6c99-169">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-169">Method</span></span> |
| [<span data-ttu-id="c6c99-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="c6c99-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="c6c99-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-171">Method</span></span> |
| [<span data-ttu-id="c6c99-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="c6c99-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="c6c99-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-173">Method</span></span> |
| [<span data-ttu-id="c6c99-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="c6c99-174">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="c6c99-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-175">Method</span></span> |
| [<span data-ttu-id="c6c99-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="c6c99-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="c6c99-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-177">Method</span></span> |
| [<span data-ttu-id="c6c99-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="c6c99-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="c6c99-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-179">Method</span></span> |
| [<span data-ttu-id="c6c99-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c6c99-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="c6c99-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-181">Method</span></span> |
| [<span data-ttu-id="c6c99-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="c6c99-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="c6c99-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-183">Method</span></span> |
| [<span data-ttu-id="c6c99-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c6c99-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="c6c99-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-185">Method</span></span> |
| [<span data-ttu-id="c6c99-186">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="c6c99-186">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="c6c99-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-187">Method</span></span> |
| [<span data-ttu-id="c6c99-188">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c6c99-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="c6c99-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-189">Method</span></span> |
| [<span data-ttu-id="c6c99-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="c6c99-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="c6c99-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-191">Method</span></span> |
| [<span data-ttu-id="c6c99-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c6c99-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="c6c99-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-193">Method</span></span> |
| [<span data-ttu-id="c6c99-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="c6c99-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="c6c99-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-195">Method</span></span> |
| [<span data-ttu-id="c6c99-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c6c99-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="c6c99-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="c6c99-198">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-198">Example</span></span>

<span data-ttu-id="c6c99-199">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c6c99-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="c6c99-200">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6c99-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="c6c99-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c6c99-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="c6c99-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c6c99-204">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="c6c99-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="c6c99-205">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c6c99-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="c6c99-206">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-206">Type</span></span>

*   <span data-ttu-id="c6c99-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c6c99-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-208">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-208">Requirements</span></span>

|<span data-ttu-id="c6c99-209">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-209">Requirement</span></span>| <span data-ttu-id="c6c99-210">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-211">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-212">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-212">1.0</span></span>|
|[<span data-ttu-id="c6c99-213">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-213">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-214">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-215">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-215">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-216">読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6c99-217">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-217">Example</span></span>

<span data-ttu-id="c6c99-218">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="c6c99-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c6c99-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="c6c99-220">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="c6c99-221">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c6c99-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c6c99-222">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-222">Type</span></span>

*   [<span data-ttu-id="c6c99-223">Recipients</span><span class="sxs-lookup"><span data-stu-id="c6c99-223">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="c6c99-224">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-224">Requirements</span></span>

|<span data-ttu-id="c6c99-225">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-225">Requirement</span></span>| <span data-ttu-id="c6c99-226">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-227">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-228">1.1</span><span class="sxs-lookup"><span data-stu-id="c6c99-228">1.1</span></span>|
|[<span data-ttu-id="c6c99-229">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-229">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-230">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-231">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-231">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-232">作成</span><span class="sxs-lookup"><span data-stu-id="c6c99-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c6c99-233">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-233">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="c6c99-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="c6c99-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="c6c99-235">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c6c99-236">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-236">Type</span></span>

*   [<span data-ttu-id="c6c99-237">Body</span><span class="sxs-lookup"><span data-stu-id="c6c99-237">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="c6c99-238">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-238">Requirements</span></span>

|<span data-ttu-id="c6c99-239">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-239">Requirement</span></span>| <span data-ttu-id="c6c99-240">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-241">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-242">1.1</span><span class="sxs-lookup"><span data-stu-id="c6c99-242">1.1</span></span>|
|[<span data-ttu-id="c6c99-243">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-243">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-244">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-245">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-246">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6c99-247">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-247">Example</span></span>

<span data-ttu-id="c6c99-248">この例では、メッセージの本文をプレーンテキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-248">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="c6c99-249">次の例は、コールバック関数に渡される result パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="c6c99-249">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="c6c99-250">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c6c99-250">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="c6c99-251">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="c6c99-252">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c6c99-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6c99-253">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c6c99-253">Read mode</span></span>

<span data-ttu-id="c6c99-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="c6c99-256">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c6c99-256">Compose mode</span></span>

<span data-ttu-id="c6c99-257">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-257">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c6c99-258">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-258">Type</span></span>

*   <span data-ttu-id="c6c99-259">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c6c99-259">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-260">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-260">Requirements</span></span>

|<span data-ttu-id="c6c99-261">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-261">Requirement</span></span>| <span data-ttu-id="c6c99-262">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-263">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-264">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-264">1.0</span></span>|
|[<span data-ttu-id="c6c99-265">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-265">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-266">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-267">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-267">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-268">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-268">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="c6c99-269">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="c6c99-269">(nullable) conversationId :String</span></span>

<span data-ttu-id="c6c99-270">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="c6c99-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="c6c99-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="c6c99-275">Type</span><span class="sxs-lookup"><span data-stu-id="c6c99-275">Type</span></span>

*   <span data-ttu-id="c6c99-276">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-277">Requirements</span><span class="sxs-lookup"><span data-stu-id="c6c99-277">Requirements</span></span>

|<span data-ttu-id="c6c99-278">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-278">Requirement</span></span>| <span data-ttu-id="c6c99-279">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-280">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-281">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-281">1.0</span></span>|
|[<span data-ttu-id="c6c99-282">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-283">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-284">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-285">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6c99-286">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-286">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="c6c99-287">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="c6c99-287">dateTimeCreated :Date</span></span>

<span data-ttu-id="c6c99-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c6c99-290">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-290">Type</span></span>

*   <span data-ttu-id="c6c99-291">日付</span><span class="sxs-lookup"><span data-stu-id="c6c99-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-292">Requirements</span><span class="sxs-lookup"><span data-stu-id="c6c99-292">Requirements</span></span>

|<span data-ttu-id="c6c99-293">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-293">Requirement</span></span>| <span data-ttu-id="c6c99-294">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-295">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-296">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-296">1.0</span></span>|
|[<span data-ttu-id="c6c99-297">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-297">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-298">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-299">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-299">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-300">Read</span><span class="sxs-lookup"><span data-stu-id="c6c99-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6c99-301">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-301">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="c6c99-302">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="c6c99-302">dateTimeModified :Date</span></span>

<span data-ttu-id="c6c99-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c6c99-305">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c6c99-305">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="c6c99-306">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-306">Type</span></span>

*   <span data-ttu-id="c6c99-307">日付</span><span class="sxs-lookup"><span data-stu-id="c6c99-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-308">Requirements</span><span class="sxs-lookup"><span data-stu-id="c6c99-308">Requirements</span></span>

|<span data-ttu-id="c6c99-309">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-309">Requirement</span></span>| <span data-ttu-id="c6c99-310">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-312">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-312">1.0</span></span>|
|[<span data-ttu-id="c6c99-313">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-313">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-314">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-315">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-316">Read</span><span class="sxs-lookup"><span data-stu-id="c6c99-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6c99-317">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-317">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="c6c99-318">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="c6c99-318">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="c6c99-319">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="c6c99-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6c99-322">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c6c99-322">Read mode</span></span>

<span data-ttu-id="c6c99-323">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-323">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="c6c99-324">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c6c99-324">Compose mode</span></span>

<span data-ttu-id="c6c99-325">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="c6c99-326">[`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c6c99-326">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="c6c99-327">次の例では、 [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) `Time`オブジェクトのメソッドを使用して予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="c6c99-328">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-328">Type</span></span>

*   <span data-ttu-id="c6c99-329">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="c6c99-329">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-330">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-330">Requirements</span></span>

|<span data-ttu-id="c6c99-331">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-331">Requirement</span></span>| <span data-ttu-id="c6c99-332">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-333">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-334">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-334">1.0</span></span>|
|[<span data-ttu-id="c6c99-335">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-335">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-336">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-337">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-337">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-338">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-338">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="c6c99-339">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c6c99-339">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="c6c99-p112">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="c6c99-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c6c99-344">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="c6c99-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c6c99-345">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-345">Type</span></span>

*   [<span data-ttu-id="c6c99-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c6c99-346">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="example"></a><span data-ttu-id="c6c99-347">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-347">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="c6c99-348">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-348">Requirements</span></span>

|<span data-ttu-id="c6c99-349">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-349">Requirement</span></span>| <span data-ttu-id="c6c99-350">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-351">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-352">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-352">1.0</span></span>|
|[<span data-ttu-id="c6c99-353">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-353">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-354">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-355">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-355">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-356">読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-356">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="c6c99-357">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="c6c99-357">internetMessageId :String</span></span>

<span data-ttu-id="c6c99-p114">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c6c99-360">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-360">Type</span></span>

*   <span data-ttu-id="c6c99-361">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-362">Requirements</span><span class="sxs-lookup"><span data-stu-id="c6c99-362">Requirements</span></span>

|<span data-ttu-id="c6c99-363">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-363">Requirement</span></span>| <span data-ttu-id="c6c99-364">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-365">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-366">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-366">1.0</span></span>|
|[<span data-ttu-id="c6c99-367">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-367">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-368">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-369">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-369">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-370">Read</span><span class="sxs-lookup"><span data-stu-id="c6c99-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6c99-371">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-371">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="c6c99-372">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="c6c99-372">itemClass :String</span></span>

<span data-ttu-id="c6c99-p115">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="c6c99-p116">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="c6c99-377">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-377">Type</span></span> | <span data-ttu-id="c6c99-378">説明</span><span class="sxs-lookup"><span data-stu-id="c6c99-378">Description</span></span> | <span data-ttu-id="c6c99-379">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="c6c99-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="c6c99-380">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="c6c99-380">Appointment items</span></span> | <span data-ttu-id="c6c99-381">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c6c99-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="c6c99-382">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="c6c99-382">Message items</span></span> | <span data-ttu-id="c6c99-383">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="c6c99-384">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="c6c99-385">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-385">Type</span></span>

*   <span data-ttu-id="c6c99-386">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-387">Requirements</span><span class="sxs-lookup"><span data-stu-id="c6c99-387">Requirements</span></span>

|<span data-ttu-id="c6c99-388">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-388">Requirement</span></span>| <span data-ttu-id="c6c99-389">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-390">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-391">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-391">1.0</span></span>|
|[<span data-ttu-id="c6c99-392">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-392">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-393">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-394">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-394">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-395">Read</span><span class="sxs-lookup"><span data-stu-id="c6c99-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6c99-396">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-396">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="c6c99-397">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="c6c99-397">(nullable) itemId :String</span></span>

<span data-ttu-id="c6c99-p117">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c6c99-400">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="c6c99-400">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c6c99-401">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="c6c99-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="c6c99-402">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c6c99-402">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c6c99-403">詳細は、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c6c99-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="c6c99-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="c6c99-406">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-406">Type</span></span>

*   <span data-ttu-id="c6c99-407">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-408">Requirements</span><span class="sxs-lookup"><span data-stu-id="c6c99-408">Requirements</span></span>

|<span data-ttu-id="c6c99-409">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-409">Requirement</span></span>| <span data-ttu-id="c6c99-410">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-411">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-412">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-412">1.0</span></span>|
|[<span data-ttu-id="c6c99-413">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-413">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-414">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-415">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-415">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-416">読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6c99-417">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-417">Example</span></span>

<span data-ttu-id="c6c99-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="c6c99-420">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="c6c99-420">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="c6c99-421">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-421">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="c6c99-422">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="c6c99-422">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c6c99-423">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-423">Type</span></span>

*   [<span data-ttu-id="c6c99-424">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="c6c99-424">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="c6c99-425">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-425">Requirements</span></span>

|<span data-ttu-id="c6c99-426">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-426">Requirement</span></span>| <span data-ttu-id="c6c99-427">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-428">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-429">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-429">1.0</span></span>|
|[<span data-ttu-id="c6c99-430">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-431">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-432">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-433">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6c99-434">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-434">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="c6c99-435">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="c6c99-435">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="c6c99-436">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-436">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6c99-437">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c6c99-437">Read mode</span></span>

<span data-ttu-id="c6c99-438">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-438">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="c6c99-439">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c6c99-439">Compose mode</span></span>

<span data-ttu-id="c6c99-440">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-440">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c6c99-441">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-441">Type</span></span>

*   <span data-ttu-id="c6c99-442">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="c6c99-442">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-443">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-443">Requirements</span></span>

|<span data-ttu-id="c6c99-444">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-444">Requirement</span></span>| <span data-ttu-id="c6c99-445">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-446">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-447">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-447">1.0</span></span>|
|[<span data-ttu-id="c6c99-448">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-449">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-450">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-451">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-451">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="c6c99-452">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="c6c99-452">normalizedSubject :String</span></span>

<span data-ttu-id="c6c99-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="c6c99-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="c6c99-457">Type</span><span class="sxs-lookup"><span data-stu-id="c6c99-457">Type</span></span>

*   <span data-ttu-id="c6c99-458">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-458">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-459">Requirements</span><span class="sxs-lookup"><span data-stu-id="c6c99-459">Requirements</span></span>

|<span data-ttu-id="c6c99-460">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-460">Requirement</span></span>| <span data-ttu-id="c6c99-461">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-462">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-463">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-463">1.0</span></span>|
|[<span data-ttu-id="c6c99-464">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-464">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-465">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-466">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-466">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-467">読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6c99-468">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-468">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="c6c99-469">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="c6c99-469">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="c6c99-470">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-470">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c6c99-471">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-471">Type</span></span>

*   [<span data-ttu-id="c6c99-472">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="c6c99-472">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="c6c99-473">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-473">Requirements</span></span>

|<span data-ttu-id="c6c99-474">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-474">Requirement</span></span>| <span data-ttu-id="c6c99-475">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-476">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-476">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-477">1.3</span><span class="sxs-lookup"><span data-stu-id="c6c99-477">1.3</span></span>|
|[<span data-ttu-id="c6c99-478">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-478">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-479">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-480">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-480">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-481">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-481">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6c99-482">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-482">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="c6c99-483">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c6c99-483">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="c6c99-484">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-484">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="c6c99-485">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c6c99-485">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6c99-486">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c6c99-486">Read mode</span></span>

<span data-ttu-id="c6c99-487">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-487">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="c6c99-488">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c6c99-488">Compose mode</span></span>

<span data-ttu-id="c6c99-489">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-489">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c6c99-490">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-490">Type</span></span>

*   <span data-ttu-id="c6c99-491">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c6c99-491">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-492">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-492">Requirements</span></span>

|<span data-ttu-id="c6c99-493">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-493">Requirement</span></span>| <span data-ttu-id="c6c99-494">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-495">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-496">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-496">1.0</span></span>|
|[<span data-ttu-id="c6c99-497">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-497">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-498">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-499">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-499">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-500">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-500">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="c6c99-501">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c6c99-501">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="c6c99-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c6c99-504">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-504">Type</span></span>

*   [<span data-ttu-id="c6c99-505">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c6c99-505">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c6c99-506">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-506">Requirements</span></span>

|<span data-ttu-id="c6c99-507">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-507">Requirement</span></span>| <span data-ttu-id="c6c99-508">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-509">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-510">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-510">1.0</span></span>|
|[<span data-ttu-id="c6c99-511">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-511">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-512">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-513">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-513">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-514">読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-514">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6c99-515">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-515">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="c6c99-516">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c6c99-516">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="c6c99-517">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-517">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="c6c99-518">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c6c99-518">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6c99-519">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c6c99-519">Read mode</span></span>

<span data-ttu-id="c6c99-520">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-520">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="c6c99-521">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c6c99-521">Compose mode</span></span>

<span data-ttu-id="c6c99-522">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-522">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="c6c99-523">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-523">Type</span></span>

*   <span data-ttu-id="c6c99-524">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c6c99-524">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-525">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-525">Requirements</span></span>

|<span data-ttu-id="c6c99-526">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-526">Requirement</span></span>| <span data-ttu-id="c6c99-527">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-527">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-528">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-528">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-529">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-529">1.0</span></span>|
|[<span data-ttu-id="c6c99-530">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-530">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-531">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-532">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-532">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-533">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-533">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="c6c99-534">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c6c99-534">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="c6c99-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="c6c99-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c6c99-539">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="c6c99-539">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c6c99-540">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-540">Type</span></span>

*   [<span data-ttu-id="c6c99-541">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c6c99-541">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c6c99-542">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-542">Requirements</span></span>

|<span data-ttu-id="c6c99-543">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-543">Requirement</span></span>| <span data-ttu-id="c6c99-544">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-545">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-546">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-546">1.0</span></span>|
|[<span data-ttu-id="c6c99-547">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-547">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-548">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-549">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-549">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-550">読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-550">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6c99-551">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-551">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="c6c99-552">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="c6c99-552">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="c6c99-553">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-553">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="c6c99-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6c99-556">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c6c99-556">Read mode</span></span>

<span data-ttu-id="c6c99-557">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-557">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="c6c99-558">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c6c99-558">Compose mode</span></span>

<span data-ttu-id="c6c99-559">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-559">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="c6c99-560">[`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c6c99-560">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="c6c99-561">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-561">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="c6c99-562">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-562">Type</span></span>

*   <span data-ttu-id="c6c99-563">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="c6c99-563">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-564">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-564">Requirements</span></span>

|<span data-ttu-id="c6c99-565">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-565">Requirement</span></span>| <span data-ttu-id="c6c99-566">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-567">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-568">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-568">1.0</span></span>|
|[<span data-ttu-id="c6c99-569">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-569">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-570">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-571">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-571">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-572">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-572">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="c6c99-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c6c99-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="c6c99-574">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="c6c99-575">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6c99-576">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c6c99-576">Read mode</span></span>

<span data-ttu-id="c6c99-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="c6c99-579">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c6c99-579">Compose mode</span></span>

<span data-ttu-id="c6c99-580">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="c6c99-581">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-581">Type</span></span>

*   <span data-ttu-id="c6c99-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c6c99-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-583">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-583">Requirements</span></span>

|<span data-ttu-id="c6c99-584">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-584">Requirement</span></span>| <span data-ttu-id="c6c99-585">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-586">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-587">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-587">1.0</span></span>|
|[<span data-ttu-id="c6c99-588">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-588">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-589">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-590">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-590">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-591">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-591">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="c6c99-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c6c99-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="c6c99-593">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="c6c99-594">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c6c99-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6c99-595">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c6c99-595">Read mode</span></span>

<span data-ttu-id="c6c99-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="c6c99-598">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c6c99-598">Compose mode</span></span>

<span data-ttu-id="c6c99-599">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c6c99-600">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-600">Type</span></span>

*   <span data-ttu-id="c6c99-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c6c99-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-602">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-602">Requirements</span></span>

|<span data-ttu-id="c6c99-603">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-603">Requirement</span></span>| <span data-ttu-id="c6c99-604">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-605">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-606">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-606">1.0</span></span>|
|[<span data-ttu-id="c6c99-607">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-607">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-608">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-609">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-609">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-610">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-610">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="c6c99-611">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6c99-611">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="c6c99-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c6c99-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c6c99-613">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c6c99-614">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="c6c99-615">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6c99-616">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6c99-616">Parameters</span></span>

|<span data-ttu-id="c6c99-617">名前</span><span class="sxs-lookup"><span data-stu-id="c6c99-617">Name</span></span>| <span data-ttu-id="c6c99-618">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-618">Type</span></span>| <span data-ttu-id="c6c99-619">属性</span><span class="sxs-lookup"><span data-stu-id="c6c99-619">Attributes</span></span>| <span data-ttu-id="c6c99-620">説明</span><span class="sxs-lookup"><span data-stu-id="c6c99-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="c6c99-621">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-621">String</span></span>||<span data-ttu-id="c6c99-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="c6c99-624">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-624">String</span></span>||<span data-ttu-id="c6c99-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="c6c99-627">Object</span><span class="sxs-lookup"><span data-stu-id="c6c99-627">Object</span></span>| <span data-ttu-id="c6c99-628">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-628">&lt;optional&gt;</span></span>|<span data-ttu-id="c6c99-629">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c6c99-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="c6c99-630">Object</span><span class="sxs-lookup"><span data-stu-id="c6c99-630">Object</span></span> | <span data-ttu-id="c6c99-631">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-631">&lt;optional&gt;</span></span> | <span data-ttu-id="c6c99-632">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="c6c99-633">Boolean</span><span class="sxs-lookup"><span data-stu-id="c6c99-633">Boolean</span></span> | <span data-ttu-id="c6c99-634">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-634">&lt;optional&gt;</span></span> | <span data-ttu-id="c6c99-635">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="c6c99-636">function</span><span class="sxs-lookup"><span data-stu-id="c6c99-636">function</span></span>| <span data-ttu-id="c6c99-637">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-637">&lt;optional&gt;</span></span>|<span data-ttu-id="c6c99-638">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c6c99-639">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c6c99-640">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c6c99-641">エラー</span><span class="sxs-lookup"><span data-stu-id="c6c99-641">Errors</span></span>

| <span data-ttu-id="c6c99-642">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c6c99-642">Error code</span></span> | <span data-ttu-id="c6c99-643">説明</span><span class="sxs-lookup"><span data-stu-id="c6c99-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="c6c99-644">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="c6c99-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="c6c99-645">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="c6c99-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="c6c99-646">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c6c99-647">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-647">Requirements</span></span>

|<span data-ttu-id="c6c99-648">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-648">Requirement</span></span>| <span data-ttu-id="c6c99-649">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-650">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-651">1.1</span><span class="sxs-lookup"><span data-stu-id="c6c99-651">1.1</span></span>|
|[<span data-ttu-id="c6c99-652">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="c6c99-654">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-655">新規作成</span><span class="sxs-lookup"><span data-stu-id="c6c99-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c6c99-656">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-656">Examples</span></span>

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

<span data-ttu-id="c6c99-657">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="c6c99-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c6c99-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c6c99-659">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="c6c99-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="c6c99-663">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="c6c99-664">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="c6c99-664">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6c99-665">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6c99-665">Parameters</span></span>

|<span data-ttu-id="c6c99-666">名前</span><span class="sxs-lookup"><span data-stu-id="c6c99-666">Name</span></span>| <span data-ttu-id="c6c99-667">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-667">Type</span></span>| <span data-ttu-id="c6c99-668">属性</span><span class="sxs-lookup"><span data-stu-id="c6c99-668">Attributes</span></span>| <span data-ttu-id="c6c99-669">説明</span><span class="sxs-lookup"><span data-stu-id="c6c99-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="c6c99-670">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-670">String</span></span>||<span data-ttu-id="c6c99-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="c6c99-673">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-673">String</span></span>||<span data-ttu-id="c6c99-674">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="c6c99-674">The subject of the item to be attached.</span></span> <span data-ttu-id="c6c99-675">最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c6c99-675">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="c6c99-676">Object</span><span class="sxs-lookup"><span data-stu-id="c6c99-676">Object</span></span>| <span data-ttu-id="c6c99-677">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-677">&lt;optional&gt;</span></span>|<span data-ttu-id="c6c99-678">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c6c99-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c6c99-679">Object</span><span class="sxs-lookup"><span data-stu-id="c6c99-679">Object</span></span>| <span data-ttu-id="c6c99-680">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-680">&lt;optional&gt;</span></span>|<span data-ttu-id="c6c99-681">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c6c99-682">function</span><span class="sxs-lookup"><span data-stu-id="c6c99-682">function</span></span>| <span data-ttu-id="c6c99-683">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-683">&lt;optional&gt;</span></span>|<span data-ttu-id="c6c99-684">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c6c99-685">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c6c99-686">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c6c99-687">エラー</span><span class="sxs-lookup"><span data-stu-id="c6c99-687">Errors</span></span>

| <span data-ttu-id="c6c99-688">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c6c99-688">Error code</span></span> | <span data-ttu-id="c6c99-689">説明</span><span class="sxs-lookup"><span data-stu-id="c6c99-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="c6c99-690">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c6c99-691">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-691">Requirements</span></span>

|<span data-ttu-id="c6c99-692">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-692">Requirement</span></span>| <span data-ttu-id="c6c99-693">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-694">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-695">1.1</span><span class="sxs-lookup"><span data-stu-id="c6c99-695">1.1</span></span>|
|[<span data-ttu-id="c6c99-696">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-696">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="c6c99-698">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-698">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-699">作成</span><span class="sxs-lookup"><span data-stu-id="c6c99-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c6c99-700">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-700">Example</span></span>

<span data-ttu-id="c6c99-701">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="c6c99-702">close()</span><span class="sxs-lookup"><span data-stu-id="c6c99-702">close()</span></span>

<span data-ttu-id="c6c99-703">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="c6c99-p137">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="c6c99-706">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="c6c99-707">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="c6c99-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-708">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-708">Requirements</span></span>

|<span data-ttu-id="c6c99-709">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-709">Requirement</span></span>| <span data-ttu-id="c6c99-710">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-711">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-712">1.3</span><span class="sxs-lookup"><span data-stu-id="c6c99-712">1.3</span></span>|
|[<span data-ttu-id="c6c99-713">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-713">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-714">制限あり</span><span class="sxs-lookup"><span data-stu-id="c6c99-714">Restricted</span></span>|
|[<span data-ttu-id="c6c99-715">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-715">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-716">作成</span><span class="sxs-lookup"><span data-stu-id="c6c99-716">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="c6c99-717">displayReplyAllForm (formdata, [callback])</span><span class="sxs-lookup"><span data-stu-id="c6c99-717">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="c6c99-718">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c6c99-719">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c6c99-719">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c6c99-720">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c6c99-721">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="c6c99-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="c6c99-p138">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6c99-725">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6c99-725">Parameters</span></span>

| <span data-ttu-id="c6c99-726">名前</span><span class="sxs-lookup"><span data-stu-id="c6c99-726">Name</span></span> | <span data-ttu-id="c6c99-727">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-727">Type</span></span> | <span data-ttu-id="c6c99-728">属性</span><span class="sxs-lookup"><span data-stu-id="c6c99-728">Attributes</span></span> | <span data-ttu-id="c6c99-729">説明</span><span class="sxs-lookup"><span data-stu-id="c6c99-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="c6c99-730">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="c6c99-730">String &#124; Object</span></span>| |<span data-ttu-id="c6c99-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c6c99-733">**または**</span><span class="sxs-lookup"><span data-stu-id="c6c99-733">**OR**</span></span><br/><span data-ttu-id="c6c99-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="c6c99-736">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-736">String</span></span> | <span data-ttu-id="c6c99-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-737">&lt;optional&gt;</span></span> | <span data-ttu-id="c6c99-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="c6c99-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="c6c99-741">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-741">&lt;optional&gt;</span></span> | <span data-ttu-id="c6c99-742">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="c6c99-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="c6c99-743">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-743">String</span></span> | | <span data-ttu-id="c6c99-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="c6c99-746">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-746">String</span></span> | | <span data-ttu-id="c6c99-747">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c6c99-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="c6c99-748">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-748">String</span></span> | | <span data-ttu-id="c6c99-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="c6c99-751">Boolean</span><span class="sxs-lookup"><span data-stu-id="c6c99-751">Boolean</span></span> | | <span data-ttu-id="c6c99-p144">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="c6c99-754">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-754">String</span></span> | | <span data-ttu-id="c6c99-p145">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="c6c99-758">function</span><span class="sxs-lookup"><span data-stu-id="c6c99-758">function</span></span> | <span data-ttu-id="c6c99-759">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-759">&lt;optional&gt;</span></span> | <span data-ttu-id="c6c99-760">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c6c99-761">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-761">Requirements</span></span>

|<span data-ttu-id="c6c99-762">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-762">Requirement</span></span>| <span data-ttu-id="c6c99-763">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-764">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-765">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-765">1.0</span></span>|
|[<span data-ttu-id="c6c99-766">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-766">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-767">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-768">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-768">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-769">読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c6c99-770">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-770">Examples</span></span>

<span data-ttu-id="c6c99-771">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="c6c99-772">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-772">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="c6c99-773">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-773">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c6c99-774">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c6c99-775">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c6c99-776">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="c6c99-777">displayReplyForm (formdata, [callback])</span><span class="sxs-lookup"><span data-stu-id="c6c99-777">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="c6c99-778">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c6c99-779">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c6c99-779">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c6c99-780">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c6c99-781">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="c6c99-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="c6c99-p146">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6c99-785">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6c99-785">Parameters</span></span>

| <span data-ttu-id="c6c99-786">名前</span><span class="sxs-lookup"><span data-stu-id="c6c99-786">Name</span></span> | <span data-ttu-id="c6c99-787">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-787">Type</span></span> | <span data-ttu-id="c6c99-788">属性</span><span class="sxs-lookup"><span data-stu-id="c6c99-788">Attributes</span></span> | <span data-ttu-id="c6c99-789">説明</span><span class="sxs-lookup"><span data-stu-id="c6c99-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="c6c99-790">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="c6c99-790">String &#124; Object</span></span>| | <span data-ttu-id="c6c99-p147">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c6c99-793">**または**</span><span class="sxs-lookup"><span data-stu-id="c6c99-793">**OR**</span></span><br/><span data-ttu-id="c6c99-p148">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="c6c99-796">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-796">String</span></span> | <span data-ttu-id="c6c99-797">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-797">&lt;optional&gt;</span></span> | <span data-ttu-id="c6c99-p149">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="c6c99-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="c6c99-801">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-801">&lt;optional&gt;</span></span> | <span data-ttu-id="c6c99-802">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="c6c99-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="c6c99-803">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-803">String</span></span> | | <span data-ttu-id="c6c99-p150">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="c6c99-806">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-806">String</span></span> | | <span data-ttu-id="c6c99-807">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c6c99-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="c6c99-808">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-808">String</span></span> | | <span data-ttu-id="c6c99-p151">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="c6c99-811">Boolean</span><span class="sxs-lookup"><span data-stu-id="c6c99-811">Boolean</span></span> | | <span data-ttu-id="c6c99-p152">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="c6c99-814">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-814">String</span></span> | | <span data-ttu-id="c6c99-p153">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="c6c99-818">function</span><span class="sxs-lookup"><span data-stu-id="c6c99-818">function</span></span> | <span data-ttu-id="c6c99-819">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-819">&lt;optional&gt;</span></span> | <span data-ttu-id="c6c99-820">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c6c99-821">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-821">Requirements</span></span>

|<span data-ttu-id="c6c99-822">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-822">Requirement</span></span>| <span data-ttu-id="c6c99-823">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-824">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-825">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-825">1.0</span></span>|
|[<span data-ttu-id="c6c99-826">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-826">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-827">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-828">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-828">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-829">読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c6c99-830">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-830">Examples</span></span>

<span data-ttu-id="c6c99-831">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="c6c99-832">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-832">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="c6c99-833">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-833">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c6c99-834">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c6c99-835">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c6c99-836">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="c6c99-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c6c99-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="c6c99-838">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-838">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c6c99-839">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c6c99-839">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-840">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-840">Requirements</span></span>

|<span data-ttu-id="c6c99-841">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-841">Requirement</span></span>| <span data-ttu-id="c6c99-842">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-843">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-844">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-844">1.0</span></span>|
|[<span data-ttu-id="c6c99-845">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-845">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-846">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-847">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-847">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-848">読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6c99-849">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c6c99-849">Returns:</span></span>

<span data-ttu-id="c6c99-850">型:[Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c6c99-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c6c99-851">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-851">Example</span></span>

<span data-ttu-id="c6c99-852">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="c6c99-852">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="c6c99-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c6c99-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c6c99-854">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-854">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c6c99-855">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c6c99-855">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6c99-856">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6c99-856">Parameters</span></span>

|<span data-ttu-id="c6c99-857">名前</span><span class="sxs-lookup"><span data-stu-id="c6c99-857">Name</span></span>| <span data-ttu-id="c6c99-858">種類</span><span class="sxs-lookup"><span data-stu-id="c6c99-858">Type</span></span>| <span data-ttu-id="c6c99-859">説明</span><span class="sxs-lookup"><span data-stu-id="c6c99-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="c6c99-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="c6c99-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="c6c99-861">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="c6c99-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6c99-862">Requirements</span><span class="sxs-lookup"><span data-stu-id="c6c99-862">Requirements</span></span>

|<span data-ttu-id="c6c99-863">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-863">Requirement</span></span>| <span data-ttu-id="c6c99-864">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-865">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-866">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-866">1.0</span></span>|
|[<span data-ttu-id="c6c99-867">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-867">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-868">制限あり</span><span class="sxs-lookup"><span data-stu-id="c6c99-868">Restricted</span></span>|
|[<span data-ttu-id="c6c99-869">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-869">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-870">読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6c99-871">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c6c99-871">Returns:</span></span>

<span data-ttu-id="c6c99-872">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="c6c99-873">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-873">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="c6c99-874">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="c6c99-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="c6c99-875">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="c6c99-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="c6c99-876">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="c6c99-876">Value of `entityType`</span></span> | <span data-ttu-id="c6c99-877">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="c6c99-877">Type of objects in returned array</span></span> | <span data-ttu-id="c6c99-878">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="c6c99-879">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-879">String</span></span> | <span data-ttu-id="c6c99-880">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c6c99-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="c6c99-881">連絡先</span><span class="sxs-lookup"><span data-stu-id="c6c99-881">Contact</span></span> | <span data-ttu-id="c6c99-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c6c99-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="c6c99-883">文字列</span><span class="sxs-lookup"><span data-stu-id="c6c99-883">String</span></span> | <span data-ttu-id="c6c99-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c6c99-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="c6c99-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="c6c99-885">MeetingSuggestion</span></span> | <span data-ttu-id="c6c99-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c6c99-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="c6c99-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="c6c99-887">PhoneNumber</span></span> | <span data-ttu-id="c6c99-888">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c6c99-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="c6c99-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="c6c99-889">TaskSuggestion</span></span> | <span data-ttu-id="c6c99-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c6c99-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="c6c99-891">文字列</span><span class="sxs-lookup"><span data-stu-id="c6c99-891">String</span></span> | <span data-ttu-id="c6c99-892">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c6c99-892">**Restricted**</span></span> |

<span data-ttu-id="c6c99-893">型:Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c6c99-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="c6c99-894">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-894">Example</span></span>

<span data-ttu-id="c6c99-895">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-895">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="c6c99-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c6c99-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c6c99-897">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c6c99-898">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c6c99-898">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c6c99-899">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6c99-900">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6c99-900">Parameters</span></span>

|<span data-ttu-id="c6c99-901">名前</span><span class="sxs-lookup"><span data-stu-id="c6c99-901">Name</span></span>| <span data-ttu-id="c6c99-902">種類</span><span class="sxs-lookup"><span data-stu-id="c6c99-902">Type</span></span>| <span data-ttu-id="c6c99-903">説明</span><span class="sxs-lookup"><span data-stu-id="c6c99-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="c6c99-904">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-904">String</span></span>|<span data-ttu-id="c6c99-905">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="c6c99-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6c99-906">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-906">Requirements</span></span>

|<span data-ttu-id="c6c99-907">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-907">Requirement</span></span>| <span data-ttu-id="c6c99-908">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-909">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-910">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-910">1.0</span></span>|
|[<span data-ttu-id="c6c99-911">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-911">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-912">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-913">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-913">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-914">読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6c99-915">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c6c99-915">Returns:</span></span>

<span data-ttu-id="c6c99-p155">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="c6c99-918">型:Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c6c99-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="c6c99-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c6c99-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="c6c99-920">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c6c99-921">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c6c99-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c6c99-p156">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c6c99-925">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="c6c99-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c6c99-926">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="c6c99-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c6c99-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-930">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-930">Requirements</span></span>

|<span data-ttu-id="c6c99-931">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-931">Requirement</span></span>| <span data-ttu-id="c6c99-932">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-933">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-934">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-934">1.0</span></span>|
|[<span data-ttu-id="c6c99-935">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-935">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-936">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-937">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-937">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-938">読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6c99-939">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c6c99-939">Returns:</span></span>

<span data-ttu-id="c6c99-p158">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="c6c99-942">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c6c99-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c6c99-943">Object</span><span class="sxs-lookup"><span data-stu-id="c6c99-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c6c99-944">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-944">Example</span></span>

<span data-ttu-id="c6c99-945">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c6c99-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="c6c99-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="c6c99-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="c6c99-947">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c6c99-948">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c6c99-948">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c6c99-949">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="c6c99-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="c6c99-p159">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6c99-952">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6c99-952">Parameters</span></span>

|<span data-ttu-id="c6c99-953">名前</span><span class="sxs-lookup"><span data-stu-id="c6c99-953">Name</span></span>| <span data-ttu-id="c6c99-954">種類</span><span class="sxs-lookup"><span data-stu-id="c6c99-954">Type</span></span>| <span data-ttu-id="c6c99-955">説明</span><span class="sxs-lookup"><span data-stu-id="c6c99-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="c6c99-956">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-956">String</span></span>|<span data-ttu-id="c6c99-957">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="c6c99-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6c99-958">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-958">Requirements</span></span>

|<span data-ttu-id="c6c99-959">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-959">Requirement</span></span>| <span data-ttu-id="c6c99-960">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-961">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-961">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-962">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-962">1.0</span></span>|
|[<span data-ttu-id="c6c99-963">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-963">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-964">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-965">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-965">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-966">読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6c99-967">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c6c99-967">Returns:</span></span>

<span data-ttu-id="c6c99-968">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="c6c99-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="c6c99-969">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c6c99-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c6c99-970">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="c6c99-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c6c99-971">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-971">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="c6c99-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="c6c99-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="c6c99-973">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="c6c99-p160">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6c99-976">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6c99-976">Parameters</span></span>

|<span data-ttu-id="c6c99-977">名前</span><span class="sxs-lookup"><span data-stu-id="c6c99-977">Name</span></span>| <span data-ttu-id="c6c99-978">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-978">Type</span></span>| <span data-ttu-id="c6c99-979">属性</span><span class="sxs-lookup"><span data-stu-id="c6c99-979">Attributes</span></span>| <span data-ttu-id="c6c99-980">説明</span><span class="sxs-lookup"><span data-stu-id="c6c99-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="c6c99-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c6c99-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="c6c99-p161">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="c6c99-985">Object</span><span class="sxs-lookup"><span data-stu-id="c6c99-985">Object</span></span>| <span data-ttu-id="c6c99-986">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-986">&lt;optional&gt;</span></span>|<span data-ttu-id="c6c99-987">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c6c99-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c6c99-988">Object</span><span class="sxs-lookup"><span data-stu-id="c6c99-988">Object</span></span>| <span data-ttu-id="c6c99-989">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-989">&lt;optional&gt;</span></span>|<span data-ttu-id="c6c99-990">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c6c99-991">function</span><span class="sxs-lookup"><span data-stu-id="c6c99-991">function</span></span>||<span data-ttu-id="c6c99-992">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c6c99-993">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="c6c99-994">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="c6c99-994">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6c99-995">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-995">Requirements</span></span>

|<span data-ttu-id="c6c99-996">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-996">Requirement</span></span>| <span data-ttu-id="c6c99-997">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-998">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-998">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-999">1.2</span><span class="sxs-lookup"><span data-stu-id="c6c99-999">1.2</span></span>|
|[<span data-ttu-id="c6c99-1000">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-1000">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="c6c99-1002">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-1002">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-1003">作成</span><span class="sxs-lookup"><span data-stu-id="c6c99-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6c99-1004">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c6c99-1004">Returns:</span></span>

<span data-ttu-id="c6c99-1005">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="c6c99-1006">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c6c99-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c6c99-1007">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c6c99-1008">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-1008">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="c6c99-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c6c99-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="c6c99-p163">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p163">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c6c99-1012">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1012">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-1013">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-1013">Requirements</span></span>

|<span data-ttu-id="c6c99-1014">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-1014">Requirement</span></span>| <span data-ttu-id="c6c99-1015">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-1016">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-1016">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="c6c99-1017">1.6</span></span> |
|[<span data-ttu-id="c6c99-1018">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-1018">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-1019">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-1020">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-1020">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-1021">読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6c99-1022">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c6c99-1022">Returns:</span></span>

<span data-ttu-id="c6c99-1023">型:[Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c6c99-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c6c99-1024">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-1024">Example</span></span>

<span data-ttu-id="c6c99-1025">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="c6c99-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c6c99-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="c6c99-p164">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c6c99-1029">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1029">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c6c99-p165">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c6c99-1033">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c6c99-1034">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c6c99-p166">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6c99-1038">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-1038">Requirements</span></span>

|<span data-ttu-id="c6c99-1039">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-1039">Requirement</span></span>| <span data-ttu-id="c6c99-1040">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-1041">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="c6c99-1042">1.6</span></span> |
|[<span data-ttu-id="c6c99-1043">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-1043">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-1044">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-1045">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-1045">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-1046">読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6c99-1047">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c6c99-1047">Returns:</span></span>

<span data-ttu-id="c6c99-p167">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="c6c99-1050">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-1050">Example</span></span>

<span data-ttu-id="c6c99-1051">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="c6c99-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c6c99-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="c6c99-1053">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="c6c99-p168">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6c99-1057">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6c99-1057">Parameters</span></span>

|<span data-ttu-id="c6c99-1058">名前</span><span class="sxs-lookup"><span data-stu-id="c6c99-1058">Name</span></span>| <span data-ttu-id="c6c99-1059">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-1059">Type</span></span>| <span data-ttu-id="c6c99-1060">属性</span><span class="sxs-lookup"><span data-stu-id="c6c99-1060">Attributes</span></span>| <span data-ttu-id="c6c99-1061">説明</span><span class="sxs-lookup"><span data-stu-id="c6c99-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="c6c99-1062">function</span><span class="sxs-lookup"><span data-stu-id="c6c99-1062">function</span></span>||<span data-ttu-id="c6c99-1063">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c6c99-1064">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c6c99-1065">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1065">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="c6c99-1066">Object</span><span class="sxs-lookup"><span data-stu-id="c6c99-1066">Object</span></span>| <span data-ttu-id="c6c99-1067">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="c6c99-1068">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1068">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="c6c99-1069">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6c99-1070">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-1070">Requirements</span></span>

|<span data-ttu-id="c6c99-1071">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-1071">Requirement</span></span>| <span data-ttu-id="c6c99-1072">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-1073">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-1073">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="c6c99-1074">1.0</span></span>|
|[<span data-ttu-id="c6c99-1075">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-1075">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-1076">ReadItem</span></span>|
|[<span data-ttu-id="c6c99-1077">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-1077">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-1078">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c6c99-1078">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6c99-1079">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-1079">Example</span></span>

<span data-ttu-id="c6c99-p171">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="c6c99-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c6c99-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="c6c99-1084">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="c6c99-p172">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6c99-1089">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6c99-1089">Parameters</span></span>

|<span data-ttu-id="c6c99-1090">名前</span><span class="sxs-lookup"><span data-stu-id="c6c99-1090">Name</span></span>| <span data-ttu-id="c6c99-1091">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-1091">Type</span></span>| <span data-ttu-id="c6c99-1092">属性</span><span class="sxs-lookup"><span data-stu-id="c6c99-1092">Attributes</span></span>| <span data-ttu-id="c6c99-1093">説明</span><span class="sxs-lookup"><span data-stu-id="c6c99-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="c6c99-1094">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-1094">String</span></span>||<span data-ttu-id="c6c99-1095">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1095">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="c6c99-1096">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c6c99-1096">Object</span></span>| <span data-ttu-id="c6c99-1097">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="c6c99-1098">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1098">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c6c99-1099">Object</span><span class="sxs-lookup"><span data-stu-id="c6c99-1099">Object</span></span>| <span data-ttu-id="c6c99-1100">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-1100">&lt;optional&gt;</span></span>|<span data-ttu-id="c6c99-1101">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1101">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c6c99-1102">function</span><span class="sxs-lookup"><span data-stu-id="c6c99-1102">function</span></span>| <span data-ttu-id="c6c99-1103">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-1103">&lt;optional&gt;</span></span>|<span data-ttu-id="c6c99-1104">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1104">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c6c99-1105">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1105">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c6c99-1106">エラー</span><span class="sxs-lookup"><span data-stu-id="c6c99-1106">Errors</span></span>

| <span data-ttu-id="c6c99-1107">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c6c99-1107">Error code</span></span> | <span data-ttu-id="c6c99-1108">説明</span><span class="sxs-lookup"><span data-stu-id="c6c99-1108">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="c6c99-1109">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1109">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c6c99-1110">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-1110">Requirements</span></span>

|<span data-ttu-id="c6c99-1111">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-1111">Requirement</span></span>| <span data-ttu-id="c6c99-1112">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-1113">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-1114">1.1</span><span class="sxs-lookup"><span data-stu-id="c6c99-1114">1.1</span></span>|
|[<span data-ttu-id="c6c99-1115">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-1115">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-1116">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-1116">ReadWriteItem</span></span>|
|[<span data-ttu-id="c6c99-1117">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-1117">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-1118">作成</span><span class="sxs-lookup"><span data-stu-id="c6c99-1118">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c6c99-1119">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-1119">Example</span></span>

<span data-ttu-id="c6c99-1120">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1120">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="c6c99-1121">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c6c99-1121">saveAsync([options], callback)</span></span>

<span data-ttu-id="c6c99-1122">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1122">Asynchronously saves an item.</span></span>

<span data-ttu-id="c6c99-p173">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p173">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="c6c99-1126">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1126">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="c6c99-1127">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1127">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="c6c99-p175">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="c6c99-1131">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1131">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="c6c99-1132">Mac Outlook では、新規作成モードの会議で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1132">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="c6c99-1133">Mac Outlook では、会議で `saveAsync` を呼び出すとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1133">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="c6c99-1134">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1134">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6c99-1135">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6c99-1135">Parameters</span></span>

|<span data-ttu-id="c6c99-1136">名前</span><span class="sxs-lookup"><span data-stu-id="c6c99-1136">Name</span></span>| <span data-ttu-id="c6c99-1137">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-1137">Type</span></span>| <span data-ttu-id="c6c99-1138">属性</span><span class="sxs-lookup"><span data-stu-id="c6c99-1138">Attributes</span></span>| <span data-ttu-id="c6c99-1139">説明</span><span class="sxs-lookup"><span data-stu-id="c6c99-1139">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="c6c99-1140">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c6c99-1140">Object</span></span>| <span data-ttu-id="c6c99-1141">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-1141">&lt;optional&gt;</span></span>|<span data-ttu-id="c6c99-1142">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1142">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c6c99-1143">Object</span><span class="sxs-lookup"><span data-stu-id="c6c99-1143">Object</span></span>| <span data-ttu-id="c6c99-1144">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-1144">&lt;optional&gt;</span></span>|<span data-ttu-id="c6c99-1145">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1145">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c6c99-1146">function</span><span class="sxs-lookup"><span data-stu-id="c6c99-1146">function</span></span>||<span data-ttu-id="c6c99-1147">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1147">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c6c99-1148">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1148">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6c99-1149">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-1149">Requirements</span></span>

|<span data-ttu-id="c6c99-1150">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-1150">Requirement</span></span>| <span data-ttu-id="c6c99-1151">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-1151">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-1152">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-1152">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-1153">1.3</span><span class="sxs-lookup"><span data-stu-id="c6c99-1153">1.3</span></span>|
|[<span data-ttu-id="c6c99-1154">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-1154">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-1155">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-1155">ReadWriteItem</span></span>|
|[<span data-ttu-id="c6c99-1156">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-1156">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-1157">新規作成</span><span class="sxs-lookup"><span data-stu-id="c6c99-1157">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c6c99-1158">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-1158">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="c6c99-p177">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="c6c99-1161">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="c6c99-1161">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="c6c99-1162">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1162">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="c6c99-p178">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6c99-1166">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6c99-1166">Parameters</span></span>

|<span data-ttu-id="c6c99-1167">名前</span><span class="sxs-lookup"><span data-stu-id="c6c99-1167">Name</span></span>| <span data-ttu-id="c6c99-1168">型</span><span class="sxs-lookup"><span data-stu-id="c6c99-1168">Type</span></span>| <span data-ttu-id="c6c99-1169">属性</span><span class="sxs-lookup"><span data-stu-id="c6c99-1169">Attributes</span></span>| <span data-ttu-id="c6c99-1170">説明</span><span class="sxs-lookup"><span data-stu-id="c6c99-1170">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="c6c99-1171">String</span><span class="sxs-lookup"><span data-stu-id="c6c99-1171">String</span></span>||<span data-ttu-id="c6c99-p179">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="c6c99-1175">Object</span><span class="sxs-lookup"><span data-stu-id="c6c99-1175">Object</span></span>| <span data-ttu-id="c6c99-1176">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-1176">&lt;optional&gt;</span></span>|<span data-ttu-id="c6c99-1177">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1177">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c6c99-1178">Object</span><span class="sxs-lookup"><span data-stu-id="c6c99-1178">Object</span></span>| <span data-ttu-id="c6c99-1179">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-1179">&lt;optional&gt;</span></span>|<span data-ttu-id="c6c99-1180">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1180">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="c6c99-1181">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c6c99-1181">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="c6c99-1182">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6c99-1182">&lt;optional&gt;</span></span>|<span data-ttu-id="c6c99-p180">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p180">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="c6c99-p181">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-p181">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="c6c99-1187">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1187">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="c6c99-1188">function</span><span class="sxs-lookup"><span data-stu-id="c6c99-1188">function</span></span>||<span data-ttu-id="c6c99-1189">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c6c99-1189">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c6c99-1190">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-1190">Requirements</span></span>

|<span data-ttu-id="c6c99-1191">要件</span><span class="sxs-lookup"><span data-stu-id="c6c99-1191">Requirement</span></span>| <span data-ttu-id="c6c99-1192">値</span><span class="sxs-lookup"><span data-stu-id="c6c99-1192">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6c99-1193">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6c99-1193">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6c99-1194">1.2</span><span class="sxs-lookup"><span data-stu-id="c6c99-1194">1.2</span></span>|
|[<span data-ttu-id="c6c99-1195">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6c99-1195">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6c99-1196">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c6c99-1196">ReadWriteItem</span></span>|
|[<span data-ttu-id="c6c99-1197">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6c99-1197">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c6c99-1198">作成</span><span class="sxs-lookup"><span data-stu-id="c6c99-1198">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c6c99-1199">例</span><span class="sxs-lookup"><span data-stu-id="c6c99-1199">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
