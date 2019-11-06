---
title: Office. メールボックス-要件セット1.6
description: ''
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: f5789037ab5486fecf6e821dc39dc4b627e7f825
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001587"
---
# <a name="item"></a><span data-ttu-id="ca3e1-102">item</span><span class="sxs-lookup"><span data-stu-id="ca3e1-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="ca3e1-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="ca3e1-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="ca3e1-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-106">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-106">Requirements</span></span>

|<span data-ttu-id="ca3e1-107">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-107">Requirement</span></span>| <span data-ttu-id="ca3e1-108">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-110">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-110">1.0</span></span>|
|[<span data-ttu-id="ca3e1-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="ca3e1-112">Restricted</span></span>|
|[<span data-ttu-id="ca3e1-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ca3e1-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ca3e1-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-115">Members and methods</span></span>

| <span data-ttu-id="ca3e1-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-116">Member</span></span> | <span data-ttu-id="ca3e1-117">種類</span><span class="sxs-lookup"><span data-stu-id="ca3e1-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ca3e1-118">attachments</span><span class="sxs-lookup"><span data-stu-id="ca3e1-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="ca3e1-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-119">Member</span></span> |
| [<span data-ttu-id="ca3e1-120">bcc</span><span class="sxs-lookup"><span data-stu-id="ca3e1-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="ca3e1-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-121">Member</span></span> |
| [<span data-ttu-id="ca3e1-122">body</span><span class="sxs-lookup"><span data-stu-id="ca3e1-122">body</span></span>](#body-body) | <span data-ttu-id="ca3e1-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-123">Member</span></span> |
| [<span data-ttu-id="ca3e1-124">cc</span><span class="sxs-lookup"><span data-stu-id="ca3e1-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ca3e1-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-125">Member</span></span> |
| [<span data-ttu-id="ca3e1-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="ca3e1-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="ca3e1-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-127">Member</span></span> |
| [<span data-ttu-id="ca3e1-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="ca3e1-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="ca3e1-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-129">Member</span></span> |
| [<span data-ttu-id="ca3e1-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="ca3e1-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="ca3e1-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-131">Member</span></span> |
| [<span data-ttu-id="ca3e1-132">end</span><span class="sxs-lookup"><span data-stu-id="ca3e1-132">end</span></span>](#end-datetime) | <span data-ttu-id="ca3e1-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-133">Member</span></span> |
| [<span data-ttu-id="ca3e1-134">from</span><span class="sxs-lookup"><span data-stu-id="ca3e1-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="ca3e1-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-135">Member</span></span> |
| [<span data-ttu-id="ca3e1-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="ca3e1-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="ca3e1-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-137">Member</span></span> |
| [<span data-ttu-id="ca3e1-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="ca3e1-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="ca3e1-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-139">Member</span></span> |
| [<span data-ttu-id="ca3e1-140">itemId</span><span class="sxs-lookup"><span data-stu-id="ca3e1-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="ca3e1-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-141">Member</span></span> |
| [<span data-ttu-id="ca3e1-142">itemType</span><span class="sxs-lookup"><span data-stu-id="ca3e1-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="ca3e1-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-143">Member</span></span> |
| [<span data-ttu-id="ca3e1-144">location</span><span class="sxs-lookup"><span data-stu-id="ca3e1-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="ca3e1-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-145">Member</span></span> |
| [<span data-ttu-id="ca3e1-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="ca3e1-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="ca3e1-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-147">Member</span></span> |
| [<span data-ttu-id="ca3e1-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="ca3e1-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="ca3e1-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-149">Member</span></span> |
| [<span data-ttu-id="ca3e1-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="ca3e1-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ca3e1-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-151">Member</span></span> |
| [<span data-ttu-id="ca3e1-152">organizer</span><span class="sxs-lookup"><span data-stu-id="ca3e1-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="ca3e1-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-153">Member</span></span> |
| [<span data-ttu-id="ca3e1-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="ca3e1-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ca3e1-155">Member</span><span class="sxs-lookup"><span data-stu-id="ca3e1-155">Member</span></span> |
| [<span data-ttu-id="ca3e1-156">sender</span><span class="sxs-lookup"><span data-stu-id="ca3e1-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="ca3e1-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-157">Member</span></span> |
| [<span data-ttu-id="ca3e1-158">start</span><span class="sxs-lookup"><span data-stu-id="ca3e1-158">start</span></span>](#start-datetime) | <span data-ttu-id="ca3e1-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-159">Member</span></span> |
| [<span data-ttu-id="ca3e1-160">subject</span><span class="sxs-lookup"><span data-stu-id="ca3e1-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="ca3e1-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-161">Member</span></span> |
| [<span data-ttu-id="ca3e1-162">to</span><span class="sxs-lookup"><span data-stu-id="ca3e1-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ca3e1-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-163">Member</span></span> |
| [<span data-ttu-id="ca3e1-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ca3e1-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="ca3e1-165">メソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-165">Method</span></span> |
| [<span data-ttu-id="ca3e1-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ca3e1-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="ca3e1-167">メソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-167">Method</span></span> |
| [<span data-ttu-id="ca3e1-168">close</span><span class="sxs-lookup"><span data-stu-id="ca3e1-168">close</span></span>](#close) | <span data-ttu-id="ca3e1-169">メソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-169">Method</span></span> |
| [<span data-ttu-id="ca3e1-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="ca3e1-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="ca3e1-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-171">Method</span></span> |
| [<span data-ttu-id="ca3e1-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="ca3e1-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="ca3e1-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-173">Method</span></span> |
| [<span data-ttu-id="ca3e1-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="ca3e1-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="ca3e1-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-175">Method</span></span> |
| [<span data-ttu-id="ca3e1-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="ca3e1-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="ca3e1-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-177">Method</span></span> |
| [<span data-ttu-id="ca3e1-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="ca3e1-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="ca3e1-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-179">Method</span></span> |
| [<span data-ttu-id="ca3e1-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="ca3e1-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="ca3e1-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-181">Method</span></span> |
| [<span data-ttu-id="ca3e1-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="ca3e1-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="ca3e1-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-183">Method</span></span> |
| [<span data-ttu-id="ca3e1-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="ca3e1-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="ca3e1-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-185">Method</span></span> |
| [<span data-ttu-id="ca3e1-186">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="ca3e1-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="ca3e1-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-187">Method</span></span> |
| [<span data-ttu-id="ca3e1-188">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="ca3e1-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="ca3e1-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-189">Method</span></span> |
| [<span data-ttu-id="ca3e1-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="ca3e1-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="ca3e1-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-191">Method</span></span> |
| [<span data-ttu-id="ca3e1-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ca3e1-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="ca3e1-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-193">Method</span></span> |
| [<span data-ttu-id="ca3e1-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="ca3e1-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="ca3e1-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-195">Method</span></span> |
| [<span data-ttu-id="ca3e1-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="ca3e1-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="ca3e1-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="ca3e1-198">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-198">Example</span></span>

<span data-ttu-id="ca3e1-199">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="ca3e1-200">Members</span><span class="sxs-lookup"><span data-stu-id="ca3e1-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-16"></a><span data-ttu-id="ca3e1-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="ca3e1-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

<span data-ttu-id="ca3e1-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ca3e1-204">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="ca3e1-205">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="ca3e1-206">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-206">Type</span></span>

*   <span data-ttu-id="ca3e1-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="ca3e1-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-208">Requirements</span></span>

|<span data-ttu-id="ca3e1-209">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-209">Requirement</span></span>| <span data-ttu-id="ca3e1-210">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-211">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-212">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-212">1.0</span></span>|
|[<span data-ttu-id="ca3e1-213">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-214">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-215">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-216">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca3e1-217">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-217">Example</span></span>

<span data-ttu-id="ca3e1-218">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="ca3e1-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="ca3e1-220">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="ca3e1-221">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-221">Compose mode only.</span></span>

<span data-ttu-id="ca3e1-222">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-222">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ca3e1-223">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-223">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ca3e1-224">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-224">Get 500 members maximum.</span></span>
- <span data-ttu-id="ca3e1-225">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-225">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="ca3e1-226">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-226">Type</span></span>

*   [<span data-ttu-id="ca3e1-227">受信者</span><span class="sxs-lookup"><span data-stu-id="ca3e1-227">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="ca3e1-228">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-228">Requirements</span></span>

|<span data-ttu-id="ca3e1-229">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-229">Requirement</span></span>| <span data-ttu-id="ca3e1-230">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-231">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-231">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-232">1.1</span><span class="sxs-lookup"><span data-stu-id="ca3e1-232">1.1</span></span>|
|[<span data-ttu-id="ca3e1-233">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-233">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-234">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-234">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-235">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-235">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-236">作成</span><span class="sxs-lookup"><span data-stu-id="ca3e1-236">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ca3e1-237">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-237">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-16"></a><span data-ttu-id="ca3e1-238">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-238">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span></span>

<span data-ttu-id="ca3e1-239">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-239">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="ca3e1-240">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-240">Type</span></span>

*   [<span data-ttu-id="ca3e1-241">Body</span><span class="sxs-lookup"><span data-stu-id="ca3e1-241">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="ca3e1-242">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-242">Requirements</span></span>

|<span data-ttu-id="ca3e1-243">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-243">Requirement</span></span>| <span data-ttu-id="ca3e1-244">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-245">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-246">1.1</span><span class="sxs-lookup"><span data-stu-id="ca3e1-246">1.1</span></span>|
|[<span data-ttu-id="ca3e1-247">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-247">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-248">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-249">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-249">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-250">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ca3e1-250">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca3e1-251">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-251">Example</span></span>

<span data-ttu-id="ca3e1-252">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-252">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="ca3e1-253">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-253">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="ca3e1-254">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-254">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="ca3e1-255">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-255">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="ca3e1-256">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-256">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ca3e1-257">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-257">Read mode</span></span>

<span data-ttu-id="ca3e1-258">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-258">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="ca3e1-259">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-259">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ca3e1-260">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-260">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="ca3e1-261">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-261">Compose mode</span></span>

<span data-ttu-id="ca3e1-262">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-262">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="ca3e1-263">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-263">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ca3e1-264">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-264">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ca3e1-265">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-265">Get 500 members maximum.</span></span>
- <span data-ttu-id="ca3e1-266">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-266">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ca3e1-267">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-267">Type</span></span>

*   <span data-ttu-id="ca3e1-268">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-268">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-269">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-269">Requirements</span></span>

|<span data-ttu-id="ca3e1-270">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-270">Requirement</span></span>| <span data-ttu-id="ca3e1-271">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-271">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-272">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-272">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-273">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-273">1.0</span></span>|
|[<span data-ttu-id="ca3e1-274">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-274">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-275">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-275">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-276">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-276">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-277">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ca3e1-277">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="ca3e1-278">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-278">(nullable) conversationId: String</span></span>

<span data-ttu-id="ca3e1-279">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-279">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="ca3e1-p109">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="ca3e1-p110">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="ca3e1-284">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-284">Type</span></span>

*   <span data-ttu-id="ca3e1-285">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-285">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-286">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-286">Requirements</span></span>

|<span data-ttu-id="ca3e1-287">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-287">Requirement</span></span>| <span data-ttu-id="ca3e1-288">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-289">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-289">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-290">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-290">1.0</span></span>|
|[<span data-ttu-id="ca3e1-291">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-291">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-292">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-292">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-293">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-293">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-294">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ca3e1-294">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca3e1-295">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-295">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="ca3e1-296">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="ca3e1-296">dateTimeCreated: Date</span></span>

<span data-ttu-id="ca3e1-p111">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ca3e1-299">種類</span><span class="sxs-lookup"><span data-stu-id="ca3e1-299">Type</span></span>

*   <span data-ttu-id="ca3e1-300">日付</span><span class="sxs-lookup"><span data-stu-id="ca3e1-300">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-301">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-301">Requirements</span></span>

|<span data-ttu-id="ca3e1-302">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-302">Requirement</span></span>| <span data-ttu-id="ca3e1-303">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-304">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-305">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-305">1.0</span></span>|
|[<span data-ttu-id="ca3e1-306">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-306">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-307">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-308">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-308">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-309">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca3e1-310">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-310">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="ca3e1-311">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="ca3e1-311">dateTimeModified: Date</span></span>

<span data-ttu-id="ca3e1-p112">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ca3e1-314">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-314">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="ca3e1-315">種類</span><span class="sxs-lookup"><span data-stu-id="ca3e1-315">Type</span></span>

*   <span data-ttu-id="ca3e1-316">日付</span><span class="sxs-lookup"><span data-stu-id="ca3e1-316">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-317">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-317">Requirements</span></span>

|<span data-ttu-id="ca3e1-318">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-318">Requirement</span></span>| <span data-ttu-id="ca3e1-319">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-320">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-321">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-321">1.0</span></span>|
|[<span data-ttu-id="ca3e1-322">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-323">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-324">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-325">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-325">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca3e1-326">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-326">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="ca3e1-327">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-327">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="ca3e1-328">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-328">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="ca3e1-p113">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ca3e1-331">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-331">Read mode</span></span>

<span data-ttu-id="ca3e1-332">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-332">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="ca3e1-333">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-333">Compose mode</span></span>

<span data-ttu-id="ca3e1-334">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-334">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="ca3e1-335">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-335">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="ca3e1-336">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-336">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="ca3e1-337">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-337">Type</span></span>

*   <span data-ttu-id="ca3e1-338">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-338">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-339">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-339">Requirements</span></span>

|<span data-ttu-id="ca3e1-340">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-340">Requirement</span></span>| <span data-ttu-id="ca3e1-341">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-342">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-342">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-343">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-343">1.0</span></span>|
|[<span data-ttu-id="ca3e1-344">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-344">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-345">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-346">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-346">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-347">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ca3e1-347">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="ca3e1-348">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-348">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="ca3e1-p114">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p114">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="ca3e1-p115">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p115">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ca3e1-353">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-353">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="ca3e1-354">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-354">Type</span></span>

*   [<span data-ttu-id="ca3e1-355">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ca3e1-355">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="example"></a><span data-ttu-id="ca3e1-356">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-356">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="ca3e1-357">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-357">Requirements</span></span>

|<span data-ttu-id="ca3e1-358">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-358">Requirement</span></span>| <span data-ttu-id="ca3e1-359">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-360">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-361">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-361">1.0</span></span>|
|[<span data-ttu-id="ca3e1-362">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-362">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-363">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-364">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-364">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-365">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-365">Read</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="ca3e1-366">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-366">internetMessageId: String</span></span>

<span data-ttu-id="ca3e1-p116">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ca3e1-369">Type</span><span class="sxs-lookup"><span data-stu-id="ca3e1-369">Type</span></span>

*   <span data-ttu-id="ca3e1-370">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-370">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-371">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-371">Requirements</span></span>

|<span data-ttu-id="ca3e1-372">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-372">Requirement</span></span>| <span data-ttu-id="ca3e1-373">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-373">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-374">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-374">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-375">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-375">1.0</span></span>|
|[<span data-ttu-id="ca3e1-376">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-376">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-377">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-378">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-378">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-379">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-379">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca3e1-380">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-380">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="ca3e1-381">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-381">itemClass: String</span></span>

<span data-ttu-id="ca3e1-p117">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="ca3e1-p118">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="ca3e1-386">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-386">Type</span></span> | <span data-ttu-id="ca3e1-387">説明</span><span class="sxs-lookup"><span data-stu-id="ca3e1-387">Description</span></span> | <span data-ttu-id="ca3e1-388">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="ca3e1-388">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="ca3e1-389">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="ca3e1-389">Appointment items</span></span> | <span data-ttu-id="ca3e1-390">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-390">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="ca3e1-391">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="ca3e1-391">Message items</span></span> | <span data-ttu-id="ca3e1-392">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-392">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="ca3e1-393">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-393">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="ca3e1-394">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-394">Type</span></span>

*   <span data-ttu-id="ca3e1-395">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-396">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-396">Requirements</span></span>

|<span data-ttu-id="ca3e1-397">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-397">Requirement</span></span>| <span data-ttu-id="ca3e1-398">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-399">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-399">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-400">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-400">1.0</span></span>|
|[<span data-ttu-id="ca3e1-401">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-401">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-402">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-403">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-403">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-404">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca3e1-405">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-405">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="ca3e1-406">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-406">(nullable) itemId: String</span></span>

<span data-ttu-id="ca3e1-407">現在のアイテムの[Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)を取得します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-407">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item.</span></span> <span data-ttu-id="ca3e1-408">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-408">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ca3e1-409">`itemId`プロパティによって返される識別子は、 [Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)と同じです。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-409">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="ca3e1-410">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-410">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="ca3e1-411">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-411">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="ca3e1-412">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-412">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="ca3e1-p121">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="ca3e1-415">Type</span><span class="sxs-lookup"><span data-stu-id="ca3e1-415">Type</span></span>

*   <span data-ttu-id="ca3e1-416">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-416">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-417">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-417">Requirements</span></span>

|<span data-ttu-id="ca3e1-418">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-418">Requirement</span></span>| <span data-ttu-id="ca3e1-419">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-420">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-420">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-421">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-421">1.0</span></span>|
|[<span data-ttu-id="ca3e1-422">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-422">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-423">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-424">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-424">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-425">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-425">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca3e1-426">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-426">Example</span></span>

<span data-ttu-id="ca3e1-p122">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-16"></a><span data-ttu-id="ca3e1-429">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-429">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span></span>

<span data-ttu-id="ca3e1-430">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-430">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="ca3e1-431">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-431">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="ca3e1-432">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-432">Type</span></span>

*   [<span data-ttu-id="ca3e1-433">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="ca3e1-433">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="ca3e1-434">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-434">Requirements</span></span>

|<span data-ttu-id="ca3e1-435">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-435">Requirement</span></span>| <span data-ttu-id="ca3e1-436">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-436">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-437">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-437">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-438">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-438">1.0</span></span>|
|[<span data-ttu-id="ca3e1-439">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-439">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-440">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-440">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-441">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-441">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-442">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ca3e1-442">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca3e1-443">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-443">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-16"></a><span data-ttu-id="ca3e1-444">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-444">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

<span data-ttu-id="ca3e1-445">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-445">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ca3e1-446">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-446">Read mode</span></span>

<span data-ttu-id="ca3e1-447">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-447">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="ca3e1-448">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-448">Compose mode</span></span>

<span data-ttu-id="ca3e1-449">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-449">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ca3e1-450">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-450">Type</span></span>

*   <span data-ttu-id="ca3e1-451">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-451">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-452">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-452">Requirements</span></span>

|<span data-ttu-id="ca3e1-453">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-453">Requirement</span></span>| <span data-ttu-id="ca3e1-454">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-455">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-455">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-456">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-456">1.0</span></span>|
|[<span data-ttu-id="ca3e1-457">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-457">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-458">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-459">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-459">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-460">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ca3e1-460">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="ca3e1-461">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-461">normalizedSubject: String</span></span>

<span data-ttu-id="ca3e1-p123">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="ca3e1-p124">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="ca3e1-466">Type</span><span class="sxs-lookup"><span data-stu-id="ca3e1-466">Type</span></span>

*   <span data-ttu-id="ca3e1-467">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-467">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-468">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-468">Requirements</span></span>

|<span data-ttu-id="ca3e1-469">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-469">Requirement</span></span>| <span data-ttu-id="ca3e1-470">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-470">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-471">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-471">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-472">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-472">1.0</span></span>|
|[<span data-ttu-id="ca3e1-473">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-473">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-474">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-474">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-475">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-475">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-476">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-476">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca3e1-477">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-477">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-16"></a><span data-ttu-id="ca3e1-478">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-478">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span></span>

<span data-ttu-id="ca3e1-479">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-479">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="ca3e1-480">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-480">Type</span></span>

*   [<span data-ttu-id="ca3e1-481">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="ca3e1-481">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="ca3e1-482">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-482">Requirements</span></span>

|<span data-ttu-id="ca3e1-483">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-483">Requirement</span></span>| <span data-ttu-id="ca3e1-484">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-484">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-485">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-485">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-486">1.3</span><span class="sxs-lookup"><span data-stu-id="ca3e1-486">1.3</span></span>|
|[<span data-ttu-id="ca3e1-487">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-487">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-488">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-488">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-489">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-489">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-490">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ca3e1-490">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca3e1-491">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-491">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="ca3e1-492">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-492">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="ca3e1-493">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-493">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="ca3e1-494">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-494">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ca3e1-495">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-495">Read mode</span></span>

<span data-ttu-id="ca3e1-496">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-496">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="ca3e1-497">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-497">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ca3e1-498">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-498">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="ca3e1-499">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-499">Compose mode</span></span>

<span data-ttu-id="ca3e1-500">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-500">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="ca3e1-501">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-501">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ca3e1-502">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-502">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ca3e1-503">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-503">Get 500 members maximum.</span></span>
- <span data-ttu-id="ca3e1-504">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-504">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ca3e1-505">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-505">Type</span></span>

*   <span data-ttu-id="ca3e1-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-507">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-507">Requirements</span></span>

|<span data-ttu-id="ca3e1-508">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-508">Requirement</span></span>| <span data-ttu-id="ca3e1-509">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-509">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-510">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-510">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-511">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-511">1.0</span></span>|
|[<span data-ttu-id="ca3e1-512">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-512">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-513">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-514">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-514">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-515">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ca3e1-515">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="ca3e1-516">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-516">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="ca3e1-p128">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ca3e1-519">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-519">Type</span></span>

*   [<span data-ttu-id="ca3e1-520">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ca3e1-520">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="ca3e1-521">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-521">Requirements</span></span>

|<span data-ttu-id="ca3e1-522">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-522">Requirement</span></span>| <span data-ttu-id="ca3e1-523">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-524">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-525">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-525">1.0</span></span>|
|[<span data-ttu-id="ca3e1-526">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-527">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-528">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-529">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-529">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca3e1-530">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-530">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="ca3e1-531">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-531">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="ca3e1-532">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-532">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="ca3e1-533">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-533">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ca3e1-534">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-534">Read mode</span></span>

<span data-ttu-id="ca3e1-535">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-535">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="ca3e1-536">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-536">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ca3e1-537">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-537">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="ca3e1-538">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-538">Compose mode</span></span>

<span data-ttu-id="ca3e1-539">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-539">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="ca3e1-540">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-540">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ca3e1-541">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-541">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ca3e1-542">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-542">Get 500 members maximum.</span></span>
- <span data-ttu-id="ca3e1-543">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-543">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="ca3e1-544">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-544">Type</span></span>

*   <span data-ttu-id="ca3e1-545">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-545">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-546">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-546">Requirements</span></span>

|<span data-ttu-id="ca3e1-547">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-547">Requirement</span></span>| <span data-ttu-id="ca3e1-548">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-549">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-550">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-550">1.0</span></span>|
|[<span data-ttu-id="ca3e1-551">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-552">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-552">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-553">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-554">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ca3e1-554">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="ca3e1-555">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-555">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="ca3e1-p132">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="ca3e1-p133">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ca3e1-560">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-560">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="ca3e1-561">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-561">Type</span></span>

*   [<span data-ttu-id="ca3e1-562">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ca3e1-562">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="ca3e1-563">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-563">Requirements</span></span>

|<span data-ttu-id="ca3e1-564">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-564">Requirement</span></span>| <span data-ttu-id="ca3e1-565">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-565">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-566">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-566">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-567">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-567">1.0</span></span>|
|[<span data-ttu-id="ca3e1-568">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-568">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-569">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-569">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-570">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-570">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-571">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-571">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca3e1-572">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-572">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="ca3e1-573">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-573">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="ca3e1-574">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-574">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="ca3e1-p134">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ca3e1-577">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-577">Read mode</span></span>

<span data-ttu-id="ca3e1-578">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-578">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="ca3e1-579">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-579">Compose mode</span></span>

<span data-ttu-id="ca3e1-580">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-580">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="ca3e1-581">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-581">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="ca3e1-582">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-582">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="ca3e1-583">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-583">Type</span></span>

*   <span data-ttu-id="ca3e1-584">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-584">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-585">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-585">Requirements</span></span>

|<span data-ttu-id="ca3e1-586">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-586">Requirement</span></span>| <span data-ttu-id="ca3e1-587">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-587">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-588">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-589">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-589">1.0</span></span>|
|[<span data-ttu-id="ca3e1-590">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-590">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-591">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-591">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-592">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-592">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-593">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ca3e1-593">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-16"></a><span data-ttu-id="ca3e1-594">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-594">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

<span data-ttu-id="ca3e1-595">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-595">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="ca3e1-596">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-596">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ca3e1-597">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-597">Read mode</span></span>

<span data-ttu-id="ca3e1-p135">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="ca3e1-600">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-600">Compose mode</span></span>

<span data-ttu-id="ca3e1-601">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-601">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="ca3e1-602">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-602">Type</span></span>

*   <span data-ttu-id="ca3e1-603">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-603">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-604">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-604">Requirements</span></span>

|<span data-ttu-id="ca3e1-605">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-605">Requirement</span></span>| <span data-ttu-id="ca3e1-606">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-607">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-608">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-608">1.0</span></span>|
|[<span data-ttu-id="ca3e1-609">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-610">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-610">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-611">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-611">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-612">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ca3e1-612">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="ca3e1-613">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-613">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="ca3e1-614">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-614">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="ca3e1-615">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-615">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ca3e1-616">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-616">Read mode</span></span>

<span data-ttu-id="ca3e1-617">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-617">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="ca3e1-618">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-618">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ca3e1-619">ただし、Windows および Mac では、500メンバーの最大値を取得するように設定できます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-619">However, on Windows and Mac, you can set up to get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="ca3e1-620">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-620">Compose mode</span></span>

<span data-ttu-id="ca3e1-621">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-621">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="ca3e1-622">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-622">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ca3e1-623">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-623">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ca3e1-624">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-624">Get 500 members maximum.</span></span>
- <span data-ttu-id="ca3e1-625">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-625">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ca3e1-626">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-626">Type</span></span>

*   <span data-ttu-id="ca3e1-627">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-627">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-628">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-628">Requirements</span></span>

|<span data-ttu-id="ca3e1-629">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-629">Requirement</span></span>| <span data-ttu-id="ca3e1-630">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-630">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-631">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-631">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-632">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-632">1.0</span></span>|
|[<span data-ttu-id="ca3e1-633">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-633">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-634">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-634">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-635">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-635">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-636">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ca3e1-636">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="ca3e1-637">メソッド</span><span class="sxs-lookup"><span data-stu-id="ca3e1-637">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="ca3e1-638">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ca3e1-638">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ca3e1-639">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-639">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="ca3e1-640">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-640">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="ca3e1-641">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-641">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ca3e1-642">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ca3e1-642">Parameters</span></span>

|<span data-ttu-id="ca3e1-643">名前</span><span class="sxs-lookup"><span data-stu-id="ca3e1-643">Name</span></span>| <span data-ttu-id="ca3e1-644">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-644">Type</span></span>| <span data-ttu-id="ca3e1-645">属性</span><span class="sxs-lookup"><span data-stu-id="ca3e1-645">Attributes</span></span>| <span data-ttu-id="ca3e1-646">説明</span><span class="sxs-lookup"><span data-stu-id="ca3e1-646">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="ca3e1-647">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-647">String</span></span>||<span data-ttu-id="ca3e1-p139">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="ca3e1-650">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-650">String</span></span>||<span data-ttu-id="ca3e1-p140">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="ca3e1-653">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ca3e1-653">Object</span></span>| <span data-ttu-id="ca3e1-654">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-654">&lt;optional&gt;</span></span>|<span data-ttu-id="ca3e1-655">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-655">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="ca3e1-656">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ca3e1-656">Object</span></span> | <span data-ttu-id="ca3e1-657">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-657">&lt;optional&gt;</span></span> | <span data-ttu-id="ca3e1-658">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-658">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="ca3e1-659">Boolean</span><span class="sxs-lookup"><span data-stu-id="ca3e1-659">Boolean</span></span> | <span data-ttu-id="ca3e1-660">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-660">&lt;optional&gt;</span></span> | <span data-ttu-id="ca3e1-661">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-661">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="ca3e1-662">function</span><span class="sxs-lookup"><span data-stu-id="ca3e1-662">function</span></span>| <span data-ttu-id="ca3e1-663">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-663">&lt;optional&gt;</span></span>|<span data-ttu-id="ca3e1-664">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-664">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ca3e1-665">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-665">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ca3e1-666">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-666">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ca3e1-667">エラー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-667">Errors</span></span>

| <span data-ttu-id="ca3e1-668">エラー コード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-668">Error code</span></span> | <span data-ttu-id="ca3e1-669">説明</span><span class="sxs-lookup"><span data-stu-id="ca3e1-669">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="ca3e1-670">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-670">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="ca3e1-671">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-671">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="ca3e1-672">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-672">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ca3e1-673">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-673">Requirements</span></span>

|<span data-ttu-id="ca3e1-674">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-674">Requirement</span></span>| <span data-ttu-id="ca3e1-675">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-675">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-676">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-676">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-677">1.1</span><span class="sxs-lookup"><span data-stu-id="ca3e1-677">1.1</span></span>|
|[<span data-ttu-id="ca3e1-678">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-678">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-679">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-679">ReadWriteItem</span></span>|
|[<span data-ttu-id="ca3e1-680">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-680">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-681">作成</span><span class="sxs-lookup"><span data-stu-id="ca3e1-681">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="ca3e1-682">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-682">Examples</span></span>

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

<span data-ttu-id="ca3e1-683">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-683">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="ca3e1-684">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ca3e1-684">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ca3e1-685">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-685">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="ca3e1-p141">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="ca3e1-689">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-689">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="ca3e1-690">Office アドインを Outlook on the web で実行している場合、編集中のアイテム以外のアイテムに `addItemAttachmentAsync` メソッドでアイテムを添付できます。ただし、これはサポートされていないため、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-690">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ca3e1-691">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ca3e1-691">Parameters</span></span>

|<span data-ttu-id="ca3e1-692">名前</span><span class="sxs-lookup"><span data-stu-id="ca3e1-692">Name</span></span>| <span data-ttu-id="ca3e1-693">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-693">Type</span></span>| <span data-ttu-id="ca3e1-694">属性</span><span class="sxs-lookup"><span data-stu-id="ca3e1-694">Attributes</span></span>| <span data-ttu-id="ca3e1-695">説明</span><span class="sxs-lookup"><span data-stu-id="ca3e1-695">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="ca3e1-696">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-696">String</span></span>||<span data-ttu-id="ca3e1-p142">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="ca3e1-699">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-699">String</span></span>||<span data-ttu-id="ca3e1-700">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-700">The subject of the item to be attached.</span></span> <span data-ttu-id="ca3e1-701">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-701">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="ca3e1-702">Object</span><span class="sxs-lookup"><span data-stu-id="ca3e1-702">Object</span></span>| <span data-ttu-id="ca3e1-703">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-703">&lt;optional&gt;</span></span>|<span data-ttu-id="ca3e1-704">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-704">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ca3e1-705">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ca3e1-705">Object</span></span>| <span data-ttu-id="ca3e1-706">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-706">&lt;optional&gt;</span></span>|<span data-ttu-id="ca3e1-707">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-707">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ca3e1-708">function</span><span class="sxs-lookup"><span data-stu-id="ca3e1-708">function</span></span>| <span data-ttu-id="ca3e1-709">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-709">&lt;optional&gt;</span></span>|<span data-ttu-id="ca3e1-710">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-710">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ca3e1-711">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-711">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ca3e1-712">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-712">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ca3e1-713">エラー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-713">Errors</span></span>

| <span data-ttu-id="ca3e1-714">エラー コード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-714">Error code</span></span> | <span data-ttu-id="ca3e1-715">説明</span><span class="sxs-lookup"><span data-stu-id="ca3e1-715">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="ca3e1-716">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-716">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ca3e1-717">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-717">Requirements</span></span>

|<span data-ttu-id="ca3e1-718">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-718">Requirement</span></span>| <span data-ttu-id="ca3e1-719">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-719">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-720">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-720">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-721">1.1</span><span class="sxs-lookup"><span data-stu-id="ca3e1-721">1.1</span></span>|
|[<span data-ttu-id="ca3e1-722">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-722">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-723">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-723">ReadWriteItem</span></span>|
|[<span data-ttu-id="ca3e1-724">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-724">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-725">作成</span><span class="sxs-lookup"><span data-stu-id="ca3e1-725">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ca3e1-726">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-726">Example</span></span>

<span data-ttu-id="ca3e1-727">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-727">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="ca3e1-728">close()</span><span class="sxs-lookup"><span data-stu-id="ca3e1-728">close()</span></span>

<span data-ttu-id="ca3e1-729">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-729">Closes the current item that is being composed.</span></span>

<span data-ttu-id="ca3e1-p144">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="ca3e1-732">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-732">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="ca3e1-733">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-733">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-734">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-734">Requirements</span></span>

|<span data-ttu-id="ca3e1-735">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-735">Requirement</span></span>| <span data-ttu-id="ca3e1-736">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-736">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-737">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-737">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-738">1.3</span><span class="sxs-lookup"><span data-stu-id="ca3e1-738">1.3</span></span>|
|[<span data-ttu-id="ca3e1-739">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-739">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-740">制限あり</span><span class="sxs-lookup"><span data-stu-id="ca3e1-740">Restricted</span></span>|
|[<span data-ttu-id="ca3e1-741">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-741">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-742">新規作成</span><span class="sxs-lookup"><span data-stu-id="ca3e1-742">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="ca3e1-743">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="ca3e1-743">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="ca3e1-744">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-744">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ca3e1-745">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-745">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ca3e1-746">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-746">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ca3e1-747">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-747">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="ca3e1-p145">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ca3e1-751">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ca3e1-751">Parameters</span></span>

| <span data-ttu-id="ca3e1-752">名前</span><span class="sxs-lookup"><span data-stu-id="ca3e1-752">Name</span></span> | <span data-ttu-id="ca3e1-753">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-753">Type</span></span> | <span data-ttu-id="ca3e1-754">属性</span><span class="sxs-lookup"><span data-stu-id="ca3e1-754">Attributes</span></span> | <span data-ttu-id="ca3e1-755">説明</span><span class="sxs-lookup"><span data-stu-id="ca3e1-755">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="ca3e1-756">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="ca3e1-756">String &#124; Object</span></span>| |<span data-ttu-id="ca3e1-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ca3e1-759">**または**</span><span class="sxs-lookup"><span data-stu-id="ca3e1-759">**OR**</span></span><br/><span data-ttu-id="ca3e1-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="ca3e1-762">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-762">String</span></span> | <span data-ttu-id="ca3e1-763">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-763">&lt;optional&gt;</span></span> | <span data-ttu-id="ca3e1-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="ca3e1-766">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-766">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="ca3e1-767">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-767">&lt;optional&gt;</span></span> | <span data-ttu-id="ca3e1-768">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-768">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="ca3e1-769">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-769">String</span></span> | | <span data-ttu-id="ca3e1-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="ca3e1-772">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-772">String</span></span> | | <span data-ttu-id="ca3e1-773">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-773">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="ca3e1-774">文字列</span><span class="sxs-lookup"><span data-stu-id="ca3e1-774">String</span></span> | | <span data-ttu-id="ca3e1-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="ca3e1-777">ブール値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-777">Boolean</span></span> | | <span data-ttu-id="ca3e1-p151">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="ca3e1-780">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-780">String</span></span> | | <span data-ttu-id="ca3e1-p152">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="ca3e1-784">function</span><span class="sxs-lookup"><span data-stu-id="ca3e1-784">function</span></span> | <span data-ttu-id="ca3e1-785">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-785">&lt;optional&gt;</span></span> | <span data-ttu-id="ca3e1-786">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-786">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ca3e1-787">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-787">Requirements</span></span>

|<span data-ttu-id="ca3e1-788">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-788">Requirement</span></span>| <span data-ttu-id="ca3e1-789">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-789">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-790">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-790">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-791">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-791">1.0</span></span>|
|[<span data-ttu-id="ca3e1-792">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-792">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-793">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-793">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-794">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-794">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-795">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-795">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ca3e1-796">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-796">Examples</span></span>

<span data-ttu-id="ca3e1-797">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-797">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="ca3e1-798">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-798">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="ca3e1-799">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-799">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ca3e1-800">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-800">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="ca3e1-801">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-801">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="ca3e1-802">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-802">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="ca3e1-803">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="ca3e1-803">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="ca3e1-804">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-804">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ca3e1-805">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-805">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ca3e1-806">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-806">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ca3e1-807">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-807">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="ca3e1-p153">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ca3e1-811">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ca3e1-811">Parameters</span></span>

| <span data-ttu-id="ca3e1-812">名前</span><span class="sxs-lookup"><span data-stu-id="ca3e1-812">Name</span></span> | <span data-ttu-id="ca3e1-813">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-813">Type</span></span> | <span data-ttu-id="ca3e1-814">属性</span><span class="sxs-lookup"><span data-stu-id="ca3e1-814">Attributes</span></span> | <span data-ttu-id="ca3e1-815">説明</span><span class="sxs-lookup"><span data-stu-id="ca3e1-815">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="ca3e1-816">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="ca3e1-816">String &#124; Object</span></span>| | <span data-ttu-id="ca3e1-p154">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ca3e1-819">**または**</span><span class="sxs-lookup"><span data-stu-id="ca3e1-819">**OR**</span></span><br/><span data-ttu-id="ca3e1-p155">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="ca3e1-822">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-822">String</span></span> | <span data-ttu-id="ca3e1-823">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-823">&lt;optional&gt;</span></span> | <span data-ttu-id="ca3e1-p156">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="ca3e1-826">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-826">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="ca3e1-827">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-827">&lt;optional&gt;</span></span> | <span data-ttu-id="ca3e1-828">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-828">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="ca3e1-829">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-829">String</span></span> | | <span data-ttu-id="ca3e1-p157">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="ca3e1-832">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-832">String</span></span> | | <span data-ttu-id="ca3e1-833">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-833">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="ca3e1-834">文字列</span><span class="sxs-lookup"><span data-stu-id="ca3e1-834">String</span></span> | | <span data-ttu-id="ca3e1-p158">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="ca3e1-837">ブール値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-837">Boolean</span></span> | | <span data-ttu-id="ca3e1-p159">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="ca3e1-840">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-840">String</span></span> | | <span data-ttu-id="ca3e1-p160">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="ca3e1-844">function</span><span class="sxs-lookup"><span data-stu-id="ca3e1-844">function</span></span> | <span data-ttu-id="ca3e1-845">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-845">&lt;optional&gt;</span></span> | <span data-ttu-id="ca3e1-846">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-846">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ca3e1-847">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-847">Requirements</span></span>

|<span data-ttu-id="ca3e1-848">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-848">Requirement</span></span>| <span data-ttu-id="ca3e1-849">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-849">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-850">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-850">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-851">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-851">1.0</span></span>|
|[<span data-ttu-id="ca3e1-852">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-852">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-853">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-853">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-854">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-854">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-855">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-855">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ca3e1-856">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-856">Examples</span></span>

<span data-ttu-id="ca3e1-857">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-857">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="ca3e1-858">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-858">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="ca3e1-859">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-859">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ca3e1-860">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-860">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="ca3e1-861">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-861">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="ca3e1-862">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-862">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="ca3e1-863">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="ca3e1-863">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="ca3e1-864">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-864">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="ca3e1-865">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-865">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-866">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-866">Requirements</span></span>

|<span data-ttu-id="ca3e1-867">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-867">Requirement</span></span>| <span data-ttu-id="ca3e1-868">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-868">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-869">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-869">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-870">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-870">1.0</span></span>|
|[<span data-ttu-id="ca3e1-871">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-871">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-872">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-872">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-873">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-873">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-874">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-874">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ca3e1-875">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ca3e1-875">Returns:</span></span>

<span data-ttu-id="ca3e1-876">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-876">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="ca3e1-877">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-877">Example</span></span>

<span data-ttu-id="ca3e1-878">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-878">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="ca3e1-879">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="ca3e1-879">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="ca3e1-880">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-880">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="ca3e1-881">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-881">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ca3e1-882">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ca3e1-882">Parameters</span></span>

|<span data-ttu-id="ca3e1-883">名前</span><span class="sxs-lookup"><span data-stu-id="ca3e1-883">Name</span></span>| <span data-ttu-id="ca3e1-884">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-884">Type</span></span>| <span data-ttu-id="ca3e1-885">説明</span><span class="sxs-lookup"><span data-stu-id="ca3e1-885">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="ca3e1-886">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="ca3e1-886">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.6)|<span data-ttu-id="ca3e1-887">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-887">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ca3e1-888">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-888">Requirements</span></span>

|<span data-ttu-id="ca3e1-889">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-889">Requirement</span></span>| <span data-ttu-id="ca3e1-890">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-890">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-891">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-891">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-892">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-892">1.0</span></span>|
|[<span data-ttu-id="ca3e1-893">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-893">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-894">制限あり</span><span class="sxs-lookup"><span data-stu-id="ca3e1-894">Restricted</span></span>|
|[<span data-ttu-id="ca3e1-895">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-895">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-896">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-896">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ca3e1-897">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ca3e1-897">Returns:</span></span>

<span data-ttu-id="ca3e1-898">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-898">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="ca3e1-899">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-899">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="ca3e1-900">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-900">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="ca3e1-901">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-901">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="ca3e1-902">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-902">Value of `entityType`</span></span> | <span data-ttu-id="ca3e1-903">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-903">Type of objects in returned array</span></span> | <span data-ttu-id="ca3e1-904">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-904">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="ca3e1-905">文字列</span><span class="sxs-lookup"><span data-stu-id="ca3e1-905">String</span></span> | <span data-ttu-id="ca3e1-906">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="ca3e1-906">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="ca3e1-907">連絡先</span><span class="sxs-lookup"><span data-stu-id="ca3e1-907">Contact</span></span> | <span data-ttu-id="ca3e1-908">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ca3e1-908">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="ca3e1-909">文字列</span><span class="sxs-lookup"><span data-stu-id="ca3e1-909">String</span></span> | <span data-ttu-id="ca3e1-910">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ca3e1-910">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="ca3e1-911">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="ca3e1-911">MeetingSuggestion</span></span> | <span data-ttu-id="ca3e1-912">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ca3e1-912">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="ca3e1-913">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="ca3e1-913">PhoneNumber</span></span> | <span data-ttu-id="ca3e1-914">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="ca3e1-914">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="ca3e1-915">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="ca3e1-915">TaskSuggestion</span></span> | <span data-ttu-id="ca3e1-916">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ca3e1-916">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="ca3e1-917">文字列</span><span class="sxs-lookup"><span data-stu-id="ca3e1-917">String</span></span> | <span data-ttu-id="ca3e1-918">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="ca3e1-918">**Restricted**</span></span> |

<span data-ttu-id="ca3e1-919">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="ca3e1-919">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

##### <a name="example"></a><span data-ttu-id="ca3e1-920">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-920">Example</span></span>

<span data-ttu-id="ca3e1-921">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-921">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="ca3e1-922">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="ca3e1-922">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="ca3e1-923">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-923">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ca3e1-924">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-924">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ca3e1-925">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-925">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ca3e1-926">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ca3e1-926">Parameters</span></span>

|<span data-ttu-id="ca3e1-927">名前</span><span class="sxs-lookup"><span data-stu-id="ca3e1-927">Name</span></span>| <span data-ttu-id="ca3e1-928">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-928">Type</span></span>| <span data-ttu-id="ca3e1-929">説明</span><span class="sxs-lookup"><span data-stu-id="ca3e1-929">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="ca3e1-930">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-930">String</span></span>|<span data-ttu-id="ca3e1-931">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-931">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ca3e1-932">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-932">Requirements</span></span>

|<span data-ttu-id="ca3e1-933">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-933">Requirement</span></span>| <span data-ttu-id="ca3e1-934">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-935">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-936">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-936">1.0</span></span>|
|[<span data-ttu-id="ca3e1-937">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-937">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-938">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-939">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-939">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-940">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-940">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ca3e1-941">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ca3e1-941">Returns:</span></span>

<span data-ttu-id="ca3e1-p162">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="ca3e1-944">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="ca3e1-944">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="ca3e1-945">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="ca3e1-945">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="ca3e1-946">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-946">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ca3e1-947">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-947">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ca3e1-p163">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="ca3e1-951">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-951">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="ca3e1-952">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-952">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="ca3e1-p164">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-956">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-956">Requirements</span></span>

|<span data-ttu-id="ca3e1-957">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-957">Requirement</span></span>| <span data-ttu-id="ca3e1-958">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-958">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-959">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-959">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-960">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-960">1.0</span></span>|
|[<span data-ttu-id="ca3e1-961">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-961">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-962">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-962">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-963">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-963">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-964">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-964">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ca3e1-965">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ca3e1-965">Returns:</span></span>

<span data-ttu-id="ca3e1-p165">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="ca3e1-968">型: Object</span><span class="sxs-lookup"><span data-stu-id="ca3e1-968">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="ca3e1-969">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-969">Example</span></span>

<span data-ttu-id="ca3e1-970">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-970">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="ca3e1-971">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="ca3e1-971">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="ca3e1-972">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-972">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ca3e1-973">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-973">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ca3e1-974">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-974">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="ca3e1-p166">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ca3e1-977">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ca3e1-977">Parameters</span></span>

|<span data-ttu-id="ca3e1-978">名前</span><span class="sxs-lookup"><span data-stu-id="ca3e1-978">Name</span></span>| <span data-ttu-id="ca3e1-979">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-979">Type</span></span>| <span data-ttu-id="ca3e1-980">説明</span><span class="sxs-lookup"><span data-stu-id="ca3e1-980">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="ca3e1-981">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-981">String</span></span>|<span data-ttu-id="ca3e1-982">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-982">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ca3e1-983">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-983">Requirements</span></span>

|<span data-ttu-id="ca3e1-984">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-984">Requirement</span></span>| <span data-ttu-id="ca3e1-985">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-986">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-987">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-987">1.0</span></span>|
|[<span data-ttu-id="ca3e1-988">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-988">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-989">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-989">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-990">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-990">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-991">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-991">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ca3e1-992">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ca3e1-992">Returns:</span></span>

<span data-ttu-id="ca3e1-993">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-993">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="ca3e1-994">型: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="ca3e1-994">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="ca3e1-995">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-995">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="ca3e1-996">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="ca3e1-996">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="ca3e1-997">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-997">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="ca3e1-p167">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p167">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="ca3e1-1000">Web 上の Outlook では、テキストが選択されておらず、カーソルが本文にある場合、このメソッドは文字列 "null" を返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1000">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="ca3e1-1001">このような状況を確認するには、次のようなコードを含めます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1001">To check for this situation, include code similar to the following:</span></span>
>
> `var selectedText = (asyncResult.value.endPosition === asyncResult.value.startPosition) ? "" : asyncResult.value.data;`

##### <a name="parameters"></a><span data-ttu-id="ca3e1-1002">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1002">Parameters</span></span>

|<span data-ttu-id="ca3e1-1003">名前</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1003">Name</span></span>| <span data-ttu-id="ca3e1-1004">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1004">Type</span></span>| <span data-ttu-id="ca3e1-1005">属性</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1005">Attributes</span></span>| <span data-ttu-id="ca3e1-1006">説明</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1006">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="ca3e1-1007">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1007">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="ca3e1-p169">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p169">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="ca3e1-1011">Object</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1011">Object</span></span>| <span data-ttu-id="ca3e1-1012">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1012">&lt;optional&gt;</span></span>|<span data-ttu-id="ca3e1-1013">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1013">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ca3e1-1014">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1014">Object</span></span>| <span data-ttu-id="ca3e1-1015">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1015">&lt;optional&gt;</span></span>|<span data-ttu-id="ca3e1-1016">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1016">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ca3e1-1017">function</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1017">function</span></span>||<span data-ttu-id="ca3e1-1018">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1018">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ca3e1-1019">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1019">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="ca3e1-1020">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1020">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ca3e1-1021">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1021">Requirements</span></span>

|<span data-ttu-id="ca3e1-1022">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1022">Requirement</span></span>| <span data-ttu-id="ca3e1-1023">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1023">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-1024">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1024">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-1025">1.2</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1025">1.2</span></span>|
|[<span data-ttu-id="ca3e1-1026">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1026">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-1027">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1027">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-1028">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1028">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-1029">作成</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1029">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="ca3e1-1030">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1030">Returns:</span></span>

<span data-ttu-id="ca3e1-1031">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1031">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="ca3e1-1032">型:String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1032">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="ca3e1-1033">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1033">Example</span></span>

```js
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

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="ca3e1-1034">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1034">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="ca3e1-1035">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1035">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="ca3e1-1036">強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1036">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="ca3e1-1037">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1037">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-1038">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1038">Requirements</span></span>

|<span data-ttu-id="ca3e1-1039">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1039">Requirement</span></span>| <span data-ttu-id="ca3e1-1040">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-1041">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1042">1.6</span></span> |
|[<span data-ttu-id="ca3e1-1043">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1043">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1044">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-1045">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1045">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-1046">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ca3e1-1047">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1047">Returns:</span></span>

<span data-ttu-id="ca3e1-1048">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1048">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="ca3e1-1049">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1049">Example</span></span>

<span data-ttu-id="ca3e1-1050">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1050">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="ca3e1-1051">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1051">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="ca3e1-p172">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p172">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="ca3e1-1054">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1054">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ca3e1-p173">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p173">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="ca3e1-1058">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1058">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="ca3e1-1059">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1059">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="ca3e1-p174">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p174">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca3e1-1063">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1063">Requirements</span></span>

|<span data-ttu-id="ca3e1-1064">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1064">Requirement</span></span>| <span data-ttu-id="ca3e1-1065">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1065">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-1066">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1066">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-1067">1.6</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1067">1.6</span></span> |
|[<span data-ttu-id="ca3e1-1068">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1068">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-1069">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1069">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-1070">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1070">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-1071">読み取り</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1071">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ca3e1-1072">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1072">Returns:</span></span>

<span data-ttu-id="ca3e1-p175">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p175">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="ca3e1-1075">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1075">Example</span></span>

<span data-ttu-id="ca3e1-1076">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1076">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="ca3e1-1077">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1077">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="ca3e1-1078">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1078">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="ca3e1-p176">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p176">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ca3e1-1082">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1082">Parameters</span></span>

|<span data-ttu-id="ca3e1-1083">名前</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1083">Name</span></span>| <span data-ttu-id="ca3e1-1084">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1084">Type</span></span>| <span data-ttu-id="ca3e1-1085">属性</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1085">Attributes</span></span>| <span data-ttu-id="ca3e1-1086">説明</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1086">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="ca3e1-1087">function</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1087">function</span></span>||<span data-ttu-id="ca3e1-1088">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1088">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ca3e1-1089">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1089">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="ca3e1-1090">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1090">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="ca3e1-1091">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1091">Object</span></span>| <span data-ttu-id="ca3e1-1092">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1092">&lt;optional&gt;</span></span>|<span data-ttu-id="ca3e1-1093">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1093">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="ca3e1-1094">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1094">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ca3e1-1095">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1095">Requirements</span></span>

|<span data-ttu-id="ca3e1-1096">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1096">Requirement</span></span>| <span data-ttu-id="ca3e1-1097">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1097">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-1098">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1098">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-1099">1.0</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1099">1.0</span></span>|
|[<span data-ttu-id="ca3e1-1100">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1100">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-1101">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1101">ReadItem</span></span>|
|[<span data-ttu-id="ca3e1-1102">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1102">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-1103">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1103">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca3e1-1104">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1104">Example</span></span>

<span data-ttu-id="ca3e1-p179">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p179">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="ca3e1-1108">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1108">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="ca3e1-1109">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1109">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="ca3e1-1110">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1110">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="ca3e1-1111">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1111">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="ca3e1-1112">Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1112">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="ca3e1-1113">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1113">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ca3e1-1114">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1114">Parameters</span></span>

|<span data-ttu-id="ca3e1-1115">名前</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1115">Name</span></span>| <span data-ttu-id="ca3e1-1116">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1116">Type</span></span>| <span data-ttu-id="ca3e1-1117">属性</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1117">Attributes</span></span>| <span data-ttu-id="ca3e1-1118">説明</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1118">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="ca3e1-1119">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1119">String</span></span>||<span data-ttu-id="ca3e1-1120">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1120">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="ca3e1-1121">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1121">Object</span></span>| <span data-ttu-id="ca3e1-1122">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1122">&lt;optional&gt;</span></span>|<span data-ttu-id="ca3e1-1123">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1123">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ca3e1-1124">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1124">Object</span></span>| <span data-ttu-id="ca3e1-1125">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1125">&lt;optional&gt;</span></span>|<span data-ttu-id="ca3e1-1126">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1126">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ca3e1-1127">関数</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1127">function</span></span>| <span data-ttu-id="ca3e1-1128">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1128">&lt;optional&gt;</span></span>|<span data-ttu-id="ca3e1-1129">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1129">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ca3e1-1130">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1130">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ca3e1-1131">エラー</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1131">Errors</span></span>

| <span data-ttu-id="ca3e1-1132">エラー コード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1132">Error code</span></span> | <span data-ttu-id="ca3e1-1133">説明</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1133">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="ca3e1-1134">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1134">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ca3e1-1135">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1135">Requirements</span></span>

|<span data-ttu-id="ca3e1-1136">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1136">Requirement</span></span>| <span data-ttu-id="ca3e1-1137">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-1138">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-1139">1.1</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1139">1.1</span></span>|
|[<span data-ttu-id="ca3e1-1140">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-1141">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1141">ReadWriteItem</span></span>|
|[<span data-ttu-id="ca3e1-1142">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-1143">作成</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1143">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ca3e1-1144">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1144">Example</span></span>

<span data-ttu-id="ca3e1-1145">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1145">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="ca3e1-1146">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1146">saveAsync([options], callback)</span></span>

<span data-ttu-id="ca3e1-1147">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1147">Asynchronously saves an item.</span></span>

<span data-ttu-id="ca3e1-1148">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1148">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="ca3e1-1149">Outlook on the web またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1149">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="ca3e1-1150">キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1150">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="ca3e1-1151">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1151">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="ca3e1-1152">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1152">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="ca3e1-p183">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="ca3e1-1156">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1156">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="ca3e1-1157">Outlook on Mac では、会議の保存はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1157">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="ca3e1-1158">`saveAsync` メソッドは、作成モードの会議から呼び出されると失敗します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1158">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="ca3e1-1159">回避策については、「[Office JS API を使用して Outlook for Mac で会議を下書きとして保存できない](https://support.microsoft.com/help/4505745)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1159">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="ca3e1-1160">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1160">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ca3e1-1161">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1161">Parameters</span></span>

|<span data-ttu-id="ca3e1-1162">名前</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1162">Name</span></span>| <span data-ttu-id="ca3e1-1163">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1163">Type</span></span>| <span data-ttu-id="ca3e1-1164">属性</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1164">Attributes</span></span>| <span data-ttu-id="ca3e1-1165">説明</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1165">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="ca3e1-1166">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1166">Object</span></span>| <span data-ttu-id="ca3e1-1167">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1167">&lt;optional&gt;</span></span>|<span data-ttu-id="ca3e1-1168">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1168">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ca3e1-1169">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1169">Object</span></span>| <span data-ttu-id="ca3e1-1170">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1170">&lt;optional&gt;</span></span>|<span data-ttu-id="ca3e1-1171">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1171">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ca3e1-1172">関数</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1172">function</span></span>||<span data-ttu-id="ca3e1-1173">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1173">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ca3e1-1174">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1174">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ca3e1-1175">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1175">Requirements</span></span>

|<span data-ttu-id="ca3e1-1176">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1176">Requirement</span></span>| <span data-ttu-id="ca3e1-1177">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1177">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-1178">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-1179">1.3</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1179">1.3</span></span>|
|[<span data-ttu-id="ca3e1-1180">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1180">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-1181">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1181">ReadWriteItem</span></span>|
|[<span data-ttu-id="ca3e1-1182">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1182">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-1183">作成</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1183">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="ca3e1-1184">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1184">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="ca3e1-p185">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="ca3e1-1187">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1187">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="ca3e1-1188">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1188">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="ca3e1-p186">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ca3e1-1192">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1192">Parameters</span></span>

|<span data-ttu-id="ca3e1-1193">名前</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1193">Name</span></span>| <span data-ttu-id="ca3e1-1194">型</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1194">Type</span></span>| <span data-ttu-id="ca3e1-1195">属性</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1195">Attributes</span></span>| <span data-ttu-id="ca3e1-1196">説明</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1196">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="ca3e1-1197">String</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1197">String</span></span>||<span data-ttu-id="ca3e1-p187">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="ca3e1-1201">Object</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1201">Object</span></span>| <span data-ttu-id="ca3e1-1202">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1202">&lt;optional&gt;</span></span>|<span data-ttu-id="ca3e1-1203">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1203">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ca3e1-1204">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1204">Object</span></span>| <span data-ttu-id="ca3e1-1205">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1205">&lt;optional&gt;</span></span>|<span data-ttu-id="ca3e1-1206">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1206">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="ca3e1-1207">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1207">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="ca3e1-1208">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1208">&lt;optional&gt;</span></span>|<span data-ttu-id="ca3e1-1209">`text` の場合、Outlook on the web とデスクトップ クライアントでは現在のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1209">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="ca3e1-1210">フィールドが HTML エディターの場合、データが HTML の場合でもテキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1210">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="ca3e1-1211">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Outlook on the web では現在のスタイルが適用され、Outlook デスクトップ クライアントでは既定のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1211">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="ca3e1-1212">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1212">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="ca3e1-1213">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1213">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="ca3e1-1214">function</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1214">function</span></span>||<span data-ttu-id="ca3e1-1215">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ca3e1-1216">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1216">Requirements</span></span>

|<span data-ttu-id="ca3e1-1217">要件</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1217">Requirement</span></span>| <span data-ttu-id="ca3e1-1218">値</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1218">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca3e1-1219">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ca3e1-1220">1.2</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1220">1.2</span></span>|
|[<span data-ttu-id="ca3e1-1221">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ca3e1-1222">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1222">ReadWriteItem</span></span>|
|[<span data-ttu-id="ca3e1-1223">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ca3e1-1224">作成</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1224">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ca3e1-1225">例</span><span class="sxs-lookup"><span data-stu-id="ca3e1-1225">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
