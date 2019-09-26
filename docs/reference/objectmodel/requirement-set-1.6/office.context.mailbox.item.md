---
title: Office. メールボックス-要件セット1.6
description: ''
ms.date: 09/23/2019
localization_priority: Normal
ms.openlocfilehash: 980135223414b58bb048dce54a9fe1446a26086c
ms.sourcegitcommit: 3c84fe6302341668c3f9f6dd64e636a97d03023c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/26/2019
ms.locfileid: "37167362"
---
# <a name="item"></a><span data-ttu-id="cf9f6-102">item</span><span class="sxs-lookup"><span data-stu-id="cf9f6-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="cf9f6-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="cf9f6-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="cf9f6-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-106">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-106">Requirements</span></span>

|<span data-ttu-id="cf9f6-107">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-107">Requirement</span></span>| <span data-ttu-id="cf9f6-108">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-110">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-110">1.0</span></span>|
|[<span data-ttu-id="cf9f6-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="cf9f6-112">Restricted</span></span>|
|[<span data-ttu-id="cf9f6-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cf9f6-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="cf9f6-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-115">Members and methods</span></span>

| <span data-ttu-id="cf9f6-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-116">Member</span></span> | <span data-ttu-id="cf9f6-117">種類</span><span class="sxs-lookup"><span data-stu-id="cf9f6-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="cf9f6-118">attachments</span><span class="sxs-lookup"><span data-stu-id="cf9f6-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="cf9f6-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-119">Member</span></span> |
| [<span data-ttu-id="cf9f6-120">bcc</span><span class="sxs-lookup"><span data-stu-id="cf9f6-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="cf9f6-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-121">Member</span></span> |
| [<span data-ttu-id="cf9f6-122">body</span><span class="sxs-lookup"><span data-stu-id="cf9f6-122">body</span></span>](#body-body) | <span data-ttu-id="cf9f6-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-123">Member</span></span> |
| [<span data-ttu-id="cf9f6-124">cc</span><span class="sxs-lookup"><span data-stu-id="cf9f6-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="cf9f6-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-125">Member</span></span> |
| [<span data-ttu-id="cf9f6-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="cf9f6-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="cf9f6-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-127">Member</span></span> |
| [<span data-ttu-id="cf9f6-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="cf9f6-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="cf9f6-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-129">Member</span></span> |
| [<span data-ttu-id="cf9f6-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="cf9f6-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="cf9f6-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-131">Member</span></span> |
| [<span data-ttu-id="cf9f6-132">end</span><span class="sxs-lookup"><span data-stu-id="cf9f6-132">end</span></span>](#end-datetime) | <span data-ttu-id="cf9f6-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-133">Member</span></span> |
| [<span data-ttu-id="cf9f6-134">from</span><span class="sxs-lookup"><span data-stu-id="cf9f6-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="cf9f6-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-135">Member</span></span> |
| [<span data-ttu-id="cf9f6-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="cf9f6-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="cf9f6-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-137">Member</span></span> |
| [<span data-ttu-id="cf9f6-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="cf9f6-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="cf9f6-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-139">Member</span></span> |
| [<span data-ttu-id="cf9f6-140">itemId</span><span class="sxs-lookup"><span data-stu-id="cf9f6-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="cf9f6-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-141">Member</span></span> |
| [<span data-ttu-id="cf9f6-142">itemType</span><span class="sxs-lookup"><span data-stu-id="cf9f6-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="cf9f6-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-143">Member</span></span> |
| [<span data-ttu-id="cf9f6-144">location</span><span class="sxs-lookup"><span data-stu-id="cf9f6-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="cf9f6-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-145">Member</span></span> |
| [<span data-ttu-id="cf9f6-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="cf9f6-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="cf9f6-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-147">Member</span></span> |
| [<span data-ttu-id="cf9f6-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="cf9f6-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="cf9f6-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-149">Member</span></span> |
| [<span data-ttu-id="cf9f6-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="cf9f6-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="cf9f6-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-151">Member</span></span> |
| [<span data-ttu-id="cf9f6-152">organizer</span><span class="sxs-lookup"><span data-stu-id="cf9f6-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="cf9f6-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-153">Member</span></span> |
| [<span data-ttu-id="cf9f6-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="cf9f6-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="cf9f6-155">Member</span><span class="sxs-lookup"><span data-stu-id="cf9f6-155">Member</span></span> |
| [<span data-ttu-id="cf9f6-156">sender</span><span class="sxs-lookup"><span data-stu-id="cf9f6-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="cf9f6-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-157">Member</span></span> |
| [<span data-ttu-id="cf9f6-158">start</span><span class="sxs-lookup"><span data-stu-id="cf9f6-158">start</span></span>](#start-datetime) | <span data-ttu-id="cf9f6-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-159">Member</span></span> |
| [<span data-ttu-id="cf9f6-160">subject</span><span class="sxs-lookup"><span data-stu-id="cf9f6-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="cf9f6-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-161">Member</span></span> |
| [<span data-ttu-id="cf9f6-162">to</span><span class="sxs-lookup"><span data-stu-id="cf9f6-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="cf9f6-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-163">Member</span></span> |
| [<span data-ttu-id="cf9f6-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="cf9f6-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="cf9f6-165">メソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-165">Method</span></span> |
| [<span data-ttu-id="cf9f6-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="cf9f6-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="cf9f6-167">メソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-167">Method</span></span> |
| [<span data-ttu-id="cf9f6-168">close</span><span class="sxs-lookup"><span data-stu-id="cf9f6-168">close</span></span>](#close) | <span data-ttu-id="cf9f6-169">メソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-169">Method</span></span> |
| [<span data-ttu-id="cf9f6-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="cf9f6-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="cf9f6-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-171">Method</span></span> |
| [<span data-ttu-id="cf9f6-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="cf9f6-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="cf9f6-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-173">Method</span></span> |
| [<span data-ttu-id="cf9f6-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="cf9f6-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="cf9f6-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-175">Method</span></span> |
| [<span data-ttu-id="cf9f6-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="cf9f6-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="cf9f6-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-177">Method</span></span> |
| [<span data-ttu-id="cf9f6-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="cf9f6-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="cf9f6-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-179">Method</span></span> |
| [<span data-ttu-id="cf9f6-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="cf9f6-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="cf9f6-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-181">Method</span></span> |
| [<span data-ttu-id="cf9f6-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="cf9f6-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="cf9f6-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-183">Method</span></span> |
| [<span data-ttu-id="cf9f6-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="cf9f6-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="cf9f6-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-185">Method</span></span> |
| [<span data-ttu-id="cf9f6-186">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="cf9f6-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="cf9f6-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-187">Method</span></span> |
| [<span data-ttu-id="cf9f6-188">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="cf9f6-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="cf9f6-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-189">Method</span></span> |
| [<span data-ttu-id="cf9f6-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="cf9f6-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="cf9f6-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-191">Method</span></span> |
| [<span data-ttu-id="cf9f6-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="cf9f6-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="cf9f6-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-193">Method</span></span> |
| [<span data-ttu-id="cf9f6-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="cf9f6-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="cf9f6-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-195">Method</span></span> |
| [<span data-ttu-id="cf9f6-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="cf9f6-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="cf9f6-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="cf9f6-198">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-198">Example</span></span>

<span data-ttu-id="cf9f6-199">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="cf9f6-200">メンバー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-16"></a><span data-ttu-id="cf9f6-201">添付ファイル: <[Attachmentdetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="cf9f6-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

<span data-ttu-id="cf9f6-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="cf9f6-204">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="cf9f6-205">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="cf9f6-206">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-206">Type</span></span>

*   <span data-ttu-id="cf9f6-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="cf9f6-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-208">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-208">Requirements</span></span>

|<span data-ttu-id="cf9f6-209">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-209">Requirement</span></span>| <span data-ttu-id="cf9f6-210">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-211">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-212">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-212">1.0</span></span>|
|[<span data-ttu-id="cf9f6-213">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-214">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-215">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-216">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf9f6-217">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-217">Example</span></span>

<span data-ttu-id="cf9f6-218">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="cf9f6-219">bcc:[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="cf9f6-220">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="cf9f6-221">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="cf9f6-222">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-222">Type</span></span>

*   [<span data-ttu-id="cf9f6-223">受信者</span><span class="sxs-lookup"><span data-stu-id="cf9f6-223">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="cf9f6-224">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-224">Requirements</span></span>

|<span data-ttu-id="cf9f6-225">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-225">Requirement</span></span>| <span data-ttu-id="cf9f6-226">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-227">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-228">1.1</span><span class="sxs-lookup"><span data-stu-id="cf9f6-228">1.1</span></span>|
|[<span data-ttu-id="cf9f6-229">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-230">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-231">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-232">作成</span><span class="sxs-lookup"><span data-stu-id="cf9f6-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="cf9f6-233">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-233">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-16"></a><span data-ttu-id="cf9f6-234">本文:[本文](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span></span>

<span data-ttu-id="cf9f6-235">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="cf9f6-236">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-236">Type</span></span>

*   [<span data-ttu-id="cf9f6-237">Body</span><span class="sxs-lookup"><span data-stu-id="cf9f6-237">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="cf9f6-238">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-238">Requirements</span></span>

|<span data-ttu-id="cf9f6-239">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-239">Requirement</span></span>| <span data-ttu-id="cf9f6-240">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-241">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-242">1.1</span><span class="sxs-lookup"><span data-stu-id="cf9f6-242">1.1</span></span>|
|[<span data-ttu-id="cf9f6-243">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-244">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-246">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cf9f6-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf9f6-247">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-247">Example</span></span>

<span data-ttu-id="cf9f6-248">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-248">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="cf9f6-249">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-249">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="cf9f6-250">cc: <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="cf9f6-251">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="cf9f6-252">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cf9f6-253">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-253">Read mode</span></span>

<span data-ttu-id="cf9f6-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="cf9f6-256">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-256">Compose mode</span></span>

<span data-ttu-id="cf9f6-257">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-257">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="cf9f6-258">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-258">Type</span></span>

*   <span data-ttu-id="cf9f6-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-260">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-260">Requirements</span></span>

|<span data-ttu-id="cf9f6-261">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-261">Requirement</span></span>| <span data-ttu-id="cf9f6-262">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-263">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-264">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-264">1.0</span></span>|
|[<span data-ttu-id="cf9f6-265">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-266">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-267">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-268">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cf9f6-268">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="cf9f6-269">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-269">(nullable) conversationId: String</span></span>

<span data-ttu-id="cf9f6-270">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="cf9f6-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="cf9f6-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="cf9f6-275">Type</span><span class="sxs-lookup"><span data-stu-id="cf9f6-275">Type</span></span>

*   <span data-ttu-id="cf9f6-276">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-277">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-277">Requirements</span></span>

|<span data-ttu-id="cf9f6-278">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-278">Requirement</span></span>| <span data-ttu-id="cf9f6-279">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-280">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-281">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-281">1.0</span></span>|
|[<span data-ttu-id="cf9f6-282">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-283">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-284">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-285">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cf9f6-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf9f6-286">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-286">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="cf9f6-287">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="cf9f6-287">dateTimeCreated: Date</span></span>

<span data-ttu-id="cf9f6-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="cf9f6-290">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-290">Type</span></span>

*   <span data-ttu-id="cf9f6-291">日付</span><span class="sxs-lookup"><span data-stu-id="cf9f6-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-292">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-292">Requirements</span></span>

|<span data-ttu-id="cf9f6-293">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-293">Requirement</span></span>| <span data-ttu-id="cf9f6-294">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-295">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-296">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-296">1.0</span></span>|
|[<span data-ttu-id="cf9f6-297">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-298">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-299">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-300">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf9f6-301">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-301">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="cf9f6-302">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="cf9f6-302">dateTimeModified: Date</span></span>

<span data-ttu-id="cf9f6-303">アイテムが最後に変更された日時を取得します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-303">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="cf9f6-304">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-304">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="cf9f6-305">このメンバーは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-305">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="cf9f6-306">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-306">Type</span></span>

*   <span data-ttu-id="cf9f6-307">日付</span><span class="sxs-lookup"><span data-stu-id="cf9f6-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-308">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-308">Requirements</span></span>

|<span data-ttu-id="cf9f6-309">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-309">Requirement</span></span>| <span data-ttu-id="cf9f6-310">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-312">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-312">1.0</span></span>|
|[<span data-ttu-id="cf9f6-313">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-314">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-316">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf9f6-317">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-317">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="cf9f6-318">終了: 日付 |[時間](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="cf9f6-319">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="cf9f6-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cf9f6-322">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-322">Read mode</span></span>

<span data-ttu-id="cf9f6-323">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-323">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="cf9f6-324">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-324">Compose mode</span></span>

<span data-ttu-id="cf9f6-325">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="cf9f6-326">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-326">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="cf9f6-327">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="cf9f6-328">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-328">Type</span></span>

*   <span data-ttu-id="cf9f6-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-330">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-330">Requirements</span></span>

|<span data-ttu-id="cf9f6-331">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-331">Requirement</span></span>| <span data-ttu-id="cf9f6-332">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-333">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-334">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-334">1.0</span></span>|
|[<span data-ttu-id="cf9f6-335">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-336">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-337">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-338">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cf9f6-338">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="cf9f6-339">from: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="cf9f6-p112">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="cf9f6-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="cf9f6-344">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="cf9f6-345">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-345">Type</span></span>

*   [<span data-ttu-id="cf9f6-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="cf9f6-346">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="example"></a><span data-ttu-id="cf9f6-347">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-347">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="cf9f6-348">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-348">Requirements</span></span>

|<span data-ttu-id="cf9f6-349">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-349">Requirement</span></span>| <span data-ttu-id="cf9f6-350">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-351">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-352">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-352">1.0</span></span>|
|[<span data-ttu-id="cf9f6-353">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-354">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-355">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-356">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-356">Read</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="cf9f6-357">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-357">internetMessageId: String</span></span>

<span data-ttu-id="cf9f6-p114">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="cf9f6-360">Type</span><span class="sxs-lookup"><span data-stu-id="cf9f6-360">Type</span></span>

*   <span data-ttu-id="cf9f6-361">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-362">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-362">Requirements</span></span>

|<span data-ttu-id="cf9f6-363">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-363">Requirement</span></span>| <span data-ttu-id="cf9f6-364">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-365">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-366">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-366">1.0</span></span>|
|[<span data-ttu-id="cf9f6-367">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-368">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-369">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-370">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf9f6-371">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-371">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="cf9f6-372">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-372">itemClass: String</span></span>

<span data-ttu-id="cf9f6-p115">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="cf9f6-p116">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="cf9f6-377">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-377">Type</span></span> | <span data-ttu-id="cf9f6-378">説明</span><span class="sxs-lookup"><span data-stu-id="cf9f6-378">Description</span></span> | <span data-ttu-id="cf9f6-379">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="cf9f6-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="cf9f6-380">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="cf9f6-380">Appointment items</span></span> | <span data-ttu-id="cf9f6-381">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="cf9f6-382">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="cf9f6-382">Message items</span></span> | <span data-ttu-id="cf9f6-383">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="cf9f6-384">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="cf9f6-385">Type</span><span class="sxs-lookup"><span data-stu-id="cf9f6-385">Type</span></span>

*   <span data-ttu-id="cf9f6-386">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-387">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-387">Requirements</span></span>

|<span data-ttu-id="cf9f6-388">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-388">Requirement</span></span>| <span data-ttu-id="cf9f6-389">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-390">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-391">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-391">1.0</span></span>|
|[<span data-ttu-id="cf9f6-392">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-393">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-394">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-395">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf9f6-396">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-396">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="cf9f6-397">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-397">(nullable) itemId: String</span></span>

<span data-ttu-id="cf9f6-p117">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="cf9f6-400">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-400">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="cf9f6-401">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="cf9f6-402">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-402">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="cf9f6-403">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="cf9f6-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="cf9f6-406">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-406">Type</span></span>

*   <span data-ttu-id="cf9f6-407">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-408">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-408">Requirements</span></span>

|<span data-ttu-id="cf9f6-409">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-409">Requirement</span></span>| <span data-ttu-id="cf9f6-410">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-411">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-412">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-412">1.0</span></span>|
|[<span data-ttu-id="cf9f6-413">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-414">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-415">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-416">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf9f6-417">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-417">Example</span></span>

<span data-ttu-id="cf9f6-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-16"></a><span data-ttu-id="cf9f6-420">itemType: [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-420">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span></span>

<span data-ttu-id="cf9f6-421">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-421">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="cf9f6-422">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-422">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="cf9f6-423">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-423">Type</span></span>

*   [<span data-ttu-id="cf9f6-424">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="cf9f6-424">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="cf9f6-425">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-425">Requirements</span></span>

|<span data-ttu-id="cf9f6-426">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-426">Requirement</span></span>| <span data-ttu-id="cf9f6-427">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-428">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-429">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-429">1.0</span></span>|
|[<span data-ttu-id="cf9f6-430">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-431">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-432">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-433">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cf9f6-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf9f6-434">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-434">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-16"></a><span data-ttu-id="cf9f6-435">場所: String |[場所](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-435">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

<span data-ttu-id="cf9f6-436">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-436">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cf9f6-437">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-437">Read mode</span></span>

<span data-ttu-id="cf9f6-438">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-438">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="cf9f6-439">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-439">Compose mode</span></span>

<span data-ttu-id="cf9f6-440">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-440">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="cf9f6-441">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-441">Type</span></span>

*   <span data-ttu-id="cf9f6-442">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-442">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-443">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-443">Requirements</span></span>

|<span data-ttu-id="cf9f6-444">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-444">Requirement</span></span>| <span data-ttu-id="cf9f6-445">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-446">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-447">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-447">1.0</span></span>|
|[<span data-ttu-id="cf9f6-448">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-448">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-449">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-450">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-450">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-451">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cf9f6-451">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="cf9f6-452">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-452">normalizedSubject: String</span></span>

<span data-ttu-id="cf9f6-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="cf9f6-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="cf9f6-457">Type</span><span class="sxs-lookup"><span data-stu-id="cf9f6-457">Type</span></span>

*   <span data-ttu-id="cf9f6-458">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-458">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-459">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-459">Requirements</span></span>

|<span data-ttu-id="cf9f6-460">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-460">Requirement</span></span>| <span data-ttu-id="cf9f6-461">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-462">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-463">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-463">1.0</span></span>|
|[<span data-ttu-id="cf9f6-464">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-465">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-466">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-467">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf9f6-468">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-468">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-16"></a><span data-ttu-id="cf9f6-469">notificationMessages: [Notificationmessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-469">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span></span>

<span data-ttu-id="cf9f6-470">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-470">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="cf9f6-471">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-471">Type</span></span>

*   [<span data-ttu-id="cf9f6-472">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="cf9f6-472">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="cf9f6-473">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-473">Requirements</span></span>

|<span data-ttu-id="cf9f6-474">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-474">Requirement</span></span>| <span data-ttu-id="cf9f6-475">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-476">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-476">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-477">1.3</span><span class="sxs-lookup"><span data-stu-id="cf9f6-477">1.3</span></span>|
|[<span data-ttu-id="cf9f6-478">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-478">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-479">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-480">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-480">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-481">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cf9f6-481">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf9f6-482">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-482">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="cf9f6-483">任意出席者: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-483">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="cf9f6-484">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-484">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="cf9f6-485">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-485">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cf9f6-486">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-486">Read mode</span></span>

<span data-ttu-id="cf9f6-487">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-487">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="cf9f6-488">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-488">Compose mode</span></span>

<span data-ttu-id="cf9f6-489">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-489">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="cf9f6-490">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-490">Type</span></span>

*   <span data-ttu-id="cf9f6-491">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-491">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-492">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-492">Requirements</span></span>

|<span data-ttu-id="cf9f6-493">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-493">Requirement</span></span>| <span data-ttu-id="cf9f6-494">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-495">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-496">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-496">1.0</span></span>|
|[<span data-ttu-id="cf9f6-497">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-498">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-499">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-500">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cf9f6-500">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="cf9f6-501">開催者: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-501">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="cf9f6-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="cf9f6-504">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-504">Type</span></span>

*   [<span data-ttu-id="cf9f6-505">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="cf9f6-505">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="cf9f6-506">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-506">Requirements</span></span>

|<span data-ttu-id="cf9f6-507">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-507">Requirement</span></span>| <span data-ttu-id="cf9f6-508">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-509">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-510">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-510">1.0</span></span>|
|[<span data-ttu-id="cf9f6-511">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-512">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-513">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-514">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-514">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf9f6-515">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-515">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="cf9f6-516">requiredatて dees: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-516">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="cf9f6-517">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-517">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="cf9f6-518">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-518">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cf9f6-519">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-519">Read mode</span></span>

<span data-ttu-id="cf9f6-520">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-520">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="cf9f6-521">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-521">Compose mode</span></span>

<span data-ttu-id="cf9f6-522">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-522">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="cf9f6-523">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-523">Type</span></span>

*   <span data-ttu-id="cf9f6-524">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-524">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-525">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-525">Requirements</span></span>

|<span data-ttu-id="cf9f6-526">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-526">Requirement</span></span>| <span data-ttu-id="cf9f6-527">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-527">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-528">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-528">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-529">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-529">1.0</span></span>|
|[<span data-ttu-id="cf9f6-530">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-530">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-531">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-532">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-532">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-533">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cf9f6-533">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="cf9f6-534">sender: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-534">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="cf9f6-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="cf9f6-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="cf9f6-539">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-539">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="cf9f6-540">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-540">Type</span></span>

*   [<span data-ttu-id="cf9f6-541">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="cf9f6-541">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="cf9f6-542">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-542">Requirements</span></span>

|<span data-ttu-id="cf9f6-543">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-543">Requirement</span></span>| <span data-ttu-id="cf9f6-544">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-545">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-546">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-546">1.0</span></span>|
|[<span data-ttu-id="cf9f6-547">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-548">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-549">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-550">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-550">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf9f6-551">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-551">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="cf9f6-552">開始: 日付 |[時間](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-552">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="cf9f6-553">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-553">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="cf9f6-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cf9f6-556">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-556">Read mode</span></span>

<span data-ttu-id="cf9f6-557">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-557">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="cf9f6-558">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-558">Compose mode</span></span>

<span data-ttu-id="cf9f6-559">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-559">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="cf9f6-560">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-560">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="cf9f6-561">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-561">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="cf9f6-562">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-562">Type</span></span>

*   <span data-ttu-id="cf9f6-563">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-563">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-564">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-564">Requirements</span></span>

|<span data-ttu-id="cf9f6-565">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-565">Requirement</span></span>| <span data-ttu-id="cf9f6-566">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-567">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-568">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-568">1.0</span></span>|
|[<span data-ttu-id="cf9f6-569">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-570">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-571">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-572">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cf9f6-572">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-16"></a><span data-ttu-id="cf9f6-573">subject: String |[件名](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-573">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

<span data-ttu-id="cf9f6-574">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="cf9f6-575">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cf9f6-576">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-576">Read mode</span></span>

<span data-ttu-id="cf9f6-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="cf9f6-579">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-579">Compose mode</span></span>

<span data-ttu-id="cf9f6-580">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="cf9f6-581">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-581">Type</span></span>

*   <span data-ttu-id="cf9f6-582">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-582">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-583">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-583">Requirements</span></span>

|<span data-ttu-id="cf9f6-584">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-584">Requirement</span></span>| <span data-ttu-id="cf9f6-585">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-586">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-587">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-587">1.0</span></span>|
|[<span data-ttu-id="cf9f6-588">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-588">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-589">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-590">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-590">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-591">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cf9f6-591">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="cf9f6-592">宛先: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-592">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="cf9f6-593">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="cf9f6-594">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cf9f6-595">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-595">Read mode</span></span>

<span data-ttu-id="cf9f6-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="cf9f6-598">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-598">Compose mode</span></span>

<span data-ttu-id="cf9f6-599">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="cf9f6-600">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-600">Type</span></span>

*   <span data-ttu-id="cf9f6-601">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-601">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-602">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-602">Requirements</span></span>

|<span data-ttu-id="cf9f6-603">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-603">Requirement</span></span>| <span data-ttu-id="cf9f6-604">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-605">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-606">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-606">1.0</span></span>|
|[<span data-ttu-id="cf9f6-607">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-608">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-609">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-610">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cf9f6-610">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="cf9f6-611">メソッド</span><span class="sxs-lookup"><span data-stu-id="cf9f6-611">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="cf9f6-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="cf9f6-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="cf9f6-613">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="cf9f6-614">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="cf9f6-615">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf9f6-616">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cf9f6-616">Parameters</span></span>

|<span data-ttu-id="cf9f6-617">名前</span><span class="sxs-lookup"><span data-stu-id="cf9f6-617">Name</span></span>| <span data-ttu-id="cf9f6-618">種類</span><span class="sxs-lookup"><span data-stu-id="cf9f6-618">Type</span></span>| <span data-ttu-id="cf9f6-619">属性</span><span class="sxs-lookup"><span data-stu-id="cf9f6-619">Attributes</span></span>| <span data-ttu-id="cf9f6-620">説明</span><span class="sxs-lookup"><span data-stu-id="cf9f6-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="cf9f6-621">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-621">String</span></span>||<span data-ttu-id="cf9f6-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="cf9f6-624">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-624">String</span></span>||<span data-ttu-id="cf9f6-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="cf9f6-627">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cf9f6-627">Object</span></span>| <span data-ttu-id="cf9f6-628">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-628">&lt;optional&gt;</span></span>|<span data-ttu-id="cf9f6-629">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="cf9f6-630">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cf9f6-630">Object</span></span> | <span data-ttu-id="cf9f6-631">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-631">&lt;optional&gt;</span></span> | <span data-ttu-id="cf9f6-632">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="cf9f6-633">Boolean</span><span class="sxs-lookup"><span data-stu-id="cf9f6-633">Boolean</span></span> | <span data-ttu-id="cf9f6-634">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-634">&lt;optional&gt;</span></span> | <span data-ttu-id="cf9f6-635">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="cf9f6-636">function</span><span class="sxs-lookup"><span data-stu-id="cf9f6-636">function</span></span>| <span data-ttu-id="cf9f6-637">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-637">&lt;optional&gt;</span></span>|<span data-ttu-id="cf9f6-638">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="cf9f6-639">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="cf9f6-640">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="cf9f6-641">エラー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-641">Errors</span></span>

| <span data-ttu-id="cf9f6-642">エラー コード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-642">Error code</span></span> | <span data-ttu-id="cf9f6-643">説明</span><span class="sxs-lookup"><span data-stu-id="cf9f6-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="cf9f6-644">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="cf9f6-645">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="cf9f6-646">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cf9f6-647">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-647">Requirements</span></span>

|<span data-ttu-id="cf9f6-648">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-648">Requirement</span></span>| <span data-ttu-id="cf9f6-649">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-650">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-651">1.1</span><span class="sxs-lookup"><span data-stu-id="cf9f6-651">1.1</span></span>|
|[<span data-ttu-id="cf9f6-652">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="cf9f6-654">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-655">作成</span><span class="sxs-lookup"><span data-stu-id="cf9f6-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="cf9f6-656">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-656">Examples</span></span>

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

<span data-ttu-id="cf9f6-657">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="cf9f6-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="cf9f6-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="cf9f6-659">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="cf9f6-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="cf9f6-663">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="cf9f6-664">Office アドインが Outlook on the web で実行されている場合、 `addItemAttachmentAsync`メソッドは、編集しているアイテム以外のアイテムにアイテムを添付できます。ただし、これはサポートされておらず、推奨されていません。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-664">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf9f6-665">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cf9f6-665">Parameters</span></span>

|<span data-ttu-id="cf9f6-666">名前</span><span class="sxs-lookup"><span data-stu-id="cf9f6-666">Name</span></span>| <span data-ttu-id="cf9f6-667">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-667">Type</span></span>| <span data-ttu-id="cf9f6-668">属性</span><span class="sxs-lookup"><span data-stu-id="cf9f6-668">Attributes</span></span>| <span data-ttu-id="cf9f6-669">説明</span><span class="sxs-lookup"><span data-stu-id="cf9f6-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="cf9f6-670">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-670">String</span></span>||<span data-ttu-id="cf9f6-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="cf9f6-673">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-673">String</span></span>||<span data-ttu-id="cf9f6-674">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-674">The subject of the item to be attached.</span></span> <span data-ttu-id="cf9f6-675">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-675">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="cf9f6-676">Object</span><span class="sxs-lookup"><span data-stu-id="cf9f6-676">Object</span></span>| <span data-ttu-id="cf9f6-677">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-677">&lt;optional&gt;</span></span>|<span data-ttu-id="cf9f6-678">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="cf9f6-679">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cf9f6-679">Object</span></span>| <span data-ttu-id="cf9f6-680">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-680">&lt;optional&gt;</span></span>|<span data-ttu-id="cf9f6-681">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="cf9f6-682">関数</span><span class="sxs-lookup"><span data-stu-id="cf9f6-682">function</span></span>| <span data-ttu-id="cf9f6-683">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-683">&lt;optional&gt;</span></span>|<span data-ttu-id="cf9f6-684">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="cf9f6-685">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="cf9f6-686">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="cf9f6-687">エラー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-687">Errors</span></span>

| <span data-ttu-id="cf9f6-688">エラー コード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-688">Error code</span></span> | <span data-ttu-id="cf9f6-689">説明</span><span class="sxs-lookup"><span data-stu-id="cf9f6-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="cf9f6-690">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cf9f6-691">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-691">Requirements</span></span>

|<span data-ttu-id="cf9f6-692">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-692">Requirement</span></span>| <span data-ttu-id="cf9f6-693">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-694">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-695">1.1</span><span class="sxs-lookup"><span data-stu-id="cf9f6-695">1.1</span></span>|
|[<span data-ttu-id="cf9f6-696">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="cf9f6-698">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-699">作成</span><span class="sxs-lookup"><span data-stu-id="cf9f6-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="cf9f6-700">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-700">Example</span></span>

<span data-ttu-id="cf9f6-701">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="cf9f6-702">close()</span><span class="sxs-lookup"><span data-stu-id="cf9f6-702">close()</span></span>

<span data-ttu-id="cf9f6-703">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="cf9f6-p137">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="cf9f6-706">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="cf9f6-707">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-708">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-708">Requirements</span></span>

|<span data-ttu-id="cf9f6-709">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-709">Requirement</span></span>| <span data-ttu-id="cf9f6-710">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-711">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-712">1.3</span><span class="sxs-lookup"><span data-stu-id="cf9f6-712">1.3</span></span>|
|[<span data-ttu-id="cf9f6-713">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-714">制限あり</span><span class="sxs-lookup"><span data-stu-id="cf9f6-714">Restricted</span></span>|
|[<span data-ttu-id="cf9f6-715">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-716">新規作成</span><span class="sxs-lookup"><span data-stu-id="cf9f6-716">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="cf9f6-717">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="cf9f6-717">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="cf9f6-718">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="cf9f6-719">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-719">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cf9f6-720">Web 上の Outlook では、返信フォームは、3列表示のポップアップフォームとして表示され、2列または1列表示のポップアップフォームとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-720">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="cf9f6-721">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="cf9f6-722">`formData.attachments`パラメーターで添付ファイルが指定されている場合、web 上の Outlook およびデスクトップクライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-722">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="cf9f6-723">添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-723">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="cf9f6-724">表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-724">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf9f6-725">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cf9f6-725">Parameters</span></span>

| <span data-ttu-id="cf9f6-726">名前</span><span class="sxs-lookup"><span data-stu-id="cf9f6-726">Name</span></span> | <span data-ttu-id="cf9f6-727">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-727">Type</span></span> | <span data-ttu-id="cf9f6-728">属性</span><span class="sxs-lookup"><span data-stu-id="cf9f6-728">Attributes</span></span> | <span data-ttu-id="cf9f6-729">説明</span><span class="sxs-lookup"><span data-stu-id="cf9f6-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="cf9f6-730">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="cf9f6-730">String &#124; Object</span></span>| |<span data-ttu-id="cf9f6-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="cf9f6-733">**または**</span><span class="sxs-lookup"><span data-stu-id="cf9f6-733">**OR**</span></span><br/><span data-ttu-id="cf9f6-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="cf9f6-736">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-736">String</span></span> | <span data-ttu-id="cf9f6-737">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-737">&lt;optional&gt;</span></span> | <span data-ttu-id="cf9f6-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="cf9f6-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="cf9f6-741">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-741">&lt;optional&gt;</span></span> | <span data-ttu-id="cf9f6-742">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="cf9f6-743">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-743">String</span></span> | | <span data-ttu-id="cf9f6-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="cf9f6-746">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-746">String</span></span> | | <span data-ttu-id="cf9f6-747">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="cf9f6-748">文字列</span><span class="sxs-lookup"><span data-stu-id="cf9f6-748">String</span></span> | | <span data-ttu-id="cf9f6-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="cf9f6-751">ブール値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-751">Boolean</span></span> | | <span data-ttu-id="cf9f6-p144">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="cf9f6-754">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-754">String</span></span> | | <span data-ttu-id="cf9f6-p145">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="cf9f6-758">function</span><span class="sxs-lookup"><span data-stu-id="cf9f6-758">function</span></span> | <span data-ttu-id="cf9f6-759">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-759">&lt;optional&gt;</span></span> | <span data-ttu-id="cf9f6-760">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cf9f6-761">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-761">Requirements</span></span>

|<span data-ttu-id="cf9f6-762">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-762">Requirement</span></span>| <span data-ttu-id="cf9f6-763">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-764">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-765">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-765">1.0</span></span>|
|[<span data-ttu-id="cf9f6-766">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-767">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-768">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-769">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="cf9f6-770">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-770">Examples</span></span>

<span data-ttu-id="cf9f6-771">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="cf9f6-772">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-772">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="cf9f6-773">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-773">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="cf9f6-774">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="cf9f6-775">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="cf9f6-776">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="cf9f6-777">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="cf9f6-777">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="cf9f6-778">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="cf9f6-779">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-779">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cf9f6-780">Web 上の Outlook では、返信フォームは、3列表示のポップアップフォームとして表示され、2列または1列表示のポップアップフォームとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-780">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="cf9f6-781">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="cf9f6-782">`formData.attachments`パラメーターで添付ファイルが指定されている場合、web 上の Outlook およびデスクトップクライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-782">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="cf9f6-783">添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-783">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="cf9f6-784">表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-784">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf9f6-785">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cf9f6-785">Parameters</span></span>

| <span data-ttu-id="cf9f6-786">名前</span><span class="sxs-lookup"><span data-stu-id="cf9f6-786">Name</span></span> | <span data-ttu-id="cf9f6-787">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-787">Type</span></span> | <span data-ttu-id="cf9f6-788">属性</span><span class="sxs-lookup"><span data-stu-id="cf9f6-788">Attributes</span></span> | <span data-ttu-id="cf9f6-789">説明</span><span class="sxs-lookup"><span data-stu-id="cf9f6-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="cf9f6-790">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="cf9f6-790">String &#124; Object</span></span>| | <span data-ttu-id="cf9f6-p147">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="cf9f6-793">**または**</span><span class="sxs-lookup"><span data-stu-id="cf9f6-793">**OR**</span></span><br/><span data-ttu-id="cf9f6-p148">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="cf9f6-796">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-796">String</span></span> | <span data-ttu-id="cf9f6-797">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-797">&lt;optional&gt;</span></span> | <span data-ttu-id="cf9f6-p149">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="cf9f6-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="cf9f6-801">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-801">&lt;optional&gt;</span></span> | <span data-ttu-id="cf9f6-802">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="cf9f6-803">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-803">String</span></span> | | <span data-ttu-id="cf9f6-p150">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="cf9f6-806">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-806">String</span></span> | | <span data-ttu-id="cf9f6-807">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="cf9f6-808">文字列</span><span class="sxs-lookup"><span data-stu-id="cf9f6-808">String</span></span> | | <span data-ttu-id="cf9f6-p151">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="cf9f6-811">ブール値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-811">Boolean</span></span> | | <span data-ttu-id="cf9f6-p152">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="cf9f6-814">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-814">String</span></span> | | <span data-ttu-id="cf9f6-p153">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="cf9f6-818">function</span><span class="sxs-lookup"><span data-stu-id="cf9f6-818">function</span></span> | <span data-ttu-id="cf9f6-819">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-819">&lt;optional&gt;</span></span> | <span data-ttu-id="cf9f6-820">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cf9f6-821">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-821">Requirements</span></span>

|<span data-ttu-id="cf9f6-822">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-822">Requirement</span></span>| <span data-ttu-id="cf9f6-823">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-824">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-825">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-825">1.0</span></span>|
|[<span data-ttu-id="cf9f6-826">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-826">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-827">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-828">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-828">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-829">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="cf9f6-830">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-830">Examples</span></span>

<span data-ttu-id="cf9f6-831">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="cf9f6-832">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-832">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="cf9f6-833">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-833">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="cf9f6-834">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="cf9f6-835">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="cf9f6-836">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="cf9f6-837">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="cf9f6-837">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="cf9f6-838">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-838">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="cf9f6-839">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-839">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-840">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-840">Requirements</span></span>

|<span data-ttu-id="cf9f6-841">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-841">Requirement</span></span>| <span data-ttu-id="cf9f6-842">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-843">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-844">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-844">1.0</span></span>|
|[<span data-ttu-id="cf9f6-845">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-846">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-847">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-848">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cf9f6-849">戻り値:</span><span class="sxs-lookup"><span data-stu-id="cf9f6-849">Returns:</span></span>

<span data-ttu-id="cf9f6-850">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-850">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="cf9f6-851">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-851">Example</span></span>

<span data-ttu-id="cf9f6-852">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-852">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="cf9f6-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="cf9f6-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="cf9f6-854">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-854">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="cf9f6-855">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-855">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf9f6-856">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cf9f6-856">Parameters</span></span>

|<span data-ttu-id="cf9f6-857">名前</span><span class="sxs-lookup"><span data-stu-id="cf9f6-857">Name</span></span>| <span data-ttu-id="cf9f6-858">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-858">Type</span></span>| <span data-ttu-id="cf9f6-859">説明</span><span class="sxs-lookup"><span data-stu-id="cf9f6-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="cf9f6-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="cf9f6-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.6)|<span data-ttu-id="cf9f6-861">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cf9f6-862">Requirements</span><span class="sxs-lookup"><span data-stu-id="cf9f6-862">Requirements</span></span>

|<span data-ttu-id="cf9f6-863">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-863">Requirement</span></span>| <span data-ttu-id="cf9f6-864">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-865">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-866">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-866">1.0</span></span>|
|[<span data-ttu-id="cf9f6-867">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-867">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-868">制限あり</span><span class="sxs-lookup"><span data-stu-id="cf9f6-868">Restricted</span></span>|
|[<span data-ttu-id="cf9f6-869">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-869">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-870">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cf9f6-871">戻り値:</span><span class="sxs-lookup"><span data-stu-id="cf9f6-871">Returns:</span></span>

<span data-ttu-id="cf9f6-872">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="cf9f6-873">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-873">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="cf9f6-874">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="cf9f6-875">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="cf9f6-876">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-876">Value of `entityType`</span></span> | <span data-ttu-id="cf9f6-877">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-877">Type of objects in returned array</span></span> | <span data-ttu-id="cf9f6-878">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="cf9f6-879">文字列</span><span class="sxs-lookup"><span data-stu-id="cf9f6-879">String</span></span> | <span data-ttu-id="cf9f6-880">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="cf9f6-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="cf9f6-881">連絡先</span><span class="sxs-lookup"><span data-stu-id="cf9f6-881">Contact</span></span> | <span data-ttu-id="cf9f6-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="cf9f6-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="cf9f6-883">文字列</span><span class="sxs-lookup"><span data-stu-id="cf9f6-883">String</span></span> | <span data-ttu-id="cf9f6-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="cf9f6-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="cf9f6-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="cf9f6-885">MeetingSuggestion</span></span> | <span data-ttu-id="cf9f6-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="cf9f6-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="cf9f6-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="cf9f6-887">PhoneNumber</span></span> | <span data-ttu-id="cf9f6-888">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="cf9f6-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="cf9f6-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="cf9f6-889">TaskSuggestion</span></span> | <span data-ttu-id="cf9f6-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="cf9f6-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="cf9f6-891">文字列</span><span class="sxs-lookup"><span data-stu-id="cf9f6-891">String</span></span> | <span data-ttu-id="cf9f6-892">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="cf9f6-892">**Restricted**</span></span> |

<span data-ttu-id="cf9f6-893">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="cf9f6-893">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

##### <a name="example"></a><span data-ttu-id="cf9f6-894">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-894">Example</span></span>

<span data-ttu-id="cf9f6-895">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-895">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="cf9f6-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="cf9f6-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="cf9f6-897">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="cf9f6-898">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-898">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cf9f6-899">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf9f6-900">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cf9f6-900">Parameters</span></span>

|<span data-ttu-id="cf9f6-901">名前</span><span class="sxs-lookup"><span data-stu-id="cf9f6-901">Name</span></span>| <span data-ttu-id="cf9f6-902">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-902">Type</span></span>| <span data-ttu-id="cf9f6-903">説明</span><span class="sxs-lookup"><span data-stu-id="cf9f6-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="cf9f6-904">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-904">String</span></span>|<span data-ttu-id="cf9f6-905">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cf9f6-906">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-906">Requirements</span></span>

|<span data-ttu-id="cf9f6-907">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-907">Requirement</span></span>| <span data-ttu-id="cf9f6-908">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-909">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-910">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-910">1.0</span></span>|
|[<span data-ttu-id="cf9f6-911">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-911">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-912">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-913">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-913">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-914">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cf9f6-915">戻り値:</span><span class="sxs-lookup"><span data-stu-id="cf9f6-915">Returns:</span></span>

<span data-ttu-id="cf9f6-p155">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="cf9f6-918">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="cf9f6-918">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="cf9f6-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="cf9f6-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="cf9f6-920">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="cf9f6-921">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-921">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cf9f6-p156">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="cf9f6-925">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="cf9f6-926">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="cf9f6-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-930">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-930">Requirements</span></span>

|<span data-ttu-id="cf9f6-931">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-931">Requirement</span></span>| <span data-ttu-id="cf9f6-932">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-933">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-934">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-934">1.0</span></span>|
|[<span data-ttu-id="cf9f6-935">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-936">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-937">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-938">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cf9f6-939">戻り値:</span><span class="sxs-lookup"><span data-stu-id="cf9f6-939">Returns:</span></span>

<span data-ttu-id="cf9f6-p158">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="cf9f6-942">型: Object</span><span class="sxs-lookup"><span data-stu-id="cf9f6-942">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="cf9f6-943">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-943">Example</span></span>

<span data-ttu-id="cf9f6-944">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-944">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="cf9f6-945">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="cf9f6-945">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="cf9f6-946">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-946">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="cf9f6-947">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-947">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cf9f6-948">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-948">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="cf9f6-p159">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf9f6-951">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cf9f6-951">Parameters</span></span>

|<span data-ttu-id="cf9f6-952">名前</span><span class="sxs-lookup"><span data-stu-id="cf9f6-952">Name</span></span>| <span data-ttu-id="cf9f6-953">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-953">Type</span></span>| <span data-ttu-id="cf9f6-954">説明</span><span class="sxs-lookup"><span data-stu-id="cf9f6-954">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="cf9f6-955">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-955">String</span></span>|<span data-ttu-id="cf9f6-956">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-956">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cf9f6-957">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-957">Requirements</span></span>

|<span data-ttu-id="cf9f6-958">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-958">Requirement</span></span>| <span data-ttu-id="cf9f6-959">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-959">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-960">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-960">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-961">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-961">1.0</span></span>|
|[<span data-ttu-id="cf9f6-962">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-962">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-963">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-963">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-964">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-964">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-965">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-965">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cf9f6-966">戻り値:</span><span class="sxs-lookup"><span data-stu-id="cf9f6-966">Returns:</span></span>

<span data-ttu-id="cf9f6-967">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-967">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="cf9f6-968">型: Array. < 文字列 ></span><span class="sxs-lookup"><span data-stu-id="cf9f6-968">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="cf9f6-969">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-969">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="cf9f6-970">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="cf9f6-970">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="cf9f6-971">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-971">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="cf9f6-p160">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf9f6-974">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cf9f6-974">Parameters</span></span>

|<span data-ttu-id="cf9f6-975">名前</span><span class="sxs-lookup"><span data-stu-id="cf9f6-975">Name</span></span>| <span data-ttu-id="cf9f6-976">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-976">Type</span></span>| <span data-ttu-id="cf9f6-977">属性</span><span class="sxs-lookup"><span data-stu-id="cf9f6-977">Attributes</span></span>| <span data-ttu-id="cf9f6-978">説明</span><span class="sxs-lookup"><span data-stu-id="cf9f6-978">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="cf9f6-979">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="cf9f6-979">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="cf9f6-p161">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="cf9f6-983">Object</span><span class="sxs-lookup"><span data-stu-id="cf9f6-983">Object</span></span>| <span data-ttu-id="cf9f6-984">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-984">&lt;optional&gt;</span></span>|<span data-ttu-id="cf9f6-985">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-985">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="cf9f6-986">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cf9f6-986">Object</span></span>| <span data-ttu-id="cf9f6-987">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-987">&lt;optional&gt;</span></span>|<span data-ttu-id="cf9f6-988">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-988">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="cf9f6-989">function</span><span class="sxs-lookup"><span data-stu-id="cf9f6-989">function</span></span>||<span data-ttu-id="cf9f6-990">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-990">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cf9f6-991">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-991">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="cf9f6-992">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-992">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cf9f6-993">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-993">Requirements</span></span>

|<span data-ttu-id="cf9f6-994">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-994">Requirement</span></span>| <span data-ttu-id="cf9f6-995">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-995">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-996">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-996">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-997">1.2</span><span class="sxs-lookup"><span data-stu-id="cf9f6-997">1.2</span></span>|
|[<span data-ttu-id="cf9f6-998">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-998">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-999">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-999">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-1000">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1000">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-1001">作成</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1001">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="cf9f6-1002">戻り値:</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1002">Returns:</span></span>

<span data-ttu-id="cf9f6-1003">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1003">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="cf9f6-1004">型:String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1004">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="cf9f6-1005">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1005">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="cf9f6-1006">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1006">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="cf9f6-1007">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1007">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="cf9f6-1008">強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1008">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="cf9f6-1009">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1009">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-1010">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1010">Requirements</span></span>

|<span data-ttu-id="cf9f6-1011">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1011">Requirement</span></span>| <span data-ttu-id="cf9f6-1012">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1012">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-1013">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1013">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-1014">1.6</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1014">1.6</span></span> |
|[<span data-ttu-id="cf9f6-1015">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1015">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-1016">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1016">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-1017">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1017">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-1018">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1018">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cf9f6-1019">戻り値:</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1019">Returns:</span></span>

<span data-ttu-id="cf9f6-1020">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1020">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="cf9f6-1021">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1021">Example</span></span>

<span data-ttu-id="cf9f6-1022">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1022">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="cf9f6-1023">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1023">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="cf9f6-p164">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="cf9f6-1026">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1026">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cf9f6-p165">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="cf9f6-1030">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1030">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="cf9f6-1031">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1031">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="cf9f6-p166">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cf9f6-1035">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1035">Requirements</span></span>

|<span data-ttu-id="cf9f6-1036">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1036">Requirement</span></span>| <span data-ttu-id="cf9f6-1037">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1037">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-1038">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1038">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-1039">1.6</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1039">1.6</span></span> |
|[<span data-ttu-id="cf9f6-1040">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1040">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-1041">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1041">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-1042">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1042">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-1043">読み取り</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1043">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cf9f6-1044">戻り値:</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1044">Returns:</span></span>

<span data-ttu-id="cf9f6-p167">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="cf9f6-1047">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1047">Example</span></span>

<span data-ttu-id="cf9f6-1048">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1048">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="cf9f6-1049">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1049">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="cf9f6-1050">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1050">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="cf9f6-p168">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf9f6-1054">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1054">Parameters</span></span>

|<span data-ttu-id="cf9f6-1055">名前</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1055">Name</span></span>| <span data-ttu-id="cf9f6-1056">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1056">Type</span></span>| <span data-ttu-id="cf9f6-1057">属性</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1057">Attributes</span></span>| <span data-ttu-id="cf9f6-1058">説明</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1058">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="cf9f6-1059">function</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1059">function</span></span>||<span data-ttu-id="cf9f6-1060">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1060">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cf9f6-1061">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1061">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="cf9f6-1062">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1062">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="cf9f6-1063">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1063">Object</span></span>| <span data-ttu-id="cf9f6-1064">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1064">&lt;optional&gt;</span></span>|<span data-ttu-id="cf9f6-1065">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1065">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="cf9f6-1066">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1066">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cf9f6-1067">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1067">Requirements</span></span>

|<span data-ttu-id="cf9f6-1068">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1068">Requirement</span></span>| <span data-ttu-id="cf9f6-1069">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1069">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-1070">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1070">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-1071">1.0</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1071">1.0</span></span>|
|[<span data-ttu-id="cf9f6-1072">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1072">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-1073">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1073">ReadItem</span></span>|
|[<span data-ttu-id="cf9f6-1074">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1074">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-1075">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1075">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cf9f6-1076">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1076">Example</span></span>

<span data-ttu-id="cf9f6-p171">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="cf9f6-1080">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1080">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="cf9f6-1081">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1081">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="cf9f6-1082">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1082">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="cf9f6-1083">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1083">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="cf9f6-1084">Outlook on the web およびモバイルデバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1084">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="cf9f6-1085">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1085">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf9f6-1086">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1086">Parameters</span></span>

|<span data-ttu-id="cf9f6-1087">名前</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1087">Name</span></span>| <span data-ttu-id="cf9f6-1088">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1088">Type</span></span>| <span data-ttu-id="cf9f6-1089">属性</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1089">Attributes</span></span>| <span data-ttu-id="cf9f6-1090">説明</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1090">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="cf9f6-1091">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1091">String</span></span>||<span data-ttu-id="cf9f6-1092">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1092">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="cf9f6-1093">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1093">Object</span></span>| <span data-ttu-id="cf9f6-1094">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1094">&lt;optional&gt;</span></span>|<span data-ttu-id="cf9f6-1095">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1095">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="cf9f6-1096">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1096">Object</span></span>| <span data-ttu-id="cf9f6-1097">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="cf9f6-1098">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1098">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="cf9f6-1099">function</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1099">function</span></span>| <span data-ttu-id="cf9f6-1100">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1100">&lt;optional&gt;</span></span>|<span data-ttu-id="cf9f6-1101">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="cf9f6-1102">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1102">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="cf9f6-1103">エラー</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1103">Errors</span></span>

| <span data-ttu-id="cf9f6-1104">エラー コード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1104">Error code</span></span> | <span data-ttu-id="cf9f6-1105">説明</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1105">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="cf9f6-1106">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1106">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cf9f6-1107">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1107">Requirements</span></span>

|<span data-ttu-id="cf9f6-1108">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1108">Requirement</span></span>| <span data-ttu-id="cf9f6-1109">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1109">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-1110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-1111">1.1</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1111">1.1</span></span>|
|[<span data-ttu-id="cf9f6-1112">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-1113">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1113">ReadWriteItem</span></span>|
|[<span data-ttu-id="cf9f6-1114">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-1115">作成</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1115">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="cf9f6-1116">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1116">Example</span></span>

<span data-ttu-id="cf9f6-1117">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1117">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="cf9f6-1118">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1118">saveAsync([options], callback)</span></span>

<span data-ttu-id="cf9f6-1119">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1119">Asynchronously saves an item.</span></span>

<span data-ttu-id="cf9f6-1120">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1120">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="cf9f6-1121">Outlook on the web または online モードの Outlook では、アイテムはサーバーに保存されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1121">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="cf9f6-1122">キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1122">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="cf9f6-1123">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1123">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="cf9f6-1124">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1124">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="cf9f6-p175">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="cf9f6-1128">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1128">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="cf9f6-1129">Outlook on Mac では、会議の保存はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1129">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="cf9f6-1130">新規`saveAsync`作成モードで会議から呼び出された場合、メソッドは失敗します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1130">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="cf9f6-1131">回避策については[、「OFFICE JS API を使用して Outlook For Mac で会議を下書きとして保存できません](https://support.microsoft.com/help/4505745)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1131">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="cf9f6-1132">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1132">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf9f6-1133">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1133">Parameters</span></span>

|<span data-ttu-id="cf9f6-1134">名前</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1134">Name</span></span>| <span data-ttu-id="cf9f6-1135">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1135">Type</span></span>| <span data-ttu-id="cf9f6-1136">属性</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1136">Attributes</span></span>| <span data-ttu-id="cf9f6-1137">説明</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1137">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="cf9f6-1138">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1138">Object</span></span>| <span data-ttu-id="cf9f6-1139">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1139">&lt;optional&gt;</span></span>|<span data-ttu-id="cf9f6-1140">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1140">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="cf9f6-1141">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1141">Object</span></span>| <span data-ttu-id="cf9f6-1142">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="cf9f6-1143">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1143">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="cf9f6-1144">関数</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1144">function</span></span>||<span data-ttu-id="cf9f6-1145">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1145">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cf9f6-1146">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1146">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cf9f6-1147">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1147">Requirements</span></span>

|<span data-ttu-id="cf9f6-1148">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1148">Requirement</span></span>| <span data-ttu-id="cf9f6-1149">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1149">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-1150">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-1151">1.3</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1151">1.3</span></span>|
|[<span data-ttu-id="cf9f6-1152">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1152">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-1153">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1153">ReadWriteItem</span></span>|
|[<span data-ttu-id="cf9f6-1154">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1154">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-1155">作成</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1155">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="cf9f6-1156">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1156">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="cf9f6-p177">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="cf9f6-1159">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1159">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="cf9f6-1160">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1160">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="cf9f6-p178">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cf9f6-1164">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1164">Parameters</span></span>

|<span data-ttu-id="cf9f6-1165">名前</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1165">Name</span></span>| <span data-ttu-id="cf9f6-1166">型</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1166">Type</span></span>| <span data-ttu-id="cf9f6-1167">属性</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1167">Attributes</span></span>| <span data-ttu-id="cf9f6-1168">説明</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1168">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="cf9f6-1169">String</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1169">String</span></span>||<span data-ttu-id="cf9f6-p179">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="cf9f6-1173">Object</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1173">Object</span></span>| <span data-ttu-id="cf9f6-1174">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1174">&lt;optional&gt;</span></span>|<span data-ttu-id="cf9f6-1175">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1175">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="cf9f6-1176">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1176">Object</span></span>| <span data-ttu-id="cf9f6-1177">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="cf9f6-1178">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1178">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="cf9f6-1179">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1179">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="cf9f6-1180">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="cf9f6-1181">の`text`場合、現在のスタイルが Outlook on the web およびデスクトップクライアントで適用されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1181">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="cf9f6-1182">フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1182">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="cf9f6-1183">フィールド`html`が HTML をサポートする場合 (件名は含まれません)、現在のスタイルが outlook on the web で適用され、既定のスタイルが outlook デスクトップクライアントで適用されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1183">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="cf9f6-1184">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1184">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="cf9f6-1185">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1185">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="cf9f6-1186">function</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1186">function</span></span>||<span data-ttu-id="cf9f6-1187">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1187">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cf9f6-1188">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1188">Requirements</span></span>

|<span data-ttu-id="cf9f6-1189">要件</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1189">Requirement</span></span>| <span data-ttu-id="cf9f6-1190">値</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1190">Value</span></span>|
|---|---|
|[<span data-ttu-id="cf9f6-1191">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1191">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cf9f6-1192">1.2</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1192">1.2</span></span>|
|[<span data-ttu-id="cf9f6-1193">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1193">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cf9f6-1194">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1194">ReadWriteItem</span></span>|
|[<span data-ttu-id="cf9f6-1195">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1195">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cf9f6-1196">作成</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1196">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="cf9f6-1197">例</span><span class="sxs-lookup"><span data-stu-id="cf9f6-1197">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
