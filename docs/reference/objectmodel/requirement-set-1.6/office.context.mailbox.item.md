---
title: Office. メールボックス-要件セット1.6
description: ''
ms.date: 05/30/2019
localization_priority: Normal
ms.openlocfilehash: 578e25b4fd7caf08087f24febdfd5b1877ed57bf
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706331"
---
# <a name="item"></a><span data-ttu-id="03e79-102">item</span><span class="sxs-lookup"><span data-stu-id="03e79-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="03e79-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="03e79-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="03e79-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="03e79-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-106">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-106">Requirements</span></span>

|<span data-ttu-id="03e79-107">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-107">Requirement</span></span>| <span data-ttu-id="03e79-108">値</span><span class="sxs-lookup"><span data-stu-id="03e79-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-110">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-110">1.0</span></span>|
|[<span data-ttu-id="03e79-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="03e79-112">Restricted</span></span>|
|[<span data-ttu-id="03e79-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="03e79-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="03e79-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-115">Members and methods</span></span>

| <span data-ttu-id="03e79-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-116">Member</span></span> | <span data-ttu-id="03e79-117">種類</span><span class="sxs-lookup"><span data-stu-id="03e79-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="03e79-118">attachments</span><span class="sxs-lookup"><span data-stu-id="03e79-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="03e79-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-119">Member</span></span> |
| [<span data-ttu-id="03e79-120">bcc</span><span class="sxs-lookup"><span data-stu-id="03e79-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="03e79-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-121">Member</span></span> |
| [<span data-ttu-id="03e79-122">body</span><span class="sxs-lookup"><span data-stu-id="03e79-122">body</span></span>](#body-body) | <span data-ttu-id="03e79-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-123">Member</span></span> |
| [<span data-ttu-id="03e79-124">cc</span><span class="sxs-lookup"><span data-stu-id="03e79-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="03e79-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-125">Member</span></span> |
| [<span data-ttu-id="03e79-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="03e79-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="03e79-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-127">Member</span></span> |
| [<span data-ttu-id="03e79-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="03e79-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="03e79-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-129">Member</span></span> |
| [<span data-ttu-id="03e79-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="03e79-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="03e79-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-131">Member</span></span> |
| [<span data-ttu-id="03e79-132">end</span><span class="sxs-lookup"><span data-stu-id="03e79-132">end</span></span>](#end-datetime) | <span data-ttu-id="03e79-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-133">Member</span></span> |
| [<span data-ttu-id="03e79-134">from</span><span class="sxs-lookup"><span data-stu-id="03e79-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="03e79-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-135">Member</span></span> |
| [<span data-ttu-id="03e79-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="03e79-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="03e79-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-137">Member</span></span> |
| [<span data-ttu-id="03e79-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="03e79-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="03e79-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-139">Member</span></span> |
| [<span data-ttu-id="03e79-140">itemId</span><span class="sxs-lookup"><span data-stu-id="03e79-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="03e79-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-141">Member</span></span> |
| [<span data-ttu-id="03e79-142">itemType</span><span class="sxs-lookup"><span data-stu-id="03e79-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="03e79-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-143">Member</span></span> |
| [<span data-ttu-id="03e79-144">location</span><span class="sxs-lookup"><span data-stu-id="03e79-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="03e79-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-145">Member</span></span> |
| [<span data-ttu-id="03e79-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="03e79-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="03e79-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-147">Member</span></span> |
| [<span data-ttu-id="03e79-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="03e79-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="03e79-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-149">Member</span></span> |
| [<span data-ttu-id="03e79-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="03e79-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="03e79-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-151">Member</span></span> |
| [<span data-ttu-id="03e79-152">organizer</span><span class="sxs-lookup"><span data-stu-id="03e79-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="03e79-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-153">Member</span></span> |
| [<span data-ttu-id="03e79-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="03e79-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="03e79-155">Member</span><span class="sxs-lookup"><span data-stu-id="03e79-155">Member</span></span> |
| [<span data-ttu-id="03e79-156">sender</span><span class="sxs-lookup"><span data-stu-id="03e79-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="03e79-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-157">Member</span></span> |
| [<span data-ttu-id="03e79-158">start</span><span class="sxs-lookup"><span data-stu-id="03e79-158">start</span></span>](#start-datetime) | <span data-ttu-id="03e79-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-159">Member</span></span> |
| [<span data-ttu-id="03e79-160">subject</span><span class="sxs-lookup"><span data-stu-id="03e79-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="03e79-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-161">Member</span></span> |
| [<span data-ttu-id="03e79-162">to</span><span class="sxs-lookup"><span data-stu-id="03e79-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="03e79-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-163">Member</span></span> |
| [<span data-ttu-id="03e79-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="03e79-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="03e79-165">メソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-165">Method</span></span> |
| [<span data-ttu-id="03e79-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="03e79-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="03e79-167">メソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-167">Method</span></span> |
| [<span data-ttu-id="03e79-168">close</span><span class="sxs-lookup"><span data-stu-id="03e79-168">close</span></span>](#close) | <span data-ttu-id="03e79-169">メソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-169">Method</span></span> |
| [<span data-ttu-id="03e79-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="03e79-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="03e79-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-171">Method</span></span> |
| [<span data-ttu-id="03e79-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="03e79-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="03e79-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-173">Method</span></span> |
| [<span data-ttu-id="03e79-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="03e79-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="03e79-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-175">Method</span></span> |
| [<span data-ttu-id="03e79-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="03e79-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="03e79-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-177">Method</span></span> |
| [<span data-ttu-id="03e79-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="03e79-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="03e79-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-179">Method</span></span> |
| [<span data-ttu-id="03e79-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="03e79-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="03e79-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-181">Method</span></span> |
| [<span data-ttu-id="03e79-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="03e79-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="03e79-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-183">Method</span></span> |
| [<span data-ttu-id="03e79-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="03e79-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="03e79-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-185">Method</span></span> |
| [<span data-ttu-id="03e79-186">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="03e79-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="03e79-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-187">Method</span></span> |
| [<span data-ttu-id="03e79-188">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="03e79-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="03e79-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-189">Method</span></span> |
| [<span data-ttu-id="03e79-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="03e79-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="03e79-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-191">Method</span></span> |
| [<span data-ttu-id="03e79-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="03e79-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="03e79-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-193">Method</span></span> |
| [<span data-ttu-id="03e79-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="03e79-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="03e79-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-195">Method</span></span> |
| [<span data-ttu-id="03e79-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="03e79-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="03e79-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="03e79-198">例</span><span class="sxs-lookup"><span data-stu-id="03e79-198">Example</span></span>

<span data-ttu-id="03e79-199">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="03e79-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="03e79-200">メンバー</span><span class="sxs-lookup"><span data-stu-id="03e79-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="03e79-201">添付ファイル: <[Attachmentdetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="03e79-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="03e79-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="03e79-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="03e79-204">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="03e79-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="03e79-205">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="03e79-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="03e79-206">型</span><span class="sxs-lookup"><span data-stu-id="03e79-206">Type</span></span>

*   <span data-ttu-id="03e79-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="03e79-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-208">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-208">Requirements</span></span>

|<span data-ttu-id="03e79-209">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-209">Requirement</span></span>| <span data-ttu-id="03e79-210">値</span><span class="sxs-lookup"><span data-stu-id="03e79-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-211">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-212">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-212">1.0</span></span>|
|[<span data-ttu-id="03e79-213">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-214">ReadItem</span></span>|
|[<span data-ttu-id="03e79-215">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-216">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03e79-217">例</span><span class="sxs-lookup"><span data-stu-id="03e79-217">Example</span></span>

<span data-ttu-id="03e79-218">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="03e79-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="03e79-219">bcc:[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="03e79-219">bcc: [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="03e79-220">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="03e79-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="03e79-221">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="03e79-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="03e79-222">型</span><span class="sxs-lookup"><span data-stu-id="03e79-222">Type</span></span>

*   [<span data-ttu-id="03e79-223">受信者</span><span class="sxs-lookup"><span data-stu-id="03e79-223">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="03e79-224">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-224">Requirements</span></span>

|<span data-ttu-id="03e79-225">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-225">Requirement</span></span>| <span data-ttu-id="03e79-226">値</span><span class="sxs-lookup"><span data-stu-id="03e79-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-227">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-228">1.1</span><span class="sxs-lookup"><span data-stu-id="03e79-228">1.1</span></span>|
|[<span data-ttu-id="03e79-229">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-230">ReadItem</span></span>|
|[<span data-ttu-id="03e79-231">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-232">作成</span><span class="sxs-lookup"><span data-stu-id="03e79-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="03e79-233">例</span><span class="sxs-lookup"><span data-stu-id="03e79-233">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="03e79-234">本文:[本文](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="03e79-234">body: [Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="03e79-235">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="03e79-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="03e79-236">型</span><span class="sxs-lookup"><span data-stu-id="03e79-236">Type</span></span>

*   [<span data-ttu-id="03e79-237">Body</span><span class="sxs-lookup"><span data-stu-id="03e79-237">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="03e79-238">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-238">Requirements</span></span>

|<span data-ttu-id="03e79-239">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-239">Requirement</span></span>| <span data-ttu-id="03e79-240">値</span><span class="sxs-lookup"><span data-stu-id="03e79-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-241">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-242">1.1</span><span class="sxs-lookup"><span data-stu-id="03e79-242">1.1</span></span>|
|[<span data-ttu-id="03e79-243">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-244">ReadItem</span></span>|
|[<span data-ttu-id="03e79-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-246">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="03e79-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03e79-247">例</span><span class="sxs-lookup"><span data-stu-id="03e79-247">Example</span></span>

<span data-ttu-id="03e79-248">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="03e79-248">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="03e79-249">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="03e79-249">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="03e79-250">cc: <[emailaddressdetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="03e79-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="03e79-251">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="03e79-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="03e79-252">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="03e79-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="03e79-253">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="03e79-253">Read mode</span></span>

<span data-ttu-id="03e79-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="03e79-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="03e79-256">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="03e79-256">Compose mode</span></span>

<span data-ttu-id="03e79-257">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-257">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="03e79-258">型</span><span class="sxs-lookup"><span data-stu-id="03e79-258">Type</span></span>

*   <span data-ttu-id="03e79-259">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="03e79-259">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-260">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-260">Requirements</span></span>

|<span data-ttu-id="03e79-261">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-261">Requirement</span></span>| <span data-ttu-id="03e79-262">値</span><span class="sxs-lookup"><span data-stu-id="03e79-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-263">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-264">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-264">1.0</span></span>|
|[<span data-ttu-id="03e79-265">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-266">ReadItem</span></span>|
|[<span data-ttu-id="03e79-267">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-268">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="03e79-268">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="03e79-269">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="03e79-269">(nullable) conversationId: String</span></span>

<span data-ttu-id="03e79-270">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="03e79-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="03e79-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="03e79-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="03e79-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="03e79-275">Type</span><span class="sxs-lookup"><span data-stu-id="03e79-275">Type</span></span>

*   <span data-ttu-id="03e79-276">String</span><span class="sxs-lookup"><span data-stu-id="03e79-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-277">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-277">Requirements</span></span>

|<span data-ttu-id="03e79-278">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-278">Requirement</span></span>| <span data-ttu-id="03e79-279">値</span><span class="sxs-lookup"><span data-stu-id="03e79-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-280">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-281">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-281">1.0</span></span>|
|[<span data-ttu-id="03e79-282">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-283">ReadItem</span></span>|
|[<span data-ttu-id="03e79-284">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-285">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="03e79-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03e79-286">例</span><span class="sxs-lookup"><span data-stu-id="03e79-286">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="03e79-287">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="03e79-287">dateTimeCreated: Date</span></span>

<span data-ttu-id="03e79-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="03e79-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="03e79-290">型</span><span class="sxs-lookup"><span data-stu-id="03e79-290">Type</span></span>

*   <span data-ttu-id="03e79-291">日付</span><span class="sxs-lookup"><span data-stu-id="03e79-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-292">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-292">Requirements</span></span>

|<span data-ttu-id="03e79-293">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-293">Requirement</span></span>| <span data-ttu-id="03e79-294">値</span><span class="sxs-lookup"><span data-stu-id="03e79-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-295">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-296">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-296">1.0</span></span>|
|[<span data-ttu-id="03e79-297">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-298">ReadItem</span></span>|
|[<span data-ttu-id="03e79-299">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-300">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03e79-301">例</span><span class="sxs-lookup"><span data-stu-id="03e79-301">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="03e79-302">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="03e79-302">dateTimeModified: Date</span></span>

<span data-ttu-id="03e79-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="03e79-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="03e79-305">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="03e79-305">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="03e79-306">型</span><span class="sxs-lookup"><span data-stu-id="03e79-306">Type</span></span>

*   <span data-ttu-id="03e79-307">日付</span><span class="sxs-lookup"><span data-stu-id="03e79-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-308">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-308">Requirements</span></span>

|<span data-ttu-id="03e79-309">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-309">Requirement</span></span>| <span data-ttu-id="03e79-310">値</span><span class="sxs-lookup"><span data-stu-id="03e79-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-312">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-312">1.0</span></span>|
|[<span data-ttu-id="03e79-313">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-314">ReadItem</span></span>|
|[<span data-ttu-id="03e79-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-316">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03e79-317">例</span><span class="sxs-lookup"><span data-stu-id="03e79-317">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="03e79-318">終了: 日付 |[時間](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="03e79-318">end: Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="03e79-319">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="03e79-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="03e79-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="03e79-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="03e79-322">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="03e79-322">Read mode</span></span>

<span data-ttu-id="03e79-323">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-323">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="03e79-324">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="03e79-324">Compose mode</span></span>

<span data-ttu-id="03e79-325">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="03e79-326">[`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="03e79-326">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="03e79-327">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="03e79-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="03e79-328">型</span><span class="sxs-lookup"><span data-stu-id="03e79-328">Type</span></span>

*   <span data-ttu-id="03e79-329">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="03e79-329">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-330">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-330">Requirements</span></span>

|<span data-ttu-id="03e79-331">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-331">Requirement</span></span>| <span data-ttu-id="03e79-332">値</span><span class="sxs-lookup"><span data-stu-id="03e79-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-333">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-334">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-334">1.0</span></span>|
|[<span data-ttu-id="03e79-335">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-336">ReadItem</span></span>|
|[<span data-ttu-id="03e79-337">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-338">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="03e79-338">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="03e79-339">from: [Emailaddressdetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="03e79-339">from: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="03e79-p112">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="03e79-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="03e79-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="03e79-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="03e79-344">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="03e79-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="03e79-345">型</span><span class="sxs-lookup"><span data-stu-id="03e79-345">Type</span></span>

*   [<span data-ttu-id="03e79-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="03e79-346">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="example"></a><span data-ttu-id="03e79-347">例</span><span class="sxs-lookup"><span data-stu-id="03e79-347">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="03e79-348">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-348">Requirements</span></span>

|<span data-ttu-id="03e79-349">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-349">Requirement</span></span>| <span data-ttu-id="03e79-350">値</span><span class="sxs-lookup"><span data-stu-id="03e79-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-351">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-352">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-352">1.0</span></span>|
|[<span data-ttu-id="03e79-353">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-354">ReadItem</span></span>|
|[<span data-ttu-id="03e79-355">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-356">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-356">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="03e79-357">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="03e79-357">internetMessageId: String</span></span>

<span data-ttu-id="03e79-p114">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="03e79-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="03e79-360">Type</span><span class="sxs-lookup"><span data-stu-id="03e79-360">Type</span></span>

*   <span data-ttu-id="03e79-361">String</span><span class="sxs-lookup"><span data-stu-id="03e79-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-362">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-362">Requirements</span></span>

|<span data-ttu-id="03e79-363">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-363">Requirement</span></span>| <span data-ttu-id="03e79-364">値</span><span class="sxs-lookup"><span data-stu-id="03e79-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-365">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-366">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-366">1.0</span></span>|
|[<span data-ttu-id="03e79-367">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-368">ReadItem</span></span>|
|[<span data-ttu-id="03e79-369">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-370">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03e79-371">例</span><span class="sxs-lookup"><span data-stu-id="03e79-371">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="03e79-372">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="03e79-372">itemClass: String</span></span>

<span data-ttu-id="03e79-p115">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="03e79-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="03e79-p116">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="03e79-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="03e79-377">型</span><span class="sxs-lookup"><span data-stu-id="03e79-377">Type</span></span> | <span data-ttu-id="03e79-378">説明</span><span class="sxs-lookup"><span data-stu-id="03e79-378">Description</span></span> | <span data-ttu-id="03e79-379">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="03e79-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="03e79-380">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="03e79-380">Appointment items</span></span> | <span data-ttu-id="03e79-381">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="03e79-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="03e79-382">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="03e79-382">Message items</span></span> | <span data-ttu-id="03e79-383">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="03e79-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="03e79-384">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="03e79-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="03e79-385">Type</span><span class="sxs-lookup"><span data-stu-id="03e79-385">Type</span></span>

*   <span data-ttu-id="03e79-386">String</span><span class="sxs-lookup"><span data-stu-id="03e79-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-387">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-387">Requirements</span></span>

|<span data-ttu-id="03e79-388">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-388">Requirement</span></span>| <span data-ttu-id="03e79-389">値</span><span class="sxs-lookup"><span data-stu-id="03e79-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-390">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-391">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-391">1.0</span></span>|
|[<span data-ttu-id="03e79-392">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-393">ReadItem</span></span>|
|[<span data-ttu-id="03e79-394">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-395">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03e79-396">例</span><span class="sxs-lookup"><span data-stu-id="03e79-396">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="03e79-397">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="03e79-397">(nullable) itemId: String</span></span>

<span data-ttu-id="03e79-p117">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="03e79-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="03e79-400">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="03e79-400">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="03e79-401">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="03e79-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="03e79-402">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="03e79-402">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="03e79-403">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="03e79-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="03e79-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="03e79-406">型</span><span class="sxs-lookup"><span data-stu-id="03e79-406">Type</span></span>

*   <span data-ttu-id="03e79-407">String</span><span class="sxs-lookup"><span data-stu-id="03e79-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-408">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-408">Requirements</span></span>

|<span data-ttu-id="03e79-409">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-409">Requirement</span></span>| <span data-ttu-id="03e79-410">値</span><span class="sxs-lookup"><span data-stu-id="03e79-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-411">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-412">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-412">1.0</span></span>|
|[<span data-ttu-id="03e79-413">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-414">ReadItem</span></span>|
|[<span data-ttu-id="03e79-415">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-416">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03e79-417">例</span><span class="sxs-lookup"><span data-stu-id="03e79-417">Example</span></span>

<span data-ttu-id="03e79-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="03e79-420">itemType: [MailboxEnums](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="03e79-420">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="03e79-421">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="03e79-421">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="03e79-422">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="03e79-422">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="03e79-423">型</span><span class="sxs-lookup"><span data-stu-id="03e79-423">Type</span></span>

*   [<span data-ttu-id="03e79-424">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="03e79-424">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="03e79-425">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-425">Requirements</span></span>

|<span data-ttu-id="03e79-426">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-426">Requirement</span></span>| <span data-ttu-id="03e79-427">値</span><span class="sxs-lookup"><span data-stu-id="03e79-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-428">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-429">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-429">1.0</span></span>|
|[<span data-ttu-id="03e79-430">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-431">ReadItem</span></span>|
|[<span data-ttu-id="03e79-432">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-433">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="03e79-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03e79-434">例</span><span class="sxs-lookup"><span data-stu-id="03e79-434">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="03e79-435">場所: String |[場所](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="03e79-435">location: String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="03e79-436">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="03e79-436">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="03e79-437">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="03e79-437">Read mode</span></span>

<span data-ttu-id="03e79-438">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-438">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="03e79-439">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="03e79-439">Compose mode</span></span>

<span data-ttu-id="03e79-440">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-440">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="03e79-441">型</span><span class="sxs-lookup"><span data-stu-id="03e79-441">Type</span></span>

*   <span data-ttu-id="03e79-442">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="03e79-442">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-443">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-443">Requirements</span></span>

|<span data-ttu-id="03e79-444">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-444">Requirement</span></span>| <span data-ttu-id="03e79-445">値</span><span class="sxs-lookup"><span data-stu-id="03e79-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-446">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-447">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-447">1.0</span></span>|
|[<span data-ttu-id="03e79-448">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-448">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-449">ReadItem</span></span>|
|[<span data-ttu-id="03e79-450">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-450">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-451">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="03e79-451">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="03e79-452">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="03e79-452">normalizedSubject: String</span></span>

<span data-ttu-id="03e79-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="03e79-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="03e79-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="03e79-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="03e79-457">Type</span><span class="sxs-lookup"><span data-stu-id="03e79-457">Type</span></span>

*   <span data-ttu-id="03e79-458">String</span><span class="sxs-lookup"><span data-stu-id="03e79-458">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-459">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-459">Requirements</span></span>

|<span data-ttu-id="03e79-460">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-460">Requirement</span></span>| <span data-ttu-id="03e79-461">値</span><span class="sxs-lookup"><span data-stu-id="03e79-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-462">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-463">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-463">1.0</span></span>|
|[<span data-ttu-id="03e79-464">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-465">ReadItem</span></span>|
|[<span data-ttu-id="03e79-466">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-467">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03e79-468">例</span><span class="sxs-lookup"><span data-stu-id="03e79-468">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="03e79-469">notificationMessages: [Notificationmessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="03e79-469">notificationMessages: [NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="03e79-470">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="03e79-470">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="03e79-471">型</span><span class="sxs-lookup"><span data-stu-id="03e79-471">Type</span></span>

*   [<span data-ttu-id="03e79-472">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="03e79-472">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="03e79-473">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-473">Requirements</span></span>

|<span data-ttu-id="03e79-474">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-474">Requirement</span></span>| <span data-ttu-id="03e79-475">値</span><span class="sxs-lookup"><span data-stu-id="03e79-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-476">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-476">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-477">1.3</span><span class="sxs-lookup"><span data-stu-id="03e79-477">1.3</span></span>|
|[<span data-ttu-id="03e79-478">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-478">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-479">ReadItem</span></span>|
|[<span data-ttu-id="03e79-480">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-480">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-481">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="03e79-481">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03e79-482">例</span><span class="sxs-lookup"><span data-stu-id="03e79-482">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="03e79-483">任意出席者: 配列. <[emailaddressdetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="03e79-483">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="03e79-484">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="03e79-484">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="03e79-485">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="03e79-485">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="03e79-486">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="03e79-486">Read mode</span></span>

<span data-ttu-id="03e79-487">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-487">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="03e79-488">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="03e79-488">Compose mode</span></span>

<span data-ttu-id="03e79-489">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-489">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="03e79-490">型</span><span class="sxs-lookup"><span data-stu-id="03e79-490">Type</span></span>

*   <span data-ttu-id="03e79-491">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="03e79-491">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-492">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-492">Requirements</span></span>

|<span data-ttu-id="03e79-493">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-493">Requirement</span></span>| <span data-ttu-id="03e79-494">値</span><span class="sxs-lookup"><span data-stu-id="03e79-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-495">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-496">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-496">1.0</span></span>|
|[<span data-ttu-id="03e79-497">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-498">ReadItem</span></span>|
|[<span data-ttu-id="03e79-499">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-500">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="03e79-500">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="03e79-501">開催者: [Emailaddressdetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="03e79-501">organizer: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="03e79-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="03e79-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="03e79-504">型</span><span class="sxs-lookup"><span data-stu-id="03e79-504">Type</span></span>

*   [<span data-ttu-id="03e79-505">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="03e79-505">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="03e79-506">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-506">Requirements</span></span>

|<span data-ttu-id="03e79-507">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-507">Requirement</span></span>| <span data-ttu-id="03e79-508">値</span><span class="sxs-lookup"><span data-stu-id="03e79-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-509">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-510">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-510">1.0</span></span>|
|[<span data-ttu-id="03e79-511">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-512">ReadItem</span></span>|
|[<span data-ttu-id="03e79-513">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-514">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-514">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03e79-515">例</span><span class="sxs-lookup"><span data-stu-id="03e79-515">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="03e79-516">requiredatて dees: 配列. <[emailaddressdetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="03e79-516">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="03e79-517">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="03e79-517">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="03e79-518">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="03e79-518">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="03e79-519">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="03e79-519">Read mode</span></span>

<span data-ttu-id="03e79-520">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-520">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="03e79-521">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="03e79-521">Compose mode</span></span>

<span data-ttu-id="03e79-522">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-522">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="03e79-523">型</span><span class="sxs-lookup"><span data-stu-id="03e79-523">Type</span></span>

*   <span data-ttu-id="03e79-524">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="03e79-524">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-525">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-525">Requirements</span></span>

|<span data-ttu-id="03e79-526">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-526">Requirement</span></span>| <span data-ttu-id="03e79-527">値</span><span class="sxs-lookup"><span data-stu-id="03e79-527">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-528">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-528">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-529">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-529">1.0</span></span>|
|[<span data-ttu-id="03e79-530">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-530">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-531">ReadItem</span></span>|
|[<span data-ttu-id="03e79-532">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-532">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-533">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="03e79-533">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="03e79-534">sender: [Emailaddressdetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="03e79-534">sender: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="03e79-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="03e79-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="03e79-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="03e79-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="03e79-539">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="03e79-539">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="03e79-540">型</span><span class="sxs-lookup"><span data-stu-id="03e79-540">Type</span></span>

*   [<span data-ttu-id="03e79-541">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="03e79-541">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="03e79-542">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-542">Requirements</span></span>

|<span data-ttu-id="03e79-543">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-543">Requirement</span></span>| <span data-ttu-id="03e79-544">値</span><span class="sxs-lookup"><span data-stu-id="03e79-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-545">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-546">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-546">1.0</span></span>|
|[<span data-ttu-id="03e79-547">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-548">ReadItem</span></span>|
|[<span data-ttu-id="03e79-549">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-550">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-550">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03e79-551">例</span><span class="sxs-lookup"><span data-stu-id="03e79-551">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="03e79-552">開始: 日付 |[時間](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="03e79-552">start: Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="03e79-553">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="03e79-553">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="03e79-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="03e79-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="03e79-556">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="03e79-556">Read mode</span></span>

<span data-ttu-id="03e79-557">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-557">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="03e79-558">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="03e79-558">Compose mode</span></span>

<span data-ttu-id="03e79-559">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-559">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="03e79-560">[`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="03e79-560">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="03e79-561">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="03e79-561">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="03e79-562">型</span><span class="sxs-lookup"><span data-stu-id="03e79-562">Type</span></span>

*   <span data-ttu-id="03e79-563">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="03e79-563">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-564">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-564">Requirements</span></span>

|<span data-ttu-id="03e79-565">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-565">Requirement</span></span>| <span data-ttu-id="03e79-566">値</span><span class="sxs-lookup"><span data-stu-id="03e79-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-567">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-568">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-568">1.0</span></span>|
|[<span data-ttu-id="03e79-569">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-570">ReadItem</span></span>|
|[<span data-ttu-id="03e79-571">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-572">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="03e79-572">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="03e79-573">subject: String |[件名](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="03e79-573">subject: String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="03e79-574">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="03e79-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="03e79-575">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="03e79-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="03e79-576">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="03e79-576">Read mode</span></span>

<span data-ttu-id="03e79-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="03e79-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="03e79-579">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="03e79-579">Compose mode</span></span>

<span data-ttu-id="03e79-580">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="03e79-581">型</span><span class="sxs-lookup"><span data-stu-id="03e79-581">Type</span></span>

*   <span data-ttu-id="03e79-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="03e79-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-583">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-583">Requirements</span></span>

|<span data-ttu-id="03e79-584">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-584">Requirement</span></span>| <span data-ttu-id="03e79-585">値</span><span class="sxs-lookup"><span data-stu-id="03e79-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-586">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-587">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-587">1.0</span></span>|
|[<span data-ttu-id="03e79-588">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-588">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-589">ReadItem</span></span>|
|[<span data-ttu-id="03e79-590">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-590">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-591">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="03e79-591">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="03e79-592">宛先: 配列. <[emailaddressdetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="03e79-592">to: Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="03e79-593">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="03e79-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="03e79-594">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="03e79-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="03e79-595">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="03e79-595">Read mode</span></span>

<span data-ttu-id="03e79-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="03e79-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="03e79-598">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="03e79-598">Compose mode</span></span>

<span data-ttu-id="03e79-599">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="03e79-600">型</span><span class="sxs-lookup"><span data-stu-id="03e79-600">Type</span></span>

*   <span data-ttu-id="03e79-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="03e79-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-602">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-602">Requirements</span></span>

|<span data-ttu-id="03e79-603">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-603">Requirement</span></span>| <span data-ttu-id="03e79-604">値</span><span class="sxs-lookup"><span data-stu-id="03e79-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-605">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-606">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-606">1.0</span></span>|
|[<span data-ttu-id="03e79-607">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-608">ReadItem</span></span>|
|[<span data-ttu-id="03e79-609">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-610">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="03e79-610">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="03e79-611">メソッド</span><span class="sxs-lookup"><span data-stu-id="03e79-611">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="03e79-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="03e79-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="03e79-613">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="03e79-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="03e79-614">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="03e79-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="03e79-615">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="03e79-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03e79-616">パラメーター</span><span class="sxs-lookup"><span data-stu-id="03e79-616">Parameters</span></span>

|<span data-ttu-id="03e79-617">名前</span><span class="sxs-lookup"><span data-stu-id="03e79-617">Name</span></span>| <span data-ttu-id="03e79-618">種類</span><span class="sxs-lookup"><span data-stu-id="03e79-618">Type</span></span>| <span data-ttu-id="03e79-619">属性</span><span class="sxs-lookup"><span data-stu-id="03e79-619">Attributes</span></span>| <span data-ttu-id="03e79-620">説明</span><span class="sxs-lookup"><span data-stu-id="03e79-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="03e79-621">String</span><span class="sxs-lookup"><span data-stu-id="03e79-621">String</span></span>||<span data-ttu-id="03e79-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="03e79-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="03e79-624">String</span><span class="sxs-lookup"><span data-stu-id="03e79-624">String</span></span>||<span data-ttu-id="03e79-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="03e79-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="03e79-627">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="03e79-627">Object</span></span>| <span data-ttu-id="03e79-628">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-628">&lt;optional&gt;</span></span>|<span data-ttu-id="03e79-629">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="03e79-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="03e79-630">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="03e79-630">Object</span></span> | <span data-ttu-id="03e79-631">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-631">&lt;optional&gt;</span></span> | <span data-ttu-id="03e79-632">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="03e79-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="03e79-633">Boolean</span><span class="sxs-lookup"><span data-stu-id="03e79-633">Boolean</span></span> | <span data-ttu-id="03e79-634">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-634">&lt;optional&gt;</span></span> | <span data-ttu-id="03e79-635">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="03e79-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="03e79-636">function</span><span class="sxs-lookup"><span data-stu-id="03e79-636">function</span></span>| <span data-ttu-id="03e79-637">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-637">&lt;optional&gt;</span></span>|<span data-ttu-id="03e79-638">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="03e79-639">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="03e79-640">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="03e79-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="03e79-641">エラー</span><span class="sxs-lookup"><span data-stu-id="03e79-641">Errors</span></span>

| <span data-ttu-id="03e79-642">エラー コード</span><span class="sxs-lookup"><span data-stu-id="03e79-642">Error code</span></span> | <span data-ttu-id="03e79-643">説明</span><span class="sxs-lookup"><span data-stu-id="03e79-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="03e79-644">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="03e79-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="03e79-645">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="03e79-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="03e79-646">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="03e79-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="03e79-647">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-647">Requirements</span></span>

|<span data-ttu-id="03e79-648">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-648">Requirement</span></span>| <span data-ttu-id="03e79-649">値</span><span class="sxs-lookup"><span data-stu-id="03e79-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-650">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-651">1.1</span><span class="sxs-lookup"><span data-stu-id="03e79-651">1.1</span></span>|
|[<span data-ttu-id="03e79-652">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="03e79-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="03e79-654">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-655">作成</span><span class="sxs-lookup"><span data-stu-id="03e79-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="03e79-656">例</span><span class="sxs-lookup"><span data-stu-id="03e79-656">Examples</span></span>

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

<span data-ttu-id="03e79-657">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="03e79-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="03e79-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="03e79-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="03e79-659">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="03e79-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="03e79-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="03e79-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="03e79-663">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="03e79-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="03e79-664">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="03e79-664">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03e79-665">パラメーター</span><span class="sxs-lookup"><span data-stu-id="03e79-665">Parameters</span></span>

|<span data-ttu-id="03e79-666">名前</span><span class="sxs-lookup"><span data-stu-id="03e79-666">Name</span></span>| <span data-ttu-id="03e79-667">型</span><span class="sxs-lookup"><span data-stu-id="03e79-667">Type</span></span>| <span data-ttu-id="03e79-668">属性</span><span class="sxs-lookup"><span data-stu-id="03e79-668">Attributes</span></span>| <span data-ttu-id="03e79-669">説明</span><span class="sxs-lookup"><span data-stu-id="03e79-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="03e79-670">String</span><span class="sxs-lookup"><span data-stu-id="03e79-670">String</span></span>||<span data-ttu-id="03e79-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="03e79-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="03e79-673">String</span><span class="sxs-lookup"><span data-stu-id="03e79-673">String</span></span>||<span data-ttu-id="03e79-674">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="03e79-674">The subject of the item to be attached.</span></span> <span data-ttu-id="03e79-675">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="03e79-675">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="03e79-676">Object</span><span class="sxs-lookup"><span data-stu-id="03e79-676">Object</span></span>| <span data-ttu-id="03e79-677">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-677">&lt;optional&gt;</span></span>|<span data-ttu-id="03e79-678">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="03e79-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="03e79-679">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="03e79-679">Object</span></span>| <span data-ttu-id="03e79-680">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-680">&lt;optional&gt;</span></span>|<span data-ttu-id="03e79-681">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="03e79-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="03e79-682">関数</span><span class="sxs-lookup"><span data-stu-id="03e79-682">function</span></span>| <span data-ttu-id="03e79-683">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-683">&lt;optional&gt;</span></span>|<span data-ttu-id="03e79-684">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="03e79-685">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="03e79-686">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="03e79-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="03e79-687">エラー</span><span class="sxs-lookup"><span data-stu-id="03e79-687">Errors</span></span>

| <span data-ttu-id="03e79-688">エラー コード</span><span class="sxs-lookup"><span data-stu-id="03e79-688">Error code</span></span> | <span data-ttu-id="03e79-689">説明</span><span class="sxs-lookup"><span data-stu-id="03e79-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="03e79-690">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="03e79-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="03e79-691">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-691">Requirements</span></span>

|<span data-ttu-id="03e79-692">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-692">Requirement</span></span>| <span data-ttu-id="03e79-693">値</span><span class="sxs-lookup"><span data-stu-id="03e79-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-694">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-695">1.1</span><span class="sxs-lookup"><span data-stu-id="03e79-695">1.1</span></span>|
|[<span data-ttu-id="03e79-696">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="03e79-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="03e79-698">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-699">作成</span><span class="sxs-lookup"><span data-stu-id="03e79-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="03e79-700">例</span><span class="sxs-lookup"><span data-stu-id="03e79-700">Example</span></span>

<span data-ttu-id="03e79-701">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="03e79-702">close()</span><span class="sxs-lookup"><span data-stu-id="03e79-702">close()</span></span>

<span data-ttu-id="03e79-703">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="03e79-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="03e79-p137">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="03e79-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="03e79-706">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="03e79-707">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="03e79-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-708">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-708">Requirements</span></span>

|<span data-ttu-id="03e79-709">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-709">Requirement</span></span>| <span data-ttu-id="03e79-710">値</span><span class="sxs-lookup"><span data-stu-id="03e79-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-711">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-712">1.3</span><span class="sxs-lookup"><span data-stu-id="03e79-712">1.3</span></span>|
|[<span data-ttu-id="03e79-713">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-714">制限あり</span><span class="sxs-lookup"><span data-stu-id="03e79-714">Restricted</span></span>|
|[<span data-ttu-id="03e79-715">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-716">新規作成</span><span class="sxs-lookup"><span data-stu-id="03e79-716">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="03e79-717">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="03e79-717">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="03e79-718">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="03e79-719">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="03e79-719">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="03e79-720">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="03e79-721">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="03e79-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="03e79-p138">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="03e79-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03e79-725">パラメーター</span><span class="sxs-lookup"><span data-stu-id="03e79-725">Parameters</span></span>

| <span data-ttu-id="03e79-726">名前</span><span class="sxs-lookup"><span data-stu-id="03e79-726">Name</span></span> | <span data-ttu-id="03e79-727">型</span><span class="sxs-lookup"><span data-stu-id="03e79-727">Type</span></span> | <span data-ttu-id="03e79-728">属性</span><span class="sxs-lookup"><span data-stu-id="03e79-728">Attributes</span></span> | <span data-ttu-id="03e79-729">説明</span><span class="sxs-lookup"><span data-stu-id="03e79-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="03e79-730">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="03e79-730">String &#124; Object</span></span>| |<span data-ttu-id="03e79-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="03e79-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="03e79-733">**または**</span><span class="sxs-lookup"><span data-stu-id="03e79-733">**OR**</span></span><br/><span data-ttu-id="03e79-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="03e79-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="03e79-736">String</span><span class="sxs-lookup"><span data-stu-id="03e79-736">String</span></span> | <span data-ttu-id="03e79-737">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-737">&lt;optional&gt;</span></span> | <span data-ttu-id="03e79-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="03e79-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="03e79-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="03e79-741">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-741">&lt;optional&gt;</span></span> | <span data-ttu-id="03e79-742">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="03e79-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="03e79-743">String</span><span class="sxs-lookup"><span data-stu-id="03e79-743">String</span></span> | | <span data-ttu-id="03e79-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="03e79-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="03e79-746">String</span><span class="sxs-lookup"><span data-stu-id="03e79-746">String</span></span> | | <span data-ttu-id="03e79-747">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="03e79-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="03e79-748">文字列</span><span class="sxs-lookup"><span data-stu-id="03e79-748">String</span></span> | | <span data-ttu-id="03e79-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="03e79-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="03e79-751">ブール値</span><span class="sxs-lookup"><span data-stu-id="03e79-751">Boolean</span></span> | | <span data-ttu-id="03e79-p144">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="03e79-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="03e79-754">String</span><span class="sxs-lookup"><span data-stu-id="03e79-754">String</span></span> | | <span data-ttu-id="03e79-p145">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="03e79-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="03e79-758">function</span><span class="sxs-lookup"><span data-stu-id="03e79-758">function</span></span> | <span data-ttu-id="03e79-759">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-759">&lt;optional&gt;</span></span> | <span data-ttu-id="03e79-760">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="03e79-761">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-761">Requirements</span></span>

|<span data-ttu-id="03e79-762">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-762">Requirement</span></span>| <span data-ttu-id="03e79-763">値</span><span class="sxs-lookup"><span data-stu-id="03e79-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-764">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-765">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-765">1.0</span></span>|
|[<span data-ttu-id="03e79-766">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-767">ReadItem</span></span>|
|[<span data-ttu-id="03e79-768">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-769">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="03e79-770">例</span><span class="sxs-lookup"><span data-stu-id="03e79-770">Examples</span></span>

<span data-ttu-id="03e79-771">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="03e79-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="03e79-772">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="03e79-772">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="03e79-773">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="03e79-773">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="03e79-774">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="03e79-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="03e79-775">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="03e79-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="03e79-776">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="03e79-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="03e79-777">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="03e79-777">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="03e79-778">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="03e79-779">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="03e79-779">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="03e79-780">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="03e79-781">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="03e79-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="03e79-p146">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="03e79-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03e79-785">パラメーター</span><span class="sxs-lookup"><span data-stu-id="03e79-785">Parameters</span></span>

| <span data-ttu-id="03e79-786">名前</span><span class="sxs-lookup"><span data-stu-id="03e79-786">Name</span></span> | <span data-ttu-id="03e79-787">型</span><span class="sxs-lookup"><span data-stu-id="03e79-787">Type</span></span> | <span data-ttu-id="03e79-788">属性</span><span class="sxs-lookup"><span data-stu-id="03e79-788">Attributes</span></span> | <span data-ttu-id="03e79-789">説明</span><span class="sxs-lookup"><span data-stu-id="03e79-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="03e79-790">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="03e79-790">String &#124; Object</span></span>| | <span data-ttu-id="03e79-p147">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="03e79-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="03e79-793">**または**</span><span class="sxs-lookup"><span data-stu-id="03e79-793">**OR**</span></span><br/><span data-ttu-id="03e79-p148">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="03e79-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="03e79-796">String</span><span class="sxs-lookup"><span data-stu-id="03e79-796">String</span></span> | <span data-ttu-id="03e79-797">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-797">&lt;optional&gt;</span></span> | <span data-ttu-id="03e79-p149">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="03e79-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="03e79-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="03e79-801">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-801">&lt;optional&gt;</span></span> | <span data-ttu-id="03e79-802">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="03e79-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="03e79-803">String</span><span class="sxs-lookup"><span data-stu-id="03e79-803">String</span></span> | | <span data-ttu-id="03e79-p150">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="03e79-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="03e79-806">String</span><span class="sxs-lookup"><span data-stu-id="03e79-806">String</span></span> | | <span data-ttu-id="03e79-807">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="03e79-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="03e79-808">文字列</span><span class="sxs-lookup"><span data-stu-id="03e79-808">String</span></span> | | <span data-ttu-id="03e79-p151">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="03e79-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="03e79-811">ブール値</span><span class="sxs-lookup"><span data-stu-id="03e79-811">Boolean</span></span> | | <span data-ttu-id="03e79-p152">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="03e79-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="03e79-814">String</span><span class="sxs-lookup"><span data-stu-id="03e79-814">String</span></span> | | <span data-ttu-id="03e79-p153">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="03e79-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="03e79-818">function</span><span class="sxs-lookup"><span data-stu-id="03e79-818">function</span></span> | <span data-ttu-id="03e79-819">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-819">&lt;optional&gt;</span></span> | <span data-ttu-id="03e79-820">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="03e79-821">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-821">Requirements</span></span>

|<span data-ttu-id="03e79-822">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-822">Requirement</span></span>| <span data-ttu-id="03e79-823">値</span><span class="sxs-lookup"><span data-stu-id="03e79-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-824">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-825">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-825">1.0</span></span>|
|[<span data-ttu-id="03e79-826">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-826">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-827">ReadItem</span></span>|
|[<span data-ttu-id="03e79-828">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-828">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-829">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="03e79-830">例</span><span class="sxs-lookup"><span data-stu-id="03e79-830">Examples</span></span>

<span data-ttu-id="03e79-831">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="03e79-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="03e79-832">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="03e79-832">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="03e79-833">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="03e79-833">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="03e79-834">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="03e79-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="03e79-835">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="03e79-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="03e79-836">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="03e79-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="03e79-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="03e79-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="03e79-838">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="03e79-838">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="03e79-839">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="03e79-839">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-840">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-840">Requirements</span></span>

|<span data-ttu-id="03e79-841">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-841">Requirement</span></span>| <span data-ttu-id="03e79-842">値</span><span class="sxs-lookup"><span data-stu-id="03e79-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-843">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-844">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-844">1.0</span></span>|
|[<span data-ttu-id="03e79-845">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-846">ReadItem</span></span>|
|[<span data-ttu-id="03e79-847">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-848">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="03e79-849">戻り値:</span><span class="sxs-lookup"><span data-stu-id="03e79-849">Returns:</span></span>

<span data-ttu-id="03e79-850">型:[Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="03e79-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="03e79-851">例</span><span class="sxs-lookup"><span data-stu-id="03e79-851">Example</span></span>

<span data-ttu-id="03e79-852">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="03e79-852">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="03e79-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="03e79-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="03e79-854">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="03e79-854">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="03e79-855">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="03e79-855">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03e79-856">パラメーター</span><span class="sxs-lookup"><span data-stu-id="03e79-856">Parameters</span></span>

|<span data-ttu-id="03e79-857">名前</span><span class="sxs-lookup"><span data-stu-id="03e79-857">Name</span></span>| <span data-ttu-id="03e79-858">型</span><span class="sxs-lookup"><span data-stu-id="03e79-858">Type</span></span>| <span data-ttu-id="03e79-859">説明</span><span class="sxs-lookup"><span data-stu-id="03e79-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="03e79-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="03e79-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="03e79-861">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="03e79-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03e79-862">Requirements</span><span class="sxs-lookup"><span data-stu-id="03e79-862">Requirements</span></span>

|<span data-ttu-id="03e79-863">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-863">Requirement</span></span>| <span data-ttu-id="03e79-864">値</span><span class="sxs-lookup"><span data-stu-id="03e79-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-865">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-866">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-866">1.0</span></span>|
|[<span data-ttu-id="03e79-867">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-867">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-868">制限あり</span><span class="sxs-lookup"><span data-stu-id="03e79-868">Restricted</span></span>|
|[<span data-ttu-id="03e79-869">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-869">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-870">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="03e79-871">戻り値:</span><span class="sxs-lookup"><span data-stu-id="03e79-871">Returns:</span></span>

<span data-ttu-id="03e79-872">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="03e79-873">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-873">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="03e79-874">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="03e79-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="03e79-875">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="03e79-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="03e79-876">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="03e79-876">Value of `entityType`</span></span> | <span data-ttu-id="03e79-877">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="03e79-877">Type of objects in returned array</span></span> | <span data-ttu-id="03e79-878">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="03e79-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="03e79-879">文字列</span><span class="sxs-lookup"><span data-stu-id="03e79-879">String</span></span> | <span data-ttu-id="03e79-880">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="03e79-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="03e79-881">連絡先</span><span class="sxs-lookup"><span data-stu-id="03e79-881">Contact</span></span> | <span data-ttu-id="03e79-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="03e79-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="03e79-883">文字列</span><span class="sxs-lookup"><span data-stu-id="03e79-883">String</span></span> | <span data-ttu-id="03e79-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="03e79-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="03e79-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="03e79-885">MeetingSuggestion</span></span> | <span data-ttu-id="03e79-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="03e79-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="03e79-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="03e79-887">PhoneNumber</span></span> | <span data-ttu-id="03e79-888">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="03e79-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="03e79-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="03e79-889">TaskSuggestion</span></span> | <span data-ttu-id="03e79-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="03e79-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="03e79-891">文字列</span><span class="sxs-lookup"><span data-stu-id="03e79-891">String</span></span> | <span data-ttu-id="03e79-892">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="03e79-892">**Restricted**</span></span> |

<span data-ttu-id="03e79-893">型:Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="03e79-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="03e79-894">例</span><span class="sxs-lookup"><span data-stu-id="03e79-894">Example</span></span>

<span data-ttu-id="03e79-895">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="03e79-895">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="03e79-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="03e79-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="03e79-897">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="03e79-898">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="03e79-898">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="03e79-899">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03e79-900">パラメーター</span><span class="sxs-lookup"><span data-stu-id="03e79-900">Parameters</span></span>

|<span data-ttu-id="03e79-901">名前</span><span class="sxs-lookup"><span data-stu-id="03e79-901">Name</span></span>| <span data-ttu-id="03e79-902">型</span><span class="sxs-lookup"><span data-stu-id="03e79-902">Type</span></span>| <span data-ttu-id="03e79-903">説明</span><span class="sxs-lookup"><span data-stu-id="03e79-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="03e79-904">String</span><span class="sxs-lookup"><span data-stu-id="03e79-904">String</span></span>|<span data-ttu-id="03e79-905">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="03e79-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03e79-906">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-906">Requirements</span></span>

|<span data-ttu-id="03e79-907">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-907">Requirement</span></span>| <span data-ttu-id="03e79-908">値</span><span class="sxs-lookup"><span data-stu-id="03e79-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-909">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-910">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-910">1.0</span></span>|
|[<span data-ttu-id="03e79-911">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-911">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-912">ReadItem</span></span>|
|[<span data-ttu-id="03e79-913">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-913">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-914">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="03e79-915">戻り値:</span><span class="sxs-lookup"><span data-stu-id="03e79-915">Returns:</span></span>

<span data-ttu-id="03e79-p155">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="03e79-918">型:Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="03e79-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="03e79-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="03e79-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="03e79-920">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="03e79-921">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="03e79-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="03e79-p156">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="03e79-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="03e79-925">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="03e79-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="03e79-926">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="03e79-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="03e79-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="03e79-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-930">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-930">Requirements</span></span>

|<span data-ttu-id="03e79-931">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-931">Requirement</span></span>| <span data-ttu-id="03e79-932">値</span><span class="sxs-lookup"><span data-stu-id="03e79-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-933">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-934">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-934">1.0</span></span>|
|[<span data-ttu-id="03e79-935">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-936">ReadItem</span></span>|
|[<span data-ttu-id="03e79-937">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-938">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="03e79-939">戻り値:</span><span class="sxs-lookup"><span data-stu-id="03e79-939">Returns:</span></span>

<span data-ttu-id="03e79-p158">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="03e79-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="03e79-942">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="03e79-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="03e79-943">Object</span><span class="sxs-lookup"><span data-stu-id="03e79-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="03e79-944">例</span><span class="sxs-lookup"><span data-stu-id="03e79-944">Example</span></span>

<span data-ttu-id="03e79-945">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="03e79-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="03e79-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="03e79-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="03e79-947">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="03e79-948">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="03e79-948">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="03e79-949">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="03e79-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="03e79-p159">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="03e79-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03e79-952">パラメーター</span><span class="sxs-lookup"><span data-stu-id="03e79-952">Parameters</span></span>

|<span data-ttu-id="03e79-953">名前</span><span class="sxs-lookup"><span data-stu-id="03e79-953">Name</span></span>| <span data-ttu-id="03e79-954">型</span><span class="sxs-lookup"><span data-stu-id="03e79-954">Type</span></span>| <span data-ttu-id="03e79-955">説明</span><span class="sxs-lookup"><span data-stu-id="03e79-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="03e79-956">String</span><span class="sxs-lookup"><span data-stu-id="03e79-956">String</span></span>|<span data-ttu-id="03e79-957">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="03e79-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03e79-958">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-958">Requirements</span></span>

|<span data-ttu-id="03e79-959">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-959">Requirement</span></span>| <span data-ttu-id="03e79-960">値</span><span class="sxs-lookup"><span data-stu-id="03e79-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-961">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-961">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-962">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-962">1.0</span></span>|
|[<span data-ttu-id="03e79-963">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-963">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-964">ReadItem</span></span>|
|[<span data-ttu-id="03e79-965">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-965">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-966">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="03e79-967">戻り値:</span><span class="sxs-lookup"><span data-stu-id="03e79-967">Returns:</span></span>

<span data-ttu-id="03e79-968">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="03e79-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="03e79-969">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="03e79-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="03e79-970">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="03e79-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="03e79-971">例</span><span class="sxs-lookup"><span data-stu-id="03e79-971">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="03e79-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="03e79-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="03e79-973">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="03e79-p160">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03e79-976">パラメーター</span><span class="sxs-lookup"><span data-stu-id="03e79-976">Parameters</span></span>

|<span data-ttu-id="03e79-977">名前</span><span class="sxs-lookup"><span data-stu-id="03e79-977">Name</span></span>| <span data-ttu-id="03e79-978">型</span><span class="sxs-lookup"><span data-stu-id="03e79-978">Type</span></span>| <span data-ttu-id="03e79-979">属性</span><span class="sxs-lookup"><span data-stu-id="03e79-979">Attributes</span></span>| <span data-ttu-id="03e79-980">説明</span><span class="sxs-lookup"><span data-stu-id="03e79-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="03e79-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="03e79-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="03e79-p161">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="03e79-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="03e79-985">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="03e79-985">Object</span></span>| <span data-ttu-id="03e79-986">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-986">&lt;optional&gt;</span></span>|<span data-ttu-id="03e79-987">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="03e79-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="03e79-988">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="03e79-988">Object</span></span>| <span data-ttu-id="03e79-989">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-989">&lt;optional&gt;</span></span>|<span data-ttu-id="03e79-990">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="03e79-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="03e79-991">function</span><span class="sxs-lookup"><span data-stu-id="03e79-991">function</span></span>||<span data-ttu-id="03e79-992">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="03e79-993">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="03e79-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="03e79-994">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="03e79-994">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03e79-995">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-995">Requirements</span></span>

|<span data-ttu-id="03e79-996">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-996">Requirement</span></span>| <span data-ttu-id="03e79-997">値</span><span class="sxs-lookup"><span data-stu-id="03e79-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-998">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-998">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-999">1.2</span><span class="sxs-lookup"><span data-stu-id="03e79-999">1.2</span></span>|
|[<span data-ttu-id="03e79-1000">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-1000">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="03e79-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="03e79-1002">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-1002">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-1003">作成</span><span class="sxs-lookup"><span data-stu-id="03e79-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="03e79-1004">戻り値:</span><span class="sxs-lookup"><span data-stu-id="03e79-1004">Returns:</span></span>

<span data-ttu-id="03e79-1005">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="03e79-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="03e79-1006">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="03e79-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="03e79-1007">String</span><span class="sxs-lookup"><span data-stu-id="03e79-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="03e79-1008">例</span><span class="sxs-lookup"><span data-stu-id="03e79-1008">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="03e79-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="03e79-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="03e79-1010">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="03e79-1010">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="03e79-1011">強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-1011">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="03e79-1012">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="03e79-1012">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-1013">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-1013">Requirements</span></span>

|<span data-ttu-id="03e79-1014">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-1014">Requirement</span></span>| <span data-ttu-id="03e79-1015">値</span><span class="sxs-lookup"><span data-stu-id="03e79-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-1016">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-1016">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="03e79-1017">1.6</span></span> |
|[<span data-ttu-id="03e79-1018">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-1018">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-1019">ReadItem</span></span>|
|[<span data-ttu-id="03e79-1020">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-1020">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-1021">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="03e79-1022">戻り値:</span><span class="sxs-lookup"><span data-stu-id="03e79-1022">Returns:</span></span>

<span data-ttu-id="03e79-1023">型:[Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="03e79-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="03e79-1024">例</span><span class="sxs-lookup"><span data-stu-id="03e79-1024">Example</span></span>

<span data-ttu-id="03e79-1025">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="03e79-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="03e79-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="03e79-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="03e79-p164">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="03e79-1029">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="03e79-1029">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="03e79-p165">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="03e79-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="03e79-1033">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="03e79-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="03e79-1034">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="03e79-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="03e79-p166">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="03e79-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="03e79-1038">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-1038">Requirements</span></span>

|<span data-ttu-id="03e79-1039">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-1039">Requirement</span></span>| <span data-ttu-id="03e79-1040">値</span><span class="sxs-lookup"><span data-stu-id="03e79-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-1041">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="03e79-1042">1.6</span></span> |
|[<span data-ttu-id="03e79-1043">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-1043">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-1044">ReadItem</span></span>|
|[<span data-ttu-id="03e79-1045">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-1045">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-1046">読み取り</span><span class="sxs-lookup"><span data-stu-id="03e79-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="03e79-1047">戻り値:</span><span class="sxs-lookup"><span data-stu-id="03e79-1047">Returns:</span></span>

<span data-ttu-id="03e79-p167">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="03e79-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="03e79-1050">例</span><span class="sxs-lookup"><span data-stu-id="03e79-1050">Example</span></span>

<span data-ttu-id="03e79-1051">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="03e79-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="03e79-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="03e79-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="03e79-1053">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="03e79-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="03e79-p168">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="03e79-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03e79-1057">パラメーター</span><span class="sxs-lookup"><span data-stu-id="03e79-1057">Parameters</span></span>

|<span data-ttu-id="03e79-1058">名前</span><span class="sxs-lookup"><span data-stu-id="03e79-1058">Name</span></span>| <span data-ttu-id="03e79-1059">型</span><span class="sxs-lookup"><span data-stu-id="03e79-1059">Type</span></span>| <span data-ttu-id="03e79-1060">属性</span><span class="sxs-lookup"><span data-stu-id="03e79-1060">Attributes</span></span>| <span data-ttu-id="03e79-1061">説明</span><span class="sxs-lookup"><span data-stu-id="03e79-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="03e79-1062">function</span><span class="sxs-lookup"><span data-stu-id="03e79-1062">function</span></span>||<span data-ttu-id="03e79-1063">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="03e79-1064">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="03e79-1065">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="03e79-1065">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="03e79-1066">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="03e79-1066">Object</span></span>| <span data-ttu-id="03e79-1067">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="03e79-1068">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="03e79-1068">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="03e79-1069">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="03e79-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03e79-1070">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-1070">Requirements</span></span>

|<span data-ttu-id="03e79-1071">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-1071">Requirement</span></span>| <span data-ttu-id="03e79-1072">値</span><span class="sxs-lookup"><span data-stu-id="03e79-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-1073">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-1073">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="03e79-1074">1.0</span></span>|
|[<span data-ttu-id="03e79-1075">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-1075">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03e79-1076">ReadItem</span></span>|
|[<span data-ttu-id="03e79-1077">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-1077">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-1078">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="03e79-1078">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03e79-1079">例</span><span class="sxs-lookup"><span data-stu-id="03e79-1079">Example</span></span>

<span data-ttu-id="03e79-p171">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="03e79-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="03e79-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="03e79-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="03e79-1084">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="03e79-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="03e79-p172">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="03e79-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03e79-1089">パラメーター</span><span class="sxs-lookup"><span data-stu-id="03e79-1089">Parameters</span></span>

|<span data-ttu-id="03e79-1090">名前</span><span class="sxs-lookup"><span data-stu-id="03e79-1090">Name</span></span>| <span data-ttu-id="03e79-1091">型</span><span class="sxs-lookup"><span data-stu-id="03e79-1091">Type</span></span>| <span data-ttu-id="03e79-1092">属性</span><span class="sxs-lookup"><span data-stu-id="03e79-1092">Attributes</span></span>| <span data-ttu-id="03e79-1093">説明</span><span class="sxs-lookup"><span data-stu-id="03e79-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="03e79-1094">String</span><span class="sxs-lookup"><span data-stu-id="03e79-1094">String</span></span>||<span data-ttu-id="03e79-1095">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="03e79-1095">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="03e79-1096">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="03e79-1096">Object</span></span>| <span data-ttu-id="03e79-1097">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="03e79-1098">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="03e79-1098">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="03e79-1099">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="03e79-1099">Object</span></span>| <span data-ttu-id="03e79-1100">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-1100">&lt;optional&gt;</span></span>|<span data-ttu-id="03e79-1101">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="03e79-1101">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="03e79-1102">function</span><span class="sxs-lookup"><span data-stu-id="03e79-1102">function</span></span>| <span data-ttu-id="03e79-1103">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-1103">&lt;optional&gt;</span></span>|<span data-ttu-id="03e79-1104">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-1104">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="03e79-1105">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="03e79-1105">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="03e79-1106">エラー</span><span class="sxs-lookup"><span data-stu-id="03e79-1106">Errors</span></span>

| <span data-ttu-id="03e79-1107">エラー コード</span><span class="sxs-lookup"><span data-stu-id="03e79-1107">Error code</span></span> | <span data-ttu-id="03e79-1108">説明</span><span class="sxs-lookup"><span data-stu-id="03e79-1108">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="03e79-1109">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="03e79-1109">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="03e79-1110">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-1110">Requirements</span></span>

|<span data-ttu-id="03e79-1111">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-1111">Requirement</span></span>| <span data-ttu-id="03e79-1112">値</span><span class="sxs-lookup"><span data-stu-id="03e79-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-1113">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-1114">1.1</span><span class="sxs-lookup"><span data-stu-id="03e79-1114">1.1</span></span>|
|[<span data-ttu-id="03e79-1115">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-1115">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-1116">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="03e79-1116">ReadWriteItem</span></span>|
|[<span data-ttu-id="03e79-1117">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-1117">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-1118">作成</span><span class="sxs-lookup"><span data-stu-id="03e79-1118">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="03e79-1119">例</span><span class="sxs-lookup"><span data-stu-id="03e79-1119">Example</span></span>

<span data-ttu-id="03e79-1120">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="03e79-1120">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="03e79-1121">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="03e79-1121">saveAsync([options], callback)</span></span>

<span data-ttu-id="03e79-1122">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="03e79-1122">Asynchronously saves an item.</span></span>

<span data-ttu-id="03e79-p173">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-p173">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="03e79-1126">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="03e79-1126">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="03e79-1127">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-1127">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="03e79-p175">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="03e79-1131">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="03e79-1131">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="03e79-1132">Outlook for Mac では、会議の保存はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="03e79-1132">Outlook for Mac does not support saving a meeting.</span></span> <span data-ttu-id="03e79-1133">新規`saveAsync`作成モードで会議から呼び出された場合、メソッドは失敗します。</span><span class="sxs-lookup"><span data-stu-id="03e79-1133">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="03e79-1134">回避策については[、「OFFICE JS API を使用して Outlook For Mac で会議を下書きとして保存できません](https://support.microsoft.com/help/4505745)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="03e79-1134">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="03e79-1135">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03e79-1136">パラメーター</span><span class="sxs-lookup"><span data-stu-id="03e79-1136">Parameters</span></span>

|<span data-ttu-id="03e79-1137">名前</span><span class="sxs-lookup"><span data-stu-id="03e79-1137">Name</span></span>| <span data-ttu-id="03e79-1138">型</span><span class="sxs-lookup"><span data-stu-id="03e79-1138">Type</span></span>| <span data-ttu-id="03e79-1139">属性</span><span class="sxs-lookup"><span data-stu-id="03e79-1139">Attributes</span></span>| <span data-ttu-id="03e79-1140">説明</span><span class="sxs-lookup"><span data-stu-id="03e79-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="03e79-1141">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="03e79-1141">Object</span></span>| <span data-ttu-id="03e79-1142">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="03e79-1143">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="03e79-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="03e79-1144">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="03e79-1144">Object</span></span>| <span data-ttu-id="03e79-1145">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="03e79-1146">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="03e79-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="03e79-1147">関数</span><span class="sxs-lookup"><span data-stu-id="03e79-1147">function</span></span>||<span data-ttu-id="03e79-1148">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="03e79-1149">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-1149">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03e79-1150">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-1150">Requirements</span></span>

|<span data-ttu-id="03e79-1151">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-1151">Requirement</span></span>| <span data-ttu-id="03e79-1152">値</span><span class="sxs-lookup"><span data-stu-id="03e79-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-1153">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-1153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="03e79-1154">1.3</span></span>|
|[<span data-ttu-id="03e79-1155">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-1155">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="03e79-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="03e79-1157">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-1157">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-1158">作成</span><span class="sxs-lookup"><span data-stu-id="03e79-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="03e79-1159">例</span><span class="sxs-lookup"><span data-stu-id="03e79-1159">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="03e79-p177">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="03e79-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="03e79-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="03e79-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="03e79-1163">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="03e79-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="03e79-p178">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="03e79-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03e79-1167">パラメーター</span><span class="sxs-lookup"><span data-stu-id="03e79-1167">Parameters</span></span>

|<span data-ttu-id="03e79-1168">名前</span><span class="sxs-lookup"><span data-stu-id="03e79-1168">Name</span></span>| <span data-ttu-id="03e79-1169">型</span><span class="sxs-lookup"><span data-stu-id="03e79-1169">Type</span></span>| <span data-ttu-id="03e79-1170">属性</span><span class="sxs-lookup"><span data-stu-id="03e79-1170">Attributes</span></span>| <span data-ttu-id="03e79-1171">説明</span><span class="sxs-lookup"><span data-stu-id="03e79-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="03e79-1172">String</span><span class="sxs-lookup"><span data-stu-id="03e79-1172">String</span></span>||<span data-ttu-id="03e79-p179">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="03e79-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="03e79-1176">Object</span><span class="sxs-lookup"><span data-stu-id="03e79-1176">Object</span></span>| <span data-ttu-id="03e79-1177">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="03e79-1178">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="03e79-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="03e79-1179">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="03e79-1179">Object</span></span>| <span data-ttu-id="03e79-1180">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="03e79-1181">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="03e79-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="03e79-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="03e79-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="03e79-1183">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="03e79-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="03e79-p180">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-p180">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="03e79-p181">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-p181">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="03e79-1188">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="03e79-1189">function</span><span class="sxs-lookup"><span data-stu-id="03e79-1189">function</span></span>||<span data-ttu-id="03e79-1190">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="03e79-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="03e79-1191">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-1191">Requirements</span></span>

|<span data-ttu-id="03e79-1192">要件</span><span class="sxs-lookup"><span data-stu-id="03e79-1192">Requirement</span></span>| <span data-ttu-id="03e79-1193">値</span><span class="sxs-lookup"><span data-stu-id="03e79-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="03e79-1194">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="03e79-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03e79-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="03e79-1195">1.2</span></span>|
|[<span data-ttu-id="03e79-1196">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="03e79-1196">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03e79-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="03e79-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="03e79-1198">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="03e79-1198">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03e79-1199">作成</span><span class="sxs-lookup"><span data-stu-id="03e79-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="03e79-1200">例</span><span class="sxs-lookup"><span data-stu-id="03e79-1200">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
