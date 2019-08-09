---
title: Office. メールボックス-要件セット1.2
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 536c8b7bece6df6f9609406f3eccc50b330d7925
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268692"
---
# <a name="item"></a><span data-ttu-id="0f8b7-102">item</span><span class="sxs-lookup"><span data-stu-id="0f8b7-102">item</span></span>

### <span data-ttu-id="0f8b7-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="0f8b7-p102">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-107">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-107">Requirements</span></span>

|<span data-ttu-id="0f8b7-108">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-108">Requirement</span></span>| <span data-ttu-id="0f8b7-109">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-111">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-111">1.0</span></span>|
|[<span data-ttu-id="0f8b7-112">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-113">制限あり</span><span class="sxs-lookup"><span data-stu-id="0f8b7-113">Restricted</span></span>|
|[<span data-ttu-id="0f8b7-114">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-115">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f8b7-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="0f8b7-116">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="0f8b7-116">Members and methods</span></span>

| <span data-ttu-id="0f8b7-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="0f8b7-117">Member</span></span> | <span data-ttu-id="0f8b7-118">種類</span><span class="sxs-lookup"><span data-stu-id="0f8b7-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="0f8b7-119">attachments</span><span class="sxs-lookup"><span data-stu-id="0f8b7-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="0f8b7-120">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-120">Member</span></span> |
| [<span data-ttu-id="0f8b7-121">bcc</span><span class="sxs-lookup"><span data-stu-id="0f8b7-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="0f8b7-122">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-122">Member</span></span> |
| [<span data-ttu-id="0f8b7-123">body</span><span class="sxs-lookup"><span data-stu-id="0f8b7-123">body</span></span>](#body-body) | <span data-ttu-id="0f8b7-124">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-124">Member</span></span> |
| [<span data-ttu-id="0f8b7-125">cc</span><span class="sxs-lookup"><span data-stu-id="0f8b7-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0f8b7-126">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-126">Member</span></span> |
| [<span data-ttu-id="0f8b7-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="0f8b7-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="0f8b7-128">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-128">Member</span></span> |
| [<span data-ttu-id="0f8b7-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="0f8b7-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="0f8b7-130">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-130">Member</span></span> |
| [<span data-ttu-id="0f8b7-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="0f8b7-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="0f8b7-132">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-132">Member</span></span> |
| [<span data-ttu-id="0f8b7-133">end</span><span class="sxs-lookup"><span data-stu-id="0f8b7-133">end</span></span>](#end-datetime) | <span data-ttu-id="0f8b7-134">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-134">Member</span></span> |
| [<span data-ttu-id="0f8b7-135">from</span><span class="sxs-lookup"><span data-stu-id="0f8b7-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="0f8b7-136">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-136">Member</span></span> |
| [<span data-ttu-id="0f8b7-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="0f8b7-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="0f8b7-138">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-138">Member</span></span> |
| [<span data-ttu-id="0f8b7-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="0f8b7-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="0f8b7-140">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-140">Member</span></span> |
| [<span data-ttu-id="0f8b7-141">itemId</span><span class="sxs-lookup"><span data-stu-id="0f8b7-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="0f8b7-142">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-142">Member</span></span> |
| [<span data-ttu-id="0f8b7-143">itemType</span><span class="sxs-lookup"><span data-stu-id="0f8b7-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="0f8b7-144">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-144">Member</span></span> |
| [<span data-ttu-id="0f8b7-145">location</span><span class="sxs-lookup"><span data-stu-id="0f8b7-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="0f8b7-146">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-146">Member</span></span> |
| [<span data-ttu-id="0f8b7-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="0f8b7-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="0f8b7-148">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-148">Member</span></span> |
| [<span data-ttu-id="0f8b7-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="0f8b7-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0f8b7-150">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-150">Member</span></span> |
| [<span data-ttu-id="0f8b7-151">organizer</span><span class="sxs-lookup"><span data-stu-id="0f8b7-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="0f8b7-152">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-152">Member</span></span> |
| [<span data-ttu-id="0f8b7-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="0f8b7-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0f8b7-154">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-154">Member</span></span> |
| [<span data-ttu-id="0f8b7-155">sender</span><span class="sxs-lookup"><span data-stu-id="0f8b7-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="0f8b7-156">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-156">Member</span></span> |
| [<span data-ttu-id="0f8b7-157">start</span><span class="sxs-lookup"><span data-stu-id="0f8b7-157">start</span></span>](#start-datetime) | <span data-ttu-id="0f8b7-158">Member</span><span class="sxs-lookup"><span data-stu-id="0f8b7-158">Member</span></span> |
| [<span data-ttu-id="0f8b7-159">subject</span><span class="sxs-lookup"><span data-stu-id="0f8b7-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="0f8b7-160">メンバー</span><span class="sxs-lookup"><span data-stu-id="0f8b7-160">Member</span></span> |
| [<span data-ttu-id="0f8b7-161">to</span><span class="sxs-lookup"><span data-stu-id="0f8b7-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0f8b7-162">メンバー</span><span class="sxs-lookup"><span data-stu-id="0f8b7-162">Member</span></span> |
| [<span data-ttu-id="0f8b7-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0f8b7-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="0f8b7-164">メソッド</span><span class="sxs-lookup"><span data-stu-id="0f8b7-164">Method</span></span> |
| [<span data-ttu-id="0f8b7-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0f8b7-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="0f8b7-166">メソッド</span><span class="sxs-lookup"><span data-stu-id="0f8b7-166">Method</span></span> |
| [<span data-ttu-id="0f8b7-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="0f8b7-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="0f8b7-168">メソッド</span><span class="sxs-lookup"><span data-stu-id="0f8b7-168">Method</span></span> |
| [<span data-ttu-id="0f8b7-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="0f8b7-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="0f8b7-170">メソッド</span><span class="sxs-lookup"><span data-stu-id="0f8b7-170">Method</span></span> |
| [<span data-ttu-id="0f8b7-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="0f8b7-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="0f8b7-172">メソッド</span><span class="sxs-lookup"><span data-stu-id="0f8b7-172">Method</span></span> |
| [<span data-ttu-id="0f8b7-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="0f8b7-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="0f8b7-174">メソッド</span><span class="sxs-lookup"><span data-stu-id="0f8b7-174">Method</span></span> |
| [<span data-ttu-id="0f8b7-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="0f8b7-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="0f8b7-176">メソッド</span><span class="sxs-lookup"><span data-stu-id="0f8b7-176">Method</span></span> |
| [<span data-ttu-id="0f8b7-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="0f8b7-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="0f8b7-178">メソッド</span><span class="sxs-lookup"><span data-stu-id="0f8b7-178">Method</span></span> |
| [<span data-ttu-id="0f8b7-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="0f8b7-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="0f8b7-180">メソッド</span><span class="sxs-lookup"><span data-stu-id="0f8b7-180">Method</span></span> |
| [<span data-ttu-id="0f8b7-181">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="0f8b7-181">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="0f8b7-182">メソッド</span><span class="sxs-lookup"><span data-stu-id="0f8b7-182">Method</span></span> |
| [<span data-ttu-id="0f8b7-183">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="0f8b7-183">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="0f8b7-184">メソッド</span><span class="sxs-lookup"><span data-stu-id="0f8b7-184">Method</span></span> |
| [<span data-ttu-id="0f8b7-185">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0f8b7-185">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="0f8b7-186">メソッド</span><span class="sxs-lookup"><span data-stu-id="0f8b7-186">Method</span></span> |
| [<span data-ttu-id="0f8b7-187">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="0f8b7-187">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="0f8b7-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="0f8b7-188">Method</span></span> |

### <a name="example"></a><span data-ttu-id="0f8b7-189">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-189">Example</span></span>

<span data-ttu-id="0f8b7-190">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-190">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="0f8b7-191">メンバー</span><span class="sxs-lookup"><span data-stu-id="0f8b7-191">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-12"></a><span data-ttu-id="0f8b7-192">添付ファイル: <[Attachmentdetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="0f8b7-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

<span data-ttu-id="0f8b7-p103">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0f8b7-195">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-195">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="0f8b7-196">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-196">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="0f8b7-197">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-197">Type</span></span>

*   <span data-ttu-id="0f8b7-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="0f8b7-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-199">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-199">Requirements</span></span>

|<span data-ttu-id="0f8b7-200">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-200">Requirement</span></span>| <span data-ttu-id="0f8b7-201">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-201">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-202">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-202">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-203">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-203">1.0</span></span>|
|[<span data-ttu-id="0f8b7-204">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-204">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-205">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-205">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-206">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-206">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-207">読み取り</span><span class="sxs-lookup"><span data-stu-id="0f8b7-207">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-208">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-208">Example</span></span>

<span data-ttu-id="0f8b7-209">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-209">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="0f8b7-210">bcc:[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0f8b7-211">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-211">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="0f8b7-212">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-212">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0f8b7-213">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-213">Type</span></span>

*   [<span data-ttu-id="0f8b7-214">受信者</span><span class="sxs-lookup"><span data-stu-id="0f8b7-214">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="0f8b7-215">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-215">Requirements</span></span>

|<span data-ttu-id="0f8b7-216">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-216">Requirement</span></span>| <span data-ttu-id="0f8b7-217">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-218">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-219">1.1</span><span class="sxs-lookup"><span data-stu-id="0f8b7-219">1.1</span></span>|
|[<span data-ttu-id="0f8b7-220">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-221">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-222">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-223">作成</span><span class="sxs-lookup"><span data-stu-id="0f8b7-223">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-224">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-224">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-12"></a><span data-ttu-id="0f8b7-225">本文:[本文](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-225">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0f8b7-226">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-226">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="0f8b7-227">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-227">Type</span></span>

*   [<span data-ttu-id="0f8b7-228">Body</span><span class="sxs-lookup"><span data-stu-id="0f8b7-228">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="0f8b7-229">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-229">Requirements</span></span>

|<span data-ttu-id="0f8b7-230">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-230">Requirement</span></span>| <span data-ttu-id="0f8b7-231">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-231">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-232">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-232">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-233">1.1</span><span class="sxs-lookup"><span data-stu-id="0f8b7-233">1.1</span></span>|
|[<span data-ttu-id="0f8b7-234">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-234">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-235">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-235">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-236">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-237">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f8b7-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-238">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-238">Example</span></span>

<span data-ttu-id="0f8b7-239">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-239">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="0f8b7-240">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-240">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="0f8b7-241">cc: <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-241">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0f8b7-242">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-242">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="0f8b7-243">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-243">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0f8b7-244">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-244">Read mode</span></span>

<span data-ttu-id="0f8b7-p107">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="0f8b7-247">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-247">Compose mode</span></span>

<span data-ttu-id="0f8b7-248">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-248">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0f8b7-249">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-249">Type</span></span>

*   <span data-ttu-id="0f8b7-250">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-250">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-251">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-251">Requirements</span></span>

|<span data-ttu-id="0f8b7-252">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-252">Requirement</span></span>| <span data-ttu-id="0f8b7-253">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-254">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-254">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-255">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-255">1.0</span></span>|
|[<span data-ttu-id="0f8b7-256">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-256">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-257">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-258">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-258">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-259">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f8b7-259">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="0f8b7-260">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-260">(nullable) conversationId: String</span></span>

<span data-ttu-id="0f8b7-261">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-261">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="0f8b7-p108">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="0f8b7-p109">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="0f8b7-266">Type</span><span class="sxs-lookup"><span data-stu-id="0f8b7-266">Type</span></span>

*   <span data-ttu-id="0f8b7-267">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-267">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-268">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-268">Requirements</span></span>

|<span data-ttu-id="0f8b7-269">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-269">Requirement</span></span>| <span data-ttu-id="0f8b7-270">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-271">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-272">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-272">1.0</span></span>|
|[<span data-ttu-id="0f8b7-273">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-274">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-275">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-276">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f8b7-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-277">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-277">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="0f8b7-278">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="0f8b7-278">dateTimeCreated: Date</span></span>

<span data-ttu-id="0f8b7-p110">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0f8b7-281">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-281">Type</span></span>

*   <span data-ttu-id="0f8b7-282">日付</span><span class="sxs-lookup"><span data-stu-id="0f8b7-282">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-283">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-283">Requirements</span></span>

|<span data-ttu-id="0f8b7-284">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-284">Requirement</span></span>| <span data-ttu-id="0f8b7-285">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-285">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-286">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-286">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-287">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-287">1.0</span></span>|
|[<span data-ttu-id="0f8b7-288">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-288">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-289">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-289">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-290">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-290">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-291">読み取り</span><span class="sxs-lookup"><span data-stu-id="0f8b7-291">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-292">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-292">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="0f8b7-293">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="0f8b7-293">dateTimeModified: Date</span></span>

<span data-ttu-id="0f8b7-294">アイテムが最後に変更された日時を取得します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-294">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="0f8b7-295">Read mode only.</span><span class="sxs-lookup"><span data-stu-id="0f8b7-295">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0f8b7-296">このメンバーは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-296">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="0f8b7-297">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-297">Type</span></span>

*   <span data-ttu-id="0f8b7-298">日付</span><span class="sxs-lookup"><span data-stu-id="0f8b7-298">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-299">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-299">Requirements</span></span>

|<span data-ttu-id="0f8b7-300">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-300">Requirement</span></span>| <span data-ttu-id="0f8b7-301">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-302">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-303">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-303">1.0</span></span>|
|[<span data-ttu-id="0f8b7-304">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-305">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-306">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-307">読み取り</span><span class="sxs-lookup"><span data-stu-id="0f8b7-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-308">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-308">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="0f8b7-309">終了: 日付 |[時間](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-309">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0f8b7-310">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-310">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="0f8b7-p112">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0f8b7-313">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-313">Read mode</span></span>

<span data-ttu-id="0f8b7-314">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-314">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="0f8b7-315">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-315">Compose mode</span></span>

<span data-ttu-id="0f8b7-316">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-316">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="0f8b7-317">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-317">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="0f8b7-318">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-318">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="0f8b7-319">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-319">Type</span></span>

*   <span data-ttu-id="0f8b7-320">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-320">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-321">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-321">Requirements</span></span>

|<span data-ttu-id="0f8b7-322">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-322">Requirement</span></span>| <span data-ttu-id="0f8b7-323">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-324">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-325">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-325">1.0</span></span>|
|[<span data-ttu-id="0f8b7-326">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-326">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-327">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-328">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-328">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-329">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f8b7-329">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="0f8b7-330">from: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-330">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0f8b7-p113">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="0f8b7-p114">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="0f8b7-335">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-335">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="0f8b7-336">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-336">Type</span></span>

*   [<span data-ttu-id="0f8b7-337">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0f8b7-337">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="0f8b7-338">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-338">Requirements</span></span>

|<span data-ttu-id="0f8b7-339">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-339">Requirement</span></span>| <span data-ttu-id="0f8b7-340">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-341">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-342">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-342">1.0</span></span>|
|[<span data-ttu-id="0f8b7-343">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-344">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-345">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-346">読み取り</span><span class="sxs-lookup"><span data-stu-id="0f8b7-346">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-347">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-347">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="0f8b7-348">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-348">internetMessageId: String</span></span>

<span data-ttu-id="0f8b7-p115">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0f8b7-351">Type</span><span class="sxs-lookup"><span data-stu-id="0f8b7-351">Type</span></span>

*   <span data-ttu-id="0f8b7-352">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-352">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-353">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-353">Requirements</span></span>

|<span data-ttu-id="0f8b7-354">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-354">Requirement</span></span>| <span data-ttu-id="0f8b7-355">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-356">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-357">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-357">1.0</span></span>|
|[<span data-ttu-id="0f8b7-358">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-359">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-360">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-361">読み取り</span><span class="sxs-lookup"><span data-stu-id="0f8b7-361">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-362">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-362">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="0f8b7-363">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-363">itemClass: String</span></span>

<span data-ttu-id="0f8b7-p116">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="0f8b7-p117">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="0f8b7-368">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-368">Type</span></span> | <span data-ttu-id="0f8b7-369">説明</span><span class="sxs-lookup"><span data-stu-id="0f8b7-369">Description</span></span> | <span data-ttu-id="0f8b7-370">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="0f8b7-370">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="0f8b7-371">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="0f8b7-371">Appointment items</span></span> | <span data-ttu-id="0f8b7-372">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-372">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="0f8b7-373">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="0f8b7-373">Message items</span></span> | <span data-ttu-id="0f8b7-374">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-374">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="0f8b7-375">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-375">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="0f8b7-376">Type</span><span class="sxs-lookup"><span data-stu-id="0f8b7-376">Type</span></span>

*   <span data-ttu-id="0f8b7-377">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-377">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-378">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-378">Requirements</span></span>

|<span data-ttu-id="0f8b7-379">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-379">Requirement</span></span>| <span data-ttu-id="0f8b7-380">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-381">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-382">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-382">1.0</span></span>|
|[<span data-ttu-id="0f8b7-383">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-383">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-384">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-385">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-385">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-386">読み取り</span><span class="sxs-lookup"><span data-stu-id="0f8b7-386">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-387">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-387">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="0f8b7-388">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-388">(nullable) itemId: String</span></span>

<span data-ttu-id="0f8b7-389">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-389">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="0f8b7-390">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-390">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0f8b7-391">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-391">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="0f8b7-392">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-392">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="0f8b7-393">この値を使用して REST API を呼び出す前に、を`Office.context.mailbox.convertToRestId`使用して変換する必要があります。これは、要件セット1.3 から開始できます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-393">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="0f8b7-394">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-394">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="0f8b7-395">Type</span><span class="sxs-lookup"><span data-stu-id="0f8b7-395">Type</span></span>

*   <span data-ttu-id="0f8b7-396">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-396">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-397">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-397">Requirements</span></span>

|<span data-ttu-id="0f8b7-398">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-398">Requirement</span></span>| <span data-ttu-id="0f8b7-399">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-399">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-400">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-401">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-401">1.0</span></span>|
|[<span data-ttu-id="0f8b7-402">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-402">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-403">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-404">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-404">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-405">読み取り</span><span class="sxs-lookup"><span data-stu-id="0f8b7-405">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-406">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-406">Example</span></span>

<span data-ttu-id="0f8b7-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-12"></a><span data-ttu-id="0f8b7-409">itemType: [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-409">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0f8b7-410">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-410">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="0f8b7-411">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-411">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="0f8b7-412">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-412">Type</span></span>

*   [<span data-ttu-id="0f8b7-413">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="0f8b7-413">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="0f8b7-414">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-414">Requirements</span></span>

|<span data-ttu-id="0f8b7-415">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-415">Requirement</span></span>| <span data-ttu-id="0f8b7-416">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-416">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-417">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-417">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-418">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-418">1.0</span></span>|
|[<span data-ttu-id="0f8b7-419">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-419">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-420">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-420">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-421">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-421">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-422">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f8b7-422">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-423">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-423">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-12"></a><span data-ttu-id="0f8b7-424">場所: String |[場所](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-424">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0f8b7-425">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-425">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0f8b7-426">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-426">Read mode</span></span>

<span data-ttu-id="0f8b7-427">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-427">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="0f8b7-428">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-428">Compose mode</span></span>

<span data-ttu-id="0f8b7-429">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-429">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0f8b7-430">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-430">Type</span></span>

*   <span data-ttu-id="0f8b7-431">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-431">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-432">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-432">Requirements</span></span>

|<span data-ttu-id="0f8b7-433">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-433">Requirement</span></span>| <span data-ttu-id="0f8b7-434">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-434">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-435">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-435">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-436">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-436">1.0</span></span>|
|[<span data-ttu-id="0f8b7-437">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-437">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-438">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-438">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-439">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-439">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-440">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f8b7-440">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="0f8b7-441">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-441">normalizedSubject: String</span></span>

<span data-ttu-id="0f8b7-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="0f8b7-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="0f8b7-446">Type</span><span class="sxs-lookup"><span data-stu-id="0f8b7-446">Type</span></span>

*   <span data-ttu-id="0f8b7-447">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-447">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-448">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-448">Requirements</span></span>

|<span data-ttu-id="0f8b7-449">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-449">Requirement</span></span>| <span data-ttu-id="0f8b7-450">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-451">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-451">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-452">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-452">1.0</span></span>|
|[<span data-ttu-id="0f8b7-453">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-453">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-454">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-455">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-455">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-456">読み取り</span><span class="sxs-lookup"><span data-stu-id="0f8b7-456">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-457">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-457">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="0f8b7-458">任意出席者: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-458">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0f8b7-459">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-459">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="0f8b7-460">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-460">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0f8b7-461">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-461">Read mode</span></span>

<span data-ttu-id="0f8b7-462">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-462">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="0f8b7-463">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-463">Compose mode</span></span>

<span data-ttu-id="0f8b7-464">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-464">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0f8b7-465">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-465">Type</span></span>

*   <span data-ttu-id="0f8b7-466">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-466">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-467">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-467">Requirements</span></span>

|<span data-ttu-id="0f8b7-468">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-468">Requirement</span></span>| <span data-ttu-id="0f8b7-469">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-470">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-471">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-471">1.0</span></span>|
|[<span data-ttu-id="0f8b7-472">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-473">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-474">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-475">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f8b7-475">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="0f8b7-476">開催者: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-476">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0f8b7-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0f8b7-479">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-479">Type</span></span>

*   [<span data-ttu-id="0f8b7-480">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0f8b7-480">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="0f8b7-481">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-481">Requirements</span></span>

|<span data-ttu-id="0f8b7-482">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-482">Requirement</span></span>| <span data-ttu-id="0f8b7-483">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-484">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-485">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-485">1.0</span></span>|
|[<span data-ttu-id="0f8b7-486">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-487">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-488">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-489">読み取り</span><span class="sxs-lookup"><span data-stu-id="0f8b7-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-490">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-490">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="0f8b7-491">requiredatて dees: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-491">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0f8b7-492">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-492">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="0f8b7-493">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-493">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0f8b7-494">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-494">Read mode</span></span>

<span data-ttu-id="0f8b7-495">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-495">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="0f8b7-496">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-496">Compose mode</span></span>

<span data-ttu-id="0f8b7-497">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-497">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="0f8b7-498">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-498">Type</span></span>

*   <span data-ttu-id="0f8b7-499">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-499">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-500">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-500">Requirements</span></span>

|<span data-ttu-id="0f8b7-501">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-501">Requirement</span></span>| <span data-ttu-id="0f8b7-502">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-503">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-504">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-504">1.0</span></span>|
|[<span data-ttu-id="0f8b7-505">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-506">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-507">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-508">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f8b7-508">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="0f8b7-509">sender: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-509">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0f8b7-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="0f8b7-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="0f8b7-514">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-514">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="0f8b7-515">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-515">Type</span></span>

*   [<span data-ttu-id="0f8b7-516">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0f8b7-516">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="0f8b7-517">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-517">Requirements</span></span>

|<span data-ttu-id="0f8b7-518">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-518">Requirement</span></span>| <span data-ttu-id="0f8b7-519">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-520">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-521">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-521">1.0</span></span>|
|[<span data-ttu-id="0f8b7-522">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-523">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-524">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-525">読み取り</span><span class="sxs-lookup"><span data-stu-id="0f8b7-525">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-526">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-526">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="0f8b7-527">開始: 日付 |[時間](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-527">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0f8b7-528">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-528">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="0f8b7-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0f8b7-531">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-531">Read mode</span></span>

<span data-ttu-id="0f8b7-532">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-532">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="0f8b7-533">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-533">Compose mode</span></span>

<span data-ttu-id="0f8b7-534">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-534">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="0f8b7-535">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-535">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="0f8b7-536">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-536">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="0f8b7-537">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-537">Type</span></span>

*   <span data-ttu-id="0f8b7-538">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-538">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-539">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-539">Requirements</span></span>

|<span data-ttu-id="0f8b7-540">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-540">Requirement</span></span>| <span data-ttu-id="0f8b7-541">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-541">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-542">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-542">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-543">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-543">1.0</span></span>|
|[<span data-ttu-id="0f8b7-544">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-544">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-545">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-545">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-546">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-546">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-547">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f8b7-547">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-12"></a><span data-ttu-id="0f8b7-548">subject: String |[件名](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-548">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0f8b7-549">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-549">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="0f8b7-550">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-550">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0f8b7-551">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-551">Read mode</span></span>

<span data-ttu-id="0f8b7-p130">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p130">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="0f8b7-554">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-554">Compose mode</span></span>

<span data-ttu-id="0f8b7-555">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-555">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="0f8b7-556">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-556">Type</span></span>

*   <span data-ttu-id="0f8b7-557">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-557">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-558">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-558">Requirements</span></span>

|<span data-ttu-id="0f8b7-559">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-559">Requirement</span></span>| <span data-ttu-id="0f8b7-560">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-561">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-562">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-562">1.0</span></span>|
|[<span data-ttu-id="0f8b7-563">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-563">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-564">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-565">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-565">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-566">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f8b7-566">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="0f8b7-567">宛先: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-567">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="0f8b7-568">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-568">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="0f8b7-569">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-569">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0f8b7-570">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-570">Read mode</span></span>

<span data-ttu-id="0f8b7-p132">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p132">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="0f8b7-573">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-573">Compose mode</span></span>

<span data-ttu-id="0f8b7-574">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-574">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0f8b7-575">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-575">Type</span></span>

*   <span data-ttu-id="0f8b7-576">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-576">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-577">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-577">Requirements</span></span>

|<span data-ttu-id="0f8b7-578">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-578">Requirement</span></span>| <span data-ttu-id="0f8b7-579">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-579">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-580">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-580">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-581">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-581">1.0</span></span>|
|[<span data-ttu-id="0f8b7-582">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-583">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-583">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-584">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-584">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-585">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f8b7-585">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="0f8b7-586">メソッド</span><span class="sxs-lookup"><span data-stu-id="0f8b7-586">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="0f8b7-587">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0f8b7-587">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0f8b7-588">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-588">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="0f8b7-589">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-589">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="0f8b7-590">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-590">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f8b7-591">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0f8b7-591">Parameters</span></span>

|<span data-ttu-id="0f8b7-592">名前</span><span class="sxs-lookup"><span data-stu-id="0f8b7-592">Name</span></span>| <span data-ttu-id="0f8b7-593">種類</span><span class="sxs-lookup"><span data-stu-id="0f8b7-593">Type</span></span>| <span data-ttu-id="0f8b7-594">属性</span><span class="sxs-lookup"><span data-stu-id="0f8b7-594">Attributes</span></span>| <span data-ttu-id="0f8b7-595">説明</span><span class="sxs-lookup"><span data-stu-id="0f8b7-595">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="0f8b7-596">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-596">String</span></span>||<span data-ttu-id="0f8b7-p133">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p133">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="0f8b7-599">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-599">String</span></span>||<span data-ttu-id="0f8b7-p134">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p134">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="0f8b7-602">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0f8b7-602">Object</span></span>| <span data-ttu-id="0f8b7-603">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-603">&lt;optional&gt;</span></span>|<span data-ttu-id="0f8b7-604">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-604">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0f8b7-605">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0f8b7-605">Object</span></span>| <span data-ttu-id="0f8b7-606">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-606">&lt;optional&gt;</span></span>|<span data-ttu-id="0f8b7-607">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-607">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0f8b7-608">function</span><span class="sxs-lookup"><span data-stu-id="0f8b7-608">function</span></span>| <span data-ttu-id="0f8b7-609">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-609">&lt;optional&gt;</span></span>|<span data-ttu-id="0f8b7-610">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-610">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0f8b7-611">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-611">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0f8b7-612">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-612">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0f8b7-613">エラー</span><span class="sxs-lookup"><span data-stu-id="0f8b7-613">Errors</span></span>

| <span data-ttu-id="0f8b7-614">エラー コード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-614">Error code</span></span> | <span data-ttu-id="0f8b7-615">説明</span><span class="sxs-lookup"><span data-stu-id="0f8b7-615">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="0f8b7-616">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-616">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="0f8b7-617">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-617">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="0f8b7-618">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-618">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0f8b7-619">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-619">Requirements</span></span>

|<span data-ttu-id="0f8b7-620">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-620">Requirement</span></span>| <span data-ttu-id="0f8b7-621">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-621">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-622">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-622">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-623">1.1</span><span class="sxs-lookup"><span data-stu-id="0f8b7-623">1.1</span></span>|
|[<span data-ttu-id="0f8b7-624">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-624">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-625">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-625">ReadWriteItem</span></span>|
|[<span data-ttu-id="0f8b7-626">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-626">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-627">作成</span><span class="sxs-lookup"><span data-stu-id="0f8b7-627">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-628">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-628">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="0f8b7-629">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0f8b7-629">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0f8b7-630">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-630">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="0f8b7-p135">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p135">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="0f8b7-634">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-634">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="0f8b7-635">Office アドインが Outlook on the web で実行されている場合、 `addItemAttachmentAsync`メソッドは、編集しているアイテム以外のアイテムにアイテムを添付できます。ただし、これはサポートされておらず、推奨されていません。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-635">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f8b7-636">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0f8b7-636">Parameters</span></span>

|<span data-ttu-id="0f8b7-637">名前</span><span class="sxs-lookup"><span data-stu-id="0f8b7-637">Name</span></span>| <span data-ttu-id="0f8b7-638">種類</span><span class="sxs-lookup"><span data-stu-id="0f8b7-638">Type</span></span>| <span data-ttu-id="0f8b7-639">属性</span><span class="sxs-lookup"><span data-stu-id="0f8b7-639">Attributes</span></span>| <span data-ttu-id="0f8b7-640">説明</span><span class="sxs-lookup"><span data-stu-id="0f8b7-640">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="0f8b7-641">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-641">String</span></span>||<span data-ttu-id="0f8b7-p136">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p136">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="0f8b7-644">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-644">String</span></span>||<span data-ttu-id="0f8b7-645">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-645">The subject of the item to be attached.</span></span> <span data-ttu-id="0f8b7-646">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-646">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="0f8b7-647">Object</span><span class="sxs-lookup"><span data-stu-id="0f8b7-647">Object</span></span>| <span data-ttu-id="0f8b7-648">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-648">&lt;optional&gt;</span></span>|<span data-ttu-id="0f8b7-649">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-649">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0f8b7-650">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0f8b7-650">Object</span></span>| <span data-ttu-id="0f8b7-651">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-651">&lt;optional&gt;</span></span>|<span data-ttu-id="0f8b7-652">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-652">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0f8b7-653">function</span><span class="sxs-lookup"><span data-stu-id="0f8b7-653">function</span></span>| <span data-ttu-id="0f8b7-654">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-654">&lt;optional&gt;</span></span>|<span data-ttu-id="0f8b7-655">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-655">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0f8b7-656">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-656">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0f8b7-657">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-657">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0f8b7-658">エラー</span><span class="sxs-lookup"><span data-stu-id="0f8b7-658">Errors</span></span>

| <span data-ttu-id="0f8b7-659">エラー コード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-659">Error code</span></span> | <span data-ttu-id="0f8b7-660">説明</span><span class="sxs-lookup"><span data-stu-id="0f8b7-660">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="0f8b7-661">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-661">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0f8b7-662">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-662">Requirements</span></span>

|<span data-ttu-id="0f8b7-663">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-663">Requirement</span></span>| <span data-ttu-id="0f8b7-664">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-664">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-665">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-665">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-666">1.1</span><span class="sxs-lookup"><span data-stu-id="0f8b7-666">1.1</span></span>|
|[<span data-ttu-id="0f8b7-667">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-667">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-668">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-668">ReadWriteItem</span></span>|
|[<span data-ttu-id="0f8b7-669">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-669">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-670">作成</span><span class="sxs-lookup"><span data-stu-id="0f8b7-670">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-671">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-671">Example</span></span>

<span data-ttu-id="0f8b7-672">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-672">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="0f8b7-673">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="0f8b7-673">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="0f8b7-674">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-674">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0f8b7-675">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-675">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0f8b7-676">Web 上の Outlook では、返信フォームは、3列表示のポップアップフォームとして表示され、2列または1列表示のポップアップフォームとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-676">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="0f8b7-677">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-677">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="0f8b7-678">`formData.attachments`パラメーターで添付ファイルが指定されている場合、web 上の Outlook およびデスクトップクライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-678">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="0f8b7-679">添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-679">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="0f8b7-680">表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-680">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f8b7-681">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0f8b7-681">Parameters</span></span>

|<span data-ttu-id="0f8b7-682">名前</span><span class="sxs-lookup"><span data-stu-id="0f8b7-682">Name</span></span>| <span data-ttu-id="0f8b7-683">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-683">Type</span></span>| <span data-ttu-id="0f8b7-684">説明</span><span class="sxs-lookup"><span data-stu-id="0f8b7-684">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="0f8b7-685">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="0f8b7-685">String &#124; Object</span></span>| |<span data-ttu-id="0f8b7-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="0f8b7-688">**または**</span><span class="sxs-lookup"><span data-stu-id="0f8b7-688">**OR**</span></span><br/><span data-ttu-id="0f8b7-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="0f8b7-691">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-691">String</span></span> | <span data-ttu-id="0f8b7-692">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-692">&lt;optional&gt;</span></span> | <span data-ttu-id="0f8b7-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="0f8b7-695">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-695">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="0f8b7-696">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-696">&lt;optional&gt;</span></span> | <span data-ttu-id="0f8b7-697">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-697">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="0f8b7-698">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-698">String</span></span> | | <span data-ttu-id="0f8b7-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="0f8b7-701">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-701">String</span></span> | | <span data-ttu-id="0f8b7-702">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-702">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="0f8b7-703">文字列</span><span class="sxs-lookup"><span data-stu-id="0f8b7-703">String</span></span> | | <span data-ttu-id="0f8b7-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="0f8b7-706">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-706">String</span></span> | | <span data-ttu-id="0f8b7-p144">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="0f8b7-710">function</span><span class="sxs-lookup"><span data-stu-id="0f8b7-710">function</span></span> | <span data-ttu-id="0f8b7-711">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-711">&lt;optional&gt;</span></span> | <span data-ttu-id="0f8b7-712">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-712">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0f8b7-713">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-713">Requirements</span></span>

|<span data-ttu-id="0f8b7-714">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-714">Requirement</span></span>| <span data-ttu-id="0f8b7-715">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-715">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-716">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-716">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-717">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-717">1.0</span></span>|
|[<span data-ttu-id="0f8b7-718">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-718">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-719">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-719">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-720">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-720">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-721">読み取り</span><span class="sxs-lookup"><span data-stu-id="0f8b7-721">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="0f8b7-722">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-722">Examples</span></span>

<span data-ttu-id="0f8b7-723">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-723">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="0f8b7-724">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-724">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="0f8b7-725">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-725">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="0f8b7-726">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-726">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="0f8b7-727">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-727">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="0f8b7-728">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-728">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="0f8b7-729">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="0f8b7-729">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="0f8b7-730">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-730">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0f8b7-731">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-731">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0f8b7-732">Web 上の Outlook では、返信フォームは、3列表示のポップアップフォームとして表示され、2列または1列表示のポップアップフォームとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-732">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="0f8b7-733">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-733">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="0f8b7-734">`formData.attachments`パラメーターで添付ファイルが指定されている場合、web 上の Outlook およびデスクトップクライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-734">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="0f8b7-735">添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-735">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="0f8b7-736">表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-736">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f8b7-737">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0f8b7-737">Parameters</span></span>

|<span data-ttu-id="0f8b7-738">名前</span><span class="sxs-lookup"><span data-stu-id="0f8b7-738">Name</span></span>| <span data-ttu-id="0f8b7-739">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-739">Type</span></span>| <span data-ttu-id="0f8b7-740">説明</span><span class="sxs-lookup"><span data-stu-id="0f8b7-740">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="0f8b7-741">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="0f8b7-741">String &#124; Object</span></span>| | <span data-ttu-id="0f8b7-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="0f8b7-744">**または**</span><span class="sxs-lookup"><span data-stu-id="0f8b7-744">**OR**</span></span><br/><span data-ttu-id="0f8b7-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="0f8b7-747">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-747">String</span></span> | <span data-ttu-id="0f8b7-748">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-748">&lt;optional&gt;</span></span> | <span data-ttu-id="0f8b7-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="0f8b7-751">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-751">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="0f8b7-752">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-752">&lt;optional&gt;</span></span> | <span data-ttu-id="0f8b7-753">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-753">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="0f8b7-754">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-754">String</span></span> | | <span data-ttu-id="0f8b7-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="0f8b7-757">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-757">String</span></span> | | <span data-ttu-id="0f8b7-758">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-758">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="0f8b7-759">文字列</span><span class="sxs-lookup"><span data-stu-id="0f8b7-759">String</span></span> | | <span data-ttu-id="0f8b7-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="0f8b7-762">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-762">String</span></span> | | <span data-ttu-id="0f8b7-p151">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="0f8b7-766">function</span><span class="sxs-lookup"><span data-stu-id="0f8b7-766">function</span></span> | <span data-ttu-id="0f8b7-767">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-767">&lt;optional&gt;</span></span> | <span data-ttu-id="0f8b7-768">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-768">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0f8b7-769">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-769">Requirements</span></span>

|<span data-ttu-id="0f8b7-770">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-770">Requirement</span></span>| <span data-ttu-id="0f8b7-771">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-771">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-772">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-772">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-773">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-773">1.0</span></span>|
|[<span data-ttu-id="0f8b7-774">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-774">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-775">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-775">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-776">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-776">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-777">読み取り</span><span class="sxs-lookup"><span data-stu-id="0f8b7-777">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="0f8b7-778">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-778">Examples</span></span>

<span data-ttu-id="0f8b7-779">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-779">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="0f8b7-780">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-780">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="0f8b7-781">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-781">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="0f8b7-782">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-782">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="0f8b7-783">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-783">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="0f8b7-784">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-784">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-12"></a><span data-ttu-id="0f8b7-785">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="0f8b7-785">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="0f8b7-786">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-786">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="0f8b7-787">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-787">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-788">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-788">Requirements</span></span>

|<span data-ttu-id="0f8b7-789">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-789">Requirement</span></span>| <span data-ttu-id="0f8b7-790">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-791">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-792">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-792">1.0</span></span>|
|[<span data-ttu-id="0f8b7-793">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-793">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-794">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-794">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-795">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-795">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-796">読み取り</span><span class="sxs-lookup"><span data-stu-id="0f8b7-796">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0f8b7-797">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0f8b7-797">Returns:</span></span>

<span data-ttu-id="0f8b7-798">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-798">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span></span>

##### <a name="example"></a><span data-ttu-id="0f8b7-799">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-799">Example</span></span>

<span data-ttu-id="0f8b7-800">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-800">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="0f8b7-801">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="0f8b7-801">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="0f8b7-802">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-802">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="0f8b7-803">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-803">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f8b7-804">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0f8b7-804">Parameters</span></span>

|<span data-ttu-id="0f8b7-805">名前</span><span class="sxs-lookup"><span data-stu-id="0f8b7-805">Name</span></span>| <span data-ttu-id="0f8b7-806">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-806">Type</span></span>| <span data-ttu-id="0f8b7-807">説明</span><span class="sxs-lookup"><span data-stu-id="0f8b7-807">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="0f8b7-808">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="0f8b7-808">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.2)|<span data-ttu-id="0f8b7-809">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-809">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0f8b7-810">Requirements</span><span class="sxs-lookup"><span data-stu-id="0f8b7-810">Requirements</span></span>

|<span data-ttu-id="0f8b7-811">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-811">Requirement</span></span>| <span data-ttu-id="0f8b7-812">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-812">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-813">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-813">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-814">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-814">1.0</span></span>|
|[<span data-ttu-id="0f8b7-815">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-815">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-816">制限あり</span><span class="sxs-lookup"><span data-stu-id="0f8b7-816">Restricted</span></span>|
|[<span data-ttu-id="0f8b7-817">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-817">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-818">読み取り</span><span class="sxs-lookup"><span data-stu-id="0f8b7-818">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0f8b7-819">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0f8b7-819">Returns:</span></span>

<span data-ttu-id="0f8b7-820">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-820">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="0f8b7-821">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-821">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="0f8b7-822">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-822">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="0f8b7-823">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-823">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="0f8b7-824">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-824">Value of `entityType`</span></span> | <span data-ttu-id="0f8b7-825">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-825">Type of objects in returned array</span></span> | <span data-ttu-id="0f8b7-826">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-826">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="0f8b7-827">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-827">String</span></span> | <span data-ttu-id="0f8b7-828">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="0f8b7-828">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="0f8b7-829">連絡先</span><span class="sxs-lookup"><span data-stu-id="0f8b7-829">Contact</span></span> | <span data-ttu-id="0f8b7-830">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0f8b7-830">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="0f8b7-831">文字列</span><span class="sxs-lookup"><span data-stu-id="0f8b7-831">String</span></span> | <span data-ttu-id="0f8b7-832">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0f8b7-832">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="0f8b7-833">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="0f8b7-833">MeetingSuggestion</span></span> | <span data-ttu-id="0f8b7-834">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0f8b7-834">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="0f8b7-835">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="0f8b7-835">PhoneNumber</span></span> | <span data-ttu-id="0f8b7-836">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="0f8b7-836">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="0f8b7-837">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="0f8b7-837">TaskSuggestion</span></span> | <span data-ttu-id="0f8b7-838">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0f8b7-838">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="0f8b7-839">文字列</span><span class="sxs-lookup"><span data-stu-id="0f8b7-839">String</span></span> | <span data-ttu-id="0f8b7-840">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="0f8b7-840">**Restricted**</span></span> |

<span data-ttu-id="0f8b7-841">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="0f8b7-841">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

##### <a name="example"></a><span data-ttu-id="0f8b7-842">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-842">Example</span></span>

<span data-ttu-id="0f8b7-843">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-843">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="0f8b7-844">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="0f8b7-844">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="0f8b7-845">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-845">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0f8b7-846">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-846">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0f8b7-847">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-847">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f8b7-848">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0f8b7-848">Parameters</span></span>

|<span data-ttu-id="0f8b7-849">名前</span><span class="sxs-lookup"><span data-stu-id="0f8b7-849">Name</span></span>| <span data-ttu-id="0f8b7-850">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-850">Type</span></span>| <span data-ttu-id="0f8b7-851">説明</span><span class="sxs-lookup"><span data-stu-id="0f8b7-851">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="0f8b7-852">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-852">String</span></span>|<span data-ttu-id="0f8b7-853">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-853">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0f8b7-854">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-854">Requirements</span></span>

|<span data-ttu-id="0f8b7-855">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-855">Requirement</span></span>| <span data-ttu-id="0f8b7-856">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-856">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-857">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-857">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-858">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-858">1.0</span></span>|
|[<span data-ttu-id="0f8b7-859">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-859">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-860">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-860">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-861">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-861">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-862">読み取り</span><span class="sxs-lookup"><span data-stu-id="0f8b7-862">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0f8b7-863">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0f8b7-863">Returns:</span></span>

<span data-ttu-id="0f8b7-p153">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="0f8b7-866">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="0f8b7-866">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="0f8b7-867">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="0f8b7-867">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="0f8b7-868">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-868">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0f8b7-869">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-869">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0f8b7-p154">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="0f8b7-873">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-873">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="0f8b7-874">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-874">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="0f8b7-p155">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f8b7-877">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-877">Requirements</span></span>

|<span data-ttu-id="0f8b7-878">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-878">Requirement</span></span>| <span data-ttu-id="0f8b7-879">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-879">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-880">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-880">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-881">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-881">1.0</span></span>|
|[<span data-ttu-id="0f8b7-882">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-882">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-883">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-883">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-884">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-884">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-885">読み取り</span><span class="sxs-lookup"><span data-stu-id="0f8b7-885">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0f8b7-886">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0f8b7-886">Returns:</span></span>

<span data-ttu-id="0f8b7-p156">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="0f8b7-889">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="0f8b7-889">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0f8b7-890">Object</span><span class="sxs-lookup"><span data-stu-id="0f8b7-890">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="0f8b7-891">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-891">Example</span></span>

<span data-ttu-id="0f8b7-892">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="0f8b7-892">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="0f8b7-893">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="0f8b7-893">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="0f8b7-894">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-894">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0f8b7-895">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-895">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0f8b7-896">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-896">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="0f8b7-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f8b7-899">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0f8b7-899">Parameters</span></span>

|<span data-ttu-id="0f8b7-900">名前</span><span class="sxs-lookup"><span data-stu-id="0f8b7-900">Name</span></span>| <span data-ttu-id="0f8b7-901">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-901">Type</span></span>| <span data-ttu-id="0f8b7-902">説明</span><span class="sxs-lookup"><span data-stu-id="0f8b7-902">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="0f8b7-903">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-903">String</span></span>|<span data-ttu-id="0f8b7-904">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-904">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0f8b7-905">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-905">Requirements</span></span>

|<span data-ttu-id="0f8b7-906">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-906">Requirement</span></span>| <span data-ttu-id="0f8b7-907">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-907">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-908">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-908">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-909">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-909">1.0</span></span>|
|[<span data-ttu-id="0f8b7-910">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-910">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-911">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-911">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-912">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-912">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-913">読み取り</span><span class="sxs-lookup"><span data-stu-id="0f8b7-913">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0f8b7-914">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0f8b7-914">Returns:</span></span>

<span data-ttu-id="0f8b7-915">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-915">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="0f8b7-916">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="0f8b7-916">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0f8b7-917">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="0f8b7-917">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="0f8b7-918">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-918">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="0f8b7-919">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="0f8b7-919">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="0f8b7-920">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-920">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="0f8b7-p158">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f8b7-923">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0f8b7-923">Parameters</span></span>

|<span data-ttu-id="0f8b7-924">名前</span><span class="sxs-lookup"><span data-stu-id="0f8b7-924">Name</span></span>| <span data-ttu-id="0f8b7-925">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-925">Type</span></span>| <span data-ttu-id="0f8b7-926">属性</span><span class="sxs-lookup"><span data-stu-id="0f8b7-926">Attributes</span></span>| <span data-ttu-id="0f8b7-927">説明</span><span class="sxs-lookup"><span data-stu-id="0f8b7-927">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="0f8b7-928">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="0f8b7-928">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="0f8b7-p159">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="0f8b7-932">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0f8b7-932">Object</span></span>| <span data-ttu-id="0f8b7-933">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-933">&lt;optional&gt;</span></span>|<span data-ttu-id="0f8b7-934">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-934">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0f8b7-935">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0f8b7-935">Object</span></span>| <span data-ttu-id="0f8b7-936">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-936">&lt;optional&gt;</span></span>|<span data-ttu-id="0f8b7-937">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-937">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0f8b7-938">function</span><span class="sxs-lookup"><span data-stu-id="0f8b7-938">function</span></span>||<span data-ttu-id="0f8b7-939">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-939">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0f8b7-940">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-940">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="0f8b7-941">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-941">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0f8b7-942">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-942">Requirements</span></span>

|<span data-ttu-id="0f8b7-943">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-943">Requirement</span></span>| <span data-ttu-id="0f8b7-944">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-944">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-945">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-945">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-946">1.2</span><span class="sxs-lookup"><span data-stu-id="0f8b7-946">1.2</span></span>|
|[<span data-ttu-id="0f8b7-947">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-947">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-948">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-948">ReadWriteItem</span></span>|
|[<span data-ttu-id="0f8b7-949">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-949">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-950">作成</span><span class="sxs-lookup"><span data-stu-id="0f8b7-950">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="0f8b7-951">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0f8b7-951">Returns:</span></span>

<span data-ttu-id="0f8b7-952">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-952">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="0f8b7-953">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="0f8b7-953">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0f8b7-954">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-954">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="0f8b7-955">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-955">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="0f8b7-956">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0f8b7-956">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="0f8b7-957">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-957">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="0f8b7-p161">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f8b7-961">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0f8b7-961">Parameters</span></span>

|<span data-ttu-id="0f8b7-962">名前</span><span class="sxs-lookup"><span data-stu-id="0f8b7-962">Name</span></span>| <span data-ttu-id="0f8b7-963">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-963">Type</span></span>| <span data-ttu-id="0f8b7-964">属性</span><span class="sxs-lookup"><span data-stu-id="0f8b7-964">Attributes</span></span>| <span data-ttu-id="0f8b7-965">説明</span><span class="sxs-lookup"><span data-stu-id="0f8b7-965">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="0f8b7-966">function</span><span class="sxs-lookup"><span data-stu-id="0f8b7-966">function</span></span>||<span data-ttu-id="0f8b7-967">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0f8b7-968">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-968">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="0f8b7-969">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-969">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="0f8b7-970">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0f8b7-970">Object</span></span>| <span data-ttu-id="0f8b7-971">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-971">&lt;optional&gt;</span></span>|<span data-ttu-id="0f8b7-972">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-972">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="0f8b7-973">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-973">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0f8b7-974">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-974">Requirements</span></span>

|<span data-ttu-id="0f8b7-975">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-975">Requirement</span></span>| <span data-ttu-id="0f8b7-976">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-976">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-977">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-977">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-978">1.0</span><span class="sxs-lookup"><span data-stu-id="0f8b7-978">1.0</span></span>|
|[<span data-ttu-id="0f8b7-979">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-979">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-980">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-980">ReadItem</span></span>|
|[<span data-ttu-id="0f8b7-981">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-981">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-982">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f8b7-982">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-983">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-983">Example</span></span>

<span data-ttu-id="0f8b7-p164">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="0f8b7-987">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0f8b7-987">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="0f8b7-988">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-988">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="0f8b7-989">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-989">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="0f8b7-990">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-990">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="0f8b7-991">Outlook on the web およびモバイルデバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-991">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="0f8b7-992">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-992">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f8b7-993">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0f8b7-993">Parameters</span></span>

|<span data-ttu-id="0f8b7-994">名前</span><span class="sxs-lookup"><span data-stu-id="0f8b7-994">Name</span></span>| <span data-ttu-id="0f8b7-995">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-995">Type</span></span>| <span data-ttu-id="0f8b7-996">属性</span><span class="sxs-lookup"><span data-stu-id="0f8b7-996">Attributes</span></span>| <span data-ttu-id="0f8b7-997">説明</span><span class="sxs-lookup"><span data-stu-id="0f8b7-997">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="0f8b7-998">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-998">String</span></span>||<span data-ttu-id="0f8b7-999">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-999">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="0f8b7-1000">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1000">Object</span></span>| <span data-ttu-id="0f8b7-1001">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1001">&lt;optional&gt;</span></span>|<span data-ttu-id="0f8b7-1002">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1002">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0f8b7-1003">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1003">Object</span></span>| <span data-ttu-id="0f8b7-1004">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="0f8b7-1005">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1005">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0f8b7-1006">関数</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1006">function</span></span>| <span data-ttu-id="0f8b7-1007">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="0f8b7-1008">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1008">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0f8b7-1009">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1009">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0f8b7-1010">エラー</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1010">Errors</span></span>

| <span data-ttu-id="0f8b7-1011">エラー コード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1011">Error code</span></span> | <span data-ttu-id="0f8b7-1012">説明</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1012">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="0f8b7-1013">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1013">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0f8b7-1014">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1014">Requirements</span></span>

|<span data-ttu-id="0f8b7-1015">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1015">Requirement</span></span>| <span data-ttu-id="0f8b7-1016">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1016">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-1017">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1017">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-1018">1.1</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1018">1.1</span></span>|
|[<span data-ttu-id="0f8b7-1019">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1019">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-1020">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1020">ReadWriteItem</span></span>|
|[<span data-ttu-id="0f8b7-1021">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1021">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-1022">作成</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1022">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-1023">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1023">Example</span></span>

<span data-ttu-id="0f8b7-1024">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1024">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="0f8b7-1025">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1025">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="0f8b7-1026">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1026">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="0f8b7-p166">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f8b7-1030">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1030">Parameters</span></span>

|<span data-ttu-id="0f8b7-1031">名前</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1031">Name</span></span>| <span data-ttu-id="0f8b7-1032">型</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1032">Type</span></span>| <span data-ttu-id="0f8b7-1033">属性</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1033">Attributes</span></span>| <span data-ttu-id="0f8b7-1034">説明</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1034">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="0f8b7-1035">String</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1035">String</span></span>||<span data-ttu-id="0f8b7-p167">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="0f8b7-1039">Object</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1039">Object</span></span>| <span data-ttu-id="0f8b7-1040">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="0f8b7-1041">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1041">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0f8b7-1042">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1042">Object</span></span>| <span data-ttu-id="0f8b7-1043">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1043">&lt;optional&gt;</span></span>|<span data-ttu-id="0f8b7-1044">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1044">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="0f8b7-1045">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1045">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="0f8b7-1046">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1046">&lt;optional&gt;</span></span>|<span data-ttu-id="0f8b7-1047">の`text`場合、現在のスタイルが Outlook on the web およびデスクトップクライアントで適用されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1047">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="0f8b7-1048">フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1048">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="0f8b7-1049">フィールド`html`が HTML をサポートする場合 (件名は含まれません)、現在のスタイルが outlook on the web で適用され、既定のスタイルが outlook デスクトップクライアントで適用されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1049">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="0f8b7-1050">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1050">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="0f8b7-1051">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1051">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="0f8b7-1052">function</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1052">function</span></span>||<span data-ttu-id="0f8b7-1053">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1053">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0f8b7-1054">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1054">Requirements</span></span>

|<span data-ttu-id="0f8b7-1055">要件</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1055">Requirement</span></span>| <span data-ttu-id="0f8b7-1056">値</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1056">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f8b7-1057">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1057">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f8b7-1058">1.2</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1058">1.2</span></span>|
|[<span data-ttu-id="0f8b7-1059">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1059">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f8b7-1060">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1060">ReadWriteItem</span></span>|
|[<span data-ttu-id="0f8b7-1061">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1061">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f8b7-1062">作成</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1062">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0f8b7-1063">例</span><span class="sxs-lookup"><span data-stu-id="0f8b7-1063">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
